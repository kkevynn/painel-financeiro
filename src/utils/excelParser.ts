import * as XLSX from 'xlsx';
import { FileData, ConsolidatedSupplier, DashboardMetrics, StageKey } from '../types';
import { toastAlertService } from './toastAlertService';
import { mlInferenceService } from './machineLearningInferenceService';

export function detectStageFromFileName(fileName: string): StageKey {
  if (!fileName) return 'outros';

  const normalized = fileName
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  // 0. DETALHAMENTO EXPLÍCITO DE ORDEM BANCÁRIA (Listar Ordem Bancária / OBs - sigef - Principal)
  if (
    normalized.includes('ordem bancaria') ||
    normalized.includes('ordens bancarias') ||
    normalized.includes('listar ordem') ||
    normalized.includes('listar ob') ||
    normalized.includes('relacao de ob') ||
    normalized.includes('relatorio ob') ||
    normalized.includes('relatorio de ob') ||
    normalized.includes('planilha ob') ||
    normalized.includes('planilha de ob') ||
    ((normalized.includes('obs') || normalized.includes('ob')) && normalized.includes('principal'))
  ) {
    return 'ordem_bancaria';
  }

  // 1. NATUREZA DA DESPESA (explícita pelo nome "natureza")
  if (normalized.includes('natureza')) {
    return 'natureza_despesa';
  }

  // 1.1 PUBLICAÇÃO / DOE / PARA PUBLICAR (prioridade sobre termos genéricos como "empenho" ou "obs")
  if (
    normalized.includes('publicacao') ||
    normalized.includes('publicac') ||
    normalized.includes('publicar') ||
    normalized.includes('publicado') ||
    normalized.includes('diario oficial') ||
    normalized.includes('para publicacao') ||
    /\bdoe\b/.test(normalized)
  ) {
    return 'publicacao';
  }

  // 2. NÃO OBEDECE (Priority check before "obedece" and before generic "empenho"/"obs")
  if (
    normalized.includes('nao obedece') ||
    normalized.includes('naoobedece') ||
    normalized.includes('nao obedec') ||
    normalized.includes('desobedece') ||
    normalized.includes('desobediencia') ||
    normalized.includes('nao cronologica') ||
    normalized.includes('naocronologica') ||
    normalized.includes('nao cronologico') ||
    normalized.includes('naocronologico') ||
    normalized.includes('nao cron') ||
    normalized.includes('fora da ordem') ||
    normalized.includes('fora de ordem') ||
    normalized.includes('sem ordem') ||
    normalized.includes('sem cronologia') ||
    normalized.includes('sem cron') ||
    normalized.includes('excecao') ||
    normalized.includes('excecoes') ||
    normalized.includes('excepcional') ||
    normalized.includes('art 141') ||
    normalized.includes('artigo 141') ||
    normalized.includes('dispensa ordem') ||
    normalized.includes('dispensa de ordem')
  ) {
    return 'nao_obedece';
  }

  // 3. OBEDECE
  if (
    normalized.includes('obedece') ||
    normalized.includes('cronologica') ||
    normalized.includes('cronologico') ||
    normalized.includes('ordem cronologica')
  ) {
    return 'obedece';
  }

  // 3.1 NATUREZA / EMPENHO (quando não é liquidação/publicação)
  if (normalized.includes('empenho')) {
    return 'natureza_despesa';
  }

  // 3.2 ORDEM BANCÁRIA por palavra isolada "ob" ou "obs" (fallback após checar etapas funcionais)
  if (/\bobs?\b/.test(normalized)) {
    return 'ordem_bancaria';
  }

  // 4. BASE DE PPS (including SIGEF reports such as Relatorio_*.xls and payment proposals)
  if (
    normalized.startsWith('relatorio') ||
    normalized.includes('relatorio') ||
    normalized.includes('relatório') ||
    normalized.includes('base') ||
    normalized.includes('sigef') ||
    normalized.includes('pp') ||
    normalized.includes('proposta') ||
    normalized.includes('programacao') ||
    normalized.includes('pagamento') ||
    normalized.includes('relacao de pp')
  ) {
    return 'base_pps';
  }

  return 'outros';
}

// Format currency as BRL (R$ 1.234.567,89)
export function formatBRL(val: number): string {
  return new Intl.NumberFormat('pt-BR', {
    style: 'currency',
    currency: 'BRL',
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(val || 0);
}

// Clean and normalize Brazilian currency values from string or numbers
export function parseCurrencyValue(raw: any): number {
  if (raw === null || raw === undefined) return 0;
  if (typeof raw === 'number') return isNaN(raw) ? 0 : raw;

  let str = String(raw).trim();
  if (!str) return 0;

  // Remove currency symbol, spaces, non-breaking spaces
  str = str.replace(/R\$\s?/gi, '').replace(/\s/g, '').trim();

  // If in format (1.234,56) meaning negative
  const isNegative = (str.startsWith('(') && str.endsWith(')')) || str.startsWith('-');
  str = str.replace(/[()\-]/g, '');

  if (str.includes(',') && str.includes('.')) {
    const lastComma = str.lastIndexOf(',');
    const lastDot = str.lastIndexOf('.');
    if (lastComma > lastDot) {
      // Brazilian: 1.234.567,89 -> remove dots, replace comma with dot
      str = str.replace(/\./g, '').replace(',', '.');
    } else {
      // US: 1,234,567.89 -> remove commas
      str = str.replace(/,/g, '');
    }
  } else if (str.includes(',')) {
    // 1234,56 -> 1234.56
    str = str.replace(',', '.');
  } else if (str.includes('.')) {
    const parts = str.split('.');
    if (parts.length > 2 || (parts.length === 2 && parts[1].length === 3)) {
      str = str.replace(/\./g, '');
    }
  }

  const parsed = parseFloat(str);
  if (isNaN(parsed)) return 0;
  return isNegative ? -parsed : parsed;
}

/**
 * Checks if a string indicates "AO" situation (Associada 2 Ordenadores - PP Pronta)
 */
export function isSituacaoAO(situacaoRaw: any): boolean {
  if (situacaoRaw === null || situacaoRaw === undefined) return false;
  const str = String(situacaoRaw)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toUpperCase();
  if (!str) return false;
  if (str === 'AO' || str === 'A.O.' || str === 'A O') return true;
  if (
    str.startsWith('AO ') ||
    str.startsWith('AO-') ||
    str.startsWith('AO/') ||
    str.startsWith('AO:') ||
    str.startsWith('AO_') ||
    str.endsWith(' AO') ||
    str.includes(' AO ') ||
    str.includes('(AO)') ||
    str.includes('[AO]')
  ) {
    return true;
  }
  if (
    str.includes('ASSOCIADA') &&
    (str.includes('ORDENADOR') || str.includes('ORDENADORES') || str.includes('2') || str.includes('DOIS'))
  ) {
    return true;
  }
  if (
    str.includes('2 ORDENADORES') ||
    str.includes('DOIS ORDENADORES') ||
    str.includes('2 ORD') ||
    str.includes('2ORD')
  ) {
    return true;
  }
  if (
    str.includes('COM OB ASSOCIADA') ||
    str.includes('ORDEM ASSOCIADA') ||
    str.includes('PP ASSOCIADA') ||
    str.includes('OB ASSOCIADA')
  ) {
    return true;
  }
  if (
    str === 'PP PRONTA' ||
    str === 'PP PRONTAS' ||
    str.includes('PRONTA P/ PGTO') ||
    str.includes('PRONTA PARA PAGAMENTO') ||
    str.includes('PRONTA P/ PAGAMENTO')
  ) {
    return true;
  }
  return false;
}

/**
 * Checks if a string indicates that a PP/OB was sent to bank (situation changed to PPCB or PPNC)
 */
export function isSituacaoEnviadaBanco(situacaoRaw: any): { isEnviada: boolean; tipo: 'PPCB' | 'PPNC' | null; label: string } {
  if (situacaoRaw === null || situacaoRaw === undefined) return { isEnviada: false, tipo: null, label: '' };
  const str = String(situacaoRaw)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toUpperCase();
  if (!str) return { isEnviada: false, tipo: null, label: '' };

  if (str === 'PPCB' || str.startsWith('PPCB ') || str.startsWith('PPCB-') || str.includes('PPCB')) {
    return { isEnviada: true, tipo: 'PPCB', label: 'Enviada a Banco (Confirmada - PPCB)' };
  }
  if (str === 'PPNC' || str.startsWith('PPNC ') || str.startsWith('PPNC-') || str.includes('PPNC') || str === 'CBNC' || str.startsWith('CBNC')) {
    return { isEnviada: true, tipo: 'PPNC', label: 'Enviada a Banco (Não Confirmada - PPNC)' };
  }
  if (str.includes('CONFIRMADA BANCO') || str === 'CB' || str === 'CM') {
    return { isEnviada: true, tipo: 'PPCB', label: 'Confirmada Banco' };
  }
  return { isEnviada: false, tipo: null, label: '' };
}

/**
 * Checks if a string indicates a conclusive payment status
 * (e.g. PPCB, CB, CM, LI, LD, II, CONFIRMADA BANCO, PAGO, EFETIVADO)
 */
export function isSituacaoConclusivaPagamento(situacaoRaw: any): boolean {
  if (situacaoRaw === null || situacaoRaw === undefined) return false;
  const str = String(situacaoRaw)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toUpperCase();
  if (!str) return false;

  // Never consider rejection/cancellation as conclusive payment
  if (
    str === 'RJ' ||
    str.startsWith('RJ') ||
    str === 'PPRJ' ||
    str.startsWith('PPRJ') ||
    str === 'CBNC' ||
    str.startsWith('CBNC') ||
    str.includes('REJEITAD') ||
    str.includes('CANCELAD') ||
    str.includes('DEVOLVID') ||
    str.includes('ESTORNAD')
  ) {
    return false;
  }

  if (
    str === 'PPCB' ||
    str.startsWith('PPCB') ||
    str === 'CB' ||
    str.startsWith('CB ') ||
    str.startsWith('CB-') ||
    str === 'CM' ||
    str.startsWith('CM ') ||
    str.startsWith('CM-') ||
    str === 'LI' ||
    str === 'LD' ||
    str === 'II' ||
    str === 'OB' ||
    str === 'PG' ||
    str === 'PAG' ||
    str.startsWith('PG ') ||
    str.startsWith('PAG ') ||
    str.includes('CONFIRMADA BANCO') ||
    str.includes('CONFIRMADO BANCO') ||
    str.includes('CONFIRMADA NO BANCO') ||
    str.includes('CONFIRMADO NO BANCO') ||
    str.includes('CONFIRMADA MANUAL') ||
    str.includes('PAGAMENTO CONFIRMADO') ||
    str.includes('PAGTO CONFIRMADO') ||
    str === 'CONFIRMADA' ||
    str === 'CONFIRMADO' ||
    str.startsWith('CONFIRMADA ') ||
    str.startsWith('CONFIRMADO ') ||
    str === 'PAGO' ||
    str === 'PAGA' ||
    str.startsWith('PAGO ') ||
    str.startsWith('PAGA ') ||
    str.includes('PAGO NO BANCO') ||
    str.includes('LIQUIDADO E PAGO') ||
    str.includes('LIQUIDADA E PAGA') ||
    str.includes('EFETIVADO') ||
    str.includes('EFETIVADA') ||
    str.includes('BAIXADO') ||
    str.includes('BAIXADA') ||
    str.includes('CREDITADO') ||
    str.includes('CREDITADA') ||
    str.includes('COMPENSADO') ||
    str.includes('COMPENSADA') ||
    str.includes('QUITADO') ||
    str.includes('QUITADA') ||
    str.includes('PROCESSADO NO BANCO') ||
    str.includes('OB EMITIDA') ||
    str.includes('ORDEM BANCARIA EMITIDA')
  ) {
    return true;
  }
  return false;
}

/**
 * Checks if a string indicates a rejected or cancelled bank status (Alert Crítico)
 * (e.g. RJ, PPRJ, REJEITADA, CANCELADA, DEVOLVIDA, ESTORNADA)
 */
export function isSituacaoRejeitadaOuCancelada(situacaoRaw: any): boolean {
  if (situacaoRaw === null || situacaoRaw === undefined) return false;
  const str = String(situacaoRaw)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toUpperCase();
  if (!str) return false;

  if (
    str === 'RJ' ||
    str.startsWith('RJ ') ||
    str.startsWith('RJ-') ||
    str === 'PPRJ' ||
    str.startsWith('PPRJ') ||
    str.includes('REJEITAD') ||
    str.includes('CANCELAD') ||
    str.includes('ESTORNAD') ||
    str.includes('DEVOLVID') ||
    str.includes('INCONSIST')
  ) {
    return true;
  }
  return false;
}

/**
 * Checks if a string or name indicates a bank, treasury or payment intermediary
 * (e.g., BANCO DO BRASIL, CAIXA, TESOURO) where the true contractor is in Favorecido NE
 */
export function isBankOrIntermediary(name: string): boolean {
  if (!name) return false;
  const upper = name.toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  return (
    upper.includes('BANCO DO BRASIL') ||
    upper.includes('BANCO') ||
    upper.includes('CAIXA ECONOMICA') ||
    upper.includes('BRADESCO') ||
    upper.includes('ITAU') ||
    upper.includes('SANTANDER') ||
    upper.includes('TESOURO') ||
    upper.includes('FOLHA DE PAGAMENTO') ||
    upper.includes('ORDEM BANCARIA')
  );
}

/**
 * Checks if a row is a summary, subtotal, grand total or report footer line
 * (which often contain multi-billion budget or aggregated totals that distort metrics)
 */
export function isTotalizadorRow(
  row: Record<string, any>,
  rawFav?: any,
  rawNumero?: any
): boolean {
  const favStr = String(rawFav || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toLowerCase();

  if (
    favStr === 'total' ||
    favStr === 'subtotal' ||
    favStr.startsWith('total ') ||
    favStr.startsWith('total:') ||
    favStr.startsWith('total -') ||
    favStr.startsWith('subtotal') ||
    favStr.startsWith('resumo') ||
    favStr.includes('total geral') ||
    favStr.includes('total do orgao') ||
    favStr.includes('total da unidade') ||
    favStr.includes('total do credor') ||
    favStr.includes('total da despesa') ||
    favStr.includes('total do elemento') ||
    favStr.includes('total da celula') ||
    favStr.includes('total do empenho') ||
    favStr.includes('saldo orcamentario') ||
    favStr.includes('dotacao orcamentaria') ||
    favStr.startsWith('celula') ||
    favStr.startsWith('empenho celula')
  ) {
    return true;
  }

  // Check if any cell in the row explicitly identifies a total or budget summary
  const cells = Object.values(row);
  for (const cell of cells) {
    if (cell === null || cell === undefined) continue;
    const s = String(cell)
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .trim()
      .toLowerCase();

    if (
      s.startsWith('c - cancelada; ap - a pagar') ||
      s.startsWith('c - cancelada') ||
      s.startsWith('relatorio emitido em') ||
      s.startsWith('relatório emitido em') ||
      s === 'total geral' ||
      s.startsWith('total geral') ||
      s.startsWith('total do orgao') ||
      s.startsWith('total da unidade') ||
      s.startsWith('total da despesa') ||
      s.startsWith('total da celula') ||
      s.startsWith('total do empenho') ||
      s.startsWith('resumo geral') ||
      s.startsWith('saldo orcamentario') ||
      s.startsWith('dotacao atualizada') ||
      s.startsWith('dotacao orcamentaria') ||
      s.startsWith('quantidade de registros')
    ) {
      return true;
    }
  }

  return false;
}

/**
 * Checks if a record meets the strict criteria for PPs Prontas:
 * 1. Must be in AO situation (Associada 2 Ordenadores / PP Pronta)
 * 2. Must have all 5 columns complete and valid:
 *    - Número (PP/documento)
 *    - OB (Ordem Bancária)
 *    - Favorecido (Credor)
 *    - Nota Empenho (NE)
 *    - Valor (> 0)
 */
export function isPPProntaAOValida(params: {
  situacaoRaw: any;
  numeroRaw: any;
  obRaw: any;
  favorecidoRaw: any;
  notaEmpenhoRaw: any;
  valor: number;
}): boolean {
  const { situacaoRaw, numeroRaw, obRaw, favorecidoRaw, notaEmpenhoRaw, valor } = params;

  // 1. Deve estar na situação AO
  if (!isSituacaoAO(situacaoRaw)) {
    return false;
  }

  // 2. Valor deve ser positivo e válido
  if (typeof valor !== 'number' || isNaN(valor) || valor <= 0) {
    return false;
  }

  // 3. Número deve estar preenchido
  const numStr = String(numeroRaw || '').trim();
  if (
    !numStr ||
    numStr === '-' ||
    numStr === '0' ||
    numStr === 'N/I' ||
    /^(numero|número|total|pp|ordem)$/i.test(numStr)
  ) {
    return false;
  }

  // 4. OB deve estar preenchida
  const obStr = String(obRaw || '').trim();
  if (
    !obStr ||
    obStr === '-' ||
    obStr === '0' ||
    obStr === 'N/I' ||
    /^(ob|ordem|total|bancaria|bancária)$/i.test(obStr)
  ) {
    return false;
  }

  // 5. Favorecido deve estar preenchido e ser válido (não resumo/total)
  const favStr = String(favorecidoRaw || '').trim();
  if (
    !favStr ||
    favStr === '-' ||
    favStr === '0' ||
    favStr === 'N/I' ||
    favStr.toUpperCase() === 'NÃO INFORMADO' ||
    favStr.toUpperCase() === 'NAO INFORMADO' ||
    /^(total|subtotal|resumo|saldo|relatorio|relatório)/i.test(favStr)
  ) {
    return false;
  }

  // 6. Nota de Empenho deve estar preenchida
  const neStr = String(notaEmpenhoRaw || '').trim();
  if (
    !neStr ||
    neStr === '-' ||
    neStr === '0' ||
    neStr === 'N/I' ||
    /^(ne|nota|empenho|total)$/i.test(neStr)
  ) {
    return false;
  }

  return true;
}

/**
 * Identifica se o nome de uma coluna se refere a Fonte de Recursos
 */
export function isFonteHeader(headerName: any): boolean {
  if (!headerName) return false;
  const norm = String(headerName)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, ' ')
    .trim();
  return (
    norm === 'fonte' ||
    norm === 'fr' ||
    norm === 'cod fonte' ||
    norm === 'cd fonte' ||
    norm === 'codigo fonte' ||
    norm === 'fonte recurso' ||
    norm === 'fonte recursos' ||
    norm === 'fonte de recurso' ||
    norm === 'fonte de recursos' ||
    norm === 'fonte detalhada' ||
    norm.startsWith('fonte ') ||
    norm.endsWith(' fonte') ||
    norm.includes('fonte de recurso') ||
    norm.includes('fonte de recursos') ||
    norm.includes('fonte recurso') ||
    norm.includes('fonte recursos') ||
    norm.includes('cd fonte') ||
    norm.includes('cod fonte')
  );
}

/**
 * Identifica categoricamente se um valor representa um código/identificador de Fonte de Recursos
 * (e.g. 0.6.00.000600, 0.5.00.000500, 0.6.00, 0.5.00, 0.600, 0.500, 0,600, 0,500, 600, 500, 1500, 1600)
 * para impedir categoricamente que seja interpretado como valor monetário da despesa.
 */
export function isFonteString(val: any): boolean {
  if (val === null || val === undefined) return false;
  if (typeof val === 'number') {
    if (val === 600 || val === 500 || val === 1500 || val === 1600 || val === 0.6 || val === 0.5 || val === 100) return true;
  }
  const s = String(val).trim().toUpperCase();
  if (!s || s === '-' || s === '0') return false;

  // Padrões estruturados do SIGEF: 0.6.00.000600, 0.5.00.000500, 0.6.00, 0.5.00, 0.1.00, 1.5.00
  if (/\b\d\.\s?\d(?:\.\d{2})?(?:\.\d{4,6})?\b/.test(s)) return true;
  // Padrões como 0.600, 0,600, 0.500, 0,500, 1.500, 1,500, ou 0.6, 0.5, 0,6, 0,5
  if (/^0?[,\.][56]0{0,3}$/.test(s) || /^1[,\.][56]0{0,3}$/.test(s)) return true;
  // Códigos de fontes federais e estaduais isolados comuns em planilhas
  if (/^(?:0500|0600|1500|1600|0100|1000|600|500)$/.test(s)) return true;
  // Textos explícitos de fontes
  if (s.startsWith('FONTE') || s.startsWith('FR ') || s.includes('FONTE DE RECURSO') || s.includes('FONTE:') || s.includes('FONTE ')) return true;

  return false;
}

// Automatically detect Favorecido/Credor, Valor, Situação, PP, OB, Número, Nota Empenho, and Favorecido Nota Empenho columns in a sheet
export function autoDetectColumns(
  headers: string[],
  sampleRows?: Record<string, any>[]
): {
  favorecidoCol: string;
  valorCol: string;
  situacaoCol?: string;
  situacaoObCol?: string;
  ppCol?: string;
  obCol?: string;
  numeroCol?: string;
  notaEmpenhoCol?: string;
  favorecidoNECol?: string;
  cnpjCol?: string;
  processoCol?: string;
  fonteCol?: string;
  naturezaCol?: string;
} {
  let favorecidoCol = '';
  let valorCol = '';
  let situacaoCol = '';
  let situacaoObCol = '';
  let ppCol = '';
  let obCol = '';
  let numeroCol = '';
  let notaEmpenhoCol = '';
  let favorecidoNECol = '';

  const normalize = (s: string) =>
    String(s || '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toLowerCase()
      .replace(/[^a-z0-9]/g, ' ')
      .trim();

  // 1. Favorecido Nota Empenho (specific first so it doesn't get picked as plain favorecido or plain nota empenho)
  const favorecidoNEKeywords = [
    'favorecido nota empenho',
    'favorecido nota de empenho',
    'favorecido da nota de empenho',
    'favorecido da ne',
    'favorecido ne',
    'credor nota empenho',
    'credor da ne',
    'credor ne',
    'fornecedor nota empenho',
    'fornecedor ne',
    'nome favorecido ne',
    'nome credor ne',
  ];

  for (const h of headers) {
    const norm = normalize(h);
    for (const kw of favorecidoNEKeywords) {
      if (norm.includes(kw) || norm === kw) {
        favorecidoNECol = h;
        break;
      }
    }
    if (favorecidoNECol) break;
  }

  // 2. Nota Empenho (excluding Favorecido Nota Empenho)
  const notaEmpenhoKeywords = [
    'nota de empenho',
    'nota empenho',
    'nota empenho ne',
    'numero empenho',
    'numero ne',
    'num ne',
    'n ne',
    'nr ne',
    'num empenho',
    'empenho',
    'ne',
  ];

  for (const h of headers) {
    if (h === favorecidoNECol) continue;
    const norm = normalize(h);
    if (norm.includes('favorecido') || norm.includes('credor') || norm.includes('fornecedor')) continue;
    for (const kw of notaEmpenhoKeywords) {
      if (norm === kw || norm.startsWith(`${kw} `) || norm.endsWith(` ${kw}`) || norm.includes(` ${kw} `)) {
        notaEmpenhoCol = h;
        break;
      }
    }
    if (notaEmpenhoCol) break;
  }

  // 3. Favorecido / Credor (excluding Favorecido Nota Empenho and NE)
  const favorecidoKeywords = [
    'favorecido',
    'credor',
    'fornecedor',
    'razao social',
    'razao',
    'beneficiario',
    'beneficiaria',
    'interessado',
    'nome credor',
    'nome favorecido',
    'nome fornecedor',
    'nome empresa',
    'empresa',
    'contratada',
    'prestador',
    'cliente',
    'pessoa',
    'nome',
    'cnpj',
    'cpf',
    'titular',
  ];

  for (const h of headers) {
    if (h === favorecidoNECol || h === notaEmpenhoCol) continue;
    const norm = normalize(h);
    if ((norm.includes('empenho') || norm.includes(' ne ') || norm.endsWith(' ne')) && !norm.includes('favorecido') && !norm.includes('credor')) {
      continue;
    }
    for (const kw of favorecidoKeywords) {
      if (norm.includes(kw)) {
        favorecidoCol = h;
        break;
      }
    }
    if (favorecidoCol) break;
  }

  // Fallback if not found: allow any header matching favorecidoKeywords
  if (!favorecidoCol) {
    for (const h of headers) {
      if (h === favorecidoNECol || h === notaEmpenhoCol) continue;
      const norm = normalize(h);
      for (const kw of favorecidoKeywords) {
        if (norm.includes(kw)) {
          favorecidoCol = h;
          break;
        }
      }
      if (favorecidoCol) break;
    }
  }

  // 4. OB (Ordem Bancária)
  const obKeywords = [
    'ob',
    'ordem bancaria',
    'ordem bancária',
    'ordem pagamento',
    'ordem de pagamento',
    'numero ob',
    'num ob',
    'n ob',
    'nr ob',
    'cod ob',
  ];

  for (const h of headers) {
    const norm = normalize(h);
    for (const kw of obKeywords) {
      if (norm === kw || norm.startsWith(`${kw} `) || norm.endsWith(` ${kw}`) || norm.includes(` ${kw} `)) {
        obCol = h;
        break;
      }
    }
    if (obCol) break;
  }

  // Fallback: scan sample rows for OB pattern \d{4}OB\d+
  if (!obCol && sampleRows && sampleRows.length > 0) {
    for (const h of headers) {
      if (h === favorecidoCol || h === notaEmpenhoCol || h === favorecidoNECol) continue;
      const count = sampleRows.filter((r) => /\b\d{4}OB\d+\b/i.test(String(r[h] || ''))).length;
      if (count >= 1) {
        obCol = h;
        break;
      }
    }
  }

  // 5. Número / PP (Programação de Pagamento) / NL (Nota de Liquidação)
  if (sampleRows && sampleRows.length > 0) {
    // Check for NL pattern (\d{4}NL\d+ or \d+NL\d+)
    for (const h of headers) {
      if (h === obCol || h === favorecidoCol || h === notaEmpenhoCol || h === favorecidoNECol) continue;
      const nlCount = sampleRows.filter((r) => /\b(?:\d{2,4})?NL[\s\-_]*\d+\b/i.test(String(r[h] || ''))).length;
      if (nlCount >= 1) {
        numeroCol = h;
        break;
      }
    }

    // Check for PP pattern (\d{4}PP\d+)
    for (const h of headers) {
      if (h === obCol || h === favorecidoCol || h === notaEmpenhoCol || h === favorecidoNECol) continue;
      const ppCount = sampleRows.filter((r) => /\b\d{4}PP\d+\b/i.test(String(r[h] || ''))).length;
      if (ppCount >= 1) {
        ppCol = h;
        if (!numeroCol) numeroCol = h;
        break;
      }
    }
  }

  // Step 5b. Header 'numero' / 'num' / 'nº' quando os dados tiverem padrão \d{4}PP\d+ ou \d{4}NL\d+
  if (!numeroCol && sampleRows && sampleRows.length > 0) {
    for (const h of headers) {
      if (h === obCol || h === favorecidoCol || h === notaEmpenhoCol || h === favorecidoNECol) continue;
      const norm = normalize(h);
      if (norm === 'numero' || norm === 'num' || norm === 'n' || norm === 'nr' || norm === 'numero doc') {
        const hasPpData = sampleRows.some((r) => /\b\d{4}PP\d+\b/i.test(String(r[h] || '')));
        const hasNlData = sampleRows.some((r) => /\b(?:\d{2,4})?NL[\s\-_]*\d+\b/i.test(String(r[h] || '')));
        if (hasPpData || hasNlData) {
          numeroCol = h;
          if (hasPpData && !ppCol) ppCol = h;
          break;
        }
      }
    }
  }

  // Step 5c. Specific NL and PP compound keywords
  if (!numeroCol) {
    const specificDocKeywords = [
      'nota de lancamento',
      'nota lancamento',
      'preparacao de pagamento',
      'preparacao pagamento',
      'prep de pagamento',
      'prep pagamento',
      'nota de liquidacao',
      'nota liquidacao',
      'numero nl',
      'num nl',
      'n nl',
      'nr nl',
      'documento habil',
      'doc habil',
      'liquidacao',
      'liquidac',
      'nl',
      'numero pp',
      'num pp',
      'n pp',
      'nr pp',
      'codigo pp',
      'pp',
      'proposta',
      'programacao',
      'programação',
    ];

    for (const h of headers) {
      if (h === obCol || h === notaEmpenhoCol) continue;
      const norm = normalize(h);
      if (
        norm.includes('valor') ||
        norm.includes('vlr') ||
        norm.includes('saldo') ||
        norm.includes('data') ||
        norm.includes('situacao') ||
        norm.includes('historico') ||
        norm.includes('status')
      ) {
        continue;
      }
      for (const kw of specificDocKeywords) {
        if (norm === kw || norm.startsWith(`${kw} `) || norm.endsWith(` ${kw}`) || norm.includes(` ${kw} `)) {
          numeroCol = h;
          if (!ppCol || kw.includes('pp') || kw.includes('preparacao') || kw.includes('prep') || kw.includes('proposta') || kw.includes('programac')) {
            ppCol = h;
          }
          break;
        }
      }
      if (numeroCol) break;
    }
  }

  // Step 5d. Fallback para cabeçalho 'numero'/'num'/'nr' com códigos numéricos não sequenciais (ex: NL numérica sem prefixo)
  if (!numeroCol) {
    for (const h of headers) {
      if (h === obCol || h === notaEmpenhoCol || h === favorecidoCol || h === favorecidoNECol) continue;
      const norm = normalize(h);
      if (norm === 'numero' || norm === 'num' || norm === 'nr' || norm === 'numero doc' || norm === 'nr doc') {
        const isRowNumberCol =
          sampleRows &&
          sampleRows.length > 1 &&
          sampleRows.every((r, idx) => {
            const val = String(r[h] || '').trim();
            return /^\d{1,3}$/.test(val) && parseInt(val, 10) === idx + 1;
          });
        if (!isRowNumberCol) {
          numeroCol = h;
          break;
        }
      }
    }
  }

  // Step 5e. Generic keywords ('documento', 'processo') but avoid row sequence lines (nº, item, ordem)
  if (!numeroCol) {
    const genericKeywords = ['documento', 'doc', 'processo'];
    for (const h of headers) {
      if (h === obCol || h === notaEmpenhoCol || h === favorecidoCol) continue;
      const norm = normalize(h);
      if (norm.includes('valor') || norm.includes('vlr') || norm.includes('data') || norm.includes('situacao')) continue;
      for (const kw of genericKeywords) {
        if (norm === kw || norm.startsWith(`${kw} `) || norm.endsWith(` ${kw}`) || norm.includes(` ${kw} `)) {
          const isRowNumberCol =
            sampleRows &&
            sampleRows.length > 0 &&
            sampleRows.every((r) => {
              const val = String(r[h] || '').trim();
              return /^\d{1,4}$/.test(val) && parseInt(val, 10) < 1000;
            });
          if (!isRowNumberCol) {
            numeroCol = h;
            break;
          }
        }
      }
      if (numeroCol) break;
    }
  }

  // 6. Valor (exact/compound keywords first, avoiding budget/dotacao/fonte columns)
  const isBudgetOrFonteHeader = (norm: string) =>
    isFonteHeader(norm) ||
    norm.includes('dotacao') ||
    norm.includes('orcamento') ||
    norm.includes('orcament') ||
    norm.includes('fixad') ||
    norm.includes('limite') ||
    norm.includes('saldo orcamentario') ||
    norm.includes('saldo dotacao') ||
    norm.includes('saldo orcamento');

  const valorKeywordsPrioritarias = [
    'valor da liquidacao',
    'valor da liq',
    'valor liquidacao',
    'valor liquidado',
    'vlr liquidacao',
    'vlr liquidado',
    'vlr da liquidacao',
    'valor da nl',
    'valor nl',
    'saldo a liquidar',
    'saldo a pagar',
    'valor a pagar',
    'valor atual',
    'saldo liquidado',
    'saldo da liquidacao',
    'total liquidado',
    'vlr liq',
    'valor liq',
    'valor da pp',
    'valor pp',
    'valor da ob',
    'valor ob',
    'valor liquido',
    'vlr liquido',
    'valor doc',
    'valor documento',
    'valor op',
    'valor original',
    'valor total',
    'valor bruto',
    'valor ne',
    'valor empenhado',
    'vlr total',
    'vlr bruto',
    'vlr doc',
    'valor pago',
    'vlr pago',
  ];

  const valorKeywordsGerais = [
    'valor',
    'liquido',
    'montante',
    'total',
    'saldo',
    'preco',
    'quantia',
    'pago',
    'vlr',
  ];

  // First pass: look for explicit compound payment/value columns
  for (const h of headers) {
    const norm = normalize(h);
    if (isBudgetOrFonteHeader(norm)) continue;
    for (const kw of valorKeywordsPrioritarias) {
      if (norm === kw || norm.startsWith(`${kw} `) || norm.endsWith(` ${kw}`) || norm.includes(` ${kw} `)) {
        valorCol = h;
        break;
      }
    }
    if (valorCol) break;
  }

  // Second pass: if not found, look for general value keywords, strictly avoiding budget and fonte columns
  if (!valorCol) {
    for (const h of headers) {
      const norm = normalize(h);
      if (isBudgetOrFonteHeader(norm)) continue;
      for (const kw of valorKeywordsGerais) {
        if (norm === kw || norm.startsWith(`${kw} `) || norm.endsWith(` ${kw}`) || norm.includes(` ${kw} `)) {
          valorCol = h;
          break;
        }
      }
      if (valorCol) break;
    }
  }

  // Third pass: contains keyword check as fallback
  if (!valorCol) {
    for (const h of headers) {
      const norm = normalize(h);
      if (isBudgetOrFonteHeader(norm)) continue;
      for (const kw of [...valorKeywordsPrioritarias, ...valorKeywordsGerais]) {
        if (norm.includes(kw)) {
          valorCol = h;
          break;
        }
      }
      if (valorCol) break;
    }
  }

  // 7. Situação
  // 7a. Situação da Ordem Bancária (específica de relatórios SIGEF com status da OB)
  const situacaoObKeywords = [
    'situacao ob',
    'situacao da ob',
    'situacao ordem bancaria',
    'situacao da ordem bancaria',
    'status ob',
    'status da ob',
    'sit ob',
    'sigla ob',
    'fase ob',
  ];
  for (const h of headers) {
    const norm = normalize(h);
    for (const kw of situacaoObKeywords) {
      if (norm === kw || norm.startsWith(`${kw} `) || norm.endsWith(` ${kw}`) || norm.includes(` ${kw} `)) {
        situacaoObCol = h;
        break;
      }
    }
    if (situacaoObCol) break;
  }

  // 7b. Situação Geral / Situação da PP
  const situacaoKeywords = [
    'situacao',
    'situacao pp',
    'situacao da pp',
    'situacao do documento',
    'situacao documento',
    'sigla',
    'sigla situacao',
    'sigla da situacao',
    'status',
    'sit pp',
    'sit',
    'estado',
    'fase',
    'fase pp',
    'fase da pp',
    'fase da despesa',
    'situacao ob',
    'situacao da ob',
  ];

  for (const h of headers) {
    if (h === situacaoObCol) continue;
    const norm = normalize(h);
    for (const kw of situacaoKeywords) {
      if (norm === kw || norm.startsWith(`${kw} `) || norm.endsWith(` ${kw}`) || norm.includes(` ${kw} `)) {
        situacaoCol = h;
        break;
      }
    }
    if (situacaoCol) break;
  }
  if (!situacaoCol && situacaoObCol) {
    situacaoCol = situacaoObCol;
  }

  // 8. CNPJ / CPF do Credor (caso venha em coluna separada do nome)
  let cnpjCol = '';
  const cnpjKeywords = ['cnpj', 'cpf', 'cnpj cpf', 'cpf cnpj', 'cnpj credor', 'cpf credor', 'documento credor', 'inscricao'];
  for (const h of headers) {
    if (h === favorecidoCol || h === notaEmpenhoCol || h === favorecidoNECol) continue;
    const norm = normalize(h);
    for (const kw of cnpjKeywords) {
      if (norm === kw || norm.startsWith(`${kw} `) || norm.endsWith(` ${kw}`) || norm.includes(` ${kw} `)) {
        cnpjCol = h;
        break;
      }
    }
    if (cnpjCol) break;
  }

  // 9. Processo Administrativo
  let processoCol = '';
  const processoKeywords = ['processo', 'num processo', 'n processo', 'nr processo', 'proc', 'sei', 'protocolo'];
  for (const h of headers) {
    const norm = normalize(h);
    for (const kw of processoKeywords) {
      if (norm === kw || norm.startsWith(`${kw} `) || norm.endsWith(` ${kw}`) || norm.includes(` ${kw} `)) {
        processoCol = h;
        break;
      }
    }
    if (processoCol) break;
  }

  // 10. Fonte de Recursos
  let fonteCol = '';
  const fonteKeywords = ['fonte', 'fonte recurso', 'fonte de recurso', 'fonte recursos', 'fonte de recursos', 'fonte rec', 'cod fonte', 'cd fonte'];
  for (const h of headers) {
    const norm = normalize(h);
    for (const kw of fonteKeywords) {
      if (norm === kw || norm.startsWith(`${kw} `) || norm.endsWith(` ${kw}`) || norm.includes(` ${kw} `)) {
        fonteCol = h;
        break;
      }
    }
    if (fonteCol) break;
  }

  // 11. Natureza da Despesa
  let naturezaCol = '';
  const naturezaKeywords = [
    'natureza da despesa',
    'natureza de despesa',
    'natureza despesa',
    'natureza',
    'nat da despesa',
    'nat despesa',
    'nat des',
    'cod natureza',
    'codigo natureza',
    'elemento de despesa',
    'elemento despesa',
  ];
  for (const h of headers) {
    const norm = normalize(h);
    for (const kw of naturezaKeywords) {
      if (norm === kw || norm.startsWith(`${kw} `) || norm.endsWith(` ${kw}`) || norm.includes(` ${kw} `)) {
        naturezaCol = h;
        break;
      }
    }
    if (naturezaCol) break;
  }

  // Pass 5: Heuristic inspection of data rows if still missing or to verify non-zero values
  if (sampleRows && sampleRows.length > 0) {
    const sample = sampleRows.slice(0, 35);

    // Heuristic for situacaoCol if not detected from headers
    // (no relatório .xls do SIGEF, a última coluna contém a sigla da situação mas o cabeçalho vem em branco, gerando Coluna_7 ou Coluna_14)
    if (!situacaoCol) {
      const siglaExact = new Set([
        'AO', 'PPCB', 'PPNC', 'PPRJ', 'CB', 'C', 'AP', 'RJ', 'PAGA', 'PAGO',
        'LD', 'LI', 'II', 'EI', 'EN', 'CM', 'CBNC', 'SD', 'SI', 'OB', 'REPP', 'RENE'
      ]);
      let bestSitCol = '';
      let bestSitMatches = 0;

      for (const h of headers) {
        if (h === valorCol || h === favorecidoCol || h === notaEmpenhoCol || h === favorecidoNECol || h === obCol || h === ppCol || h === numeroCol) continue;
        let matches = 0;
        sample.forEach((row) => {
          const val = String(row[h] ?? '').trim().toUpperCase();
          if (siglaExact.has(val) || isSituacaoAO(val) || isSituacaoEnviadaBanco(val).isEnviada) {
            matches++;
          }
        });
        if (matches > bestSitMatches && matches >= 1) {
          bestSitMatches = matches;
          bestSitCol = h;
        }
      }

      if (bestSitCol) {
        situacaoCol = bestSitCol;
      }
    }

    // Heuristic for valorCol if missing
    if (!valorCol) {
      let bestValCol = '';
      let bestValMatches = 0;
      for (const h of headers) {
        if (isBudgetOrFonteHeader(normalize(h))) continue;
        if (h === fonteCol) continue;
        const isColFonte = sample.some((row) => isFonteString(row[h]));
        if (isColFonte) continue;

        let numCount = 0;
        sample.forEach((row) => {
          const v = parseCurrencyValue(row[h]);
          if (v > 0) numCount++;
        });
        if (numCount > bestValMatches && numCount >= 2) {
          bestValMatches = numCount;
          bestValCol = h;
        }
      }
      if (bestValCol) valorCol = bestValCol;
    }

    // Smart verification: If valorCol was found but has ONLY 0 in all sample rows (or if it erroneously points to a fonte),
    // check if there's another column that has actual positive amounts (> 0),
    // avoiding the common trap where "Valor Pago = 0,00" overrides "Valor da Liquidação = 45.000,00"
    if (valorCol && sample.length > 0) {
      let positiveCount = 0;
      sample.forEach((row) => {
        if (parseCurrencyValue(row[valorCol]) > 0) positiveCount++;
      });

      if (positiveCount === 0 || isBudgetOrFonteHeader(normalize(valorCol))) {
        let altValCol = '';
        let altMatches = 0;
        for (const h of headers) {
          if (h === valorCol || isBudgetOrFonteHeader(normalize(h)) || h === fonteCol) continue;
          const isColFonte = sample.some((row) => isFonteString(row[h]));
          if (isColFonte) continue;

          let count = 0;
          sample.forEach((row) => {
            if (parseCurrencyValue(row[h]) > 0) count++;
          });
          if (count > altMatches && count >= 1) {
            altMatches = count;
            altValCol = h;
          }
        }
        if (altValCol) {
          valorCol = altValCol;
        }
      }
    }

    if (!favorecidoCol) {
      let bestFavCol = '';
      let bestFavMatches = 0;
      for (const h of headers) {
        if (h === valorCol) continue;
        let matchCount = 0;
        sample.forEach((row) => {
          const str = String(row[h] || '');
          if (/\d{2}\.?\d{3}\.?\d{3}\/?\d{4}-?\d{2}|\b\d{14}\b/.test(str) || str.length > 8) {
            matchCount++;
          }
        });
        if (matchCount > bestFavMatches && matchCount >= 2) {
          bestFavMatches = matchCount;
          bestFavCol = h;
        }
      }
      if (bestFavCol) favorecidoCol = bestFavCol;
    }
  }

  if (!favorecidoCol && headers.length > 0) favorecidoCol = headers[0];

  // Verificação de segurança estrita: valorCol NUNCA pode ser coluna de Fonte de Recursos
  if (valorCol && (isBudgetOrFonteHeader(normalize(valorCol)) || valorCol === fonteCol)) {
    valorCol = '';
  }

  if (!valorCol && headers.length > 0) {
    const safeValHeader = headers.find(
      (h) => !isBudgetOrFonteHeader(normalize(h)) && h !== fonteCol && h !== favorecidoCol
    );
    if (safeValHeader) valorCol = safeValHeader;
  }

  return {
    favorecidoCol,
    valorCol,
    situacaoCol: situacaoCol || undefined,
    situacaoObCol: situacaoObCol || undefined,
    ppCol: ppCol || undefined,
    obCol: obCol || undefined,
    numeroCol: numeroCol || undefined,
    notaEmpenhoCol: notaEmpenhoCol || undefined,
    favorecidoNECol: favorecidoNECol || undefined,
    cnpjCol: cnpjCol || undefined,
    processoCol: processoCol || undefined,
    fonteCol: fonteCol || undefined,
    naturezaCol: naturezaCol || undefined,
  };
}

/**
 * Analyzes file contents (headers and actual row data) specifically tailored for a given classification stage.
 * Used by ColumnMappingModal whenever the user changes "Classificação no Painel".
 */
export function analyzeColumnsForStage(
  headers: string[],
  rows?: Record<string, any>[],
  stage: StageKey = 'outros'
): {
  favorecidoCol: string;
  valorCol: string;
  situacaoCol?: string;
  situacaoObCol?: string;
  ppCol?: string;
  obCol?: string;
  numeroCol?: string;
  notaEmpenhoCol?: string;
  favorecidoNECol?: string;
  cnpjCol?: string;
  processoCol?: string;
  fonteCol?: string;
  naturezaCol?: string;
  dataCol?: string;
} {
  // First run the general baseline auto-detection
  const base = autoDetectColumns(headers, rows);
  let {
    favorecidoCol,
    valorCol,
    situacaoCol,
    situacaoObCol,
    ppCol,
    obCol,
    numeroCol,
    notaEmpenhoCol,
    favorecidoNECol,
    cnpjCol,
    processoCol,
    fonteCol,
    naturezaCol,
  } = base;

  let dataCol: string | undefined = undefined;

  const normalize = (s: string) =>
    String(s || '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toLowerCase()
      .replace(/[^a-z0-9]/g, ' ')
      .trim();

  const sample = (rows || []).slice(0, 40);

  // Scan for dataCol in headers and row contents
  const dataKeywords = [
    'data referencia',
    'data de referencia',
    'data aceite',
    'data vencimento',
    'data emissao',
    'data de emissao',
    'data liquidacao',
    'data de liquidacao',
    'data pagamento',
    'data pagto',
    'data',
    'dt',
    'emissao',
    'referencia',
  ];
  for (const h of headers) {
    const norm = normalize(h);
    for (const kw of dataKeywords) {
      if (norm === kw || norm.startsWith(`${kw} `) || norm.endsWith(` ${kw}`) || norm.includes(` ${kw} `) || norm.includes(kw)) {
        dataCol = h;
        break;
      }
    }
    if (dataCol) break;
  }
  if (!dataCol && sample.length > 0) {
    for (const h of headers) {
      if (sample.some((r) => /\b\d{2}\/\d{2}\/\d{4}\b/.test(String(r[h] || '')))) {
        dataCol = h;
        break;
      }
    }
  }

  // STAGE SPECIFIC ANALYSIS:
  if (stage === 'ordem_bancaria') {
    // 1. OB Column: se stage === 'ordem_bancaria' e o cabeçalho for 'numero'/'ob', ou se as células tiverem o padrão \d{4}OB\d+, mapeie como obCol
    for (const h of headers) {
      const norm = normalize(h);
      if (
        norm === 'numero' ||
        norm === 'num' ||
        norm === 'n' ||
        norm === 'ob' ||
        norm.startsWith('ob ') ||
        norm.endsWith(' ob') ||
        norm.includes(' ob ') ||
        norm.includes('ordem bancaria')
      ) {
        obCol = h;
        break;
      }
    }
    if (!obCol || !sample.some((r) => /\b\d{4}OB\d+\b/i.test(String(r[obCol] || '')))) {
      for (const h of headers) {
        const matchCount = sample.filter((r) => /\b\d{4}OB\d+\b/i.test(String(r[h] || ''))).length;
        if (matchCount >= 1) {
          obCol = h;
          break;
        }
      }
    }
    // 2. Fonte Column: Check content for 0.6.xx, .6., .5. or any isFonteString pattern
    for (const h of headers) {
      const matchCount = sample.filter((r) => {
        const v = String(r[h] || '');
        return /\b\d\.\s?\d\.\d{2}\.\d{6}\b/.test(v) || (v.includes('.6.') && !v.includes('/')) || (v.includes('.5.') && !v.includes('/'));
      }).length;
      if (matchCount >= 1) {
        fonteCol = h;
        break;
      }
    }
    if (!fonteCol) {
      for (const h of headers) {
        if (h === obCol || h === dataCol) continue;
        const matchCount = sample.filter((r) => isFonteString(r[h])).length;
        if (matchCount >= 2) {
          fonteCol = h;
          break;
        }
      }
    }
    // 3. PP Column: Check content for \d{4}PP\d+
    for (const h of headers) {
      const matchCount = sample.filter((r) => /\b\d{4}PP\d+\b/i.test(String(r[h] || ''))).length;
      if (matchCount >= 1) {
        ppCol = h;
        numeroCol = h;
        break;
      }
    }
    // 4. Favorecido / CNPJ: Check content for CNPJ or Company names (>10 chars, not OB/PP/date/fonte/bank)
    for (const h of headers) {
      if (h === obCol || h === fonteCol || h === ppCol || h === dataCol) continue;
      const cnpjMatches = sample.filter((r) => /\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}/.test(String(r[h] || ''))).length;
      if (cnpjMatches >= 1) {
        const hasCompanyNameInSameCol = sample.some((r) => {
          const s = String(r[h] || '').replace(/\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}/g, '').replace(/[^A-Za-zÀ-ÿ]/g, '').trim();
          return s.length >= 4;
        });
        if (hasCompanyNameInSameCol) {
          favorecidoCol = h;
        } else if (!cnpjCol) {
          cnpjCol = h;
        }
        break;
      }
    }
    if (!favorecidoCol || favorecidoCol === obCol || favorecidoCol === fonteCol || favorecidoCol === ppCol || favorecidoCol === dataCol || favorecidoCol === cnpjCol) {
      for (const h of headers) {
        if (h === obCol || h === fonteCol || h === ppCol || h === dataCol || h === cnpjCol || h === situacaoCol) continue;
        const nameMatches = sample.filter((r) => {
          const s = String(r[h] || '').trim();
          if (s.length <= 6 || !/[A-Za-zÀ-ÿ]{3,}/.test(s)) return false;
          if (/\b\d{4}(?:OB|PP|NL|NE)\d+\b/i.test(s) || isFonteString(s)) return false;
          if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) return false;
          return true;
        }).length;
        if (nameMatches >= 2) {
          favorecidoCol = h;
          break;
        }
      }
    }
    // 5. Situação: Check for CB, PPCB, RJ, etc.
    for (const h of headers) {
      if (h === obCol || h === fonteCol || h === ppCol || h === favorecidoCol || h === cnpjCol) continue;
      const sitMatches = sample.filter((r) => /\b(CB|RJ|PPCB|PPRJ|AO|C|AP|PAGA)\b/i.test(String(r[h] || '').trim())).length;
      if (sitMatches >= 1) {
        situacaoCol = h;
        break;
      }
    }
  } else if (stage === 'base_pps' || stage === 'pp_prontas') {
    // Base de PPs: Prioritize PP column in headers or content (incluindo preparacao pagamento e preparacao de pagamento)
    for (const h of headers) {
      const norm = normalize(h);
      if (
        norm === 'pp' ||
        norm.includes('numero pp') ||
        norm.includes('preparacao de pagamento') ||
        norm.includes('preparacao pagamento') ||
        norm.includes('prep de pagamento') ||
        norm.includes('prep pagamento') ||
        norm.includes('programacao') ||
        norm.includes('proposta')
      ) {
        ppCol = h;
        numeroCol = h;
        break;
      }
    }
    if (!ppCol && sample.length > 0) {
      for (const h of headers) {
        if (sample.some((r) => /\b\d{4}PP\d+\b/i.test(String(r[h] || '')))) {
          ppCol = h;
          numeroCol = h;
          break;
        }
      }
    }

    // In Base de PPs, ALSO actively detect the OB column if present in the SIGEF spreadsheet
    if (!obCol) {
      for (const h of headers) {
        const norm = normalize(h);
        if (norm === 'ob' || norm.includes('ordem bancaria') || norm.includes('numero ob') || norm.includes('n ob') || norm.includes('doc pagamento')) {
          obCol = h;
          break;
        }
      }
    }
    if (!obCol && sample.length > 0) {
      for (const h of headers) {
        if (sample.some((r) => /\b\d{4}OB\d+\b/i.test(String(r[h] || '')))) {
          obCol = h;
          break;
        }
      }
    }

    // Situação da OB específica no relatório de PP
    for (const h of headers) {
      const norm = normalize(h);
      if (
        norm.includes('situacao ob') ||
        norm.includes('situacao da ob') ||
        norm.includes('status ob') ||
        norm.includes('sit ob')
      ) {
        situacaoObCol = h;
        break;
      }
    }

    // Situação geral / Situação da PP
    for (const h of headers) {
      if (h === situacaoObCol) continue;
      const norm = normalize(h);
      if (norm.includes('situacao') || norm.includes('status') || norm.includes('sigla')) {
        situacaoCol = h;
        break;
      }
    }
    if (!situacaoCol && situacaoObCol) {
      situacaoCol = situacaoObCol;
    }
  } else if (stage === 'obedece' || stage === 'nao_obedece') {
    // Liquidação: Prioritize "valor da liquidacao" or "saldo a liquidar"
    for (const h of headers) {
      const norm = normalize(h);
      if (
        (norm.includes('liquidacao') || norm.includes('liquidac') || norm.includes('liquidar')) &&
        (norm.includes('valor') || norm.includes('vlr') || norm.includes('saldo'))
      ) {
        valorCol = h;
        break;
      }
    }
    // Número da liquidação (NL): priorizar padrão \d{4}NL\d+ nas linhas ou cabeçalhos explícitos de documento/NL (nunca Valor/Data/Situação)
    let matchedNlDataCol = '';
    if (sample.length > 0) {
      for (const h of headers) {
        if (h === obCol || h === notaEmpenhoCol || h === favorecidoCol || h === valorCol) continue;
        if (sample.some((r) => /\b(?:\d{2,4})?NL[\s\-_]*\d+\b/i.test(String(r[h] || '')))) {
          matchedNlDataCol = h;
          break;
        }
      }
    }
    if (matchedNlDataCol) {
      numeroCol = matchedNlDataCol;
    } else {
      for (const h of headers) {
        if (h === obCol || h === notaEmpenhoCol || h === favorecidoCol || h === valorCol) continue;
        const norm = normalize(h);
        if (
          norm.includes('valor') ||
          norm.includes('vlr') ||
          norm.includes('saldo') ||
          norm.includes('data') ||
          norm.includes('dt ') ||
          norm.includes('situacao') ||
          norm.includes('historico') ||
          norm.includes('status')
        ) {
          continue;
        }
        if (
          norm === 'nl' ||
          norm.includes(' nl') ||
          norm.startsWith('nl ') ||
          norm.includes('nota de liquidacao') ||
          norm.includes('nota liquidacao') ||
          norm.includes('nota de lancamento') ||
          norm.includes('nota lancamento') ||
          norm.includes('liquidacao') ||
          norm.includes('documento') ||
          norm === 'doc'
        ) {
          numeroCol = h;
          break;
        }
      }
    }
  } else if (stage === 'publicacao') {
    // Publicação: priorizar NL nos dados ou cabeçalho de documento/NL (nunca Valor/Data/Empenho)
    let matchedNlDataCol = '';
    if (sample.length > 0) {
      for (const h of headers) {
        if (h === obCol || h === notaEmpenhoCol || h === favorecidoCol || h === valorCol) continue;
        if (sample.some((r) => /\b(?:\d{2,4})?NL[\s\-_]*\d+\b/i.test(String(r[h] || '')))) {
          matchedNlDataCol = h;
          break;
        }
      }
    }
    if (matchedNlDataCol) {
      numeroCol = matchedNlDataCol;
    } else {
      for (const h of headers) {
        if (h === obCol || h === notaEmpenhoCol || h === favorecidoCol || h === valorCol) continue;
        const norm = normalize(h);
        if (
          norm.includes('valor') ||
          norm.includes('vlr') ||
          norm.includes('saldo') ||
          norm.includes('data') ||
          norm.includes('dt ') ||
          norm.includes('situacao') ||
          norm.includes('processo')
        ) {
          continue;
        }
        if (
          norm === 'nl' ||
          norm.includes(' nl') ||
          norm.startsWith('nl ') ||
          norm.includes('nota de lancamento') ||
          norm.includes('nota lancamento') ||
          norm.includes('liquidacao') ||
          norm.includes('documento') ||
          norm === 'doc' ||
          norm === 'numero' ||
          norm === 'num'
        ) {
          numeroCol = h;
          break;
        }
      }
    }
  } else if (stage === 'natureza_despesa') {
    // Natureza: Look for natureza column in headers or 3.3.90... / 3390... in content
    for (const h of headers) {
      const norm = normalize(h);
      if (norm.includes('natureza') || norm.includes('elemento')) {
        naturezaCol = h;
        break;
      }
    }
    if (!naturezaCol && sample.length > 0) {
      for (const h of headers) {
        if (sample.some((r) => /\b3\.\d\.\d{2}\.\d{2}\b|\b3390\d+\b/.test(String(r[h] || '')))) {
          naturezaCol = h;
          break;
        }
      }
    }
  }

  return {
    favorecidoCol: favorecidoCol || headers[0] || '',
    valorCol: valorCol || headers[1] || headers[0] || '',
    situacaoCol: situacaoCol || undefined,
    situacaoObCol: situacaoObCol || undefined,
    ppCol: ppCol || undefined,
    obCol: obCol || undefined,
    numeroCol: numeroCol || undefined,
    notaEmpenhoCol: notaEmpenhoCol || undefined,
    favorecidoNECol: favorecidoNECol || undefined,
    cnpjCol: cnpjCol || undefined,
    processoCol: processoCol || undefined,
    fonteCol: fonteCol || undefined,
    naturezaCol: naturezaCol || undefined,
    dataCol: dataCol || undefined,
  };
}

// Split a line from CSV/TSV respecting double quotes
function splitCsvLine(line: string, delim: string): string[] {
  const result: string[] = [];
  let cur = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (c === '"') {
      inQuotes = !inQuotes;
    } else if (c === delim && !inQuotes) {
      result.push(cur.trim());
      cur = '';
    } else {
      cur += c;
    }
  }
  result.push(cur.trim());
  return result.map((col) => col.replace(/^["']|["']$/g, '').trim());
}

// Parse text lines when SheetJS did not split columns
function parseDelimitedText(text: string): any[][] {
  const lines = text.split(/\r?\n/).filter((l) => l.trim().length > 0);
  if (lines.length === 0) return [];

  const head = lines.slice(0, 10);
  let semi = 0;
  let tab = 0;
  let pipe = 0;
  let comma = 0;

  head.forEach((l) => {
    semi += (l.match(/;/g) || []).length;
    tab += (l.match(/\t/g) || []).length;
    pipe += (l.match(/\|/g) || []).length;
    comma += (l.match(/,/g) || []).length;
  });

  let delim = ';';
  let best = semi;
  if (tab > best) {
    delim = '\t';
    best = tab;
  }
  if (pipe > best) {
    delim = '|';
    best = pipe;
  }
  if (comma > best && best === 0) {
    delim = ',';
  }

  return lines.map((l) => splitCsvLine(l, delim));
}

// If SheetJS loaded a CSV into a single column with delimiters, fix it
function fixDelimitedColumnsIfNeeded(matrix: any[][]): any[][] {
  if (!matrix || matrix.length === 0) return matrix;

  const sample = matrix.slice(0, 15);
  const singleColRows = sample.filter(
    (r) => Array.isArray(r) && r.length === 1 && typeof r[0] === 'string'
  );

  if (singleColRows.length / sample.length > 0.5) {
    let countSemicolon = 0;
    let countTab = 0;
    let countPipe = 0;
    let countComma = 0;

    singleColRows.forEach((r) => {
      const s = String(r[0]);
      countSemicolon += (s.match(/;/g) || []).length;
      countTab += (s.match(/\t/g) || []).length;
      countPipe += (s.match(/\|/g) || []).length;
      countComma += (s.match(/,/g) || []).length;
    });

    let bestDelim = ';';
    let maxCount = countSemicolon;
    if (countTab > maxCount) {
      bestDelim = '\t';
      maxCount = countTab;
    }
    if (countPipe > maxCount) {
      bestDelim = '|';
      maxCount = countPipe;
    }
    if (countComma > maxCount && maxCount === 0) {
      bestDelim = ',';
      maxCount = countComma;
    }

    if (maxCount > 0) {
      return matrix.map((r) => {
        if (Array.isArray(r) && r.length === 1 && typeof r[0] === 'string') {
          return splitCsvLine(r[0], bestDelim);
        }
        return r;
      });
    }
  }

  return matrix;
}

// Fast HTML table parser that avoids heavy DOMParser DOM allocation on large files
function parseHtmlTableFast(html: string): any[][] {
  const rows: any[][] = [];
  const trRegex = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
  const tdRegex = /<(?:td|th)[^>]*>([\s\S]*?)<\/(?:td|th)>/gi;
  let trMatch: RegExpExecArray | null;
  while ((trMatch = trRegex.exec(html)) !== null) {
    const trContent = trMatch[1];
    const rowData: any[] = [];
    let tdMatch: RegExpExecArray | null;
    while ((tdMatch = tdRegex.exec(trContent)) !== null) {
      let cell = tdMatch[1];
      cell = cell
        .replace(/<[^>]*>/g, '')
        .replace(/&nbsp;/gi, ' ')
        .replace(/&amp;/gi, '&')
        .replace(/&lt;/gi, '<')
        .replace(/&gt;/gi, '>')
        .replace(/&quot;/gi, '"')
        .trim();
      rowData.push(cell);
    }
    if (rowData.some((c) => c !== '')) {
      rows.push(rowData);
    }
  }
  return rows;
}

// Detect and remove completely empty columns across a matrix
export function removeEmptyColumnsFromMatrix(matrix: any[][]): any[][] {
  if (!matrix || matrix.length === 0) return matrix;

  let maxCols = 0;
  for (const row of matrix) {
    if (Array.isArray(row) && row.length > maxCols) {
      maxCols = row.length;
    }
  }
  if (maxCols === 0) return matrix;

  const colHasData = new Array(maxCols).fill(false);
  for (const row of matrix) {
    if (!Array.isArray(row)) continue;
    for (let c = 0; c < row.length; c++) {
      const val = row[c];
      if (val !== null && val !== undefined && String(val).trim() !== '') {
        colHasData[c] = true;
      }
    }
  }

  // If all columns have data or none does, return as is
  if (colHasData.every(Boolean) || !colHasData.some(Boolean)) {
    return matrix;
  }

  return matrix.map((row) => {
    if (!Array.isArray(row)) return row;
    return row.filter((_, colIdx) => colHasData[colIdx]);
  });
}

// Detect purely administrative or page break rows from SIGEF
export function isSigefPageOrAdministrativeRow(row: any[]): boolean {
  if (!Array.isArray(row) || row.length === 0) return true;

  const nonEmptyCells = row
    .map((c) => String(c ?? '').trim())
    .filter((c) => c.length > 0);

  if (nonEmptyCells.length === 0) return true;

  // Single cell or small group of cells containing page/system text
  if (nonEmptyCells.length <= 3) {
    const hasDocOrMoney = nonEmptyCells.some(
      (c) =>
        /\b\d{2,4}(?:NL|PP|OB|NE)\d+\b/i.test(c) ||
        /\d{2}\.\d{3}\.\d{3}/.test(c) ||
        /R\$\s*\d|\d{1,3}(?:\.\d{3})*,\d{2}/.test(c)
    );
    if (hasDocOrMoney) return false;

    const combined = nonEmptyCells.join(' ').toLowerCase();
    if (
      combined.startsWith('total da pagina') ||
      combined.startsWith('total da página') ||
      combined.startsWith('subtotal da pagina') ||
      combined.startsWith('subtotal da página') ||
      combined.startsWith('página') ||
      combined.startsWith('pagina') ||
      combined.startsWith('pag.') ||
      combined.startsWith('emitido em') ||
      combined.startsWith('data da emiss') ||
      combined.startsWith('data/hora') ||
      combined.startsWith('hora:') ||
      combined.startsWith('sistema integrado') ||
      combined.startsWith('sigef') ||
      combined.startsWith('governo do estado') ||
      combined.startsWith('secretaria de') ||
      combined.startsWith('secretaria da') ||
      combined.startsWith('relatório') ||
      combined.startsWith('relatorio') ||
      combined.startsWith('quantidade de registros') ||
      /^[-=_]{3,}$/.test(combined)
    ) {
      return true;
    }
  }

  return false;
}

export interface RawSheetResult {
  sheetName: string;
  matrix: any[][];
  stage: StageKey;
}

function cleanAndFilterMatrix(rawRows: any[][]): any[][] {
  let cleanMatrix = rawRows.filter((row) => {
    if (!Array.isArray(row)) return false;
    const nonEmpty = row.filter((c) => c !== null && c !== undefined && String(c).trim() !== '');
    if (nonEmpty.length === 0) return false;
    if (isSigefPageOrAdministrativeRow(row)) return false;
    return true;
  });
  cleanMatrix = removeEmptyColumnsFromMatrix(cleanMatrix);
  return cleanMatrix;
}

/**
 * Verifica com rigor se uma linha é de fato um cabeçalho repetido da planilha,
 * evitando descartar linhas de dados legítimas que começam com 2025PP... ou 2025OB...
 */
function isSheetHeaderRow(row: any[]): boolean {
  if (!Array.isArray(row) || row.length === 0) return false;

  const hasDataCodes = row.some((c) => {
    const s = String(c ?? '').trim();
    return (
      /\b\d{2,4}(?:PP|OB|NL|NE|OP)\d+\b/i.test(s) ||
      /^\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}$/.test(s) ||
      /^\d{3}\.\d{3}\.\d{3}-\d{2}$/.test(s)
    );
  });
  if (hasDataCodes) return false;

  const cleanTexts = row.map((c) =>
    String(c ?? '')
      .trim()
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
  );

  const hasFavCol = cleanTexts.some(
    (c) => c === 'favorecido' || c === 'credor' || c === 'fornecedor' || c === 'interessado' || c.startsWith('favorecido ')
  );
  const hasValCol = cleanTexts.some((c) => c === 'valor' || c === 'vlr' || c.startsWith('valor ') || c === 'total');
  const hasDocCol = cleanTexts.some(
    (c) => c === 'numero' || c === 'pp' || c === 'ob' || c === 'nl' || c === 'ne' || c === 'nota empenho' || c.includes('preparacao')
  );

  return (hasFavCol && (hasValCol || hasDocCol)) || (hasDocCol && hasValCol);
}

// Reads all sheets from a spreadsheet, segregating pages with distinct stages (e.g. "Obedece - DD/MM" and "Não Obedece - DD/MM")
export async function readSpreadsheetSheets(file: File): Promise<RawSheetResult[]> {
  const lowerName = file.name.toLowerCase();

  // 1. Friendly check for PDF or Image
  if (lowerName.endsWith('.pdf') || file.type === 'application/pdf') {
    throw new Error(
      'O arquivo selecionado é um documento PDF. O sistema processa planilhas eletrônicas e arquivos de dados (.xlsx, .xls, .ods, .csv, .tsv, .txt, .xlsm, .xlsb). Exporte o relatório do seu sistema financeiro (SIAFI/SIGEF/Banco) em formato Excel ou CSV e tente novamente.'
    );
  }
  if (file.type.startsWith('image/')) {
    throw new Error(
      'O arquivo selecionado é uma imagem. Por favor, envie uma planilha em formato Excel (.xlsx, .xls, .ods) ou CSV.'
    );
  }

  // Small asynchronous pause to allow the browser to paint the loading spinner
  await new Promise((resolve) => setTimeout(resolve, 30));

  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = async (e) => {
      try {
        const buffer = e.target?.result as ArrayBuffer;
        if (!buffer || buffer.byteLength === 0) {
          throw new Error('O arquivo selecionado está completamente vazio (0 bytes).');
        }

        const uint8 = new Uint8Array(buffer);

        // Check if file starts with %PDF magic bytes
        if (uint8[0] === 0x25 && uint8[1] === 0x50 && uint8[2] === 0x44 && uint8[3] === 0x46) {
          throw new Error(
            'O arquivo selecionado é um documento PDF. Por favor, exporte o relatório do sistema financeiro em formato Excel (.xlsx/.xls) ou CSV.'
          );
        }

        const isZip = uint8[0] === 0x50 && uint8[1] === 0x4B; // PK (xlsx, ods, xlsm)
        const isOle = uint8[0] === 0xD0 && uint8[1] === 0xCF && uint8[2] === 0x11 && uint8[3] === 0xE0; // OLE2 (xls binário)
        const isCsvOrText =
          lowerName.endsWith('.csv') ||
          lowerName.endsWith('.tsv') ||
          lowerName.endsWith('.txt') ||
          file.type.includes('text/csv') ||
          file.type.includes('text/plain');

        // Fast path for CSV/TSV/TXT
        if (isCsvOrText && !isZip && !isOle) {
          let text = '';
          try {
            text = new TextDecoder('utf-8', { fatal: false }).decode(uint8);
          } catch {
            text = new TextDecoder('iso-8859-1').decode(uint8);
          }

          if (text.includes('\uFFFD') || (text.includes('Ã') && !text.includes('ão') && !text.includes('ões'))) {
            try {
              const textLatin1 = new TextDecoder('iso-8859-1').decode(uint8);
              if (textLatin1 && textLatin1.length > 0) text = textLatin1;
            } catch {
              // ignore
            }
          }

          const parsed = parseDelimitedText(text);
          if (parsed.length > 0) {
            const matrix = cleanAndFilterMatrix(fixDelimitedColumnsIfNeeded(parsed));
            if (matrix.length > 0) {
              resolve([{
                sheetName: file.name,
                matrix,
                stage: detectStageFromFileName(file.name),
              }]);
              return;
            }
          }
        }

        // Strategy A: SheetJS read with maximum performance & low-memory flags
        let workbook: XLSX.WorkBook | null = null;
        try {
          workbook = XLSX.read(uint8, {
            type: 'array',
            dense: true,
            cellFormula: false,
            cellHTML: false,
            cellStyles: false,
            cellText: false,
            raw: true,
            codepage: 1252,
          });
        } catch {
          workbook = null;
        }

        if (workbook && workbook.SheetNames && workbook.SheetNames.length > 0) {
          const normalizedLowerName = lowerName
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .replace(/[^a-z0-9]/g, ' ')
            .replace(/\s+/g, ' ')
            .trim();

          // STRICT FILTER FOR "OBs - sigef - Principal":
          // Lê exclusivamente a página "Página3" (ignorando demais abas da planilha).
          // Se houver cabeçalhos/resumos nas linhas 1 a 6, inicia na linha 7 física do Excel (índice 6);
          // se a aba já iniciar com dados de OBs antes da linha 7, lê desde o início da aba.
          const isObsSigefPrincipal =
            (normalizedLowerName.includes('obs') || normalizedLowerName.includes('ob')) &&
            normalizedLowerName.includes('sigef') &&
            normalizedLowerName.includes('principal');

          if (isObsSigefPrincipal) {
            const normalizeSheet = (sName: string) =>
              sName
                .toLowerCase()
                .normalize('NFD')
                .replace(/[\u0300-\u036f]/g, '')
                .trim();

            const targetSheetName =
              workbook.SheetNames.find((sName) => normalizeSheet(sName).replace(/\s+/g, '') === 'pagina3') ||
              workbook.SheetNames.find((sName) => /\bpagina\s*3\b/.test(normalizeSheet(sName)));

            if (targetSheetName) {
              const ws = workbook.Sheets[targetSheetName];
              if (!ws || !ws['!ref']) {
                throw new Error(`A aba "${targetSheetName}" da planilha "${file.name}" está vazia.`);
              }

              const ref = XLSX.utils.decode_range(ws['!ref']);
              const EXCEL_ROW_7_INDEX = 6; // Linha 7 do Excel (1-indexed) = índice 6 (0-indexed)

              // Verifica nas linhas iniciais (1 a 5) se a "Página3" já traz registros reais de OB/PP antes da linha 7
              let hasEarlyObData = ref.e.r < EXCEL_ROW_7_INDEX;
              if (!hasEarlyObData && ref.s.r < EXCEL_ROW_7_INDEX - 1) {
                const earlyRows: any[][] = XLSX.utils.sheet_to_json(ws, {
                  header: 1,
                  defval: '',
                  blankrows: false,
                  range: {
                    s: { r: ref.s.r, c: ref.s.c },
                    e: { r: Math.min(ref.e.r, EXCEL_ROW_7_INDEX - 2), c: ref.e.c },
                  },
                });
                hasEarlyObData = earlyRows.some(
                  (row) =>
                    Array.isArray(row) &&
                    row.some((cell) => /\b\d{4}(?:OB|PP)\d{3,}\b/i.test(String(cell ?? '').trim()))
                );
              }

              const startRowIndex = hasEarlyObData ? ref.s.r : Math.max(ref.s.r, EXCEL_ROW_7_INDEX);
              const rowsFromStart: any[][] = XLSX.utils.sheet_to_json(ws, {
                header: 1,
                defval: '',
                blankrows: false,
                range: {
                  s: { r: startRowIndex, c: ref.s.c },
                  e: ref.e,
                },
              });

              const cleanRows = rowsFromStart.filter(
                (row) => Array.isArray(row) && row.some((c) => c !== null && c !== undefined && String(c).trim() !== '')
              );

              if (cleanRows.length === 0) {
                throw new Error(
                  `A aba "${targetSheetName}" da planilha "${file.name}" não contém informações de OBs.`
                );
              }

              // Se a primeira linha lida já for um registro de OB (ex: 2026OB..., PP ou CNPJ), preserva-a como dado:
              // utiliza os nomes das colunas da linha 6 (quando lendo a partir da linha 7) se for cabeçalho válido, ou gera Coluna_1..N.
              const firstRow = cleanRows[0];
              const firstRowIsObData = firstRow.some((cell) => {
                const s = String(cell ?? '').trim();
                return (
                  /\b\d{4}OB\d+\b/i.test(s) ||
                  /\b\d{4}PP\d+\b/i.test(s) ||
                  /\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}/.test(s)
                );
              });

              let finalRawRows = cleanRows;
              if (firstRowIsObData) {
                let prevHeader: any[] | null = null;
                if (!hasEarlyObData && ref.s.r <= EXCEL_ROW_7_INDEX - 1) {
                  const row6Arr: any[][] = XLSX.utils.sheet_to_json(ws, {
                    header: 1,
                    defval: '',
                    blankrows: true,
                    range: {
                      s: { r: EXCEL_ROW_7_INDEX - 1, c: ref.s.c },
                      e: { r: EXCEL_ROW_7_INDEX - 1, c: ref.e.c },
                    },
                  });
                  const candidate = row6Arr[0];
                  if (Array.isArray(candidate) && isSheetHeaderRow(candidate)) {
                    prevHeader = candidate;
                  }
                }

                if (prevHeader) {
                  finalRawRows = [prevHeader, ...cleanRows];
                } else {
                  const maxCols = cleanRows.reduce(
                    (max, r) => Math.max(max, Array.isArray(r) ? r.length : 0),
                    0
                  );
                  const syntheticHeaders = Array.from({ length: maxCols }, (_, idx) => `Coluna_${idx + 1}`);
                  finalRawRows = [syntheticHeaders, ...cleanRows];
                }
              }

              const cleanMatrix = cleanAndFilterMatrix(finalRawRows);
              if (cleanMatrix.length > 0) {
                resolve([{
                  sheetName: `${file.name} - ${targetSheetName}`,
                  matrix: cleanMatrix,
                  stage: 'ordem_bancaria',
                }]);
                return;
              }
            }
          }

          const sheetItems: { name: string; rows: any[][]; stage: StageKey }[] = [];

          for (const sName of workbook.SheetNames) {
            const ws = workbook.Sheets[sName];
            if (!ws || !ws['!ref']) continue;
            try {
              const sheetRows: any[][] = XLSX.utils.sheet_to_json(ws, {
                header: 1,
                defval: '',
                blankrows: false,
              });
              if (!sheetRows || sheetRows.length === 0) continue;

              const cleanSheetRows = sheetRows.filter(
                (row) => Array.isArray(row) && row.some((c) => c !== null && c !== undefined && String(c).trim() !== '')
              );
              if (cleanSheetRows.length === 0) continue;

              const detected = detectStageFromFileName(sName);
              sheetItems.push({
                name: sName,
                rows: cleanSheetRows,
                stage: detected,
              });
            } catch {
              // ignore sheet read error
            }
          }

          // Check for distinct functional stages across sheets
          // (e.g. "Processamento de dados" having "Obedece - DD/MM" and "Não Obedece - DD/MM")
          const isProcessamentoDeDados =
            lowerName.includes('processamento') && lowerName.includes('dados');

          const hasDistinctStages =
            sheetItems.length > 1 &&
            (sheetItems.some((s) => s.stage === 'obedece') || sheetItems.some((s) => s.name.toLowerCase().includes('obedece'))) &&
            (sheetItems.some((s) => s.stage === 'nao_obedece') || sheetItems.some((s) => s.name.toLowerCase().includes('nao') || s.name.toLowerCase().includes('não')));

          const hasMultipleKnownStages =
            sheetItems.length > 1 &&
            sheetItems.some((s) => s.stage !== 'outros');

          if (isProcessamentoDeDados) {
            // STRICT FILTER FOR "Processamento de dados":
            // Keeps only the latest valid liquidation sheets for "Obedece" and "Não Obedece"
            const obedeceCandidates = sheetItems.filter((s) => {
              if (s.stage === 'obedece') return true;
              const sn = s.name.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
              return sn.includes('obedece') && !sn.includes('nao');
            });

            const naoObedeceCandidates = sheetItems.filter((s) => {
              if (s.stage === 'nao_obedece') return true;
              const sn = s.name.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
              return sn.includes('nao obedece') || (sn.includes('nao') && sn.includes('obedec'));
            });

            const pickLastValidMatrix = (candidates: typeof sheetItems) => {
              for (let i = candidates.length - 1; i >= 0; i--) {
                const cleanMatrix = cleanAndFilterMatrix(candidates[i].rows);
                if (cleanMatrix.length >= 2) {
                  return { item: candidates[i], matrix: cleanMatrix };
                }
              }
              if (candidates.length > 0) {
                const last = candidates[candidates.length - 1];
                const cleanMatrix = cleanAndFilterMatrix(last.rows);
                if (cleanMatrix.length > 0) {
                  return { item: last, matrix: cleanMatrix };
                }
              }
              return null;
            };

            const pickedObedece = pickLastValidMatrix(obedeceCandidates);
            const pickedNaoObedece = pickLastValidMatrix(naoObedeceCandidates);

            const results: RawSheetResult[] = [];
            if (pickedObedece) {
              results.push({
                sheetName: pickedObedece.item.name,
                matrix: pickedObedece.matrix,
                stage: 'obedece',
              });
            }

            if (pickedNaoObedece) {
              results.push({
                sheetName: pickedNaoObedece.item.name,
                matrix: pickedNaoObedece.matrix,
                stage: 'nao_obedece',
              });
            }

            if (results.length > 0) {
              resolve(results);
              return;
            }
          }

          if (hasDistinctStages || hasMultipleKnownStages) {
            const stageGroups: Record<string, { sheetName: string; rows: any[][]; stage: StageKey }> = {};

            sheetItems.forEach((item) => {
              let finalStage = item.stage;
              const sn = item.name.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');

              if (finalStage === 'outros') {
                if ((sn.includes('nao') && sn.includes('obedec')) || sn.includes('nao cron')) {
                  finalStage = 'nao_obedece';
                } else if (sn.includes('obedece') || sn.includes('cronolog')) {
                  finalStage = 'obedece';
                } else if (sn.includes('publicac') || sn.includes('publicar') || /\bdoe\b/.test(sn)) {
                  finalStage = 'publicacao';
                }
              }

              // Se uma planilha nomeada explicitamente (ex: Publicação) tiver abas de capa/resumo sem etapa conhecida e menos de 2 linhas úteis, ignorar ruído
              const groupKey = `${finalStage}_${item.name}`;

              if (!stageGroups[groupKey]) {
                stageGroups[groupKey] = {
                  sheetName: item.name,
                  rows: [...item.rows],
                  stage: finalStage,
                };
              } else {
                const existing = stageGroups[groupKey];
                let headerRowIdx = -1;
                for (let r = 0; r < Math.min(item.rows.length, 5); r++) {
                  if (isSheetHeaderRow(item.rows[r])) {
                    headerRowIdx = r;
                    break;
                  }
                }
                const rowsToAdd = headerRowIdx >= 0 ? item.rows.slice(headerRowIdx + 1) : item.rows;
                existing.rows.push(...rowsToAdd);
              }
            });

            const results: RawSheetResult[] = [];
            for (const g of Object.values(stageGroups)) {
              const cleanMatrix = cleanAndFilterMatrix(g.rows);
              if (cleanMatrix.length >= 2) {
                results.push({
                  sheetName: g.sheetName,
                  matrix: cleanMatrix,
                  stage: g.stage,
                });
              }
            }

            if (results.length > 0) {
              resolve(results);
              return;
            }
          }

          // Default fallback: combine pages as continuation rows of the same single report
          // (ex: Sheet1, Sheet2, Sheet3 no relatório .xls do SIGEF onde apenas a primeira aba tem cabeçalho)
          const combinedRows: any[][] = [];
          for (const item of sheetItems) {
            if (combinedRows.length === 0) {
              combinedRows.push(...item.rows);
            } else {
              // Verifica se a aba de continuação repete o cabeçalho
              let headerRowIdx = -1;
              for (let r = 0; r < Math.min(item.rows.length, 5); r++) {
                if (isSheetHeaderRow(item.rows[r])) {
                  headerRowIdx = r;
                  break;
                }
              }

              // Se a aba de continuação não possui cabeçalho (caso do Sheet2/Sheet3 do SIGEF),
              // inclui todas as linhas a partir do índice 0 sem perder nenhum registro de dados
              const rowsToAdd = headerRowIdx >= 0
                ? item.rows.slice(headerRowIdx + 1)
                : item.rows;
              combinedRows.push(...rowsToAdd);
            }
          }

          if (combinedRows.length > 0) {
            const cleanMatrix = cleanAndFilterMatrix(combinedRows);
            if (cleanMatrix.length > 0) {
              resolve([{
                sheetName: file.name,
                matrix: cleanMatrix,
                stage: detectStageFromFileName(file.name),
              }]);
              return;
            }
          }
        }

        // Strategy B: For non-binary files (HTML tables or XML disguised as .xls)
        let text = '';
        try {
          text = new TextDecoder('utf-8', { fatal: false }).decode(uint8);
        } catch {
          text = new TextDecoder('iso-8859-1').decode(uint8);
        }

        if (text.includes('\uFFFD') || (text.includes('Ã') && !text.includes('ão') && !text.includes('ões'))) {
          try {
            const textLatin1 = new TextDecoder('iso-8859-1').decode(uint8);
            if (textLatin1 && textLatin1.length > 0) text = textLatin1;
          } catch {
            // ignore
          }
        }

        const lowerText = text.toLowerCase();
        let htmlOrXmlMatrix: any[][] = [];

        if (lowerText.includes('<table') || lowerText.includes('<tr')) {
          htmlOrXmlMatrix = parseHtmlTableFast(text);
        }

        if (htmlOrXmlMatrix.length === 0 && (lowerText.includes('<?xml') || lowerText.includes('<workbook'))) {
          try {
            const parser = new DOMParser();
            const doc = parser.parseFromString(text, 'application/xml');
            const rows = doc.querySelectorAll('Row');
            if (rows.length > 0) {
              const xmlMatrix: any[][] = [];
              rows.forEach((r) => {
                const rowData: any[] = [];
                r.querySelectorAll('Cell').forEach((c) => {
                  const dataElem = c.querySelector('Data');
                  rowData.push((dataElem ? dataElem.textContent : c.textContent) || '');
                });
                if (rowData.some((cell) => cell.trim() !== '')) {
                  xmlMatrix.push(rowData);
                }
              });
              if (xmlMatrix.length > 0) htmlOrXmlMatrix = xmlMatrix;
            }
          } catch (errXml) {
            console.warn('Falha ao processar XML Spreadsheet:', errXml);
          }
        }

        if (htmlOrXmlMatrix.length === 0 && text.trim().length > 0) {
          htmlOrXmlMatrix = parseDelimitedText(text);
        }

        if (htmlOrXmlMatrix.length > 0) {
          const cleanMatrix = cleanAndFilterMatrix(fixDelimitedColumnsIfNeeded(htmlOrXmlMatrix));
          if (cleanMatrix.length > 0) {
            resolve([{
              sheetName: file.name,
              matrix: cleanMatrix,
              stage: detectStageFromFileName(file.name),
            }]);
            return;
          }
        }

        throw new Error('A planilha inserida não contém linhas ou dados válidos. Verifique se o arquivo possui conteúdo.');
      } catch (err: any) {
        reject(err);
      }
    };

    reader.onerror = () =>
      reject(new Error('Erro de I/O ao ler o arquivo selecionado no disco. Verifique as permissões.'));
    reader.readAsArrayBuffer(file);
  });
}

// Universal, high-performance spreadsheet and matrix reader (optimized for large files)
export async function readSpreadsheetAsMatrix(file: File): Promise<any[][]> {
  const sheets = await readSpreadsheetSheets(file);
  return sheets[0]?.matrix || [];
}

export function cleanDigits(val: any): string {
  return String(val || '').replace(/\D/g, '');
}

export function formatCnpjOrCpf(input: string): string {
  if (!input) return 'N/I';
  const clean = cleanDigits(input);
  if (clean.length === 14) {
    return clean.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/, '$1.$2.$3/$4-$5');
  }
  if (clean.length === 13) {
    const padded = '0' + clean;
    return padded.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/, '$1.$2.$3/$4-$5');
  }
  if (clean.length === 12) {
    const padded = '00' + clean;
    return padded.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/, '$1.$2.$3/$4-$5');
  }
  if (clean.length === 11) {
    return clean.replace(/^(\d{3})(\d{3})(\d{3})(\d{2})$/, '$1.$2.$3-$4');
  }
  if (clean.length === 10) {
    const padded = '0' + clean;
    return padded.replace(/^(\d{3})(\d{3})(\d{3})(\d{2})$/, '$1.$2.$3-$4');
  }
  if (clean.length === 8) {
    return clean.replace(/^(\d{2})(\d{3})(\d{3})$/, '$1.$2.$3');
  }
  return input.trim();
}

// Parse CNPJ and Name using Regex from string like "08.234.112/0001-45 - HOSPITAL SANTA MARIA LTDA"
export function parseFavorecidoField(rawInput: any): { cnpj: string; name: string } {
  if (rawInput === null || rawInput === undefined) {
    return { cnpj: 'N/I', name: 'NÃO INFORMADO' };
  }

  const str = String(rawInput).trim();
  if (!str) return { cnpj: 'N/I', name: 'NÃO INFORMADO' };

  // 1. Full or shortened formatted CNPJ: 00.000.000/0000-00 or 00.000.000/0000
  const cnpjFormattedRegex = /(\d{2}\.\d{3}\.\d{3}\/\d{4}-?\d{2}|\d{2}\.\d{3}\.\d{3}\/\d{4})/;
  // 2. Formatted CPF: 000.000.000-00
  const cpfFormattedRegex = /(\d{3}\.\d{3}\.\d{3}-\d{2})/;
  // 3. Formatted 8-digit root CNPJ: 00.000.000
  const cnpjRootFormattedRegex = /(\d{2}\.\d{3}\.\d{3})/;
  // 4. Raw digits for CNPJ (12-14 digits) or CPF (10-11 digits)
  const cnpjDigitsRegex = /\b(\d{12,14})\b/;
  const cpfDigitsRegex = /\b(\d{10,11})\b/;
  // 5. Raw 8 digits (CNPJ base)
  const raw8DigitsRegex = /\b(\d{8})\b/;

  const match =
    str.match(cnpjFormattedRegex) ||
    str.match(cpfFormattedRegex) ||
    str.match(cnpjDigitsRegex) ||
    str.match(cpfDigitsRegex) ||
    str.match(cnpjRootFormattedRegex) ||
    str.match(raw8DigitsRegex);

  if (match) {
    const rawMatch = match[1];
    const formatted = formatCnpjOrCpf(rawMatch);
    let cleanName = str
      .replace(match[0], '')
      .replace(/^[\s\-\/\:\.\,]+|[\s\-\/\:\.\,]+$/g, '')
      .trim();
    if (!cleanName) cleanName = str;

    return {
      cnpj: formatted,
      name: cleanName.toUpperCase(),
    };
  }

  return {
    cnpj: 'N/I',
    name: str.toUpperCase(),
  };
}

// Converts an in-memory 2D matrix (from any spreadsheet sheet or CSV) into FileData
export function parseMatrixToFileData(
  matrix: any[][],
  fileName: string,
  forcedStage?: StageKey
): FileData {
  if (!matrix || matrix.length === 0) {
    throw new Error(`A planilha "${fileName}" não possui linhas de dados.`);
  }

  // Search for the header row in the first 50 rows
  let bestHeaderRowIndex = -1;
  let bestHeaderScore = -1;

  const headerKeywords = [
    'favorecido',
    'credor',
    'fornecedor',
    'beneficiario',
    'beneficiaria',
    'razao social',
    'razao',
    'nome',
    'interessado',
    'contratado',
    'cnpj',
    'cpf',
    'valor',
    'vlr',
    'liquido',
    'pago',
    'total',
    'bruto',
    'saldo',
    'pagamento',
    'liquidacao',
    'documento',
    'doc',
    'numero',
    'empenho',
    'ne',
    'data',
    'processo',
    'pp',
    'ob',
    'op',
    'nl',
    'situacao',
    'sigla',
    'status',
    'descricao',
    'fonte',
    'natureza',
  ];

  const hasSyntheticHeaders =
    Array.isArray(matrix[0]) && String(matrix[0][0] ?? '').trim() === 'Coluna_1';

  if (hasSyntheticHeaders) {
    bestHeaderRowIndex = 0;
  } else {
    for (let r = 0; r < Math.min(matrix.length, 50); r++) {
      const row = matrix[r];
      if (!Array.isArray(row)) continue;
      const cleanCells = row.map((c) => String(c ?? '').trim());
      const nonEmptyCells = cleanCells.filter((c) => c.length > 0);

      if (nonEmptyCells.length < 2) continue;

      const hasDataCodeInRow = cleanCells.some(
        (c) =>
          /\b\d{4}(?:OB|PP|NL|NE|OP)\d+\b/i.test(c) ||
          /\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}/.test(c)
      );
      if (hasDataCodeInRow) continue;

      let score = 0;
      let hasFavKw = false;
      let hasValKw = false;

      cleanCells.forEach((cell) => {
        const norm = cell
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .toLowerCase();
        for (const kw of headerKeywords) {
          if (norm.includes(kw)) {
            score += 4;
            if (
              kw.includes('favorecido') ||
              kw.includes('credor') ||
              kw.includes('fornecedor') ||
              kw.includes('nome') ||
              kw.includes('cnpj')
            ) {
              hasFavKw = true;
            }
            if (
              kw.includes('valor') ||
              kw.includes('vlr') ||
              kw.includes('liquido') ||
              kw.includes('pago') ||
              kw.includes('total')
            ) {
              hasValKw = true;
            }
            break;
          }
        }
      });

      score += nonEmptyCells.length;
      if (hasFavKw && hasValKw) score += 15;

      if (score > bestHeaderScore) {
        bestHeaderScore = score;
        bestHeaderRowIndex = r;
      }
    }

    // Fallback: if no keyword score was found, pick the first row with >= 2 non-empty cells
    if (bestHeaderRowIndex === -1 || bestHeaderScore < 4) {
      for (let r = 0; r < Math.min(matrix.length, 10); r++) {
        const row = matrix[r];
        if (Array.isArray(row) && row.filter((c) => String(c ?? '').trim() !== '').length >= 2) {
          bestHeaderRowIndex = r;
          break;
        }
      }
      if (bestHeaderRowIndex === -1) bestHeaderRowIndex = 0;
    }
  }

  const rawHeaderRow = matrix[bestHeaderRowIndex] || [];
  const headerCountMap: Record<string, number> = {};
  const headers: string[] = rawHeaderRow.map((cell, idx) => {
    let name = String(cell ?? '').trim();
    if (!name) name = `Coluna_${idx + 1}`;
    if (headerCountMap[name]) {
      headerCountMap[name]++;
      return `${name}_${headerCountMap[name]}`;
    }
    headerCountMap[name] = 1;
    return name;
  });

  const jsonRows: Record<string, any>[] = [];
  for (let r = bestHeaderRowIndex + 1; r < matrix.length; r++) {
    const row = matrix[r];
    if (!Array.isArray(row) || row.length === 0) continue;

    // Check if this row is a repeated table header row (common in multi-page SIGEF exports)
    const isRepeatedHeader =
      row.some((cell) => {
        const s = String(cell || '').toLowerCase().trim();
        return (
          (s === 'favorecido' || s === 'credor' || s === 'fornecedor') &&
          row.some((c2) => {
            const s2 = String(c2 || '').toLowerCase();
            return s2.includes('valor') || s2.includes('vlr') || s2 === 'ob' || s2 === 'pp';
          })
        );
      }) ||
      row.some((cell) => {
        const s = String(cell || '').toLowerCase().trim();
        return (s === 'pp' || s.includes('numero pp')) && row.some((c2) => String(c2 || '').toLowerCase().includes('situacao'));
      });
    if (isRepeatedHeader) continue;

    const rowObj: Record<string, any> = {};
    let hasAnyData = false;
    headers.forEach((h, colIdx) => {
      const val = row[colIdx] ?? '';
      rowObj[h] = val;
      if (val !== null && val !== undefined && String(val).trim() !== '') {
        hasAnyData = true;
      }
    });

    const isTotalOrSummaryRow =
      isSigefPageOrAdministrativeRow(row) ||
      row.some((cell) => {
        if (cell === null || cell === undefined) return false;
        const s = String(cell)
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .trim()
          .toLowerCase();
        return (
          s.startsWith('c - cancelada; ap - a pagar') ||
          s.startsWith('c - cancelada') ||
          s.startsWith('relatorio emitido em') ||
          s.startsWith('relatório emitido em') ||
          s === 'total' ||
          s === 'subtotal' ||
          s === 'total geral' ||
          s.startsWith('total geral') ||
          s.startsWith('total do orgao') ||
          s.startsWith('total da unidade') ||
          s.startsWith('total da despesa') ||
          s.startsWith('total do elemento') ||
          s.startsWith('total da fonte') ||
          s.startsWith('total da celula') ||
          s.startsWith('total do empenho') ||
          s.startsWith('resumo geral') ||
          s.startsWith('saldo orcamentario') ||
          s.startsWith('dotacao orcamentaria') ||
          s.startsWith('dotacao atualizada') ||
          s.startsWith('quantidade de')
        );
      });

    if (hasAnyData && !isTotalOrSummaryRow) {
      jsonRows.push(rowObj);
    }
  }

  if (jsonRows.length === 0) {
    throw new Error(`A planilha "${fileName}" não possui registros de dados válidos após o cabeçalho.`);
  }

  let autoStage: StageKey = forcedStage || detectStageFromFileName(fileName);
  const baselineCols = autoDetectColumns(headers, jsonRows);
  if (autoStage === 'outros') {
    const sampleSlice = jsonRows.slice(0, 30);
    const hasObRows = sampleSlice.some((r) =>
      Object.values(r).some((v) => /\b\d{4}OB\d+\b/i.test(String(v || '')))
    );
    const hasPpRows = sampleSlice.some((r) =>
      Object.values(r).some((v) => /\b\d{4}PP\d+\b/i.test(String(v || '')))
    );
    if (hasObRows && (baselineCols.fonteCol || !hasPpRows)) {
      autoStage = 'ordem_bancaria';
    } else if (hasPpRows) {
      autoStage = 'base_pps';
    }
  }

  const {
    favorecidoCol,
    valorCol,
    situacaoCol,
    situacaoObCol,
    ppCol,
    obCol,
    numeroCol,
    notaEmpenhoCol,
    favorecidoNECol,
    cnpjCol,
    processoCol,
    fonteCol,
    naturezaCol,
    dataCol,
  } = analyzeColumnsForStage(headers, jsonRows, autoStage);

  // Filtro estrito de subtotais ocultos do SIGEF EXCLUSIVAMENTE para Base de PPs / PPs Prontas.
  // Planilhas funcionais (obedece, nao_obedece, publicacao, natureza_despesa, ordem_bancaria) NÃO exigem PP ou OB preenchidos por linha.
  const isBasePpsStage = autoStage === 'base_pps' || autoStage === 'pp_prontas';
  let finalRows = jsonRows;
  if (isBasePpsStage) {
    finalRows = jsonRows.filter((r) => {
      const rawPp = ppCol ? String(r[ppCol] ?? '').trim() : '';
      const rawNum = numeroCol ? String(r[numeroCol] ?? '').trim() : '';
      const rawOb = obCol ? String(r[obCol] ?? '').trim() : '';
      const rawNe = notaEmpenhoCol ? String(r[notaEmpenhoCol] ?? '').trim() : '';

      // Descartar imediatamente se tanto Número (PP) quanto a OB estiverem vazios
      if (!rawPp && !rawNum && !rawOb) {
        return false;
      }
      // Descartar se a coluna Nota Empenho contiver apenas os 4 dígitos de um ano (ex: 2025, 2026)
      if (rawNe && /^\d{4}$/.test(rawNe)) {
        return false;
      }
      return true;
    });
  }

  return {
    id: `file_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`,
    fileName,
    headers,
    rows: finalRows,
    detectedFavorecidoCol: favorecidoCol,
    detectedValorCol: valorCol,
    detectedSituacaoCol: situacaoCol,
    detectedSituacaoObCol: situacaoObCol || undefined,
    detectedPpCol: ppCol || numeroCol,
    detectedObCol: obCol,
    detectedNumeroCol: numeroCol,
    detectedNotaEmpenhoCol: notaEmpenhoCol,
    detectedFavorecidoNECol: favorecidoNECol,
    detectedCnpjCol: cnpjCol,
    detectedProcessoCol: processoCol,
    detectedFonteCol: fonteCol,
    detectedNaturezaCol: naturezaCol,
    detectedDataCol: dataCol,
    stage: autoStage,
    totalCount: finalRows.length,
  };
}

// Parses workbook, extracting distinct pages with their individual stages (e.g. "Obedece - DD/MM" and "Não Obedece - DD/MM" from "Processamento de dados")
export async function parseWorkbookIntoFileDatas(
  file: File,
  fallbackStage?: StageKey
): Promise<FileData[]> {
  const rawSheets = await readSpreadsheetSheets(file);
  if (!rawSheets || rawSheets.length === 0) {
    throw new Error('A planilha selecionada está vazia ou não contém dados legíveis.');
  }

  const cleanBaseName = file.name.replace(/\.[^/.]+$/, '').trim();
  const fileExt = (file.name.match(/\.[^/.]+$/) || ['.xlsx'])[0];

  const parsedFiles: FileData[] = [];

  for (const s of rawSheets) {
    // When file contains multiple distinct pages, name them clearly, e.g. "Processamento de dados - Obedece - 06/11.xlsx"
    const finalFileName =
      rawSheets.length === 1
        ? file.name
        : `${cleanBaseName} - ${s.sheetName}${fileExt}`;

    const effectiveStage =
      s.stage !== 'outros'
        ? s.stage
        : fallbackStage || detectStageFromFileName(finalFileName);

    try {
      const fileData = parseMatrixToFileData(s.matrix, finalFileName, effectiveStage);
      parsedFiles.push(fileData);
    } catch (err: any) {
      console.warn(`Aviso ao processar aba "${s.sheetName}" de "${file.name}":`, err?.message || err);
    }
  }

  if (parsedFiles.length === 0) {
    throw new Error(`Nenhum registro válido pôde ser extraído de "${file.name}".`);
  }

  return parsedFiles;
}

// Parse Excel (.xlsx/.xls/.ods) or CSV/TSV/TXT file in browser
export async function parseExcelFile(file: File): Promise<FileData> {
  const list = await parseWorkbookIntoFileDatas(file);
  return list[0];
}

// Consolidate data from all uploaded files into supplier list
export function processUploadedFiles(files: FileData[]): DashboardMetrics {
  if (!files || files.length === 0) {
    return {
      grandTotalValue: 0,
      grandTotalCount: 0,
      totalBasePps: 0,
      totalPpProntas: 0,
      totalLiqObedece: 0,
      totalLiqNaoObedece: 0,
      totalOutros: 0,
      totalEnviadasBanco: 0,
      countPpProntasAO: 0,
      countEnviadasBanco: 0,
      consolidatedSuppliers: [],
      isProcessed: false,
    };
  }

  // Canonical transaction store across all files:
  // Ensures that when the user keeps previous files loaded, identical records/documents (PP, OB, NL)
  // are updated in the database rather than duplicated in financial totals.
  const transactionMap = new Map<
    string,
    {
      id: string;
      stage: StageKey;
      groupKey: string;
      name: string;
      cnpj: string;
      val: number;
      rawNumero: string;
      rawOB: string;
      rawNE: string;
      rawSit: string;
      isAO: boolean;
      enviadaInfo: ReturnType<typeof isSituacaoEnviadaBanco>;
      isPPProntaValida: boolean;
      hasAllSixColumns: boolean;
      fileHasSituacaoValues: boolean;
      isConclusivoPago: boolean;
      shouldCountInTotals: boolean;
      favNEInfo?: { cnpj: string; name: string } | null;
      obsList: string[];
      ppsList: string[];
      notasEmpenhoList: string[];
      naturezasList: string[];
    }
  >();

  // Maps to capture supplier classifications from natureza_despesa and fonte_recurso from ordem_bancaria without altering official statuses or financial totals
  const supplierNaturezasMap: Record<string, Set<string>> = {};
  const supplierEmpenhosMap: Record<string, Set<string>> = {};
  const supplierFontesAuxMap: Record<string, Record<string, number>> = {};
  const fonteByObMap = new Map<string, string>();
  const fonteByPpMap = new Map<string, string>();
  const neToGroupKey = new Map<string, string>();
  const docToGroupKey = new Map<string, string>();

  // Pass 1: Extract, deduplicate, and upsert records into canonical transaction database
  files.forEach((file) => {
    const favCol = file.detectedFavorecidoCol;
    const valCol = file.detectedValorCol;
    const cnpjCol = file.detectedCnpjCol;
    const numCol = file.detectedNumeroCol || file.detectedPpCol;
    const obCol = file.detectedObCol;
    const neCol = file.detectedNotaEmpenhoCol;
    const favNeCol = file.detectedFavorecidoNECol;
    const stage: StageKey = file.stage || detectStageFromFileName(file.fileName);
    const isObSourceFile = stage === 'ordem_bancaria' || detectStageFromFileName(file.fileName) === 'ordem_bancaria';

    // Pre-cache fallback columns from file.headers (or first row) once per file to avoid Object.keys inside row loops
    const headersList = file.headers && file.headers.length > 0
      ? file.headers
      : (file.rows && file.rows.length > 0 ? Object.keys(file.rows[0]) : []);

    let favFallbackCol = '';
    let sitFallbackCol = '';
    let natFallbackCol = '';
    let historicoCol = '';
    let fonteFallbackCol = '';

    for (const h of headersList) {
      const norm = h.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
      if (!favFallbackCol && (norm.includes('favorecido') || norm.includes('credor') || norm.includes('fornecedor') || norm.includes('nome'))) {
        favFallbackCol = h;
      }
      if (!sitFallbackCol && (norm.includes('situa') || norm.includes('status') || norm.includes('sigla') || norm === 'sit')) {
        sitFallbackCol = h;
      }
      if (!natFallbackCol && norm.includes('natureza')) {
        natFallbackCol = h;
      }
      if (!historicoCol && (norm.includes('historico') || norm.includes('histórico'))) {
        historicoCol = h;
      }
      if (!fonteFallbackCol && isFonteHeader(h)) {
        fonteFallbackCol = h;
      }
    }

    let natCol = file.detectedNaturezaCol || natFallbackCol;
    let sitCol = file.detectedSituacaoCol || sitFallbackCol;
    let fonteCol = file.detectedFonteCol || fonteFallbackCol;

    let fileHasSituacaoValues = false;
    if (sitCol) {
      for (let i = 0; i < Math.min(file.rows.length, 50); i++) {
        const val = file.rows[i]?.[sitCol];
        if (val !== null && val !== undefined && String(val).trim() !== '') {
          fileHasSituacaoValues = true;
          break;
        }
      }
    }

    const isBasePpsStage = stage === 'base_pps' || stage === 'pp_prontas';

    file.rows.forEach((row, rowIdx) => {
      let rawFav = row[favCol];
      let rawFavNE = favNeCol ? row[favNeCol] : undefined;
      const rawNumero = numCol ? String(row[numCol] || '').trim() : '';
      const rawOB = obCol ? String(row[obCol] || '').trim() : '';
      const rawNE = neCol ? String(row[neCol] || '').trim() : '';
      const val = parseCurrencyValue(row[valCol]);

      // Descartar linhas de rodapé/legenda do SIGEF (C - Cancelada; AP - A pagar / Relatório Emitido em)
      const isLegendOrFooter = Object.values(row).some((cell) => {
        if (cell === null || cell === undefined) return false;
        const s = String(cell).normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toLowerCase();
        return (
          s.startsWith('c - cancelada; ap - a pagar') ||
          s.startsWith('c - cancelada') ||
          s.startsWith('relatorio emitido em') ||
          s.startsWith('relatório emitido em')
        );
      });
      if (isLegendOrFooter) return;

      // Regra estrita SIGEF apenas para Base de PPs / PPs Prontas: descarta subtotais ocultos sem Número/PP e OB
      if (isBasePpsStage) {
        if (!rawNumero && !rawOB) {
          return;
        }
        if (rawNE && /^\d{4}$/.test(rawNE)) {
          return;
        }
      }

      // Descartar imediatamente linhas de totalizadores, resumos gerais e saldos de dotação
      if (isTotalizadorRow(row, rawFav, rawNumero)) {
        return;
      }

      if (rawFav === null || rawFav === undefined || String(rawFav).trim() === '') {
        if (rawFavNE && String(rawFavNE).trim() !== '') {
          rawFav = rawFavNE;
        } else if (favFallbackCol && row[favFallbackCol]) {
          rawFav = row[favFallbackCol];
        }
      }

      const favInfo = parseFavorecidoField(rawFav);
      const favNEInfo = rawFavNE ? parseFavorecidoField(rawFavNE) : null;

      let cnpj = favInfo.cnpj;
      let name = favInfo.name;

      // Se foi informada coluna explícita de CNPJ/CPF no mapeamento
      if (cnpjCol && row[cnpjCol]) {
        const rawCnpj = String(row[cnpjCol]).trim();
        if (rawCnpj && rawCnpj !== 'N/I') {
          const parsedDirect = parseFavorecidoField(rawCnpj);
          if (parsedDirect.cnpj !== 'N/I') {
            cnpj = parsedDirect.cnpj;
          } else {
            const digits = rawCnpj.replace(/\D/g, '');
            if (digits.length === 14) {
              cnpj = digits.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/, '$1.$2.$3/$4-$5');
            } else if (digits.length === 11) {
              cnpj = digits.replace(/^(\d{3})(\d{3})(\d{3})(\d{2})$/, '$1.$2.$3-$4');
            }
          }
        }
      }

      if (favNEInfo && favNEInfo.cnpj !== 'N/I' && (cnpj === 'N/I' || isBankOrIntermediary(name))) {
        cnpj = favNEInfo.cnpj;
        name = favNEInfo.name;
      }

      const hasAllSixColumns = !!(
        rawNumero &&
        rawOB &&
        rawFav &&
        rawNE &&
        rawFavNE &&
        val > 0
      );

      // Consolidar favorecidos principalmente a partir do CNPJ/CPF
      const digits = cleanDigits(cnpj);
      let groupKey = '';
      if (digits.length >= 12 && digits.length <= 14) {
        groupKey = 'CNPJ_' + formatCnpjOrCpf(digits).replace(/\D/g, '');
      } else if (digits.length >= 10 && digits.length <= 11) {
        groupKey = 'CPF_' + formatCnpjOrCpf(digits).replace(/\D/g, '');
      } else if (digits.length === 8) {
        groupKey = 'ROOT_' + digits;
      } else if (digits.length > 0) {
        groupKey = 'CNPJ_RAW_' + digits;
      } else if (
        name &&
        name !== 'NÃO INFORMADO' &&
        name.trim() !== '' &&
        name !== '-' &&
        name !== '0' &&
        name !== 'N/I' &&
        !name.startsWith('FORNECEDOR_REG') &&
        !/^(TOTAL|SUBTOTAL|RESUMO|QUANTIDADE|EMITIDO|SALDO|DOTAÇÃO|DOTACAO|ORGAO|UNIDADE|FONTE|ELEMENTO|CELULA|CÉLULA)/i.test(name.trim())
      ) {
        groupKey = 'NAME_' + name.trim().toUpperCase();
      } else {
        let matchedKey = '';
        if (rawNE) {
          matchedKey = neToGroupKey.get(rawNE) || '';
        }
        if (!matchedKey && rawNumero) {
          matchedKey = docToGroupKey.get(rawNumero) || '';
        }
        if (matchedKey) {
          groupKey = matchedKey;
        } else {
          return;
        }
      }

      if (rawNE && !neToGroupKey.has(rawNE)) {
        neToGroupKey.set(rawNE, groupKey);
      }
      if (rawNumero && !docToGroupKey.has(rawNumero)) {
        docToGroupKey.set(rawNumero, groupKey);
      }

      // Capturar Natureza da Despesa
      let natText = '';
      if (natCol && row[natCol]) {
        natText = String(row[natCol]).trim();
      } else if (natFallbackCol && row[natFallbackCol]) {
        natText = String(row[natFallbackCol]).trim();
      }

      // Se for arquivo exclusivo de natureza de despesa ou controle de publicação, não somar no saldo geral de credores
      if (stage === 'natureza_despesa') {
        if (natText && natText !== '-' && natText !== '0') {
          if (!supplierNaturezasMap[groupKey]) supplierNaturezasMap[groupKey] = new Set();
          supplierNaturezasMap[groupKey].add(natText);
        }
        if (rawNE && rawNE !== '-' && rawNE !== '0') {
          if (!supplierEmpenhosMap[groupKey]) supplierEmpenhosMap[groupKey] = new Set();
          supplierEmpenhosMap[groupKey].add(rawNE);
        }
        return;
      }

      if (stage === 'publicacao') {
        // Planilha dedicada ao monitoramento da publicação
        return;
      }

      if (isObSourceFile) {
        // Planilha de Ordem Bancária ("OBs - sigef - Principal"): atua exclusivamente como provedora de Fonte de Recursos e chaves OB/PP.
        // Nunca sobrescreve situações oficiais (rawSit, isAO, enviadaInfo, isPPProntaValida, isConclusivoPago) nem soma nos totais financeiros.
        let rawFonte = fonteCol && row[fonteCol] ? String(row[fonteCol]).trim().toUpperCase() : '';
        if (!rawFonte || rawFonte === '-' || rawFonte === '0' || !isFonteString(rawFonte)) {
          for (const cellVal of Object.values(row)) {
            if (typeof cellVal === 'string' && isFonteString(cellVal)) {
              rawFonte = cellVal.trim().toUpperCase();
              break;
            }
          }
        }
        const cleanDocKey = (s?: string) => String(s || '').trim().toUpperCase().replace(/[\s\.\-\/]/g, '');
        const obKey = cleanDocKey(rawOB);
        const ppKey = cleanDocKey(rawNumero);
        if (rawFonte && rawFonte !== '-' && rawFonte !== '0' && isFonteString(rawFonte)) {
          if (obKey) fonteByObMap.set(obKey, rawFonte);
          if (ppKey) fonteByPpMap.set(ppKey, rawFonte);
          if (groupKey && val > 0) {
            if (!supplierFontesAuxMap[groupKey]) supplierFontesAuxMap[groupKey] = {};
            supplierFontesAuxMap[groupKey][rawFonte] = (supplierFontesAuxMap[groupKey][rawFonte] || 0) + val;
          }
        }
        return;
      }

      if (name === 'NÃO INFORMADO' && val === 0 && String(rawFav || '').trim() === '') {
        return;
      }

      let rawSit = sitCol ? row[sitCol] : undefined;
      if ((rawSit === undefined || rawSit === null || String(rawSit).trim() === '') && sitFallbackCol) {
        rawSit = row[sitFallbackCol];
      }

      const isAO = isSituacaoAO(rawSit);
      const enviadaInfo = isSituacaoEnviadaBanco(rawSit);
      const isPPProntaValida = isPPProntaAOValida({
        situacaoRaw: rawSit,
        numeroRaw: rawNumero,
        obRaw: rawOB,
        favorecidoRaw: rawFav,
        notaEmpenhoRaw: rawNE,
        valor: val,
      });

      let shouldCountInTotals = false;
      if (stage === 'nao_obedece' || stage === 'obedece') {
        shouldCountInTotals = true;
      } else if (stage === 'base_pps' || stage === 'pp_prontas') {
        if (isAO || !fileHasSituacaoValues) {
          shouldCountInTotals = true;
        }
      } else if (!enviadaInfo.isEnviada) {
        shouldCountInTotals = true;
      }

      const isConclusivoPago = isSituacaoConclusivaPagamento(rawSit);

      // Gerar chave canônica de identificação do documento / transação
      const cleanDoc = (s?: string) => String(s || '').trim().toUpperCase().replace(/[\s\.\-\/]/g, '');
      let canonicalNum = rawNumero;
      if (!canonicalNum || /^\d{1,4}$/.test(canonicalNum) || /^linha\s*\d+$/i.test(canonicalNum)) {
        if (historicoCol && row[historicoCol]) {
          const s = String(row[historicoCol]).trim();
          const m = s.match(/\b(?:\d{2,4})?NL[\s\-_]*\d+\b/i) || s.match(/\b\d{4}PP\d+\b/i);
          if (m) {
            canonicalNum = m[0].toUpperCase().replace(/\s+/g, '');
          }
        }
      }
      const numDoc = cleanDoc(canonicalNum);
      const obDoc = cleanDoc(rawOB);
      const neDoc = cleanDoc(rawNE);
      const valCents = Math.round(val * 100);

      let txKey = '';
      if (stage === 'obedece' || stage === 'nao_obedece') {
        const uniquenessSuffix =
          !numDoc && !neDoc
            ? `_${groupKey}_${file.id}_${rowIdx}`
            : !numDoc
            ? `_${groupKey}`
            : '';
        txKey =
          `LIQ_${stage}_${numDoc || 'S_N'}` +
          (neDoc ? `_NE_${neDoc}` : '') +
          (valCents > 0 ? `_${valCents}` : '') +
          uniquenessSuffix;
      } else if (numDoc && (/\b\d{4}PP\d+\b/i.test(numDoc) || stage === 'base_pps' || stage === 'pp_prontas')) {
        // Cada proposta de pagamento é única no SIGEF. Se houver desdobramento por NE, inclui a NE
        txKey = `PP_${numDoc}` + (neDoc ? `_NE_${neDoc}` : '') + (valCents > 0 ? `_${valCents}` : '');
      } else if (obDoc && /\b\d{4}OB\d+\b/i.test(obDoc)) {
        txKey = `OB_${obDoc}` + (valCents > 0 ? `_${valCents}` : '');
      } else if (numDoc) {
        txKey = `DOC_${stage}_${numDoc}` + (neDoc ? `_NE_${neDoc}` : '') + (valCents > 0 ? `_${valCents}` : '');
      } else if (neDoc) {
        txKey = `NE_${neDoc}_${digits || cleanDoc(name)}_${valCents}`;
      } else {
        txKey = `TX_${stage}_${digits || cleanDoc(name)}_${valCents}_${file.id}_${rowIdx}`;
      }

      const existingTx = transactionMap.get(txKey);
      if (existingTx) {
        // A transação já existe em arquivo anterior!
        // Atualiza a linha no banco de dados (upsert) com as informações mais recentes, SEM DUPLICAR VALORES!
        // Planilhas de ordem_bancaria nunca devem sobrescrever situações oficiais.
        const isNewConclusivo = isSituacaoConclusivaPagamento(rawSit) || /^(PPCB|CB|PAGO|PAGA|RJ|PPRJ|C|CANCELAD|REJEITAD)/i.test(String(rawSit || '').trim());
        const isOldConclusivo = isSituacaoConclusivaPagamento(existingTx.rawSit) || /^(PPCB|CB|PAGO|PAGA|RJ|PPRJ|C|CANCELAD|REJEITAD)/i.test(String(existingTx.rawSit || '').trim());

        if (!isObSourceFile && (isNewConclusivo || !isOldConclusivo || (rawSit && rawSit !== '-' && rawSit !== '0'))) {
          existingTx.rawSit = rawSit || existingTx.rawSit;
          existingTx.isAO = isSituacaoAO(existingTx.rawSit);
          existingTx.enviadaInfo = isSituacaoEnviadaBanco(existingTx.rawSit);
          existingTx.isConclusivoPago = isSituacaoConclusivaPagamento(existingTx.rawSit);
          existingTx.isPPProntaValida = isPPProntaAOValida({
            situacaoRaw: existingTx.rawSit,
            numeroRaw: rawNumero || existingTx.rawNumero,
            obRaw: rawOB || existingTx.rawOB,
            favorecidoRaw: rawFav || existingTx.name,
            notaEmpenhoRaw: rawNE || existingTx.rawNE,
            valor: existingTx.val,
          });
        }

        if (rawOB) {
          if (!existingTx.rawOB) existingTx.rawOB = rawOB;
          if (!existingTx.obsList.includes(rawOB)) existingTx.obsList.push(rawOB);
        }
        if (rawNumero) {
          if (!existingTx.rawNumero) existingTx.rawNumero = rawNumero;
          if (!existingTx.ppsList.includes(rawNumero)) existingTx.ppsList.push(rawNumero);
        }
        if (rawNE) {
          if (!existingTx.rawNE) existingTx.rawNE = rawNE;
          if (!existingTx.notasEmpenhoList.includes(rawNE)) existingTx.notasEmpenhoList.push(rawNE);
        }
        if (name && name !== 'NÃO INFORMADO' && name.trim().length > (existingTx.name || '').length) {
          existingTx.name = name.trim().toUpperCase();
        }
        if (cnpj && cnpj !== 'N/I' && cleanDigits(cnpj).length > cleanDigits(existingTx.cnpj).length) {
          existingTx.cnpj = formatCnpjOrCpf(cnpj);
        }
        if (hasAllSixColumns) {
          existingTx.hasAllSixColumns = true;
        }
        if (favNEInfo && favNEInfo.name !== 'NÃO INFORMADO' && !existingTx.favNEInfo) {
          existingTx.favNEInfo = favNEInfo;
        }
        if (natText && !existingTx.naturezasList.includes(natText)) {
          existingTx.naturezasList.push(natText);
        }
      } else {
        // Nova transação identificada no banco de dados
        transactionMap.set(txKey, {
          id: txKey,
          stage,
          groupKey,
          name: name !== 'NÃO INFORMADO' ? name.trim().toUpperCase() : groupKey,
          cnpj: cnpj !== 'N/I' ? formatCnpjOrCpf(cnpj) : 'N/I',
          val,
          rawNumero,
          rawOB,
          rawNE,
          rawSit: rawSit || '',
          isAO,
          enviadaInfo,
          isPPProntaValida,
          hasAllSixColumns,
          fileHasSituacaoValues,
          isConclusivoPago,
          shouldCountInTotals,
          favNEInfo,
          obsList: rawOB ? [rawOB] : [],
          ppsList: rawNumero ? [rawNumero] : [],
          notasEmpenhoList: rawNE ? [rawNE] : [],
          naturezasList: natText ? [natText] : [],
        });
      }
    });
  });

  // Pass 2: Consolidar métricas gerais e totais dos fornecedores a partir das transações deduplicadas
  const supplierMap: Record<string, ConsolidatedSupplier> = {};
  let grandTotalValue = 0;
  let grandTotalCount = 0;

  let totalBasePps = 0;
  let totalPpProntas = 0;
  let totalLiqObedece = 0;
  let totalLiqNaoObedece = 0;
  let totalOutros = 0;
  let totalEnviadasBanco = 0;
  let countPpProntasAO = 0;
  let countEnviadasBanco = 0;
  let totalPagoAno = 0;
  let countPagoAnoTotal = 0;

  for (const tx of transactionMap.values()) {
    const val = tx.val;
    const stage = tx.stage;
    const isAO = tx.isAO;
    const enviadaInfo = tx.enviadaInfo;
    const fileHasSituacaoValues = tx.fileHasSituacaoValues;

    let isRowEnviadaBanco = false;
    let shouldCountInTotals = false;

    if (stage === 'nao_obedece') {
      totalLiqNaoObedece += val;
      shouldCountInTotals = true;
    } else if (stage === 'obedece') {
      totalLiqObedece += val;
      shouldCountInTotals = true;
    } else if (stage === 'base_pps' || stage === 'pp_prontas') {
      if (isAO) {
        totalBasePps += val;
        totalPpProntas += val;
        countPpProntasAO += 1;
        shouldCountInTotals = true;
      } else if (!fileHasSituacaoValues) {
        totalBasePps += val;
        totalPpProntas += val;
        shouldCountInTotals = true;
      }
      if (enviadaInfo.isEnviada) {
        totalEnviadasBanco += val;
        countEnviadasBanco += 1;
      }
    } else if (stage === 'ordem_bancaria') {
      // Alimentam o Detalhamento por Fonte via applyEnrichment
    } else if (stage === 'natureza_despesa') {
      shouldCountInTotals = false;
    } else if (enviadaInfo.isEnviada) {
      isRowEnviadaBanco = true;
      totalEnviadasBanco += val;
      countEnviadasBanco += 1;
    } else {
      totalOutros += val;
      shouldCountInTotals = true;
    }

    if (shouldCountInTotals) {
      grandTotalValue += val;
      grandTotalCount += 1;
    }

    const groupKey = tx.groupKey;
    if (!supplierMap[groupKey]) {
      supplierMap[groupKey] = {
        name: tx.name,
        cnpj: tx.cnpj,
        basePpsValue: 0,
        ppProntasValue: 0,
        liqObedeceValue: 0,
        liqNaoObedeceValue: 0,
        outrosValue: 0,
        enviadasBancoValue: 0,
        countPpProntasAO: 0,
        countEnviadasBanco: 0,
        valorPagoAno: 0,
        countPagoAno: 0,
        totalValue: 0,
        totalCount: 0,
        obs: [],
        pps: [],
        notasEmpenho: [],
        naturezasDespesa: [],
        favorecidoNE: tx.favNEInfo && tx.favNEInfo.name !== 'NÃO INFORMADO' ? tx.favNEInfo.name : undefined,
        hasFullSigefFields: tx.hasAllSixColumns,
      };
    }

    if (tx.hasAllSixColumns) {
      supplierMap[groupKey].hasFullSigefFields = true;
    }

    if (tx.name && tx.name !== 'NÃO INFORMADO') {
      const curName = supplierMap[groupKey].name || '';
      if (!curName || curName === 'NÃO INFORMADO' || tx.name.length > curName.length) {
        supplierMap[groupKey].name = tx.name;
      }
    }

    if (tx.cnpj && tx.cnpj !== 'N/I') {
      const curDigits = cleanDigits(supplierMap[groupKey].cnpj);
      const newDigits = cleanDigits(tx.cnpj);
      if (newDigits.length > curDigits.length) {
        supplierMap[groupKey].cnpj = tx.cnpj;
      }
    }

    tx.obsList.forEach((ob) => {
      supplierMap[groupKey].obs = supplierMap[groupKey].obs || [];
      if (!supplierMap[groupKey].obs!.includes(ob)) supplierMap[groupKey].obs!.push(ob);
    });
    tx.ppsList.forEach((pp) => {
      supplierMap[groupKey].pps = supplierMap[groupKey].pps || [];
      if (!supplierMap[groupKey].pps!.includes(pp)) supplierMap[groupKey].pps!.push(pp);
    });
    tx.notasEmpenhoList.forEach((ne) => {
      supplierMap[groupKey].notasEmpenho = supplierMap[groupKey].notasEmpenho || [];
      if (!supplierMap[groupKey].notasEmpenho!.includes(ne)) supplierMap[groupKey].notasEmpenho!.push(ne);
    });
    tx.naturezasList.forEach((nat) => {
      supplierMap[groupKey].naturezasDespesa = supplierMap[groupKey].naturezasDespesa || [];
      if (!supplierMap[groupKey].naturezasDespesa!.includes(nat)) supplierMap[groupKey].naturezasDespesa!.push(nat);
    });

    if (tx.favNEInfo && tx.favNEInfo.name !== 'NÃO INFORMADO' && !supplierMap[groupKey].favorecidoNE) {
      supplierMap[groupKey].favorecidoNE = tx.favNEInfo.name;
    }

    if (tx.isConclusivoPago && stage !== 'natureza_despesa' && stage !== 'ordem_bancaria') {
      supplierMap[groupKey].valorPagoAno = (supplierMap[groupKey].valorPagoAno || 0) + val;
      supplierMap[groupKey].countPagoAno = (supplierMap[groupKey].countPagoAno || 0) + 1;
      totalPagoAno += val;
      countPagoAnoTotal += 1;
    }

    if (stage === 'nao_obedece') {
      supplierMap[groupKey].liqNaoObedeceValue += val;
    } else if (stage === 'obedece') {
      supplierMap[groupKey].liqObedeceValue += val;
    } else if (stage === 'base_pps' || stage === 'pp_prontas') {
      if (isAO || !fileHasSituacaoValues) {
        supplierMap[groupKey].basePpsValue = (supplierMap[groupKey].basePpsValue || 0) + val;
        supplierMap[groupKey].ppProntasValue += val;
      }
      if (isAO) {
        supplierMap[groupKey].countPpProntasAO = (supplierMap[groupKey].countPpProntasAO || 0) + 1;
      }
      if (enviadaInfo.isEnviada) {
        supplierMap[groupKey].enviadasBancoValue = (supplierMap[groupKey].enviadasBancoValue || 0) + val;
        supplierMap[groupKey].countEnviadasBanco = (supplierMap[groupKey].countEnviadasBanco || 0) + 1;
      }
    } else if (stage === 'natureza_despesa') {
      // Natureza não soma financeiro
    } else if (isRowEnviadaBanco) {
      supplierMap[groupKey].enviadasBancoValue = (supplierMap[groupKey].enviadasBancoValue || 0) + val;
      supplierMap[groupKey].countEnviadasBanco = (supplierMap[groupKey].countEnviadasBanco || 0) + 1;
    } else {
      supplierMap[groupKey].outrosValue += val;
    }

    if (shouldCountInTotals) {
      supplierMap[groupKey].totalValue += val;
      supplierMap[groupKey].totalCount += 1;
    }
  }

  // Pass 3: Anexar classificações de natureza de despesa e empenhos aos fornecedores correspondentes
  for (const [gKey, natSet] of Object.entries(supplierNaturezasMap)) {
    if (supplierMap[gKey]) {
      supplierMap[gKey].naturezasDespesa = supplierMap[gKey].naturezasDespesa || [];
      natSet.forEach((nat) => {
        if (!supplierMap[gKey].naturezasDespesa!.includes(nat)) {
          supplierMap[gKey].naturezasDespesa!.push(nat);
        }
      });
    }
  }
  for (const [gKey, neSet] of Object.entries(supplierEmpenhosMap)) {
    if (supplierMap[gKey]) {
      supplierMap[gKey].notasEmpenho = supplierMap[gKey].notasEmpenho || [];
      neSet.forEach((ne) => {
        if (!supplierMap[gKey].notasEmpenho!.includes(ne)) {
          supplierMap[gKey].notasEmpenho!.push(ne);
        }
      });
    }
  }
  for (const [gKey, fontesObj] of Object.entries(supplierFontesAuxMap)) {
    if (supplierMap[gKey] && !supplierMap[gKey].pagamentosDetalhePorFonte) {
      supplierMap[gKey].pagamentosDetalhePorFonte = { ...fontesObj };
    }
  }
  if (fonteByObMap.size > 0 || fonteByPpMap.size > 0) {
    for (const s of Object.values(supplierMap)) {
      if (!s.pagamentosDetalhePorFonte) {
        const matchedFontes = new Set<string>();
        (s.obs || []).forEach((ob) => {
          const f = fonteByObMap.get(String(ob || '').trim().toUpperCase().replace(/[\s\.\-\/]/g, ''));
          if (f) matchedFontes.add(f);
        });
        (s.pps || []).forEach((pp) => {
          const f = fonteByPpMap.get(String(pp || '').trim().toUpperCase().replace(/[\s\.\-\/]/g, ''));
          if (f) matchedFontes.add(f);
        });
        if (matchedFontes.size > 0 && !(s as any).fonteRecurso) {
          (s as any).fonteRecurso = Array.from(matchedFontes)[0];
        }
      }
    }
  }

  // Reconciliação e consolidação inteligente por CNPJ/CPF
  // 1. Vincula CNPJs encurtados (como no relatório de liquidação) ao CNPJ completo correspondente
  // 2. Garante que para o mesmo favorecido o nome com maior número de caracteres seja sempre preservado
  const supplierKeys = Object.keys(supplierMap);
  const fullEntries: Array<{ key: string; digits: string; root: string }> = [];
  const shortenedOrNameEntries: Array<{ key: string; digits: string }> = [];

  const fullEntriesByRoot = new Map<string, string[]>();
  const fullEntriesByName = new Map<string, string>();

  for (const k of supplierKeys) {
    const s = supplierMap[k];
    const d = cleanDigits(s.cnpj);
    if (d.length === 14 || d.length === 11) {
      const root = d.substring(0, 8);
      fullEntries.push({ key: k, digits: d, root });
      let rootList = fullEntriesByRoot.get(root);
      if (!rootList) {
        rootList = [];
        fullEntriesByRoot.set(root, rootList);
      }
      rootList.push(k);
      const cleanName = (s.name || '').trim().toUpperCase();
      if (cleanName && cleanName !== 'NÃO INFORMADO') {
        fullEntriesByName.set(cleanName, k);
      }
    } else {
      shortenedOrNameEntries.push({ key: k, digits: d });
    }
  }

  for (const shortItem of shortenedOrNameEntries) {
    if (!supplierMap[shortItem.key]) continue;
    const source = supplierMap[shortItem.key];
    const srcDigits = shortItem.digits;
    const srcName = (source.name || '').trim().toUpperCase();

    let matchedTargetKey: string | null = null;

    // Caso 1: Possui 8 dígitos da raiz do CNPJ -> busca correspondente indexado O(1)
    if (srcDigits.length >= 8) {
      const srcRoot = srcDigits.substring(0, 8);
      const rootMatches = fullEntriesByRoot.get(srcRoot);
      if (rootMatches && rootMatches.length === 1) {
        matchedTargetKey = rootMatches[0];
      } else if (rootMatches && rootMatches.length > 1) {
        // Se houver mais de uma filial, correlaciona pelo nome
        const byName = rootMatches.find((k) => {
          const tName = (supplierMap[k]?.name || '').trim().toUpperCase();
          return tName.includes(srcName) || srcName.includes(tName);
        });
        matchedTargetKey = byName || rootMatches[0];
      }
    }

    // Caso 2: Se não localizou por dígitos, correlaciona pelo nome exato ou prefixo
    if (!matchedTargetKey && srcName && srcName !== 'NÃO INFORMADO') {
      matchedTargetKey = fullEntriesByName.get(srcName) || null;
      if (!matchedTargetKey && srcName.length >= 6) {
        // Busca apenas nos primeiros 200 para manter UI 60fps
        for (let i = 0; i < Math.min(fullEntries.length, 250); i++) {
          const f = fullEntries[i];
          const tName = (supplierMap[f.key]?.name || '').trim().toUpperCase();
          if (tName.length >= 6 && (tName.startsWith(srcName) || srcName.startsWith(tName))) {
            matchedTargetKey = f.key;
            break;
          }
        }
      }
    }

    // Se encontrou o favorecido com CNPJ completo, unifica os dados
    if (matchedTargetKey && supplierMap[matchedTargetKey]) {
      const target = supplierMap[matchedTargetKey];

      // Regra: levar em consideração o que tiver mais caracteres de nome
      if (source.name.trim().length > target.name.trim().length) {
        target.name = source.name.trim().toUpperCase();
      }

      target.basePpsValue = (target.basePpsValue || 0) + (source.basePpsValue || 0);
      target.ppProntasValue = (target.ppProntasValue || 0) + (source.ppProntasValue || 0);
      target.liqObedeceValue = (target.liqObedeceValue || 0) + (source.liqObedeceValue || 0);
      target.liqNaoObedeceValue = (target.liqNaoObedeceValue || 0) + (source.liqNaoObedeceValue || 0);
      target.outrosValue = (target.outrosValue || 0) + (source.outrosValue || 0);
      target.enviadasBancoValue = (target.enviadasBancoValue || 0) + (source.enviadasBancoValue || 0);
      target.valorPagoAno = (target.valorPagoAno || 0) + (source.valorPagoAno || 0);
      target.countPagoAno = (target.countPagoAno || 0) + (source.countPagoAno || 0);
      target.totalValue = (target.totalValue || 0) + (source.totalValue || 0);
      target.totalCount = (target.totalCount || 0) + (source.totalCount || 0);
      target.countPpProntasAO = (target.countPpProntasAO || 0) + (source.countPpProntasAO || 0);
      target.countEnviadasBanco = (target.countEnviadasBanco || 0) + (source.countEnviadasBanco || 0);

      if (source.hasFullSigefFields) target.hasFullSigefFields = true;
      if (source.favorecidoNE && (!target.favorecidoNE || source.favorecidoNE.length > target.favorecidoNE.length)) {
        target.favorecidoNE = source.favorecidoNE;
      }

      if (source.obs) {
        target.obs = target.obs || [];
        source.obs.forEach((ob) => {
          if (!target.obs!.includes(ob)) target.obs!.push(ob);
        });
      }
      if (source.pps) {
        target.pps = target.pps || [];
        source.pps.forEach((p) => {
          if (!target.pps!.includes(p)) target.pps!.push(p);
        });
      }
      if (source.notasEmpenho) {
        target.notasEmpenho = target.notasEmpenho || [];
        source.notasEmpenho.forEach((ne) => {
          if (!target.notasEmpenho!.includes(ne)) target.notasEmpenho!.push(ne);
        });
      }
      if (source.naturezasDespesa && source.naturezasDespesa.length > 0) {
        target.naturezasDespesa = target.naturezasDespesa || [];
        source.naturezasDespesa.forEach((nat) => {
          if (!target.naturezasDespesa!.includes(nat)) target.naturezasDespesa!.push(nat);
        });
      }

      delete supplierMap[shortItem.key];
    }
  }

  // Filtrar favorecidos válidos (removendo qualquer registro fantasma, totalizador ou cabeçalho)
  const consolidatedSuppliers: ConsolidatedSupplier[] = Object.values(supplierMap).filter((s) => {
    const cleanName = (s.name || '').trim().toUpperCase();
    if (
      !cleanName ||
      cleanName === 'NÃO INFORMADO' ||
      cleanName === 'N/I' ||
      cleanName.startsWith('FORNECEDOR_REG') ||
      cleanName.startsWith('CNPJ_RAW_') ||
      cleanName.startsWith('NAME_') ||
      /^(TOTAL|SUBTOTAL|RESUMO|QUANTIDADE|EMITIDO|SALDO|DOTAÇÃO|DOTACAO|ORGAO|UNIDADE|FONTE|ELEMENTO|CELULA|CÉLULA)/i.test(cleanName)
    ) {
      return false;
    }
    return (
      (s.totalValue || 0) > 0 ||
      (s.totalCount || 0) > 0 ||
      (s.basePpsValue || 0) > 0 ||
      (s.liqObedeceValue || 0) > 0 ||
      (s.liqNaoObedeceValue || 0) > 0 ||
      (s.outrosValue || 0) > 0 ||
      (s.valorPagoAno || 0) > 0 ||
      Boolean(s.naturezasDespesa && s.naturezasDespesa.length > 0)
    );
  });
  consolidatedSuppliers.sort((a, b) => b.totalValue - a.totalValue);

  return {
    grandTotalValue,
    grandTotalCount,
    totalBasePps,
    totalPpProntas,
    totalLiqObedece,
    totalLiqNaoObedece,
    totalOutros,
    totalEnviadasBanco,
    countPpProntasAO,
    countEnviadasBanco,
    totalPagoAno,
    countPagoAnoTotal,
    consolidatedSuppliers,
    isProcessed: grandTotalValue > 0 || grandTotalCount > 0 || consolidatedSuppliers.length > 0,
  };
}

// Empty files default (sample data disabled per user request)
export function generateSampleFiles(): FileData[] {
  return [];
}
