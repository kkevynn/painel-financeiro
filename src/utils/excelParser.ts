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

  // 1. NÃO OBEDECE (Priority check before "obedece")
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

  // 2. OBEDECE
  if (
    normalized.includes('obedece') ||
    normalized.includes('cronologica') ||
    normalized.includes('cronologico') ||
    normalized.includes('ordem cronologica')
  ) {
    return 'obedece';
  }

  // 3. BASE DE PPS (including SIGEF reports such as Relatorio_*.xls and payment proposals)
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
    normalized.includes('ordem bancaria') ||
    normalized.includes('ordens bancarias') ||
    normalized.includes('relacao de ob') ||
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
    favStr.includes('saldo orcamentario') ||
    favStr.includes('dotacao orcamentaria')
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
      s === 'total geral' ||
      s.startsWith('total geral') ||
      s.startsWith('total do orgao') ||
      s.startsWith('total da unidade') ||
      s.startsWith('total da despesa') ||
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

// Automatically detect Favorecido/Credor, Valor, Situação, PP, OB, Número, Nota Empenho, and Favorecido Nota Empenho columns in a sheet
export function autoDetectColumns(
  headers: string[],
  sampleRows?: Record<string, any>[]
): {
  favorecidoCol: string;
  valorCol: string;
  situacaoCol?: string;
  ppCol?: string;
  obCol?: string;
  numeroCol?: string;
  notaEmpenhoCol?: string;
  favorecidoNECol?: string;
  cnpjCol?: string;
  processoCol?: string;
  fonteCol?: string;
} {
  let favorecidoCol = '';
  let valorCol = '';
  let situacaoCol = '';
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

  // 5. Número / PP (Programação de Pagamento)
  const numeroKeywords = [
    'numero pp',
    'num pp',
    'n pp',
    'nr pp',
    'proposta',
    'programacao',
    'programação',
    'codigo pp',
    'pp',
    'numero',
    'número',
    'num',
    'nº',
  ];

  for (const h of headers) {
    if (h === obCol || h === notaEmpenhoCol) continue;
    const norm = normalize(h);
    for (const kw of numeroKeywords) {
      if (norm === kw || norm.startsWith(`${kw} `) || norm.endsWith(` ${kw}`) || norm.includes(` ${kw} `)) {
        numeroCol = h;
        ppCol = h;
        break;
      }
    }
    if (numeroCol) break;
  }

  // 6. Valor (exact/compound keywords first, avoiding budget/dotacao columns with multi-billion totals)
  const isBudgetHeader = (norm: string) =>
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
    if (isBudgetHeader(norm)) continue;
    for (const kw of valorKeywordsPrioritarias) {
      if (norm === kw || norm.startsWith(`${kw} `) || norm.endsWith(` ${kw}`) || norm.includes(` ${kw} `)) {
        valorCol = h;
        break;
      }
    }
    if (valorCol) break;
  }

  // Second pass: if not found, look for general value keywords, strictly avoiding budget columns
  if (!valorCol) {
    for (const h of headers) {
      const norm = normalize(h);
      if (isBudgetHeader(norm)) continue;
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
      if (isBudgetHeader(norm)) continue;
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
  const situacaoKeywords = [
    'situacao',
    'situacao pp',
    'situacao ob',
    'situacao da pp',
    'situacao da ob',
    'situacao do documento',
    'situacao documento',
    'sigla',
    'sigla situacao',
    'sigla da situacao',
    'status',
    'sit pp',
    'sit ob',
    'sit',
    'estado',
    'fase',
    'fase pp',
    'fase da pp',
    'fase da despesa',
  ];

  for (const h of headers) {
    const norm = normalize(h);
    for (const kw of situacaoKeywords) {
      if (norm === kw || norm.startsWith(`${kw} `) || norm.endsWith(` ${kw}`) || norm.includes(` ${kw} `)) {
        situacaoCol = h;
        break;
      }
    }
    if (situacaoCol) break;
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

  // Pass 5: Heuristic inspection of data rows if still missing or to verify non-zero values
  if (sampleRows && sampleRows.length > 0) {
    const sample = sampleRows.slice(0, 35);

    // Heuristic for situacaoCol if not detected from headers
    if (!situacaoCol) {
      for (const h of headers) {
        if (h === valorCol || h === favorecidoCol || h === notaEmpenhoCol || h === favorecidoNECol) continue;
        let aoMatches = 0;
        sample.forEach((row) => {
          const val = row[h];
          if (isSituacaoAO(val) || isSituacaoEnviadaBanco(val).isEnviada) {
            aoMatches++;
          }
        });
        if (aoMatches >= 1) {
          situacaoCol = h;
          break;
        }
      }
    }

    // Heuristic for valorCol if missing
    if (!valorCol) {
      let bestValCol = '';
      let bestValMatches = 0;
      for (const h of headers) {
        if (isBudgetHeader(normalize(h))) continue;
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

    // Smart verification: If valorCol was found but has ONLY 0 in all sample rows,
    // check if there's another column that has actual positive amounts (> 0),
    // avoiding the common trap where "Valor Pago = 0,00" overrides "Valor da Liquidação = 45.000,00"
    if (valorCol && sample.length > 0) {
      let positiveCount = 0;
      sample.forEach((row) => {
        if (parseCurrencyValue(row[valorCol]) > 0) positiveCount++;
      });

      if (positiveCount === 0) {
        let altValCol = '';
        let altMatches = 0;
        for (const h of headers) {
          if (h === valorCol || isBudgetHeader(normalize(h))) continue;
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
  if (!valorCol && headers.length > 1) valorCol = headers[1] || headers[0];

  return {
    favorecidoCol,
    valorCol,
    situacaoCol: situacaoCol || undefined,
    ppCol: ppCol || undefined,
    obCol: obCol || undefined,
    numeroCol: numeroCol || undefined,
    notaEmpenhoCol: notaEmpenhoCol || undefined,
    favorecidoNECol: favorecidoNECol || undefined,
    cnpjCol: cnpjCol || undefined,
    processoCol: processoCol || undefined,
    fonteCol: fonteCol || undefined,
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

// Universal, high-performance spreadsheet and matrix reader (optimized for large files)
export async function readSpreadsheetAsMatrix(file: File): Promise<any[][]> {
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

        let matrix: any[][] = [];

        // Fast path for CSV/TSV/TXT (orders of magnitude faster than SheetJS for large files)
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

          matrix = parseDelimitedText(text);
          if (matrix.length > 0) {
            resolve(fixDelimitedColumnsIfNeeded(matrix));
            return;
          }
        }

        // Strategy A: SheetJS read with maximum performance & low-memory flags
        let workbook: XLSX.WorkBook | null = null;
        try {
          workbook = XLSX.read(uint8, {
            type: 'array',
            dense: true,        // Uses array-of-arrays internally: saves ~85% RAM on large files!
            cellFormula: false, // Don't parse formulas
            cellHTML: false,    // Don't format HTML
            cellStyles: false,  // Don't parse styles
            cellText: false,    // Don't format text strings for every cell
            raw: true,          // Direct raw primitive values
            codepage: 1252,
          });
        } catch {
          workbook = null;
        }

        if (workbook && workbook.SheetNames && workbook.SheetNames.length > 0) {
          // Process all sheets in multi-page workbooks (e.g. daily SIGEF runs spanning multiple pages)
          const combinedRows: any[][] = [];

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

              if (combinedRows.length === 0) {
                combinedRows.push(...cleanSheetRows);
              } else {
                // If subsequent sheet/page repeats the header row, skip its first row
                const firstRow = cleanSheetRows[0] || [];
                const isRepeatedHeader = Array.isArray(firstRow) && firstRow.some((c) => {
                  const s = String(c || '').toLowerCase().trim();
                  return s.includes('favorecido') || s.includes('credor') || s.includes('valor') || s === 'ob' || s === 'pp' || s.includes('empenho');
                });
                const rowsToAdd = isRepeatedHeader ? cleanSheetRows.slice(1) : cleanSheetRows;
                combinedRows.push(...rowsToAdd);
              }
            } catch {
              // ignore sheet read error
            }
          }

          if (combinedRows.length > 0) {
            matrix = combinedRows;
          }
        }

        // Strategy B: For non-binary files (e.g. HTML tables or XML disguised as .xls)
        if ((!matrix || matrix.length === 0) && !isZip && !isOle) {
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

          // HTML Table check (common in public finance system exports)
          if (lowerText.includes('<table') || lowerText.includes('<tr')) {
            const htmlRows = parseHtmlTableFast(text);
            if (htmlRows.length > 0) {
              matrix = htmlRows;
            }
          }

          // XML Spreadsheet 2003 check
          if (matrix.length === 0 && (lowerText.includes('<?xml') || lowerText.includes('<workbook'))) {
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
                if (xmlMatrix.length > 0) matrix = xmlMatrix;
              }
            } catch (errXml) {
              console.warn('Falha ao processar XML Spreadsheet:', errXml);
            }
          }

          // Fallback: delimited text parser
          if (matrix.length === 0 && text.trim().length > 0) {
            matrix = parseDelimitedText(text);
          }
        }

        // Post-processing: fix un-split columns if needed
        if (matrix.length > 0) {
          matrix = fixDelimitedColumnsIfNeeded(matrix);
        }

        // 1. Filter out completely blank lines and administrative page break lines
        let cleanMatrix = matrix.filter((row) => {
          if (!Array.isArray(row)) return false;
          const nonEmpty = row.filter((c) => c !== null && c !== undefined && String(c).trim() !== '');
          if (nonEmpty.length === 0) return false;
          if (isSigefPageOrAdministrativeRow(row)) return false;
          return true;
        });

        // 2. Remove entirely empty columns
        cleanMatrix = removeEmptyColumnsFromMatrix(cleanMatrix);

        if (cleanMatrix.length === 0) {
          throw new Error(
            'A planilha inserida não contém linhas ou dados válidos. Verifique se o arquivo possui conteúdo ou tente exportá-lo novamente.'
          );
        }

        resolve(cleanMatrix);
      } catch (err: any) {
        reject(err);
      }
    };

    reader.onerror = () =>
      reject(new Error('Erro de I/O ao ler o arquivo selecionado no disco. Verifique as permissões.'));
    reader.readAsArrayBuffer(file);
  });
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

// Parse Excel (.xlsx/.xls/.ods) or CSV/TSV/TXT file in browser
export async function parseExcelFile(file: File): Promise<FileData> {
  const matrix = await readSpreadsheetAsMatrix(file);

  // Search for the header row in the first 35 rows
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
    'empenho',
    'ne',
    'data',
    'processo',
    'pp',
    'ob',
    'op',
    'situacao',
    'sigla',
    'status',
    'descricao',
    'fonte',
  ];

  for (let r = 0; r < Math.min(matrix.length, 35); r++) {
    const row = matrix[r];
    if (!Array.isArray(row)) continue;
    const cleanCells = row.map((c) => String(c ?? '').trim());
    const nonEmptyCells = cleanCells.filter((c) => c.length > 0);

    if (nonEmptyCells.length < 2) continue;

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
    const isRepeatedHeader = row.some((cell) => {
      const s = String(cell || '').toLowerCase().trim();
      return (s === 'favorecido' || s === 'credor') && row.some((c2) => String(c2 || '').toLowerCase().includes('valor'));
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

    const isTotalOrSummaryRow = row.some((cell) => {
      if (cell === null || cell === undefined) return false;
      const s = String(cell)
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .trim()
        .toLowerCase();
      return (
        s === 'total' ||
        s === 'subtotal' ||
        s === 'total geral' ||
        s.startsWith('total geral') ||
        s.startsWith('total do orgao') ||
        s.startsWith('total da unidade') ||
        s.startsWith('total da despesa') ||
        s.startsWith('total do elemento') ||
        s.startsWith('total da fonte') ||
        s.startsWith('resumo geral') ||
        s.startsWith('saldo orcamentario') ||
        s.startsWith('dotacao orcamentaria') ||
        s.startsWith('dotacao atualizada') ||
        s.startsWith('quantidade de') ||
        s.startsWith('emitido em') ||
        s.startsWith('pagina') ||
        s.startsWith('relatorio') ||
        s.startsWith('sigef')
      );
    });

    if (hasAnyData && !isTotalOrSummaryRow) {
      jsonRows.push(rowObj);
    }
  }

  if (jsonRows.length === 0) {
    throw new Error('A planilha selecionada não possui registros de dados válidos após o cabeçalho.');
  }

  const {
    favorecidoCol,
    valorCol,
    situacaoCol,
    ppCol,
    obCol,
    numeroCol,
    notaEmpenhoCol,
    favorecidoNECol,
    cnpjCol,
    processoCol,
    fonteCol,
  } = autoDetectColumns(headers, jsonRows);

  return {
    id: `file_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`,
    fileName: file.name,
    headers,
    rows: jsonRows,
    detectedFavorecidoCol: favorecidoCol,
    detectedValorCol: valorCol,
    detectedSituacaoCol: situacaoCol,
    detectedPpCol: ppCol || numeroCol,
    detectedObCol: obCol,
    detectedNumeroCol: numeroCol,
    detectedNotaEmpenhoCol: notaEmpenhoCol,
    detectedFavorecidoNECol: favorecidoNECol,
    detectedCnpjCol: cnpjCol,
    detectedProcessoCol: processoCol,
    detectedFonteCol: fonteCol,
    totalCount: jsonRows.length,
  };
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

  files.forEach((file) => {
    const favCol = file.detectedFavorecidoCol;
    const valCol = file.detectedValorCol;
    const cnpjCol = file.detectedCnpjCol;
    const numCol = file.detectedNumeroCol || file.detectedPpCol;
    const obCol = file.detectedObCol;
    const neCol = file.detectedNotaEmpenhoCol;
    const favNeCol = file.detectedFavorecidoNECol;
    const stage: StageKey = file.stage || detectStageFromFileName(file.fileName);

    // Check if the file has situation column
    let sitCol = file.detectedSituacaoCol;
    if (!sitCol && file.headers) {
      for (const h of file.headers) {
        const norm = h.toLowerCase().trim();
        if (norm.includes('situacao') || norm.includes('situação') || norm.includes('sigla') || norm.includes('status') || norm === 'sit') {
          sitCol = h;
          break;
        }
      }
    }

    // Check if any row in this file actually has a situation value
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

    file.rows.forEach((row) => {
      let rawFav = row[favCol];
      let rawFavNE = favNeCol ? row[favNeCol] : undefined;
      const rawNumero = numCol ? String(row[numCol] || '').trim() : '';
      const rawOB = obCol ? String(row[obCol] || '').trim() : '';
      const rawNE = neCol ? String(row[neCol] || '').trim() : '';
      const val = parseCurrencyValue(row[valCol]);

      // Descartar imediatamente linhas de totalizadores, resumos gerais e saldos de dotação
      if (isTotalizadorRow(row, rawFav, rawNumero)) {
        return;
      }

      if (rawFav === null || rawFav === undefined || String(rawFav).trim() === '') {
        if (rawFavNE && String(rawFavNE).trim() !== '') {
          rawFav = rawFavNE;
        } else {
          const keys = Object.keys(row);
          for (const k of keys) {
            const lk = k.toLowerCase();
            if ((lk.includes('favorecido') || lk.includes('credor') || lk.includes('fornecedor') || lk.includes('nome')) && row[k]) {
              rawFav = row[k];
              break;
            }
          }
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

      // In SIGEF reports, when Favorecido is a bank or payment intermediary and Favorecido Nota Empenho has the real contractor:
      if (favNEInfo && favNEInfo.cnpj !== 'N/I' && (cnpj === 'N/I' || isBankOrIntermediary(name))) {
        cnpj = favNEInfo.cnpj;
        name = favNEInfo.name;
      }

      // Check if this row has all 6 specific SIGEF columns filled:
      // Número / OB / Favorecido / Nota Empenho / Favorecido Nota Empenho / Valor
      const hasAllSixColumns = !!(
        rawNumero &&
        rawOB &&
        rawFav &&
        rawNE &&
        rawFavNE &&
        val > 0
      );

      if (name !== 'NÃO INFORMADO' || val !== 0 || String(rawFav || '').trim() !== '') {
        let rawSit = sitCol ? row[sitCol] : undefined;
        if (rawSit === undefined || rawSit === null || String(rawSit).trim() === '') {
          for (const k of Object.keys(row)) {
            const lk = k.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
            if (lk.includes('situa') || lk.includes('status') || lk.includes('sigla') || lk === 'sit') {
              if (row[k] !== undefined && row[k] !== null && String(row[k]).trim() !== '') {
                rawSit = row[k];
                break;
              }
            }
          }
        }

        const isAO = isSituacaoAO(rawSit);
        const enviadaInfo = isSituacaoEnviadaBanco(rawSit);

        // Regra estrita de PPs Prontas:
        // 1. Deve estar na situação AO
        // 2. Deve ter todas as 5 colunas preenchidas: Número, OB, Favorecido, Nota Empenho e Valor > 0
        const isPPProntaValida = isPPProntaAOValida({
          situacaoRaw: rawSit,
          numeroRaw: rawNumero,
          obRaw: rawOB,
          favorecidoRaw: rawFav,
          notaEmpenhoRaw: rawNE,
          valor: val,
        });

        let isRowPPPronta = false;
        let isRowEnviadaBanco = false;
        let shouldCountInTotals = false;

        // Se a planilha foi classificada ou mapeada explicitamente como Liquidação (Não Obedece / Obedece) ou Base de PPs:
        if (stage === 'nao_obedece') {
          totalLiqNaoObedece += val;
          shouldCountInTotals = true;
        } else if (stage === 'obedece') {
          totalLiqObedece += val;
          shouldCountInTotals = true;
        } else if (stage === 'base_pps' || stage === 'pp_prontas') {
          // Arquivo Base de PPs (ex: Relatório SIGEF com todas as informações das PPs)
          // Somente somar no painel inicial e na coluna PPs os registros na situação "AO"
          if (isAO) {
            totalBasePps += val;
            totalPpProntas += val;
            countPpProntasAO += 1;
            shouldCountInTotals = true;
          } else if (!fileHasSituacaoValues) {
            // Se o arquivo não contiver coluna ou dados de situação, soma normalmente
            totalBasePps += val;
            totalPpProntas += val;
            shouldCountInTotals = true;
          }
          if (enviadaInfo.isEnviada) {
            totalEnviadasBanco += val;
            countEnviadasBanco += 1;
          }
        } else if (enviadaInfo.isEnviada) {
          // Enviada a banco (PPCB / PPNC)
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

      // Consolidar favorecidos principalmente a partir do CNPJ/CPF
      const digits = cleanDigits(cnpj);
      let groupKey = '';
      if (digits.length >= 12 && digits.length <= 14) {
        // CNPJ completo ou com zeros suprimidos
        groupKey = 'CNPJ_' + formatCnpjOrCpf(digits).replace(/\D/g, '');
      } else if (digits.length >= 10 && digits.length <= 11) {
        // CPF completo ou com zero à esquerda suprimido
        groupKey = 'CPF_' + formatCnpjOrCpf(digits).replace(/\D/g, '');
      } else if (digits.length === 8) {
        // CNPJ encurtado (raiz) como ocorre em relatórios de liquidação
        groupKey = 'ROOT_' + digits;
      } else if (digits.length > 0) {
        groupKey = 'CNPJ_RAW_' + digits;
      } else if (name !== 'NÃO INFORMADO' && name.trim() !== '') {
        groupKey = 'NAME_' + name.trim().toUpperCase();
      } else {
        groupKey = `FORNECEDOR_REG_${grandTotalCount}`;
      }

      if (!supplierMap[groupKey]) {
        supplierMap[groupKey] = {
          name: name !== 'NÃO INFORMADO' ? name : groupKey,
          cnpj: cnpj !== 'N/I' ? formatCnpjOrCpf(cnpj) : 'N/I',
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
          favorecidoNE: favNEInfo && favNEInfo.name !== 'NÃO INFORMADO' ? favNEInfo.name : undefined,
          hasFullSigefFields: hasAllSixColumns,
        };
      }

      if (hasAllSixColumns) {
        supplierMap[groupKey].hasFullSigefFields = true;
      }

      // Priorizar o nome com mais caracteres para o favorecido
      if (name && name !== 'NÃO INFORMADO') {
        const curName = supplierMap[groupKey].name || '';
        if (!curName || curName === 'NÃO INFORMADO' || name.trim().length > curName.trim().length) {
          supplierMap[groupKey].name = name.trim().toUpperCase();
        }
      }

      // Atualizar CNPJ caso o novo seja mais completo
      if (cnpj && cnpj !== 'N/I') {
        const curDigits = cleanDigits(supplierMap[groupKey].cnpj);
        const newDigits = cleanDigits(cnpj);
        if (newDigits.length > curDigits.length) {
          supplierMap[groupKey].cnpj = formatCnpjOrCpf(cnpj);
        }
      }

      if (rawOB && (!supplierMap[groupKey].obs || !supplierMap[groupKey].obs!.includes(rawOB))) {
        supplierMap[groupKey].obs = supplierMap[groupKey].obs || [];
        supplierMap[groupKey].obs!.push(rawOB);
      }
      if (rawNumero && (!supplierMap[groupKey].pps || !supplierMap[groupKey].pps!.includes(rawNumero))) {
        supplierMap[groupKey].pps = supplierMap[groupKey].pps || [];
        supplierMap[groupKey].pps!.push(rawNumero);
      }
      if (rawNE && (!supplierMap[groupKey].notasEmpenho || !supplierMap[groupKey].notasEmpenho!.includes(rawNE))) {
        supplierMap[groupKey].notasEmpenho = supplierMap[groupKey].notasEmpenho || [];
        supplierMap[groupKey].notasEmpenho!.push(rawNE);
      }
      if (favNEInfo && favNEInfo.name !== 'NÃO INFORMADO' && !supplierMap[groupKey].favorecidoNE) {
        supplierMap[groupKey].favorecidoNE = favNEInfo.name;
      }

      // Histórico de pagamentos conclusivos do ano (PPCB, CB, etc.)
      const isConclusivoPago = isSituacaoConclusivaPagamento(rawSit);
      if (isConclusivoPago) {
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
  });
});

  // Reconciliação e consolidação inteligente por CNPJ/CPF
  // 1. Vincula CNPJs encurtados (como no relatório de liquidação) ao CNPJ completo correspondente
  // 2. Garante que para o mesmo favorecido o nome com maior número de caracteres seja sempre preservado
  const supplierKeys = Object.keys(supplierMap);
  const fullEntries: Array<{ key: string; digits: string; root: string }> = [];
  const shortenedOrNameEntries: Array<{ key: string; digits: string }> = [];

  for (const k of supplierKeys) {
    const s = supplierMap[k];
    const d = cleanDigits(s.cnpj);
    if (d.length === 14 || d.length === 11) {
      fullEntries.push({ key: k, digits: d, root: d.substring(0, 8) });
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

    // Caso 1: Possui 8 dígitos da raiz do CNPJ -> busca correspondente com a mesma raiz
    if (srcDigits.length >= 8) {
      const srcRoot = srcDigits.substring(0, 8);
      const matches = fullEntries.filter((f) => f.root === srcRoot);
      if (matches.length === 1) {
        matchedTargetKey = matches[0].key;
      } else if (matches.length > 1) {
        // Se houver mais de uma filial, correlaciona pelo nome
        const byName = matches.find((m) => {
          const tName = (supplierMap[m.key]?.name || '').trim().toUpperCase();
          return tName.includes(srcName) || srcName.includes(tName);
        });
        matchedTargetKey = byName ? byName.key : matches[0].key;
      }
    }

    // Caso 2: Se não localizou por dígitos, correlaciona pelo nome do favorecido
    if (!matchedTargetKey && srcName && srcName !== 'NÃO INFORMADO') {
      const byName = fullEntries.filter((f) => {
        const tName = (supplierMap[f.key]?.name || '').trim().toUpperCase();
        if (tName === srcName) return true;
        if (srcName.length >= 5 && tName.length >= 5) {
          return tName.startsWith(srcName) || srcName.startsWith(tName) || tName.includes(srcName) || srcName.includes(tName);
        }
        return false;
      });
      if (byName.length === 1) {
        matchedTargetKey = byName[0].key;
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

      delete supplierMap[shortItem.key];
    }
  }

  // Filtrar favorecidos válidos
  const consolidatedSuppliers: ConsolidatedSupplier[] = Object.values(supplierMap).filter((s) => {
    return (
      (s.totalValue || 0) > 0 ||
      (s.totalCount || 0) > 0 ||
      (s.basePpsValue || 0) > 0 ||
      (s.liqObedeceValue || 0) > 0 ||
      (s.liqNaoObedeceValue || 0) > 0 ||
      (s.outrosValue || 0) > 0 ||
      (s.valorPagoAno || 0) > 0
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
