import * as XLSX from 'xlsx';
import {
  FileData,
  CompanyAuditSummary,
  AuditRowItem,
  ConsolidatedSupplier,
  QuebraOrdemItem,
} from '../types';
import {
  parseCurrencyValue,
  parseFavorecidoField,
  formatBRL,
  isSituacaoAO,
  isSituacaoEnviadaBanco,
  isPPProntaAOValida,
  isTotalizadorRow,
} from './excelParser';
import { SITUACAO_SIGLAS } from './liberacaoParser';

/**
 * Normalizes string for fuzzy / robust matching
 */
function normalizeText(text: string): string {
  if (!text) return '';
  return text
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Checks if a row belongs to the target supplier by CNPJ or Name
 */
function matchesSupplier(
  targetCnpjDigits: string,
  targetNormName: string,
  rowCnpj: string,
  rowFavRaw: string
): boolean {
  const rowCnpjDigits = (rowCnpj || '').replace(/\D/g, '');
  if (targetCnpjDigits && targetCnpjDigits.length >= 11 && rowCnpjDigits.length >= 11) {
    if (targetCnpjDigits === rowCnpjDigits) return true;
  }

  const { name: rowCleanName } = parseFavorecidoField(rowFavRaw);
  const rowNormName = normalizeText(rowCleanName || rowFavRaw);

  if (targetNormName && rowNormName) {
    if (targetNormName === rowNormName) return true;
    if (targetNormName.length > 6 && rowNormName.length > 6) {
      if (targetNormName.includes(rowNormName) || rowNormName.includes(targetNormName)) return true;
    }
  }

  return false;
}

export interface BankAuditRecord {
  pp: string;
  ob: string;
  favorecido: string;
  notaEmpenho: string;
  valor: number;
  sigla: string;
  origemLinha?: number;
}

/**
 * Builds a comprehensive audit summary for a specific company / supplier
 * across all uploaded files, bank status records, and emitted payment records.
 */
export function buildCompanyAuditSummary(
  supplier: { name: string; cnpj: string },
  files: FileData[],
  quebraOrdemItems?: QuebraOrdemItem[],
  bankRecords?: BankAuditRecord[]
): CompanyAuditSummary {
  const targetCnpjDigits = (supplier.cnpj || '').replace(/\D/g, '');
  const targetNormName = normalizeText(supplier.name);

  const allRows: AuditRowItem[] = [];

  let totalConsolidatedValue = 0;
  let ppProntasValue = 0;
  let ppProntasCount = 0;
  let ppProntasAOValue = 0;
  let ppProntasAOCount = 0;
  let enviadasBancoValue = 0;
  let enviadasBancoCount = 0;
  let liqObedeceValue = 0;
  let liqObedeceCount = 0;
  let liqNaoObedeceValue = 0;
  let liqNaoObedeceCount = 0;

  // 1. Scan all regular uploaded files (PPs, Obedece, Não Obedece, Outros)
  files.forEach((file) => {
    const favCol = file.detectedFavorecidoCol;
    const valCol = file.detectedValorCol;
    const stage = file.stage || 'outros';

    file.rows.forEach((row, rowIndex) => {
      const rawFav = row[favCol];
      if (isTotalizadorRow(row, rawFav)) {
        return;
      }

      const { cnpj: rowCnpj, name: rowName } = parseFavorecidoField(rawFav);

      if (!matchesSupplier(targetCnpjDigits, targetNormName, rowCnpj, rawFav)) {
        return;
      }

      const val = parseCurrencyValue(row[valCol]);

      // Extract document identifiers
      let docNumber = '';
      let obNumber = '';
      let neNumber = '';
      let dateVal = '';
      let fonteRecurso = '';
      let situacao = '';

      for (const [key, value] of Object.entries(row)) {
        const lk = key.toLowerCase();
        const strVal = String(value || '').trim();

        if (!obNumber && (lk.includes('ob') || lk.includes('ordem bancaria') || lk.includes('ordem bancária')) && strVal) {
          obNumber = strVal;
        }
        if (!neNumber && (lk.includes('empenho') || lk === 'ne' || lk.includes('nota empenho')) && strVal) {
          neNumber = strVal;
        }
        if (!docNumber && (lk.includes('documento') || lk.includes('número') || lk.includes('numero') || lk.includes('pp') || lk.includes('processo')) && strVal) {
          docNumber = strVal;
        }
        if (!dateVal && (lk.includes('data') || lk.includes('emiss') || lk.includes('venc')) && strVal) {
          dateVal = strVal;
        }
        if (!fonteRecurso && (lk.includes('fonte') || lk.includes('recurso')) && strVal) {
          fonteRecurso = strVal;
        }
        if (!situacao && (lk.includes('situa') || lk.includes('status') || lk.includes('sigla')) && strVal) {
          situacao = strVal;
        }
      }

      totalConsolidatedValue += val;

      const isAO = isSituacaoAO(situacao);
      const enviadaInfo = isSituacaoEnviadaBanco(situacao);

      // Verificação estrita para PPs Prontas: Situação AO e todas as 5 colunas completas
      const isPPPronta = isPPProntaAOValida({
        situacaoRaw: situacao,
        numeroRaw: docNumber,
        obRaw: obNumber,
        favorecidoRaw: rawFav,
        notaEmpenhoRaw: neNumber,
        valor: val,
      });

      let sourceLabel = 'Outros';
      let rowStatus = situacao || sourceLabel;
      let effectiveSourceType: typeof stage | 'status_banco' | 'base_pps' = stage;

      if (stage === 'base_pps' || stage === 'pp_prontas' || isPPPronta) {
        sourceLabel = 'Base de PPs';
        ppProntasValue += val;
        ppProntasCount++;
        if (isAO) {
          ppProntasAOValue += val;
          ppProntasAOCount++;
          rowStatus = 'AO - Associada Ordenadores';
        }
        effectiveSourceType = 'base_pps';
      } else if (enviadaInfo.isEnviada) {
        sourceLabel = 'Enviada a Banco';
        enviadasBancoValue += val;
        enviadasBancoCount++;
        rowStatus = enviadaInfo.label;
        effectiveSourceType = 'status_banco';
      } else if (stage === 'obedece') {
        sourceLabel = 'Liquidado - Obedece';
        liqObedeceValue += val;
        liqObedeceCount++;
        effectiveSourceType = 'obedece';
      } else if (stage === 'nao_obedece') {
        sourceLabel = 'Liquidado - Não Obedece';
        liqNaoObedeceValue += val;
        liqNaoObedeceCount++;
        effectiveSourceType = 'nao_obedece';
      }

      allRows.push({
        id: `aud_${file.id}_${rowIndex}`,
        sourceType: effectiveSourceType,
        sourceLabel,
        sourceFileName: file.fileName,
        documentNumber: docNumber || `Linha ${rowIndex + 1}`,
        date: dateVal || undefined,
        value: val,
        status: rowStatus,
        fonteRecurso: fonteRecurso || undefined,
        rawData: row,
      });
    });
  });

  // 2. Scan Quebra de Ordem / Pagamentos Emitidos (if provided)
  let quebraOrdemCount = 0;
  let quebraOrdemValue = 0;

  if (quebraOrdemItems && quebraOrdemItems.length > 0) {
    quebraOrdemItems.forEach((qo, qoIdx) => {
      const { cnpj: qoCnpj } = parseFavorecidoField(qo.favorecido);
      if (matchesSupplier(targetCnpjDigits, targetNormName, qo.cnpj || qoCnpj, qo.favorecido)) {
        quebraOrdemCount++;
        quebraOrdemValue += qo.valorOB;

        const doc = [qo.pp ? `PP: ${qo.pp}` : '', qo.ob ? `OB: ${qo.ob}` : ''].filter(Boolean).join(' / ') || qo.processo || 'OB Emitida';

        allRows.push({
          id: `aud_qo_${qoIdx}`,
          sourceType: 'pagamentos_emitidos',
          sourceLabel: 'Pagamentos Emitidos (AO)',
          sourceFileName: 'Relatório de Pagamentos Emitidos',
          documentNumber: doc,
          value: qo.valorOB,
          status: 'AO - Quebra de Ordem',
          fonteRecurso: qo.fonteRecurso,
          details: qo.liquidacaoCorrespondente ? `Liq. vinculada: ${formatBRL(qo.liquidacaoCorrespondente.valorLiquidacao)} (Diferença Impostos: ${formatBRL(qo.liquidacaoCorrespondente.diferencaImpostos)})` : 'Com OB Associada dois Ordenadores',
          rawData: qo.rawRow || {},
        });
      }
    });
  }

  // 3. Scan Bank Status Records (if provided)
  let bancoRecordsCount = 0;
  let bancoConfirmedCount = 0;

  if (bankRecords && bankRecords.length > 0) {
    bankRecords.forEach((rec, bIdx) => {
      const { cnpj: recCnpj } = parseFavorecidoField(rec.favorecido);
      if (matchesSupplier(targetCnpjDigits, targetNormName, recCnpj, rec.favorecido)) {
        bancoRecordsCount++;
        const siglaUpper = (rec.sigla || '').trim().toUpperCase();
        const traduzida = SITUACAO_SIGLAS[siglaUpper] || siglaUpper || 'Sem Situação';

        if (['CB', 'PPCB', 'CM', 'LI', 'LD', 'II'].includes(siglaUpper)) {
          bancoConfirmedCount++;
        }

        const doc = [rec.pp ? `PP: ${rec.pp}` : '', rec.ob ? `OB: ${rec.ob}` : ''].filter(Boolean).join(' / ') || rec.notaEmpenho || 'Registro Banco';

        allRows.push({
          id: `aud_banco_${bIdx}`,
          sourceType: 'status_banco',
          sourceLabel: 'Status Banco',
          sourceFileName: 'Base de Dados Bancária',
          documentNumber: doc,
          value: rec.valor || 0,
          status: `${siglaUpper} - ${traduzida}`,
          details: rec.notaEmpenho ? `NE: ${rec.notaEmpenho}` : undefined,
          rawData: rec as any,
        });
      }
    });
  }

  return {
    supplierName: supplier.name,
    cnpj: supplier.cnpj || 'N/I',
    totalConsolidatedValue,
    ppProntasValue,
    ppProntasCount,
    ppProntasAOValue,
    ppProntasAOCount,
    enviadasBancoValue,
    enviadasBancoCount,
    liqObedeceValue,
    liqObedeceCount,
    liqNaoObedeceValue,
    liqNaoObedeceCount,
    bancoRecordsCount,
    bancoConfirmedCount,
    quebraOrdemCount,
    quebraOrdemValue,
    allRows,
  };
}

/**
 * Exports complete company audit report to formatted Excel sheet
 */
export function exportCompanyAuditToExcel(summary: CompanyAuditSummary) {
  const metaSheetData = [
    ['EXTRATO GERAL DE AUDITORIA DE FORNECEDOR', ''],
    ['Governo do Estado do Rio Grande do Norte — Painel Financeiro - FES/SESAP', ''],
    ['', ''],
    ['FAVORECIDO / RAZÃO SOCIAL', summary.supplierName],
    ['CNPJ', summary.cnpj],
    ['DATA DA AUDITORIA', new Date().toLocaleString('pt-BR')],
    ['TOTAL CONSOLIDADO (R$)', summary.totalConsolidatedValue],
    ['TOTAL PPS PRONTAS (AO) (R$)', summary.ppProntasAOValue || summary.ppProntasValue],
    ['TOTAL ENVIADAS A BANCO (PPCB/PPNC) (R$)', summary.enviadasBancoValue || 0],
    ['TOTAL LIQUIDADO - OBEDECE (R$)', summary.liqObedeceValue],
    ['TOTAL LIQUIDADO - NÃO OBEDECE (R$)', summary.liqNaoObedeceValue],
    ['TOTAL QUEBRA DE ORDEM AO (R$)', summary.quebraOrdemValue],
    ['REGISTROS NO BANCO', summary.bancoRecordsCount],
    ['', ''],
    ['DETALHAMENTO DE OCORRÊNCIAS EM TODAS AS PLANILHAS', ''],
  ];

  const rowsData = summary.allRows.map((r, i) => ({
    'ITEM': i + 1,
    'ORIGEM / ETAPA': r.sourceLabel,
    'ARQUIVO DE ORIGEM': r.sourceFileName,
    'DOCUMENTO / PP / OB': r.documentNumber,
    'DATA': r.date || '-',
    'VALOR (R$)': r.value,
    'SITUAÇÃO / STATUS': r.status || '-',
    'FONTE DE RECURSO': r.fonteRecurso || '-',
    'DETALHES / NOTAS': r.details || '-',
  }));

  const workbook = XLSX.utils.book_new();

  const summaryWs = XLSX.utils.aoa_to_sheet(metaSheetData);
  XLSX.utils.book_append_sheet(workbook, summaryWs, 'Resumo');

  const rowsWs = XLSX.utils.json_to_sheet(rowsData);
  XLSX.utils.book_append_sheet(workbook, rowsWs, 'Ocorrencias');

  const safeFileName = `Extrato_Auditoria_${summary.supplierName.substring(0, 30).replace(/[^a-zA-Z0-9]/g, '_')}_${new Date().toISOString().slice(0, 10)}.xlsx`;
  XLSX.writeFile(workbook, safeFileName);
}
