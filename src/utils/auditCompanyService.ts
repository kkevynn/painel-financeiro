import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
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
  isSituacaoConclusivaPagamento,
} from './excelParser';
import { SITUACAO_SIGLAS } from './liberacaoParser';
import {
  parseDateToMonthInfo,
  OrdemBancariaItem,
  loadOrdensBancariasFromStorage,
  isObFileName,
  isValidFonteRecursoCode,
  findFonteRecursoInRow,
} from './ordemBancariaParser';
import {
  extractRealNlFromRow,
  isValidNL,
  resolveCnpj,
  resolveSupplierName,
} from './appKnowledgeDirectory';

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
 * Normalizes document number to clean NL identifier without row numbers or prefix repetition
 */
export function formatCleanNL(doc: string | undefined): string {
  if (!doc) return 'NL Pendente';
  const clean = String(doc).trim();
  if (/^linha\s*\d+$/i.test(clean) || /^#?\d{1,4}$/.test(clean)) return 'NL Pendente';
  if (/\b\d{4}NL\d+\b/i.test(clean)) {
    const match = clean.match(/\b\d{4}NL\d+\b/i);
    return match ? match[0].toUpperCase() : clean.toUpperCase();
  }
  if (/\b\d+NL\d+\b/i.test(clean)) {
    const match = clean.match(/\b\d+NL\d+\b/i);
    return match ? match[0].toUpperCase() : clean.toUpperCase();
  }
  if (/^nl\s*[:#]?\s*([A-Z0-9]+)/i.test(clean)) {
    const match = clean.match(/^nl\s*[:#]?\s*([A-Z0-9]+)/i);
    return match ? `NL ${match[1].toUpperCase()}` : clean.toUpperCase();
  }
  return clean;
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
  bankRecords?: BankAuditRecord[],
  ordensBancariasList?: OrdemBancariaItem[] | null
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

  const seenFileDocKeys = new Set<string>();
  const cleanDocKey = (v?: string) => String(v || '').trim().toUpperCase().replace(/[\s\.\-\/]/g, '');

  // Effective OBs from "OBs - sigef - Principal" / ordensBancariasList
  const loadedOBs = loadOrdensBancariasFromStorage();
  const effectiveOBList: OrdemBancariaItem[] =
    ordensBancariasList && ordensBancariasList.length > 0
      ? ordensBancariasList
      : (loadedOBs && loadedOBs.items) || [];

  // Mapas de cruzamento estritamente para anexar Fonte de Recursos (sem nunca substituir status oficiais)
  const obFonteByOb = new Map<string, string>();
  const obFonteByPp = new Map<string, string>();
  const obFonteBySupplierVal = new Map<string, string>();

  if (effectiveOBList && effectiveOBList.length > 0) {
    effectiveOBList.forEach((obItem) => {
      if (!isValidFonteRecursoCode(obItem.fonteRecurso)) return;
      const validFonte = obItem.fonteRecurso.trim().toUpperCase();
      const obK = cleanDocKey(obItem.ob);
      const ppK = cleanDocKey(obItem.pp);
      if (obK && !obFonteByOb.has(obK)) obFonteByOb.set(obK, validFonte);
      if (ppK && !obFonteByPp.has(ppK)) obFonteByPp.set(ppK, validFonte);

      const match = matchesSupplier(
        targetCnpjDigits,
        targetNormName,
        obItem.favorecidoCnpj,
        obItem.favorecidoRaw || obItem.favorecidoName
      );
      if (match && (obItem.valor || 0) > 0) {
        const valKey = String(Math.round((obItem.valor || 0) * 100));
        if (!obFonteBySupplierVal.has(valKey)) obFonteBySupplierVal.set(valKey, validFonte);
      }
    });
  }

  // Pré-indexa fontes a partir de eventuais planilhas stage === 'ordem_bancaria' presentes em files
  files.forEach((file) => {
    const fileStage = file.stage || 'outros';
    if (fileStage !== 'ordem_bancaria' && !isObFileName(file.fileName || '')) return;
    const obCol = file.detectedObCol;
    const ppCol = file.detectedPpCol || file.detectedNumeroCol;
    const fonteCol = file.detectedFonteCol;
    const valCol = file.detectedValorCol;
    const favCol = file.detectedFavorecidoCol;

    file.rows.forEach((row) => {
      const rawFav = favCol ? row[favCol] : undefined;
      if (isTotalizadorRow(row, rawFav)) return;
      let rawFonte = fonteCol ? String(row[fonteCol] || '').trim().toUpperCase() : '';
      if (!isValidFonteRecursoCode(rawFonte)) {
        rawFonte = findFonteRecursoInRow(Object.values(row));
      }
      if (!isValidFonteRecursoCode(rawFonte)) return;

      let obStr = obCol ? String(row[obCol] || '').trim() : '';
      let ppStr = ppCol ? String(row[ppCol] || '').trim() : '';
      if (!obStr || !ppStr) {
        for (const v of Object.values(row)) {
          if (v === null || v === undefined) continue;
          const s = String(v).trim();
          if (!obStr) {
            const mOb = s.match(/\b\d{4}OB\d+\b/i);
            if (mOb) obStr = mOb[0];
          }
          if (!ppStr) {
            const mPp = s.match(/\b\d{4}PP\d+\b/i);
            if (mPp) ppStr = mPp[0];
          }
        }
      }
      const obK = cleanDocKey(obStr);
      const ppK = cleanDocKey(ppStr);
      if (obK && !obFonteByOb.has(obK)) obFonteByOb.set(obK, rawFonte);
      if (ppK && !obFonteByPp.has(ppK)) obFonteByPp.set(ppK, rawFonte);

      const { cnpj: rowCnpj } = parseFavorecidoField(rawFav);
      if (matchesSupplier(targetCnpjDigits, targetNormName, rowCnpj, rawFav)) {
        const v = valCol ? parseCurrencyValue(row[valCol]) : 0;
        if (v > 0) {
          const valKey = String(Math.round(v * 100));
          if (!obFonteBySupplierVal.has(valKey)) obFonteBySupplierVal.set(valKey, rawFonte);
        }
      }
    });
  });

  // Registra situações oficiais lidas de obedece, nao_obedece e base_pps para nunca serem substituídas por OBs
  const officialSituacaoByOb = new Map<string, string>();
  const officialSituacaoByPp = new Map<string, string>();

  // 1. Scan all regular uploaded files (PPs, Obedece, Não Obedece, Outros) — ignorando ordem_bancaria como fonte de status/valores
  files.forEach((file) => {
    const favCol = file.detectedFavorecidoCol;
    const valCol = file.detectedValorCol;
    const stage = file.stage || 'outros';

    if (stage === 'ordem_bancaria' || isObFileName(file.fileName || '')) {
      return;
    }

    file.rows.forEach((row, rowIndex) => {
      const rawFav = row[favCol];
      if (isTotalizadorRow(row, rawFav)) {
        return;
      }

      const { cnpj: rowCnpj } = parseFavorecidoField(rawFav);

      if (!matchesSupplier(targetCnpjDigits, targetNormName, rowCnpj, rawFav)) {
        return;
      }

      const val = parseCurrencyValue(row[valCol]);

      // Extract document identifiers accurately without row index fallback
      let docNumber = '';
      let obNumber = '';
      let ppNumber = '';
      let neNumber = '';
      let dateVal = '';
      let fonteRecurso = '';
      let situacao = '';

      // 1. First priority: extract real NL pattern from row (strictly rejecting line numbers like 14 or Linha 14)
      const realExtractedNl = extractRealNlFromRow(row, file.detectedNumeroCol || file.detectedPpCol);
      if (realExtractedNl) {
        docNumber = realExtractedNl;
      }

      // 2. Direct check from mapped or detected columns if not already found and not a line number
      if (!docNumber && file.detectedNumeroCol && row[file.detectedNumeroCol]) {
        const valCandidate = String(row[file.detectedNumeroCol] || '').trim();
        if (isValidNL(valCandidate)) {
          docNumber = formatCleanNL(valCandidate);
        } else if (!/^linha\s*\d+$/i.test(valCandidate) && !/^#?\d{1,4}$/.test(valCandidate)) {
          docNumber = valCandidate;
        }
      }
      if (file.detectedPpCol && row[file.detectedPpCol]) {
        const valCandidate = String(row[file.detectedPpCol] || '').trim();
        if (!ppNumber && !/^linha\s*\d+$/i.test(valCandidate) && !/^#?\d{1,4}$/.test(valCandidate)) {
          ppNumber = valCandidate;
        }
        if (!docNumber) {
          if (isValidNL(valCandidate)) {
            docNumber = formatCleanNL(valCandidate);
          } else if (!/^linha\s*\d+$/i.test(valCandidate) && !/^#?\d{1,4}$/.test(valCandidate)) {
            docNumber = valCandidate;
          }
        }
      }

      // Check specific keys
      for (const [key, value] of Object.entries(row)) {
        const lk = key.toLowerCase().trim();
        const strVal = String(value || '').trim();
        if (!strVal) continue;

        if (!obNumber && (lk.includes('ob') || lk.includes('ordem bancaria') || lk.includes('ordem bancária'))) {
          obNumber = strVal;
        }
        if (!neNumber && (lk.includes('empenho') || lk === 'ne' || lk.includes('nota empenho'))) {
          neNumber = strVal;
        }
        if (!docNumber) {
          if (
            lk === 'nl' ||
            lk.includes('numero nl') ||
            lk.includes('número nl') ||
            lk.includes('nº nl') ||
            lk.includes('n° nl') ||
            lk.includes('liquidacao') ||
            lk.includes('liquidação')
          ) {
            if (isValidNL(strVal)) {
              docNumber = formatCleanNL(strVal);
            } else if (!/^linha\s*\d+$/i.test(strVal) && !/^#?\d{1,4}$/.test(strVal)) {
              docNumber = strVal;
            }
          }
        }
        if (!dateVal && (lk.includes('data') || lk.includes('emiss') || lk.includes('venc'))) {
          dateVal = strVal;
        }
        if (!fonteRecurso && (lk.includes('fonte') || lk.includes('recurso'))) {
          fonteRecurso = strVal;
        }
        if (!situacao && (lk.includes('situa') || lk.includes('status') || lk.includes('sigla'))) {
          situacao = strVal;
        }
        if (!ppNumber && (lk === 'pp' || lk.includes('proposta') || lk.includes('preparacao'))) {
          ppNumber = strVal;
        }
      }
      if (!situacao && file.detectedSituacaoCol && row[file.detectedSituacaoCol]) {
        situacao = String(row[file.detectedSituacaoCol] || '').trim();
      }
      if (!obNumber && file.detectedObCol && row[file.detectedObCol]) {
        obNumber = String(row[file.detectedObCol] || '').trim();
      }

      // Regex fallback across all cell values to discover any NL identifier
      if (!docNumber) {
        for (const [, value] of Object.entries(row)) {
          const strVal = String(value || '').trim();
          const matchNL =
            strVal.match(/\b(?:\d{2,4})?NL[\s\-_]*\d+\b/i) ||
            strVal.match(/\b\d+NL\d+\b/i) ||
            strVal.match(/\bNL\s*\d+\b/i);
          if (matchNL) {
            docNumber = formatCleanNL(matchNL[0]);
            break;
          }
        }
      }

      // Fallback to generic document / numero / processo
      if (!docNumber) {
        for (const [key, value] of Object.entries(row)) {
          const lk = key.toLowerCase().trim();
          const strVal = String(value || '').trim();
          if (
            (lk.includes('documento') || lk.includes('número') || lk.includes('numero') || lk.includes('processo')) &&
            strVal &&
            !/^linha\s*\d+$/i.test(strVal) &&
            !/^#?\d{1,4}$/.test(strVal)
          ) {
            docNumber = strVal;
            break;
          }
        }
      }

      // If docNumber is "Linha X" or small row number, strip it
      if (/^linha\s*\d+$/i.test(docNumber) || /^#?\d{1,4}$/.test(docNumber)) {
        docNumber = '';
      }

      // De-duplicate rows across uploaded files so identical rows are not repeated and values not doubled
      const rowKey = `${stage}_${docNumber || 'SN'}_${obNumber || ''}_${neNumber || ''}_${Math.round(val * 100)}`;
      if (seenFileDocKeys.has(rowKey)) {
        return;
      }
      seenFileDocKeys.add(rowKey);

      totalConsolidatedValue += val;

      const isAO = isSituacaoAO(situacao);
      const enviadaInfo = isSituacaoEnviadaBanco(situacao);

      // Verificação estrita para PPs Prontas: Situação AO e todas as colunas
      const isPPPronta = isPPProntaAOValida({
        situacaoRaw: situacao,
        numeroRaw: docNumber,
        obRaw: obNumber,
        favorecidoRaw: rawFav,
        notaEmpenhoRaw: neNumber,
        valor: val,
      });

      let sourceLabel = 'Outros';
      const sitUpper = (situacao || '').trim().toUpperCase();
      let rowStatus = sitUpper || sourceLabel;
      let effectiveSourceType: typeof stage | 'status_banco' | 'base_pps' = stage;

      if (stage === 'base_pps' || stage === 'pp_prontas') {
        sourceLabel = 'Base de PPs';
        ppProntasValue += val;
        ppProntasCount++;
        if (isAO) {
          ppProntasAOValue += val;
          ppProntasAOCount++;
          rowStatus = 'AO - Associada Ordenadores';
        } else if (enviadaInfo.isEnviada) {
          enviadasBancoValue += val;
          enviadasBancoCount++;
          rowStatus = sitUpper && !enviadaInfo.label.toUpperCase().includes(sitUpper)
            ? `${sitUpper} - ${enviadaInfo.label}`
            : enviadaInfo.label;
        } else {
          rowStatus = sitUpper || sourceLabel;
        }
        effectiveSourceType = 'base_pps';
      } else if (stage === 'obedece') {
        sourceLabel = 'Liquidado - Obedece';
        liqObedeceValue += val;
        liqObedeceCount++;
        effectiveSourceType = 'obedece';
        if (isAO) {
          rowStatus = 'AO - Associada Ordenadores';
        } else if (enviadaInfo.isEnviada) {
          rowStatus = sitUpper && !enviadaInfo.label.toUpperCase().includes(sitUpper)
            ? `${sitUpper} - ${enviadaInfo.label}`
            : enviadaInfo.label;
        } else {
          rowStatus = sitUpper || sourceLabel;
        }
      } else if (stage === 'nao_obedece') {
        sourceLabel = 'Liquidado - Não Obedece';
        liqNaoObedeceValue += val;
        liqNaoObedeceCount++;
        effectiveSourceType = 'nao_obedece';
        if (isAO) {
          rowStatus = 'AO - Associada Ordenadores';
        } else if (enviadaInfo.isEnviada) {
          rowStatus = sitUpper && !enviadaInfo.label.toUpperCase().includes(sitUpper)
            ? `${sitUpper} - ${enviadaInfo.label}`
            : enviadaInfo.label;
        } else {
          rowStatus = sitUpper || sourceLabel;
        }
      } else if (enviadaInfo.isEnviada) {
        sourceLabel = 'Enviada a Banco';
        enviadasBancoValue += val;
        enviadasBancoCount++;
        rowStatus = sitUpper && !enviadaInfo.label.toUpperCase().includes(sitUpper)
          ? `${sitUpper} - ${enviadaInfo.label}`
          : enviadaInfo.label;
        effectiveSourceType = 'status_banco';
      } else if (stage === 'natureza_despesa') {
        sourceLabel = 'Natureza da Despesa';
        effectiveSourceType = 'natureza_despesa';
        rowStatus = sitUpper || sourceLabel;
      }

      // Guarda situação oficial por OB / PP para impedir sobrescrita por metadados de OBs
      if (sitUpper) {
        const cleanOb = cleanDocKey(obNumber);
        const cleanPp = cleanDocKey(ppNumber || docNumber);
        if (cleanOb && !officialSituacaoByOb.has(cleanOb)) officialSituacaoByOb.set(cleanOb, sitUpper);
        if (cleanPp && !officialSituacaoByPp.has(cleanPp)) officialSituacaoByPp.set(cleanPp, sitUpper);
      }

      // Cruzamento pela OB / PP (ou fornecedor+valor) na planilha de OBs apenas para anexar Fonte de Recursos
      const obLookupKey = cleanDocKey(obNumber);
      const ppLookupKey = cleanDocKey(ppNumber);
      const docLookupKey = cleanDocKey(docNumber);
      const valLookupKey = val > 0 ? String(Math.round(val * 100)) : '';
      const matchedObFonte =
        (obLookupKey && obFonteByOb.get(obLookupKey)) ||
        (ppLookupKey && obFonteByPp.get(ppLookupKey)) ||
        (docLookupKey && obFonteByPp.get(docLookupKey)) ||
        (valLookupKey && obFonteBySupplierVal.get(valLookupKey)) ||
        '';
      const resolvedFonte =
        (isValidFonteRecursoCode(matchedObFonte) ? matchedObFonte : '') ||
        (isValidFonteRecursoCode(fonteRecurso) ? fonteRecurso.trim().toUpperCase() : '') ||
        fonteRecurso ||
        '0.6.00';

      allRows.push({
        id: `aud_${file.id}_${rowIndex}`,
        sourceType: effectiveSourceType,
        sourceLabel,
        sourceFileName: file.fileName,
        documentNumber: docNumber || 'NL Pendente',
        date: dateVal || undefined,
        value: val,
        status: rowStatus,
        fonteRecurso: resolvedFonte,
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

        const doc =
          [qo.pp ? `PP: ${qo.pp}` : '', qo.ob ? `OB: ${qo.ob}` : ''].filter(Boolean).join(' / ') ||
          qo.processo ||
          'OB Emitida';

        allRows.push({
          id: `aud_qo_${qoIdx}`,
          sourceType: 'pagamentos_emitidos',
          sourceLabel: 'Pagamentos Emitidos (AO)',
          sourceFileName: 'Relatório de Pagamentos Emitidos',
          documentNumber: doc,
          value: qo.valorOB,
          status: 'AO - Quebra de Ordem',
          fonteRecurso: qo.fonteRecurso || '0.6.00',
          details: qo.liquidacaoCorrespondente
            ? `Liq. vinculada: ${formatBRL(qo.liquidacaoCorrespondente.valorLiquidacao)} (Diferença Impostos: ${formatBRL(qo.liquidacaoCorrespondente.diferencaImpostos)})`
            : 'Com OB Associada dois Ordenadores',
          rawData: qo.rawRow || {},
        });
      }
    });
  }

  // 3. Scan Bank Status Records & OrdemBancariaItem from "OBs - sigef - Principal"
  let bancoRecordsCount = 0;
  let bancoConfirmedCount = 0;
  let bancoConfirmedValue = 0;
  let bancoPendingAOCount = 0;
  let bancoPendingAOValue = 0;

  const seenBankDocKeys = new Set<string>();

  if (effectiveOBList && effectiveOBList.length > 0) {
    effectiveOBList.forEach((obItem, obIdx) => {
      const match = matchesSupplier(
        targetCnpjDigits,
        targetNormName,
        obItem.favorecidoCnpj,
        obItem.favorecidoRaw || obItem.favorecidoName
      );
      if (!match) return;

      const dedupeKey = `${obItem.ob || ''}_${obItem.pp || ''}_${obItem.valor || 0}`;
      if (seenBankDocKeys.has(dedupeKey)) return;
      seenBankDocKeys.add(dedupeKey);

      bancoRecordsCount++;

      // Prioriza sempre a situação oficial registrada em Obedece / Não Obedece / Base de PPs; nunca permite sobrescrita pela planilha de OBs.
      const obK = cleanDocKey(obItem.ob);
      const ppK = cleanDocKey(obItem.pp);
      const officialSit =
        (obK && officialSituacaoByOb.get(obK)) ||
        (ppK && officialSituacaoByPp.get(ppK)) ||
        '';
      const effectiveSit = officialSit || (obItem.situacao || '').trim().toUpperCase();

      const isAO = isSituacaoAO(effectiveSit);
      const isConfirmed =
        isSituacaoConclusivaPagamento(effectiveSit) ||
        ['CB', 'PPCB', 'CM', 'LI', 'LD', 'II', 'PAGA', 'PAGO'].includes(effectiveSit);

      if (isConfirmed) {
        bancoConfirmedCount++;
        bancoConfirmedValue += obItem.valor || 0;
      } else if (isAO) {
        bancoPendingAOCount++;
        bancoPendingAOValue += obItem.valor || 0;
      }

      const siglaUpper = effectiveSit;
      const traduzida = siglaUpper
        ? (SITUACAO_SIGLAS[siglaUpper] || siglaUpper)
        : 'Fonte de Recursos (OB)';

      const docDisplay =
        [obItem.ob ? `OB: ${obItem.ob}` : '', obItem.pp ? `PP: ${obItem.pp}` : ''].filter(Boolean).join(' / ') ||
        'OB Bancária';

      const validObFonte = isValidFonteRecursoCode(obItem.fonteRecurso)
        ? obItem.fonteRecurso.trim().toUpperCase()
        : (obItem.fonteRecurso || '0.6.00');

      allRows.push({
        id: `aud_ob_${obItem.id || obIdx}`,
        sourceType: 'status_banco',
        sourceLabel: isAO ? 'Ordem Bancária (AO)' : isConfirmed ? 'Ordem Bancária (Paga)' : 'Ordem Bancária (Fonte)',
        sourceFileName: 'OBs - sigef - Principal (Fonte de Recursos)',
        documentNumber: docDisplay,
        date: obItem.dataReferencia || obItem.dataIso || undefined,
        value: obItem.valor || 0,
        status: siglaUpper ? `${siglaUpper} - ${traduzida}` : 'OB - Fonte de Recursos',
        fonteRecurso: validObFonte,
        details: [
          obItem.ob ? `OB: ${obItem.ob}` : '',
          obItem.pp ? `PP: ${obItem.pp}` : '',
          obItem.mesAnoLabel ? `Mês Ref: ${obItem.mesAnoLabel}` : '',
        ].filter(Boolean).join(' • '),
        rawData: { ...(obItem as any), situacao: siglaUpper },
      });
    });
  }

  // Also include bankRecords if passed and not duplicate
  if (bankRecords && bankRecords.length > 0) {
    bankRecords.forEach((rec, bIdx) => {
      const { cnpj: recCnpj } = parseFavorecidoField(rec.favorecido);
      if (matchesSupplier(targetCnpjDigits, targetNormName, recCnpj, rec.favorecido)) {
        const dedupeKey = `${rec.ob || ''}_${rec.pp || ''}_${rec.valor || 0}`;
        if (seenBankDocKeys.has(dedupeKey)) return;
        seenBankDocKeys.add(dedupeKey);

        bancoRecordsCount++;
        const siglaUpper = (rec.sigla || '').trim().toUpperCase();
        const traduzida = SITUACAO_SIGLAS[siglaUpper] || siglaUpper || 'Sem Situação';
        const isAO = isSituacaoAO(siglaUpper);
        const isConfirmed = ['CB', 'PPCB', 'CM', 'LI', 'LD', 'II', 'PAGA', 'PAGO'].includes(siglaUpper);

        if (isConfirmed) {
          bancoConfirmedCount++;
          bancoConfirmedValue += rec.valor || 0;
        } else if (isAO) {
          bancoPendingAOCount++;
          bancoPendingAOValue += rec.valor || 0;
        }

        const doc =
          [rec.pp ? `PP: ${rec.pp}` : '', rec.ob ? `OB: ${rec.ob}` : ''].filter(Boolean).join(' / ') ||
          rec.notaEmpenho ||
          'Registro Banco';

        allRows.push({
          id: `aud_banco_${bIdx}`,
          sourceType: 'status_banco',
          sourceLabel: isAO ? 'Status Banco (AO)' : isConfirmed ? 'Status Banco (Pago)' : 'Status Banco',
          sourceFileName: 'Base de Dados Bancária',
          documentNumber: doc,
          value: rec.valor || 0,
          status: `${siglaUpper} - ${traduzida}`,
          fonteRecurso: '0.6.00',
          details: rec.notaEmpenho ? `NE: ${rec.notaEmpenho}` : undefined,
          rawData: rec as any,
        });
      }
    });
  }

  const resolvedCnpj =
    supplier.cnpj && supplier.cnpj !== 'N/I' && supplier.cnpj !== '-'
      ? supplier.cnpj
      : resolveCnpj(supplier.name, undefined, supplier.cnpj);

  return {
    supplierName: supplier.name,
    cnpj: resolvedCnpj || supplier.cnpj || 'N/I',
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
    bancoConfirmedValue,
    bancoPendingAOCount,
    bancoPendingAOValue,
    quebraOrdemCount,
    quebraOrdemValue,
    allRows,
  };
}

export interface CompanyPaymentsAuditBreakdown {
  monthlyList: Array<{ mesAnoKey: string; mesAnoLabel: string; count: number; total: number }>;
  confirmedBankOrders: Array<{
    id: string;
    ob: string;
    pp: string;
    data: string;
    valor: number;
    situacao: string;
    fonteRecurso: string;
  }>;
  pendingAoBankOrders: Array<{
    id: string;
    ob: string;
    pp: string;
    data: string;
    valor: number;
    situacao: string;
    fonteRecurso: string;
  }>;
  totalPago: number;
  totalPendenteAO: number;
  totalLiquidado: number;
}

/**
 * Groups paid bank orders by month and separates confirmed payments from pending AO bank orders
 */
export function computeMonthlyPaymentsBreakdown(summary: CompanyAuditSummary): CompanyPaymentsAuditBreakdown {
  const map = new Map<string, { mesAnoKey: string; mesAnoLabel: string; count: number; total: number }>();
  let totalPago = 0;
  let totalPendenteAO = 0;

  const confirmedBankOrders: Array<{
    id: string;
    ob: string;
    pp: string;
    data: string;
    valor: number;
    situacao: string;
    fonteRecurso: string;
  }> = [];

  const pendingAoBankOrders: Array<{
    id: string;
    ob: string;
    pp: string;
    data: string;
    valor: number;
    situacao: string;
    fonteRecurso: string;
  }> = [];

  summary.allRows.forEach((r) => {
    const statusUpper = (r.status || '').toUpperCase();
    const rawData = r.rawData || {};

    const isConfirmed =
      statusUpper.includes('CB') ||
      statusUpper.includes('PPCB') ||
      statusUpper.includes('PAGA') ||
      statusUpper.includes('PAGO') ||
      statusUpper.includes('LI') ||
      statusUpper.includes('LD') ||
      statusUpper.includes('CM') ||
      statusUpper.includes('ENVIADA A BANCO') ||
      isSituacaoConclusivaPagamento(r.status || '') ||
      isSituacaoConclusivaPagamento(rawData.situacao || '');

    const isAO =
      !isConfirmed &&
      (isSituacaoAO(r.status || '') ||
        isSituacaoAO(rawData.situacao || '') ||
        statusUpper.includes('AO') ||
        r.sourceType === 'pagamentos_emitidos' ||
        r.sourceType === 'base_pps');

    // Extract clean OB and PP numbers
    let obNum = rawData.ob || '';
    let ppNum = rawData.pp || '';
    if (!obNum) {
      const obMatch =
        r.documentNumber.match(/\b\d{4}OB\d+\b/i) ||
        r.documentNumber.match(/\bOB\s*[:#]?\s*([A-Z0-9]+)\b/i);
      if (obMatch) obNum = obMatch[1] || obMatch[0];
    }
    if (!ppNum) {
      const ppMatch =
        r.documentNumber.match(/\b\d{4}PP\d+\b/i) ||
        r.documentNumber.match(/\bPP\s*[:#]?\s*([A-Z0-9]+)\b/i);
      if (ppMatch) ppNum = ppMatch[1] || ppMatch[0];
    }

    const rowDate =
      r.date ||
      rawData.dataReferencia ||
      rawData.dataIso ||
      rawData.data ||
      rawData.dataPagamento ||
      '-';
    const rowFonte = r.fonteRecurso || rawData.fonteRecurso || '0.6.00';

    if (isConfirmed && r.value > 0) {
      totalPago += r.value;
      confirmedBankOrders.push({
        id: r.id,
        ob: obNum || r.documentNumber || 'OB Confirmada',
        pp: ppNum || '-',
        data: rowDate,
        valor: r.value,
        situacao: r.status || 'Confirmada no Banco',
        fonteRecurso: rowFonte,
      });

      // Monthly aggregation for confirmed payments
      const dInfo = parseDateToMonthInfo(rowDate);
      const key = dInfo.mesAnoKey || 'sem_data';
      const label = dInfo.mesAnoLabel || 'Pagamentos sem data informada';

      const existing = map.get(key) || {
        mesAnoKey: key,
        mesAnoLabel: label,
        count: 0,
        total: 0,
      };
      existing.count += 1;
      existing.total += r.value;
      map.set(key, existing);
    } else if (isAO && r.value > 0) {
      totalPendenteAO += r.value;
      pendingAoBankOrders.push({
        id: r.id,
        ob: obNum || (r.documentNumber.includes('OB') ? r.documentNumber : '-'),
        pp: ppNum || (r.documentNumber.includes('PP') ? r.documentNumber : r.documentNumber),
        data: rowDate,
        valor: r.value,
        situacao: r.status || 'AO - Associada a 2 Ordenadores',
        fonteRecurso: rowFonte,
      });
    }
  });

  const monthlyList = Array.from(map.values()).sort((a, b) => {
    if (a.mesAnoKey === 'sem_data') return 1;
    if (b.mesAnoKey === 'sem_data') return -1;
    return a.mesAnoKey.localeCompare(b.mesAnoKey);
  });

  const totalLiquidado = (summary.liqObedeceValue || 0) + (summary.liqNaoObedeceValue || 0);

  return {
    monthlyList,
    confirmedBankOrders,
    pendingAoBankOrders,
    totalPago,
    totalPendenteAO,
    totalLiquidado,
  };
}

/**
 * Exports complete company audit report to formatted Excel sheet with colors and merged headers
 */
export function exportCompanyAuditToExcel(summary: CompanyAuditSummary) {
  const { monthlyList, confirmedBankOrders, pendingAoBankOrders, totalPago, totalPendenteAO, totalLiquidado } =
    computeMonthlyPaymentsBreakdown(summary);

  const metaSheetData: any[][] = [
    ['GOVERNO DO ESTADO DO RIO GRANDE DO NORTE — SESAP / FES'],
    ['EXTRATO FINANCEIRO E AUDITORIA DE FORNECEDOR'],
    ['Emitido em: ' + new Date().toLocaleString('pt-BR')],
    [],
    ['DADOS DO FORNECEDOR / CREDOR', ''],
    ['Razão Social / Favorecido:', summary.supplierName],
    ['CNPJ / CPF:', summary.cnpj],
    [],
    ['RESUMO GERAL', 'VALOR (R$)', 'QTD REGISTROS'],
    ['Total', summary.totalConsolidatedValue, summary.allRows.length],
    ['Total Liquidado (Obedece)', summary.liqObedeceValue, summary.liqObedeceCount],
    ['Total Liquidado (Não Obedece)', summary.liqNaoObedeceValue, summary.liqNaoObedeceCount],
    ['Total Geral Liquidado (Obedece + Não Obedece)', totalLiquidado, summary.liqObedeceCount + summary.liqNaoObedeceCount],
    ['Total Ordens Bancárias e PPs Pendentes (AO)', totalPendenteAO, pendingAoBankOrders.length],
    ['Total Pago Confirmado no Banco', totalPago, confirmedBankOrders.length],
    [],
    ['DIVISÃO MENSAL DOS PAGAMENTOS EFETUADOS', 'QTD PAGAMENTOS', 'TOTAL PAGO (R$)'],
  ];

  if (monthlyList.length > 0) {
    monthlyList.forEach((m) => {
      metaSheetData.push([m.mesAnoLabel, m.count, m.total]);
    });
    metaSheetData.push(['TOTAL PAGO ACUMULADO', monthlyList.reduce((acc, it) => acc + it.count, 0), totalPago]);
  } else {
    metaSheetData.push(['Nenhum pagamento bancário registrado com data', 0, 0]);
  }

  // Confirmed Bank Orders sheet data
  const obsConfirmadasRows = confirmedBankOrders.map((ob, i) => ({
    'ITEM': i + 1,
    'DATA': ob.data,
    'ORDEM BANCÁRIA (OB)': ob.ob,
    'PREPARAÇÃO (PP)': ob.pp,
    'VALOR PAGO (R$)': ob.valor,
    'SITUAÇÃO': ob.situacao,
    'FONTE DE RECURSO': ob.fonteRecurso,
  }));

  // Pending AO Bank Orders sheet data
  const obsAORows = pendingAoBankOrders.map((ob, i) => ({
    'ITEM': i + 1,
    'DATA': ob.data,
    'ORDEM BANCÁRIA / DOC': ob.ob,
    'PREPARAÇÃO (PP)': ob.pp,
    'VALOR PENDENTE (R$)': ob.valor,
    'SITUAÇÃO': ob.situacao,
    'FONTE DE RECURSO': ob.fonteRecurso,
  }));

  // Liquidations sheet data
  const liquidacoesRows = summary.allRows
    .filter((r) => r.sourceType === 'obedece' || r.sourceType === 'nao_obedece')
    .map((r, i) => ({
      'ITEM': i + 1,
      'FILA / LIQUIDAÇÃO': r.sourceLabel,
      'NÚMERO DA LIQUIDAÇÃO (NL)': formatCleanNL(r.documentNumber),
      'VALOR BRUTO (R$)': r.value,
      'SITUAÇÃO': r.status || 'Liquidado',
      'FONTE DE RECURSO': r.fonteRecurso || '0.6.00',
      'ARQUIVO DE ORIGEM': r.sourceFileName,
    }));

  // All records sheet data
  const rowsData = summary.allRows.map((r, i) => ({
    'ITEM': i + 1,
    'ORIGEM / LIQUIDAÇÃO': r.sourceLabel,
    'NÚMERO DO DOCUMENTO / NL': formatCleanNL(r.documentNumber),
    'DATA': r.date || '-',
    'VALOR (R$)': r.value,
    'SITUAÇÃO / STATUS': r.status || '-',
    'FONTE DE RECURSO': r.fonteRecurso || '-',
    'ARQUIVO DE ORIGEM': r.sourceFileName,
  }));

  const workbook = XLSX.utils.book_new();

  // Create Resumo Sheet
  const summaryWs = XLSX.utils.aoa_to_sheet(metaSheetData);
  summaryWs['!cols'] = [{ wch: 45 }, { wch: 25 }, { wch: 20 }];
  summaryWs['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 2 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 2 } },
    { s: { r: 2, c: 0 }, e: { r: 2, c: 2 } },
  ];
  XLSX.utils.book_append_sheet(workbook, summaryWs, 'Resumo Financeiro');

  // Create Confirmed OBs Sheet
  if (obsConfirmadasRows.length > 0) {
    const obWs = XLSX.utils.json_to_sheet(obsConfirmadasRows);
    obWs['!cols'] = [
      { wch: 6 },
      { wch: 15 },
      { wch: 26 },
      { wch: 24 },
      { wch: 20 },
      { wch: 28 },
      { wch: 20 },
    ];
    XLSX.utils.book_append_sheet(workbook, obWs, 'OBs Pagas Confirmadas');
  }

  // Create Pending AO OBs Sheet
  if (obsAORows.length > 0) {
    const aoWs = XLSX.utils.json_to_sheet(obsAORows);
    aoWs['!cols'] = [
      { wch: 6 },
      { wch: 15 },
      { wch: 26 },
      { wch: 24 },
      { wch: 20 },
      { wch: 28 },
      { wch: 20 },
    ];
    XLSX.utils.book_append_sheet(workbook, aoWs, 'OBs Pendentes (AO)');
  }

  // Create Liquidations Sheet
  if (liquidacoesRows.length > 0) {
    const liqWs = XLSX.utils.json_to_sheet(liquidacoesRows);
    liqWs['!cols'] = [
      { wch: 6 },
      { wch: 26 },
      { wch: 28 },
      { wch: 20 },
      { wch: 18 },
      { wch: 20 },
      { wch: 30 },
    ];
    XLSX.utils.book_append_sheet(workbook, liqWs, 'Liquidações');
  }

  // Create All Occurrences Sheet
  const rowsWs = XLSX.utils.json_to_sheet(rowsData);
  rowsWs['!cols'] = [
    { wch: 6 },
    { wch: 26 },
    { wch: 28 },
    { wch: 15 },
    { wch: 20 },
    { wch: 22 },
    { wch: 20 },
    { wch: 30 },
  ];
  XLSX.utils.book_append_sheet(workbook, rowsWs, 'Extrato Completo');

  const safeFileName = `Extrato_${summary.supplierName.substring(0, 30).replace(/[^a-zA-Z0-9]/g, '_')}_${new Date().toISOString().slice(0, 10)}.xlsx`;
  XLSX.writeFile(workbook, safeFileName);
}

/**
 * Generates an official, publication-ready PDF report for the Supplier statement
 * showing total liquidated, confirmed paid bank orders, pending AO orders, and monthly breakdown.
 */
export function exportCompanyAuditToPdf(summary: CompanyAuditSummary) {
  const doc = new jsPDF({
    orientation: 'portrait',
    unit: 'mm',
    format: 'a4',
  });

  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const marginX = 14;

  const { monthlyList, confirmedBankOrders, pendingAoBankOrders, totalPago, totalPendenteAO, totalLiquidado } =
    computeMonthlyPaymentsBreakdown(summary);

  // 1. Institutional Header
  doc.setFillColor(6, 29, 21); // #061d15
  doc.rect(0, 0, pageWidth, 28, 'F');

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(11);
  doc.setTextColor(255, 255, 255);
  doc.text('GOVERNO DO ESTADO DO RIO GRANDE DO NORTE', marginX, 9);

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(8.5);
  doc.setTextColor(167, 243, 208); // Emerald 200
  doc.text('SECRETARIA DE ESTADO DA SAÚDE PÚBLICA — SESAP / FES', marginX, 15);

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(10);
  doc.setTextColor(255, 255, 255);
  doc.text('EXTRATO DE LIQUIDAÇÕES E PAGAMENTOS POR FORNECEDOR', marginX, 22);

  const emissaoStr = new Date().toLocaleDateString('pt-BR') + ' às ' + new Date().toLocaleTimeString('pt-BR');
  doc.setFontSize(7.5);
  doc.setTextColor(209, 250, 229);
  doc.text(`Emissão: ${emissaoStr}`, pageWidth - marginX, 22, { align: 'right' });

  let curY = 34;

  // 2. Creditor Details Block - "TOTAL:" strictly without "CONSOLIDADO"
  doc.setFillColor(248, 250, 252); // Slate 50
  doc.setDrawColor(226, 232, 240); // Slate 200
  doc.roundedRect(marginX, curY, pageWidth - marginX * 2, 20, 2, 2, 'FD');

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(8.5);
  doc.setTextColor(15, 23, 42);
  doc.text('CREDOR / FAVORECIDO:', marginX + 4, curY + 6);
  doc.setFont('helvetica', 'normal');
  doc.text(summary.supplierName, marginX + 44, curY + 6);

  doc.setFont('helvetica', 'bold');
  doc.text('CNPJ / CPF:', marginX + 4, curY + 13);
  doc.setFont('helvetica', 'normal');
  doc.text(summary.cnpj, marginX + 24, curY + 13);

  doc.setFont('helvetica', 'bold');
  doc.text('TOTAL:', marginX + 115, curY + 13);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(6, 78, 59); // Emerald 900
  doc.text(formatBRL(summary.totalConsolidatedValue), marginX + 130, curY + 13);

  curY += 25;

  // 3. Consolidated Totals KPI Boxes (4 boxes side-by-side)
  const kpiBoxWidth = (pageWidth - marginX * 2 - 9) / 4;

  // Box 1: Liquidado Obedece
  doc.setFillColor(239, 246, 255); // Blue 50
  doc.setDrawColor(191, 219, 254);
  doc.roundedRect(marginX, curY, kpiBoxWidth, 16, 2, 2, 'FD');
  doc.setFontSize(6.5);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(30, 64, 175); // Blue 800
  doc.text('LIQUIDADO - OBEDECE', marginX + 2.5, curY + 5);
  doc.setFontSize(8.5);
  doc.text(formatBRL(summary.liqObedeceValue), marginX + 2.5, curY + 12);

  // Box 2: Liquidado Não Obedece
  doc.setFillColor(254, 243, 199); // Amber 50
  doc.setDrawColor(253, 230, 138);
  doc.roundedRect(marginX + kpiBoxWidth + 3, curY, kpiBoxWidth, 16, 2, 2, 'FD');
  doc.setFontSize(6.5);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(146, 64, 14); // Amber 800
  doc.text('LIQ. NÃO OBEDECE', marginX + kpiBoxWidth + 5.5, curY + 5);
  doc.setFontSize(8.5);
  doc.text(formatBRL(summary.liqNaoObedeceValue), marginX + kpiBoxWidth + 5.5, curY + 12);

  // Box 3: Pendente (AO / PPs)
  doc.setFillColor(255, 247, 237); // Orange 50
  doc.setDrawColor(254, 215, 170);
  doc.roundedRect(marginX + (kpiBoxWidth + 3) * 2, curY, kpiBoxWidth, 16, 2, 2, 'FD');
  doc.setFontSize(6.5);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(194, 65, 12); // Orange 700
  doc.text('PENDENTE (AO / PPS)', marginX + (kpiBoxWidth + 3) * 2 + 2.5, curY + 5);
  doc.setFontSize(8.5);
  doc.text(formatBRL(totalPendenteAO), marginX + (kpiBoxWidth + 3) * 2 + 2.5, curY + 12);

  // Box 4: Total Pago Confirmado no Banco
  doc.setFillColor(236, 253, 245); // Emerald 50
  doc.setDrawColor(167, 243, 208);
  doc.roundedRect(marginX + (kpiBoxWidth + 3) * 3, curY, kpiBoxWidth, 16, 2, 2, 'FD');
  doc.setFontSize(6.5);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(6, 95, 70); // Emerald 800
  doc.text('PAGO NO BANCO', marginX + (kpiBoxWidth + 3) * 3 + 2.5, curY + 5);
  doc.setFontSize(8.5);
  doc.text(formatBRL(totalPago), marginX + (kpiBoxWidth + 3) * 3 + 2.5, curY + 12);

  curY += 21;

  // Helper for checking page breaks cleanly
  function ensurePageSpace(neededHeight: number) {
    if (curY + neededHeight > pageHeight - 20) {
      doc.addPage();
      curY = 20;
    }
  }

  // 4. Section 1: ORDENS BANCÁRIAS CONFIRMADAS NO BANCO (PAGO)
  ensurePageSpace(35);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(8.5);
  doc.setTextColor(6, 95, 70);
  doc.text('1. ORDENS BANCÁRIAS CONFIRMADAS NO BANCO (VALORES PAGOS)', marginX, curY);
  curY += 3.5;

  const confirmedBody: any[] = confirmedBankOrders.map((ob, idx) => [
    idx + 1,
    ob.data || '-',
    ob.ob || 'OB Confirmada',
    ob.pp || '-',
    formatBRL(ob.valor),
    ob.situacao || 'Confirmada no Banco',
    ob.fonteRecurso || '0.6.00',
  ]);

  if (confirmedBody.length === 0) {
    confirmedBody.push([
      { content: 'Nenhuma ordem bancária confirmada registrada para este fornecedor', colSpan: 7, styles: { halign: 'center', textColor: [100, 116, 139] } }
    ]);
  } else {
    // Totalizer row with wide colSpan to eliminate word breaks (no "TO TAL")
    confirmedBody.push([
      { content: 'TOTAL PAGO NO BANCO:', colSpan: 4, styles: { halign: 'right', fontStyle: 'bold', fillColor: [236, 253, 245], textColor: [6, 95, 70] } },
      { content: formatBRL(totalPago), styles: { halign: 'right', fontStyle: 'bold', fillColor: [236, 253, 245], textColor: [6, 95, 70] } },
      { content: `${confirmedBankOrders.length} OB(s)`, colSpan: 2, styles: { halign: 'center', fontStyle: 'bold', fillColor: [236, 253, 245], textColor: [6, 95, 70] } },
    ]);
  }

  autoTable(doc, {
    startY: curY,
    head: [['#', 'DATA', 'ORDEM BANCÁRIA (OB)', 'PREPARAÇÃO (PP)', 'VALOR PAGO (R$)', 'SITUAÇÃO', 'FONTE']],
    body: confirmedBody,
    theme: 'grid',
    headStyles: {
      fillColor: [6, 78, 59], // Emerald 900
      textColor: [255, 255, 255],
      fontSize: 7,
      fontStyle: 'bold',
      halign: 'center',
    },
    bodyStyles: {
      fontSize: 7,
      textColor: [30, 41, 59],
    },
    columnStyles: {
      0: { cellWidth: 8, halign: 'center' },
      1: { cellWidth: 22, halign: 'center' },
      2: { cellWidth: 38, halign: 'left', fontStyle: 'bold' },
      3: { cellWidth: 34, halign: 'left' },
      4: { cellWidth: 34, halign: 'right', fontStyle: 'bold' },
      5: { cellWidth: 26, halign: 'center' },
      6: { cellWidth: 20, halign: 'center' },
    },
    margin: { left: marginX, right: marginX },
  });

  curY = (doc as any).lastAutoTable.finalY + 7;

  // 5. Section 2: ORDENS BANCÁRIAS E PPS EM SITUAÇÃO AO (PENDENTE)
  ensurePageSpace(35);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(8.5);
  doc.setTextColor(194, 65, 12);
  doc.text('2. ORDENS BANCÁRIAS E PPS EM SITUAÇÃO AO (PENDENTE)', marginX, curY);
  curY += 3.5;

  const pendingBody: any[] = pendingAoBankOrders.map((ob, idx) => [
    idx + 1,
    ob.data || '-',
    ob.ob || '-',
    ob.pp || '-',
    formatBRL(ob.valor),
    ob.situacao || 'AO - Associada Ordenadores',
    ob.fonteRecurso || '0.6.00',
  ]);

  if (pendingBody.length === 0) {
    pendingBody.push([
      { content: 'Nenhuma ordem bancária ou PP pendente em situação AO', colSpan: 7, styles: { halign: 'center', textColor: [100, 116, 139] } }
    ]);
  } else {
    // Totalizer row with wide colSpan to eliminate word breaks
    pendingBody.push([
      { content: 'TOTAL PENDENTE (AO):', colSpan: 4, styles: { halign: 'right', fontStyle: 'bold', fillColor: [254, 243, 199], textColor: [146, 64, 14] } },
      { content: formatBRL(totalPendenteAO), styles: { halign: 'right', fontStyle: 'bold', fillColor: [254, 243, 199], textColor: [146, 64, 14] } },
      { content: `${pendingAoBankOrders.length} item(s)`, colSpan: 2, styles: { halign: 'center', fontStyle: 'bold', fillColor: [254, 243, 199], textColor: [146, 64, 14] } },
    ]);
  }

  autoTable(doc, {
    startY: curY,
    head: [['#', 'DATA', 'ORDEM BANCÁRIA / DOC', 'PREPARAÇÃO (PP)', 'VALOR PENDENTE (R$)', 'SITUAÇÃO', 'FONTE']],
    body: pendingBody,
    theme: 'grid',
    headStyles: {
      fillColor: [180, 83, 9], // Amber 700
      textColor: [255, 255, 255],
      fontSize: 7,
      fontStyle: 'bold',
      halign: 'center',
    },
    bodyStyles: {
      fontSize: 7,
      textColor: [30, 41, 59],
    },
    columnStyles: {
      0: { cellWidth: 8, halign: 'center' },
      1: { cellWidth: 22, halign: 'center' },
      2: { cellWidth: 38, halign: 'left', fontStyle: 'bold' },
      3: { cellWidth: 34, halign: 'left' },
      4: { cellWidth: 34, halign: 'right', fontStyle: 'bold' },
      5: { cellWidth: 26, halign: 'center' },
      6: { cellWidth: 20, halign: 'center' },
    },
    margin: { left: marginX, right: marginX },
  });

  curY = (doc as any).lastAutoTable.finalY + 7;

  // 6. Section 3: DIVISÃO MENSAL DOS PAGAMENTOS EFETUADOS
  ensurePageSpace(35);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(8.5);
  doc.setTextColor(15, 23, 42);
  doc.text('3. DIVISÃO MENSAL DOS PAGAMENTOS EFETUADOS (ORDENS BANCÁRIAS)', marginX, curY);
  curY += 3.5;

  const monthlyTableBody: any[] = monthlyList.map((m) => [
    m.mesAnoLabel,
    String(m.count) + ' pagto(s)',
    formatBRL(m.total),
    totalPago > 0 ? ((m.total / totalPago) * 100).toFixed(1).replace('.', ',') + '%' : '0%',
  ]);

  if (monthlyTableBody.length === 0) {
    monthlyTableBody.push([
      { content: 'Nenhum pagamento registrado com data', colSpan: 4, styles: { halign: 'center', textColor: [100, 116, 139] } }
    ]);
  } else {
    // Totalizer row with colSpan 2 so label never wraps
    monthlyTableBody.push([
      { content: 'TOTAL GERAL PAGO NO BANCO:', colSpan: 2, styles: { halign: 'right', fontStyle: 'bold', fillColor: [236, 253, 245], textColor: [6, 95, 70] } },
      { content: formatBRL(totalPago), styles: { halign: 'right', fontStyle: 'bold', fillColor: [236, 253, 245], textColor: [6, 95, 70] } },
      { content: '100.0%', styles: { halign: 'center', fontStyle: 'bold', fillColor: [236, 253, 245], textColor: [6, 95, 70] } },
    ]);
  }

  autoTable(doc, {
    startY: curY,
    head: [['MÊS / ANO DE PAGAMENTO', 'QUANTIDADE', 'VALOR PAGO (R$)', '% DO PAGO']],
    body: monthlyTableBody,
    theme: 'grid',
    headStyles: {
      fillColor: [6, 29, 21],
      textColor: [255, 255, 255],
      fontSize: 7.5,
      fontStyle: 'bold',
      halign: 'center',
    },
    bodyStyles: {
      fontSize: 7.5,
      textColor: [30, 41, 59],
    },
    columnStyles: {
      0: { cellWidth: 70, halign: 'left' },
      1: { cellWidth: 35, halign: 'center' },
      2: { cellWidth: 45, halign: 'right', fontStyle: 'bold' },
      3: { cellWidth: 32, halign: 'center' },
    },
    margin: { left: marginX, right: marginX },
  });

  curY = (doc as any).lastAutoTable.finalY + 7;

  // 7. Section 4: DEMONSTRATIVO DE LIQUIDAÇÕES (OBEDECE E NÃO OBEDECE)
  ensurePageSpace(35);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(8.5);
  doc.setTextColor(15, 23, 42);
  doc.text('4. DEMONSTRATIVO DE LIQUIDAÇÕES (OBEDECE E NÃO OBEDECE)', marginX, curY);
  curY += 3.5;

  const liquidacoesRows = summary.allRows.filter(
    (r) => r.sourceType === 'obedece' || r.sourceType === 'nao_obedece'
  );

  const liqTableBody: any[] = liquidacoesRows.map((r, idx) => [
    idx + 1,
    r.sourceLabel,
    formatCleanNL(r.documentNumber),
    formatBRL(r.value),
    r.status || 'Liquidado',
    r.fonteRecurso || '0.6.00',
  ]);

  if (liqTableBody.length === 0) {
    liqTableBody.push([
      { content: 'Nenhuma liquidação em fila registrada', colSpan: 6, styles: { halign: 'center', textColor: [100, 116, 139] } }
    ]);
  } else {
    // Totalizer row with colSpan 3 (96mm width) so "TOTAL:" NEVER wraps into "TO TAL"
    liqTableBody.push([
      { content: 'TOTAL:', colSpan: 3, styles: { halign: 'right', fontStyle: 'bold', fillColor: [241, 245, 249], textColor: [15, 23, 42] } },
      { content: formatBRL(totalLiquidado), styles: { halign: 'right', fontStyle: 'bold', fillColor: [241, 245, 249], textColor: [15, 23, 42] } },
      { content: `${liquidacoesRows.length} liquidação(ões)`, colSpan: 2, styles: { halign: 'center', fontStyle: 'bold', fillColor: [241, 245, 249], textColor: [15, 23, 42] } },
    ]);
  }

  autoTable(doc, {
    startY: curY,
    head: [['#', 'FILA / LIQUIDAÇÃO', 'NÚMERO DA LIQUIDAÇÃO (NL)', 'VALOR BRUTO (R$)', 'SITUAÇÃO', 'FONTE']],
    body: liqTableBody,
    theme: 'grid',
    headStyles: {
      fillColor: [15, 23, 42], // Slate 900
      textColor: [255, 255, 255],
      fontSize: 7.5,
      fontStyle: 'bold',
      halign: 'center',
    },
    bodyStyles: {
      fontSize: 7,
      textColor: [30, 41, 59],
    },
    columnStyles: {
      0: { cellWidth: 8, halign: 'center' },
      1: { cellWidth: 42, halign: 'left' },
      2: { cellWidth: 46, halign: 'left', fontStyle: 'bold' },
      3: { cellWidth: 36, halign: 'right', fontStyle: 'bold' },
      4: { cellWidth: 26, halign: 'center' },
      5: { cellWidth: 24, halign: 'center' },
    },
    margin: { left: marginX, right: marginX },
  });

  // Footer / Page numbers on all pages
  const totalPages = doc.getNumberOfPages();
  for (let i = 1; i <= totalPages; i++) {
    doc.setPage(i);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7);
    doc.setTextColor(148, 163, 184); // Slate 400
    doc.text(
      `Painel Financeiro FES/SESAP — Página ${i} de ${totalPages}`,
      marginX,
      pageHeight - 8
    );
    doc.text(
      `Relatório de Extrato de Liquidações e Pagamentos`,
      pageWidth - marginX,
      pageHeight - 8,
      { align: 'right' }
    );
  }

  const safeFileName = `Extrato_Fornecedor_${summary.supplierName.substring(0, 25).replace(/[^a-zA-Z0-9]/g, '_')}_${new Date().toISOString().slice(0, 10)}.pdf`;
  doc.save(safeFileName);
}

