import { QuebraOrdemItem, QuebraOrdemResult, FileData, StageKey } from '../types';
import {
  parseCurrencyValue,
  parseFavorecidoField,
  readSpreadsheetAsMatrix,
  isSituacaoAO,
  isSituacaoEnviadaBanco,
  isBankOrIntermediary,
} from './excelParser';

export { readSpreadsheetAsMatrix, isSituacaoAO, isSituacaoEnviadaBanco };

/**
 * Checks if a Fonte de Recurso contains .600 or 0.6.00 (which must be ignored according to rules)
 */
export function isFonte600(fonteRaw: any): boolean {
  if (fonteRaw === null || fonteRaw === undefined) return false;
  const str = String(fonteRaw).trim().toLowerCase().replace(/\s+/g, '');
  if (!str) return false;

  // Checks for suffix or presence of .600, 0.6.00, .6.00, etc.
  if (str.includes('.600') || str.includes('0.6.00') || str.includes('.6.00')) return true;
  if (str.endsWith('600') || str.endsWith('0600')) return true;
  return false;
}

/**
 * Fontes de recurso monitoradas para verificação de quebra de ordem cronológica:
 * - 0. 5.00.000000
 * - 0. 7.00.000035
 * - 0. 7.04.000121
 * - 0. 7.07.000000
 * - 0. 7.20.000720
 */
export const FONTES_MONITORADAS_OBEDECE = [
  '0. 5.00.000000',
  '0. 7.00.000035',
  '0. 7.04.000121',
  '0. 7.07.000000',
  '0. 7.20.000720',
] as const;

export function isFonteMonitoradaObedece(fonteRaw: any, fileName?: string): { isMonitorada: boolean; label: string } {
  const str = String(fonteRaw || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toLowerCase();

  const fn = String(fileName || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toLowerCase();

  const digitsOnly = str.replace(/\D/g, '');

  // 1. Fonte 0. 5.00.000000
  if (
    str.includes('5.00.000000') ||
    str.includes('500000000') ||
    str.includes('0.5.00') ||
    str.includes('0. 5.00') ||
    digitsOnly === '0500000000' ||
    digitsOnly === '500000000' ||
    digitsOnly === '500' ||
    fn.includes('fonte 500') ||
    fn.includes('fonte_500') ||
    fn.includes('0.5.00')
  ) {
    return { isMonitorada: true, label: '0. 5.00.000000' };
  }

  // 2. Fonte 0. 7.00.000035
  if (
    str.includes('7.00.000035') ||
    str.includes('700000035') ||
    str.includes('0.7.00.000035') ||
    str.includes('0. 7.00.000035') ||
    digitsOnly === '0700000035' ||
    digitsOnly === '700000035' ||
    fn.includes('700000035') ||
    fn.includes('7.00.000035') ||
    fn.includes('000035')
  ) {
    return { isMonitorada: true, label: '0. 7.00.000035' };
  }

  // 3. Fonte 0. 7.04.000121
  if (
    str.includes('7.04.000121') ||
    str.includes('704000121') ||
    str.includes('0.7.04.000121') ||
    str.includes('0. 7.04.000121') ||
    digitsOnly === '0704000121' ||
    digitsOnly === '704000121' ||
    fn.includes('704000121') ||
    fn.includes('7.04.000121') ||
    fn.includes('000121')
  ) {
    return { isMonitorada: true, label: '0. 7.04.000121' };
  }

  // 4. Fonte 0. 7.07.000000
  if (
    str.includes('7.07.000000') ||
    str.includes('707000000') ||
    str.includes('0.7.07.000000') ||
    str.includes('0. 7.07.000000') ||
    digitsOnly === '0707000000' ||
    digitsOnly === '707000000' ||
    fn.includes('707000000') ||
    fn.includes('7.07.000000')
  ) {
    return { isMonitorada: true, label: '0. 7.07.000000' };
  }

  // 5. Fonte 0. 7.20.000720
  if (
    str.includes('7.20.000720') ||
    str.includes('720000720') ||
    str.includes('0.7.20.000720') ||
    str.includes('0. 7.20.000720') ||
    digitsOnly === '0720000720' ||
    digitsOnly === '720000720' ||
    fn.includes('720000720') ||
    fn.includes('7.20.000720') ||
    fn.includes('000720')
  ) {
    return { isMonitorada: true, label: '0. 7.20.000720' };
  }

  return { isMonitorada: false, label: '' };
}

/**
 * 22 Fontes de "Fluxo Rápido"
 * Regra de Negócio: Liquidações originadas dessas fontes devem receber a badge visual
 * azul "⚡ Fluxo Rápido" e NÃO DEVEM ser consideradas como quebra de ordem cronológica
 * (pois não precisam de publicação de quebra de ordem cronológica e são pagas mais rapidamente).
 */
export const FONTES_FLUXO_RAPIDO = [
  '0.6.00.0006000',
  '0.6.31.000078',
  '0.6.31.011647',
  '0.6.31.090657',
  '0.6.31.120005',
  '0.6.31.881147',
  '0.6.31.891007',
  '0.6.31.891149',
  '0.6.31.891150',
  '0.6.31.891202',
  '0.6.31.893093',
  '0.6.31.893099',
  '0.6.31.893105',
  '0.6.31.893584',
  '0.6.31.893594',
  '0.6.31.894955',
  '0.6.31.906575',
  '0.6.31.906577',
  '0.6.31.907791',
  '0.6.31.907976',
  '0.6.31.907980',
  '0.6.31.907982',
] as const;

export function isFonteFluxoRapido(fonteRaw: any, fileName?: string): { isFluxoRapido: boolean; label: string } {
  const str = String(fonteRaw || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toLowerCase();

  const fn = String(fileName || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toLowerCase();

  const digitsOnly = str.replace(/\D/g, '');
  const fnDigits = fn.replace(/\D/g, '');

  for (const f of FONTES_FLUXO_RAPIDO) {
    const fClean = f.toLowerCase();
    const fDigits = f.replace(/\D/g, '');

    if (
      str.includes(fClean) ||
      (digitsOnly.length >= 6 && (fDigits.includes(digitsOnly) || digitsOnly.includes(fDigits))) ||
      fn.includes(fClean) ||
      (fnDigits.length >= 6 && (fnDigits.includes(fDigits) || fDigits.includes(fnDigits)))
    ) {
      return { isFluxoRapido: true, label: f };
    }
  }

  // Broad check for 0.6.31 or 0.6.00 pattern
  if (str.includes('0.6.31.') || str.includes('0. 6.31.') || str.includes('6.31.')) {
    return { isFluxoRapido: true, label: '0.6.31' };
  }
  if (str.includes('0.6.00.0006000') || str.includes('0006000')) {
    return { isFluxoRapido: true, label: '0.6.00.0006000' };
  }

  return { isFluxoRapido: false, label: '' };
}

export interface RawPagamentoEmitido {
  favorecidoRaw: string;
  cnpj: string;
  name: string;
  valorOB: number;
  pp: string;
  ob: string;
  processo: string;
  notaEmpenho: string;
  fonteRecurso: string;
  situacao: string;
  rawRow: Record<string, any>;
  isAO: boolean;
  isIgnoradaFonte600: boolean;
}

/**
 * Parses the "Relatório de Pagamentos Emitidos" matrix
 */
export function parsePagamentosEmitidosMatrix(matrix: any[][]): RawPagamentoEmitido[] {
  if (!matrix || matrix.length === 0) return [];

  // Find header row with resilient fuzzy matching
  let headerRowIndex = -1;
  let favCol = -1;
  let favNECol = -1;
  let valCol = -1;
  let sitCol = -1;
  let fonteCol = -1;
  let ppCol = -1;
  let obCol = -1;
  let procCol = -1;
  let neCol = -1;

  for (let r = 0; r < Math.min(matrix.length, 35); r++) {
    const row = matrix[r];
    if (!Array.isArray(row)) continue;
    const rowStrings = row.map((c) =>
      String(c || '')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toLowerCase()
        .trim()
    );

    const hasSit = rowStrings.some(
      (c) => c.includes('situa') || c.includes('sigla') || c === 'sit' || c.includes('status') || c.includes('estado')
    );
    const hasFav = rowStrings.some(
      (c) =>
        c.includes('favorecido') ||
        c.includes('credor') ||
        c.includes('fornecedor') ||
        c.includes('benefici') ||
        c.includes('razao') ||
        c.includes('nome')
    );
    const hasVal = rowStrings.some(
      (c) =>
        c.includes('valor') ||
        c.includes('vlr') ||
        c.includes('liquido') ||
        c.includes('pago') ||
        c.includes('total')
    );
    const hasPP = rowStrings.some(
      (c) =>
        c === 'pp' ||
        c.includes('proposta') ||
        c.includes('programacao') ||
        c.includes('numero pp') ||
        c.includes('num pp') ||
        c.includes('n pp') ||
        c.includes('nr pp')
    );
    const hasOB = rowStrings.some(
      (c) =>
        c === 'ob' ||
        c.includes('ordem banc') ||
        c.includes('ordem de pag') ||
        c.includes('numero ob') ||
        c.includes('num ob') ||
        c.includes('n ob') ||
        c.includes('nr ob')
    );

    const matches = [hasSit, hasFav, hasVal, hasPP, hasOB].filter(Boolean).length;
    if (matches >= 2 || (hasFav && hasVal) || (hasPP && hasOB) || (hasOB && hasVal)) {
      headerRowIndex = r;
      rowStrings.forEach((c, idx) => {
        if (
          (c.includes('favorecido ne') ||
            c.includes('favorecido nota') ||
            c.includes('credor ne') ||
            c.includes('credor nota') ||
            c.includes('fornecedor ne') ||
            c.includes('fornecedor nota')) &&
          favNECol === -1
        ) {
          favNECol = idx;
        } else if (
          (c.includes('favorecido') ||
            c.includes('credor') ||
            c.includes('fornecedor') ||
            c.includes('benefici') ||
            c.includes('razao')) &&
          !c.includes('empenho') &&
          favCol === -1
        ) {
          favCol = idx;
        } else if (
          (c.includes('situa') || c.includes('sigla') || c === 'sit' || c.includes('status') || c.includes('estado')) &&
          sitCol === -1
        ) {
          sitCol = idx;
        } else if ((c.includes('fonte') || c.includes('recurso')) && fonteCol === -1) {
          fonteCol = idx;
        } else if (
          (c.includes('valor') || c.includes('vlr') || c.includes('liquido') || c.includes('pago') || c.includes('total')) &&
          valCol === -1
        ) {
          valCol = idx;
        } else if (
          (c === 'pp' ||
            c.includes('proposta') ||
            c.includes('programacao') ||
            c.includes('numero pp') ||
            c.includes('num pp') ||
            c.includes('n pp') ||
            c.includes('nr pp')) &&
          ppCol === -1
        ) {
          ppCol = idx;
        } else if (
          (c === 'ob' ||
            c.includes('ordem banc') ||
            c.includes('ordem de pag') ||
            c.includes('numero ob') ||
            c.includes('num ob') ||
            c.includes('n ob') ||
            c.includes('nr ob')) &&
          obCol === -1
        ) {
          obCol = idx;
        } else if ((c.includes('processo') || c.includes('proc')) && procCol === -1) {
          procCol = idx;
        } else if ((c.includes('empenho') || c.includes('ne')) && neCol === -1) {
          neCol = idx;
        }
      });
      break;
    }
  }

  // Fallbacks if not detected
  if (headerRowIndex === -1) {
    headerRowIndex = 0;
  }

  const rawHeader = matrix[headerRowIndex] || [];
  const headers = rawHeader.map((h, i) => String(h || `Coluna_${i + 1}`).trim());

  // Additional fuzzy pass on headers if columns are still -1
  if (favCol === -1 || valCol === -1 || sitCol === -1 || ppCol === -1 || obCol === -1 || favNECol === -1) {
    headers.forEach((h, idx) => {
      const nh = h
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toLowerCase()
        .trim();
      if (
        favNECol === -1 &&
        (nh.includes('favorecido ne') ||
          nh.includes('favorecido nota') ||
          nh.includes('credor ne') ||
          nh.includes('fornecedor ne'))
      ) {
        favNECol = idx;
      }
      if (favCol === -1 && (nh.includes('favorecido') || nh.includes('credor') || nh.includes('nome') || nh.includes('fornecedor'))) {
        favCol = idx;
      }
      if (valCol === -1 && (nh.includes('valor') || nh.includes('pago') || nh.includes('liquido') || nh.includes('vlr') || nh.includes('total'))) {
        valCol = idx;
      }
      if (sitCol === -1 && (nh.includes('situa') || nh.includes('sigla') || nh === 'sit' || nh.includes('status'))) {
        sitCol = idx;
      }
      if (fonteCol === -1 && (nh.includes('fonte') || nh.includes('recurso'))) {
        fonteCol = idx;
      }
      if (ppCol === -1 && (nh.includes('pp') || nh.includes('proposta') || nh.includes('programacao'))) {
        ppCol = idx;
      }
      if (obCol === -1 && (nh.includes('ob') || nh.includes('ordem'))) {
        obCol = idx;
      }
      if (procCol === -1 && (nh.includes('processo') || nh.includes('proc'))) {
        procCol = idx;
      }
      if (neCol === -1 && (nh.includes('empenho') || nh.includes('ne'))) {
        neCol = idx;
      }
    });
  }

  // Heuristic inspection if valCol, favCol, or sitCol still -1
  if (valCol === -1 || favCol === -1 || sitCol === -1) {
    const sample = matrix.slice(headerRowIndex + 1, headerRowIndex + 35);
    if (sitCol === -1) {
      for (let c = 0; c < headers.length; c++) {
        if (c === valCol || c === favCol) continue;
        let aoFound = 0;
        sample.forEach((row) => {
          if (Array.isArray(row) && (isSituacaoAO(row[c]) || isSituacaoEnviadaBanco(row[c]).isEnviada)) {
            aoFound++;
          }
        });
        if (aoFound >= 1) {
          sitCol = c;
          break;
        }
      }
    }
    if (valCol === -1) {
      let bestIdx = -1;
      let maxNums = 0;
      for (let c = 0; c < headers.length; c++) {
        let count = 0;
        sample.forEach((row) => {
          if (Array.isArray(row) && parseCurrencyValue(row[c]) > 0) count++;
        });
        if (count > maxNums && count >= 2) {
          maxNums = count;
          bestIdx = c;
        }
      }
      if (bestIdx !== -1) valCol = bestIdx;
    }

    if (favCol === -1) {
      let bestIdx = -1;
      let maxText = 0;
      for (let c = 0; c < headers.length; c++) {
        if (c === valCol) continue;
        let count = 0;
        sample.forEach((row) => {
          const s = String(row[c] || '');
          if (s.length > 6 && !/^\d+$/.test(s)) count++;
        });
        if (count > maxText && count >= 2) {
          maxText = count;
          bestIdx = c;
        }
      }
      if (bestIdx !== -1) favCol = bestIdx;
    }
  }

  // Pre-check if any row in the file has "AO" situation
  let anyRowHasAO = false;
  if (sitCol !== -1) {
    for (let r = headerRowIndex + 1; r < matrix.length; r++) {
      const row = matrix[r];
      if (Array.isArray(row) && isSituacaoAO(row[sitCol])) {
        anyRowHasAO = true;
        break;
      }
    }
  }

  const results: RawPagamentoEmitido[] = [];

  for (let r = headerRowIndex + 1; r < matrix.length; r++) {
    const row = matrix[r];
    if (!Array.isArray(row) || row.length === 0) continue;

    // Check if entire row is empty
    const isEmpty = row.every((c) => c === null || c === undefined || String(c).trim() === '');
    if (isEmpty) continue;

    // Skip total or summary footer rows
    const firstCell = String(row[0] || '').trim().toLowerCase();
    if (firstCell.startsWith('total') || firstCell.startsWith('quantidade de') || firstCell.startsWith('emitido em')) {
      continue;
    }

    const rowObj: Record<string, any> = {};
    headers.forEach((hdr, idx) => {
      rowObj[hdr] = row[idx] ?? '';
    });

    const favRaw = favCol !== -1 ? row[favCol] : '';
    const favNERaw = favNECol !== -1 ? row[favNECol] : '';
    let { cnpj, name } = parseFavorecidoField(favRaw);

    if (favNERaw) {
      const favNEInfo = parseFavorecidoField(favNERaw);
      if (favNEInfo && favNEInfo.cnpj !== 'N/I' && (cnpj === 'N/I' || isBankOrIntermediary(name))) {
        cnpj = favNEInfo.cnpj;
        name = favNEInfo.name;
      }
    }

    const situacaoRaw = sitCol !== -1 ? String(row[sitCol] || '').trim() : '';
    const fonteRaw = fonteCol !== -1 ? String(row[fonteCol] || '').trim() : '';
    const valorOB = valCol !== -1 ? parseCurrencyValue(row[valCol]) : 0;
    const pp = ppCol !== -1 ? String(row[ppCol] || '').trim() : '';
    const ob = obCol !== -1 ? String(row[obCol] || '').trim() : '';
    const processo = procCol !== -1 ? String(row[procCol] || '').trim() : '';
    const notaEmpenho = neCol !== -1 ? String(row[neCol] || '').trim() : '';

    // If file has AO situations, filter strictly by AO. If no row has AO (e.g. status is "CONFIRMADA", "PAGA", or no sit column), treat non-empty payments as AO
    const isAO = anyRowHasAO ? isSituacaoAO(situacaoRaw) : (valorOB > 0 || ob.length > 0 || pp.length > 0);
    const isIgnoradaFonte600 = isFonte600(fonteRaw);

    results.push({
      favorecidoRaw: String(favRaw || name).trim(),
      cnpj,
      name,
      valorOB,
      pp,
      ob,
      processo,
      notaEmpenho,
      fonteRecurso: fonteRaw,
      situacao: situacaoRaw,
      rawRow: rowObj,
      isAO,
      isIgnoradaFonte600,
    });
  }

  return results;
}

export interface LiquidationCandidate {
  favorecido: string;
  cnpj: string;
  cnpjDigits: string;
  normalizedName: string;
  valor: number;
  stage: StageKey;
  fileName: string;
  documento: string;
  fonteRecurso: string;
  isFonteMonitorada: boolean;
  fonteMonitoradaLabel?: string;
  isFluxoRapido?: boolean;
  fonteFluxoRapido?: string;
  rawRow: Record<string, any>;
  isUsed?: boolean;
}

/**
 * Extracts candidate liquidations from all loaded files in the workspace
 */
export function extractLiquidationCandidates(files: FileData[]): LiquidationCandidate[] {
  const candidates: LiquidationCandidate[] = [];

  files.forEach((file) => {
    const favCol = file.detectedFavorecidoCol;
    const valCol = file.detectedValorCol;
    const stage = file.stage || 'outros';
    const isObedeceStage = stage === 'obedece';

    const fonteCol =
      file.detectedFonteCol ||
      file.headers.find((h) => {
        const nh = h.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
        return nh.includes('fonte') || nh.includes('recurso');
      });

    file.rows.forEach((row) => {
      const rawFav = row[favCol];
      const { cnpj, name } = parseFavorecidoField(rawFav);
      const valor = parseCurrencyValue(row[valCol]);

      if (valor <= 0 && (!name || name === 'NÃO INFORMADO')) return;

      const cnpjDigits = (cnpj || '').replace(/\D/g, '');
      const normalizedName = (name || '')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toUpperCase()
        .replace(/[^A-Z0-9]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();

      // Look for document / PP / processo number in row
      let doc = '';
      for (const [k, v] of Object.entries(row)) {
        const lk = k.toLowerCase();
        if ((lk.includes('documento') || lk.includes('número') || lk.includes('numero') || lk.includes('pp') || lk.includes('processo') || lk.includes('empenho')) && v) {
          doc = String(v).trim();
          break;
        }
      }

      // Check Fonte de Recurso
      const rawFonte = fonteCol ? row[fonteCol] : '';
      const fluxoCheck = isFonteFluxoRapido(rawFonte, file.fileName);
      const fonteCheck = isFonteMonitoradaObedece(rawFonte, file.fileName);

      let isFluxoRapido = false;
      let fonteFluxoRapido = '';
      let isFonteMonitorada = false;
      let fonteMonitoradaLabel = '';
      let fonteRecurso = String(rawFonte || '').trim();

      if (fluxoCheck.isFluxoRapido) {
        // Regra de Negócio: Fontes de Fluxo Rápido têm prioridade e NÃO são monitoradas para quebra de ordem
        isFluxoRapido = true;
        fonteFluxoRapido = fluxoCheck.label;
        if (!fonteRecurso) fonteRecurso = fluxoCheck.label;
      } else if (fonteCheck.isMonitorada) {
        isFonteMonitorada = true;
        fonteMonitoradaLabel = fonteCheck.label;
      } else if (isObedeceStage) {
        // When in Obedece stage and no contrary source is specified, associate to monitored queue
        isFonteMonitorada = true;
        fonteMonitoradaLabel = '0. 5.00.000000';
        if (!fonteRecurso) {
          fonteRecurso = '0. 5.00.000000';
        }
      }

      candidates.push({
        favorecido: name,
        cnpj,
        cnpjDigits,
        normalizedName,
        valor,
        stage,
        fileName: file.fileName,
        documento: doc,
        fonteRecurso,
        isFonteMonitorada,
        fonteMonitoradaLabel,
        isFluxoRapido,
        fonteFluxoRapido,
        rawRow: row,
        isUsed: false,
      });
    });
  });

  return candidates;
}

/**
 * Main cross-referencing function for Quebra de Ordem (Monitoramento de Publicações)
 * Regra de Negócio:
 * Se uma liquidação que estava em aberto (relatório de obedece das fontes 0. 5.00.000000,
 * 0. 7.00.000035, 0. 7.04.000121, 0. 7.07.000000, 0. 7.20.000720) e um valor igual ou equivalente
 * de despesa daquele fornecedor apareceu no relatório geral de pagamento, é possível que tenha
 * ocorrido a quebra de ordem cronológica.
 */
export function crossReferenceQuebraOrdem(
  pagamentosRaw: RawPagamentoEmitido[],
  liquidationCandidates: LiquidationCandidate[],
  fileName?: string
): QuebraOrdemResult {
  const totalProcessados = pagamentosRaw.length;
  let totalAOEncontrados = 0;
  let totalIgnoradosFonte600 = 0;
  let totalValidos = 0;
  let totalComLiquidacaoCasada = 0;
  let totalPossivelQuebraOrdem = 0;
  let valorPossivelQuebraOrdem = 0;
  let totalFluxoRapido = 0;
  let valorFluxoRapido = 0;
  let valorTotalOB = 0;
  let valorTotalLiquidacao = 0;
  let valorTotalImpostos = 0;

  // Make a working copy of liquidation candidates
  const availableCandidates = liquidationCandidates.map((c) => ({ ...c }));

  const items: QuebraOrdemItem[] = [];

  pagamentosRaw.forEach((item, index) => {
    if (!item.isAO) {
      return;
    }
    totalAOEncontrados++;

    if (item.isIgnoradaFonte600) {
      totalIgnoradosFonte600++;
      return;
    }

    totalValidos++;
    valorTotalOB += item.valorOB;

    const itemCnpjDigits = (item.cnpj || '').replace(/\D/g, '');
    const itemNormName = (item.name || '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toUpperCase()
      .replace(/[^A-Z0-9]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();

    // Find candidate matches for this creditor
    const matchedCandidates = availableCandidates.filter((cand) => {
      if (cand.isUsed) return false;

      // 1. Match by CNPJ (if valid 14 digits or CPF 11 digits)
      if (itemCnpjDigits.length >= 11 && cand.cnpjDigits.length >= 11) {
        if (itemCnpjDigits === cand.cnpjDigits) return true;
      }

      // 2. Match by Name if CNPJ not available or matched
      if (itemNormName && cand.normalizedName) {
        if (itemNormName === cand.normalizedName) return true;
        if (itemNormName.length > 8 && cand.normalizedName.length > 8) {
          if (itemNormName.includes(cand.normalizedName) || cand.normalizedName.includes(itemNormName)) return true;
        }
      }

      return false;
    });

    let bestCandidate: LiquidationCandidate | null = null;

    if (matchedCandidates.length > 0) {
      // Prioritize candidates from OBEDECE stage with monitored sources (0.5.00, 0.7.00.035, etc.)
      const obedeceMonitoredCandidates = matchedCandidates.filter(
        (c) => c.stage === 'obedece' && c.isFonteMonitorada
      );

      // Function to sort candidates by value equivalence (favoring values equal or with plausible tax retention)
      const sortByValueEquivalence = (cands: LiquidationCandidate[]) => {
        return [...cands].sort((a, b) => {
          // Exact or near-exact match gets highest priority
          const diffA = Math.abs(a.valor - item.valorOB);
          const diffB = Math.abs(b.valor - item.valorOB);
          if (diffA < 0.05 && diffB >= 0.05) return -1;
          if (diffB < 0.05 && diffA >= 0.05) return 1;

          // Tax retention: liquidation >= OB and difference <= 40%
          const isTaxPlausibleA = a.valor >= item.valorOB && (a.valor - item.valorOB) / a.valor <= 0.40;
          const isTaxPlausibleB = b.valor >= item.valorOB && (b.valor - item.valorOB) / b.valor <= 0.40;
          if (isTaxPlausibleA && !isTaxPlausibleB) return -1;
          if (isTaxPlausibleB && !isTaxPlausibleA) return 1;

          return diffA - diffB;
        });
      };

      if (obedeceMonitoredCandidates.length > 0) {
        const sortedObedece = sortByValueEquivalence(obedeceMonitoredCandidates);
        bestCandidate = sortedObedece[0];
      } else {
        const otherObedece = matchedCandidates.filter((c) => c.stage === 'obedece');
        if (otherObedece.length > 0) {
          bestCandidate = sortByValueEquivalence(otherObedece)[0];
        } else {
          bestCandidate = sortByValueEquivalence(matchedCandidates)[0];
        }
      }
    }

    let liquidacaoCorrespondente: QuebraOrdemItem['liquidacaoCorrespondente'] = undefined;
    let isPossivelQuebraOrdem = false;
    let fonteMonitorada: string | undefined = undefined;
    let isFluxoRapido = false;
    let fonteFluxoRapido: string | undefined = undefined;

    // Check if the payment itself is from a Fast Flow source
    const paymentFluxoCheck = isFonteFluxoRapido(item.fonteRecurso);
    if (paymentFluxoCheck.isFluxoRapido) {
      isFluxoRapido = true;
      fonteFluxoRapido = paymentFluxoCheck.label;
    }

    if (bestCandidate) {
      bestCandidate.isUsed = true;
      totalComLiquidacaoCasada++;

      if (bestCandidate.isFluxoRapido) {
        isFluxoRapido = true;
        fonteFluxoRapido = bestCandidate.fonteFluxoRapido || paymentFluxoCheck.label || bestCandidate.fonteRecurso;
      }

      const diferenca = Math.max(0, bestCandidate.valor - item.valorOB);
      const percentual = bestCandidate.valor > 0 ? (diferenca / bestCandidate.valor) * 100 : 0;

      valorTotalLiquidacao += bestCandidate.valor;
      valorTotalImpostos += diferenca;

      // Detection of potential chronological order break:
      // Liquidation was open in Obedece stage under monitored sources, and an equivalent payment appeared in the general report
      // Regra de Negócio: Fontes de "Fluxo Rápido" NÃO DEVEM ser consideradas como quebra de ordem cronológica.
      const isEquivalentValue =
        Math.abs(bestCandidate.valor - item.valorOB) < 0.05 ||
        (bestCandidate.valor >= item.valorOB && (bestCandidate.valor - item.valorOB) / bestCandidate.valor <= 0.40) ||
        Math.abs(bestCandidate.valor - item.valorOB) / Math.max(bestCandidate.valor, 1) <= 0.15;

      if (!isFluxoRapido && bestCandidate.stage === 'obedece' && bestCandidate.isFonteMonitorada && isEquivalentValue) {
        isPossivelQuebraOrdem = true;
        fonteMonitorada = bestCandidate.fonteMonitoradaLabel || bestCandidate.fonteRecurso || '0. 5.00.000000';
        totalPossivelQuebraOrdem++;
        valorPossivelQuebraOrdem += item.valorOB;
      }

      liquidacaoCorrespondente = {
        valorLiquidacao: bestCandidate.valor,
        etapaOrigem: bestCandidate.stage,
        nomePlanilha: bestCandidate.fileName,
        documento: bestCandidate.documento,
        diferencaImpostos: diferenca,
        percentualDiferenca: percentual,
        matched: true,
        fonteRecurso: bestCandidate.fonteRecurso,
        isFonteMonitorada: bestCandidate.isFonteMonitorada,
        isFluxoRapido,
        fonteFluxoRapido,
      };
    } else {
      // If no matching liquidation was found, default liquidation value is equal to OB
      valorTotalLiquidacao += item.valorOB;
    }

    if (isFluxoRapido) {
      totalFluxoRapido++;
      valorFluxoRapido += item.valorOB;
    }

    items.push({
      id: `qo_${index}_${Date.now()}`,
      favorecido: item.name || item.favorecidoRaw || 'NÃO INFORMADO',
      cnpj: item.cnpj,
      valorOB: item.valorOB,
      pp: item.pp,
      ob: item.ob,
      processo: item.processo,
      notaEmpenho: item.notaEmpenho,
      fonteRecurso: item.fonteRecurso || bestCandidate?.fonteRecurso || 'N/I',
      situacao: item.situacao || 'AO',
      isPossivelQuebraOrdem,
      fonteMonitorada,
      isFluxoRapido,
      fonteFluxoRapido,
      liquidacaoCorrespondente,
      rawRow: item.rawRow,
    });
  });

  return {
    totalProcessados,
    totalAOEncontrados,
    totalIgnoradosFonte600,
    totalValidos,
    totalComLiquidacaoCasada,
    totalPossivelQuebraOrdem,
    valorPossivelQuebraOrdem,
    totalFluxoRapido,
    valorFluxoRapido,
    valorTotalOB,
    valorTotalLiquidacao,
    valorTotalImpostos,
    items,
    fileName,
    processedAt: new Date(),
  };
}

/**
 * Generates sample demo data for quick test and verification of Quebra de Ordem
 */
export function generateSampleQuebraOrdemData(): RawPagamentoEmitido[] {
  return [
    {
      favorecidoRaw: '08.234.112/0001-45 - HOSPITAL SANTA MARIA LTDA',
      cnpj: '08.234.112/0001-45',
      name: 'HOSPITAL SANTA MARIA LTDA',
      valorOB: 428500.0,
      pp: '2024PP00142',
      ob: '2024OB008912',
      processo: '00610023.001923/2024-11',
      notaEmpenho: '2024NE001402',
      fonteRecurso: '0. 5.00.000000',
      situacao: 'AO',
      rawRow: {},
      isAO: true,
      isIgnoradaFonte600: false,
    },
    {
      favorecidoRaw: '12.890.345/0001-90 - MEDCLIN DISTRIBUIDORA DE MEDICAMENTOS S/A',
      cnpj: '12.890.345/0001-90',
      name: 'MEDCLIN DISTRIBUIDORA DE MEDICAMENTOS S/A',
      valorOB: 185200.0,
      pp: '2024PP00188',
      ob: '2024OB008915',
      processo: '00610045.002100/2024-82',
      notaEmpenho: '2024NE001844',
      fonteRecurso: '0. 7.00.000035',
      situacao: 'AO',
      rawRow: {},
      isAO: true,
      isIgnoradaFonte600: false,
    },
    {
      favorecidoRaw: '03.456.789/0001-12 - SERVICOS MEDICOS POTIGUAR LTDA',
      cnpj: '03.456.789/0001-12',
      name: 'SERVICOS MEDICOS POTIGUAR LTDA',
      valorOB: 92400.0,
      pp: '2024PP00210',
      ob: '2024OB008920',
      processo: '00610012.000845/2024-19',
      notaEmpenho: '2024NE002190',
      fonteRecurso: '0. 7.04.000121',
      situacao: 'AO',
      rawRow: {},
      isAO: true,
      isIgnoradaFonte600: false,
    },
    {
      // Ignored sample with .600
      favorecidoRaw: '24.111.222/0001-33 - ENERGIA E GAS DO NORDESTE S/A',
      cnpj: '24.111.222/0001-33',
      name: 'ENERGIA E GAS DO NORDESTE S/A',
      valorOB: 310000.0,
      pp: '2024PP00330',
      ob: '2024OB008930',
      processo: '00610088.003400/2024-55',
      notaEmpenho: '2024NE003310',
      fonteRecurso: '1.500.600', // Contains .600!
      situacao: 'AO',
      rawRow: {},
      isAO: true,
      isIgnoradaFonte600: true,
    },
    {
      favorecidoRaw: '19.876.543/0001-21 - OXIGENIO NATAL GASES HOSPITALARES LTDA',
      cnpj: '19.876.543/0001-21',
      name: 'OXIGENIO NATAL GASES HOSPITALARES LTDA',
      valorOB: 63150.0,
      pp: '2024PP00245',
      ob: '2024OB008935',
      processo: '00610099.004120/2024-33',
      notaEmpenho: '2024NE002498',
      fonteRecurso: '0. 7.20.000720',
      situacao: 'AO',
      rawRow: {},
      isAO: true,
      isIgnoradaFonte600: false,
    },
  ];
}

/**
 * Scans all files imported into the workspace (including general / large-scale SIGEF reports)
 * and extracts all rows with "AO" situation (PPs Prontas) as RawPagamentoEmitido records.
 */
export function extractAORowsFromWorkspaceFiles(files: FileData[]): {
  sourceFileName: string;
  records: RawPagamentoEmitido[];
} {
  const records: RawPagamentoEmitido[] = [];
  const fileNames: string[] = [];

  files.forEach((file) => {
    const sitCol = file.detectedSituacaoCol;
    const favCol = file.detectedFavorecidoCol;
    const valCol = file.detectedValorCol;
    const ppCol = file.detectedNumeroCol || file.detectedPpCol;
    const obCol = file.detectedObCol;
    const neCol = file.detectedNotaEmpenhoCol;
    const favNeCol = file.detectedFavorecidoNECol;
    const procCol =
      file.detectedProcessoCol ||
      file.headers.find((h) => {
        const nh = h.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
        return nh.includes('processo') || nh.includes('proc');
      });
    const fonteCol =
      file.detectedFonteCol ||
      file.headers.find((h) => {
        const nh = h.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
        return nh.includes('fonte') || nh.includes('recurso');
      });

    let fileHasAO = false;

    file.rows.forEach((row) => {
      const rawSit = sitCol ? row[sitCol] : undefined;
      const isAO = isSituacaoAO(rawSit);

      if (isAO) {
        fileHasAO = true;
        let rawFav = favCol ? row[favCol] : '';
        let rawFavNE = favNeCol ? row[favNeCol] : undefined;
        let { cnpj, name } = parseFavorecidoField(rawFav);
        const favNEInfo = rawFavNE ? parseFavorecidoField(rawFavNE) : null;

        if (favNEInfo && favNEInfo.cnpj !== 'N/I' && (cnpj === 'N/I' || isBankOrIntermediary(name))) {
          cnpj = favNEInfo.cnpj;
          name = favNEInfo.name;
        }

        const valorOB = valCol ? parseCurrencyValue(row[valCol]) : 0;
        const pp = ppCol ? String(row[ppCol] || '').trim() : '';
        const ob = obCol ? String(row[obCol] || '').trim() : '';
        const processo = procCol ? String(row[procCol] || '').trim() : '';
        const notaEmpenho = neCol ? String(row[neCol] || '').trim() : '';
        const fonteRaw = fonteCol ? String(row[fonteCol] || '').trim() : '';
        const isIgnoradaFonte600 = isFonte600(fonteRaw);

        records.push({
          favorecidoRaw: String(rawFav || name).trim(),
          cnpj,
          name,
          valorOB,
          pp,
          ob,
          processo,
          notaEmpenho,
          fonteRecurso: fonteRaw,
          situacao: String(rawSit || 'AO').trim(),
          rawRow: row,
          isAO: true,
          isIgnoradaFonte600,
        });
      }
    });

    if (fileHasAO) {
      fileNames.push(file.fileName);
    }
  });

  return {
    sourceFileName: fileNames.join(', ') || 'Relatório SIGEF (Base de PPs)',
    records,
  };
}
