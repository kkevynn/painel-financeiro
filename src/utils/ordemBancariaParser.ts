import { parseCurrencyValue, parseFavorecidoField, isSituacaoConclusivaPagamento, readSpreadsheetAsMatrix, isTotalizadorRow, isFonteHeader, isFonteString } from './excelParser';
import { FileData } from '../types';

/**
 * Checks if a filename indicates an Ordem Bancária (OB) spreadsheet
 */
export function isObFileName(fileName: string): boolean {
  if (!fileName) return false;
  const norm = fileName
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, ' ')
    .trim();

  return (
    norm.includes('ordem bancaria') ||
    norm.includes('ordens bancarias') ||
    norm.includes('listar ordem') ||
    norm.includes('listar ob') ||
    norm.includes('relacao de ob') ||
    norm.includes('relatorio ob') ||
    norm.includes('relatorio de ob') ||
    norm.includes('planilha ob') ||
    norm.includes('planilha de ob') ||
    ((norm.includes('obs') || norm.includes('ob')) && norm.includes('principal')) ||
    /\bobs?\b/.test(norm)
  );
}

/**
 * Detects if a Fonte de Recursos is F600 or any of its derivatives
 * (e.g. 0.6.00.000600, 0.6.00, 0.6.02..., 0.6.20..., 0.6.xx, 600, F600, etc.)
 */
export function isFonte600OuDerivada(fonteRaw: any): boolean {
  if (fonteRaw === null || fonteRaw === undefined) return false;
  const str = String(fonteRaw)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toUpperCase();
  if (!str || str === '-' || str === '0') return false;

  // Direct check
  if (str.includes('F600') || str.includes('F.600') || str.includes('FONTE 600') || str.includes('FONTE600') || str.includes('FONTE_600')) {
    return true;
  }

  // Dot format: 0.6.00, 0.6.02, 0.6.20, etc.
  if (str.includes('0.6.') || str.includes('0. 6.') || str.includes('.600') || str.includes('0.600')) {
    return true;
  }

  // Exact 600 or starting with 600
  if (str === '600' || str.startsWith('600 ') || str.startsWith('600-') || str.startsWith('600.')) {
    return true;
  }

  // Numbers only analysis: 0600, 600, 000600...
  const digits = str.replace(/\D/g, '');
  if (digits === '600' || digits === '0600' || digits.startsWith('0600') || digits.startsWith('06') || digits.endsWith('000600') || digits.endsWith('600')) {
    if (digits.length <= 12) {
      return true;
    }
  }

  return false;
}

/**
 * Strictly validates whether a cell value represents a valid Fonte de Recursos code
 * (rejecting monetary values like 1.500,00 or 1.620,37).
 */
export function isValidFonteRecursoCode(val: any): boolean {
  if (val === null || val === undefined) return false;
  const s = String(val).trim().toUpperCase();
  if (!s || s === '-' || s === '0') return false;
  if (/\d+,\d{2}$/.test(s) || /^\d+\.\d{1,2}$/.test(s) || s.includes('R$')) return false;
  if (/^[012]\.\d\.\d{2}(?:\.\d{6})?$/.test(s)) return true;
  if (/^(?:0500|0600|1500|1600|0100|1000|600|500|F600|F500)$/i.test(s)) return true;
  if (/^FONTE\s+[0-9.]+$/i.test(s)) return true;
  return false;
}

/**
 * Searches a row for any cell that is a strictly valid Fonte de Recurso code
 */
export function findFonteRecursoInRow(row: any[]): string {
  if (!Array.isArray(row)) return '';
  for (const cell of row) {
    if (cell === null || cell === undefined) continue;
    if (isValidFonteRecursoCode(cell)) {
      return String(cell).trim().toUpperCase();
    }
  }
  return '';
}

export interface OrdemBancariaItem {
  id: string;
  ob: string; // Ex: 2026OB001251
  dataReferencia: string; // Ex: 21/01/2026
  dataIso?: string; // Ex: 2026-01-21
  mesAnoKey: string; // Ex: 2026-01 (format YYYY-MM)
  mesAnoLabel: string; // Ex: Janeiro/2026
  valor: number; // Ex: 1620.37
  situacao: string; // Ex: CB
  pp: string; // Preparação de Pagamento
  fonteRecurso: string; // Ex: 0.6.00.000600
  favorecidoRaw: string; // Ex: 30.386.911/0001-60 NOME DA EMPRESA
  favorecidoCnpj: string;
  favorecidoName: string;
  linhaOrigem?: number;
}

const MONTH_NAMES_PT = [
  'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
  'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
];

/**
 * Parses Brazilian date strings (DD/MM/YYYY or YYYY-MM-DD) or Excel numeric serial dates
 * into monthKey ("YYYY-MM"), monthLabel ("Janeiro/2026"), and dateIso ("YYYY-MM-DD").
 */
export function parseDateToMonthInfo(rawDate: any): {
  mesAnoKey: string;
  mesAnoLabel: string;
  dateIso?: string;
} {
  if (rawDate === null || rawDate === undefined || rawDate === '') {
    return { mesAnoKey: '', mesAnoLabel: 'Sem Data' };
  }

  // Handle Excel numeric serial dates (e.g. 46043 -> 2026-01-21)
  const numSerial =
    typeof rawDate === 'number'
      ? rawDate
      : typeof rawDate === 'string' && /^\d{5}(\.\d+)?$/.test(rawDate.trim())
      ? parseFloat(rawDate.trim())
      : NaN;

  if (!isNaN(numSerial) && numSerial > 30000 && numSerial < 70000) {
    try {
      const utcDays = Math.floor(numSerial - 25569);
      const utcValue = utcDays * 86400 * 1000;
      const dateObj = new Date(utcValue);
      const year = dateObj.getUTCFullYear();
      const monthNum = dateObj.getUTCMonth() + 1;
      const dayNum = dateObj.getUTCDate();
      const monthStr = String(monthNum).padStart(2, '0');
      const dayStr = String(dayNum).padStart(2, '0');
      const key = `${year}-${monthStr}`;
      const label = `${MONTH_NAMES_PT[monthNum - 1]}/${year}`;
      return { mesAnoKey: key, mesAnoLabel: label, dateIso: `${year}-${monthStr}-${dayStr}` };
    } catch {
      // Fallback
    }
  }

  const str = String(rawDate).trim();

  // Pattern YYYY-MM-DD or YYYY/MM/DD (tested FIRST to avoid matching DD-MM-YY on ISO strings)
  const isoMatch = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (isoMatch) {
    const year = parseInt(isoMatch[1], 10);
    const month = parseInt(isoMatch[2], 10);
    const day = parseInt(isoMatch[3], 10);
    if (month >= 1 && month <= 12 && year >= 2000 && year <= 2100) {
      const monthStr = String(month).padStart(2, '0');
      const dayStr = String(day).padStart(2, '0');
      const key = `${year}-${monthStr}`;
      const label = `${MONTH_NAMES_PT[month - 1]}/${year}`;
      return { mesAnoKey: key, mesAnoLabel: label, dateIso: `${year}-${monthStr}-${dayStr}` };
    }
  }

  // Pattern DD/MM/YYYY or DD-MM-YYYY or DD/MM/YY
  const brMatch = str.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (brMatch) {
    const day = parseInt(brMatch[1], 10);
    const month = parseInt(brMatch[2], 10);
    let year = parseInt(brMatch[3], 10);
    if (year < 100) {
      year = year >= 50 ? 1900 + year : 2000 + year;
    }
    if (month >= 1 && month <= 12 && year >= 2000 && year <= 2100) {
      const monthStr = String(month).padStart(2, '0');
      const dayStr = String(day).padStart(2, '0');
      const key = `${year}-${monthStr}`;
      const label = `${MONTH_NAMES_PT[month - 1]}/${year}`;
      return { mesAnoKey: key, mesAnoLabel: label, dateIso: `${year}-${monthStr}-${dayStr}` };
    }
  }

  return { mesAnoKey: '', mesAnoLabel: str || 'Sem Data' };
}

/**
 * Checks if a string looks like an Ordem Bancária (OB) number,
 * e.g., 2026OB001251, 2025OB..., OB123456, 2026-OB-001251, etc.
 */
export function isObNumber(val: any): boolean {
  if (!val) return false;
  const str = String(val).trim().toUpperCase();
  const clean = str.replace(/[\s\.\-\/]/g, '');
  return (
    /\b\d{4}OB\d+\b/i.test(clean) ||
    /\bOB\d{3,}\b/i.test(clean) ||
    /^\d{4}[A-Z]{2}\d{3,}$/.test(clean) ||
    /^OB[\s\-_]?\d+/i.test(str)
  );
}

/**
 * Checks if a situation string is explicitly rejected or cancelled.
 * All other non-rejected statuses in an OB spreadsheet represent valid payments.
 */
function isRejectedOb(sitRaw: any): boolean {
  const str = String(sitRaw ?? '').trim().toUpperCase();
  if (!str) return false;
  return (
    str.startsWith('RJ') ||
    str.startsWith('PPRJ') ||
    str.includes('CANCEL') ||
    str.includes('REJEIT') ||
    str.includes('ESTORN') ||
    str.includes('DEVOLV') ||
    str === 'CBNC'
  );
}

/**
 * Extracts OrdemBancariaItem[] directly from an uploaded FileData.
 */
export function extractOrdensBancariasFromFileData(file: FileData): OrdemBancariaItem[] {
  if (!file || !file.rows || file.rows.length === 0) return [];
  const items: OrdemBancariaItem[] = [];

  const headers = Object.keys(file.rows[0] || {});
  const obCol =
    file.detectedObCol ||
    headers.find((k) => {
      const lk = k.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
      return lk.includes('ordem') || lk.includes('ob') || lk.includes('documento');
    });

  const fonteCol =
    file.detectedFonteCol ||
    headers.find((k) => isFonteHeader(k) || k.toLowerCase().includes('fonte'));

  const valCol =
    (file.detectedValorCol && !isFonteHeader(file.detectedValorCol) && file.detectedValorCol !== fonteCol ? file.detectedValorCol : null) ||
    headers.find((k) => {
      if (isFonteHeader(k) || k === fonteCol) return false;
      const lk = k.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
      return lk.includes('valor') || lk.includes('pago') || lk.includes('vlr');
    });

  const favCol =
    file.detectedFavorecidoCol ||
    headers.find((k) => {
      const lk = k.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
      return lk.includes('favorecido') || lk.includes('credor') || lk.includes('fornecedor');
    });
  const cnpjCol =
    file.detectedCnpjCol ||
    headers.find((k) => {
      const lk = k.toLowerCase();
      return lk.includes('cnpj') || lk.includes('cpf');
    });
  const sitCol =
    file.detectedSituacaoCol ||
    headers.find((k) => {
      const lk = k.toLowerCase();
      return lk.includes('situacao') || lk.includes('status');
    });
  const ppCol =
    file.detectedPpCol ||
    file.detectedNumeroCol ||
    headers.find((k) => {
      const lk = k.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
      return lk === 'pp' || lk.includes('proposta') || lk.includes('preparacao') || lk.includes('num pp') || lk.includes('no pp') || lk.includes('num. pp');
    });

  // Prioritize Payment/OB/Confirmation date column over Emission/Proposal date column
  const paymentDateKeywords = [
    'data pagamento', 'dt pagamento', 'data do pagamento', 'data pgto', 'dt pgto',
    'data ob', 'data da ob', 'dt ob', 'data confirmacao', 'data confirmacao banco',
    'data situacao', 'data da situacao', 'dt situacao', 'data movimento', 'data baixa', 'data efetivacao'
  ];

  const paymentDateCol = headers.find((k) => {
    const lk = k.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
    return paymentDateKeywords.some((kw) => lk.includes(kw));
  });

  const generalDateCol =
    paymentDateCol ||
    file.detectedDataCol ||
    headers.find((k) => {
      const lk = k.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
      return lk.includes('data') || lk.includes('dt');
    });

  for (let idx = 0; idx < file.rows.length; idx++) {
    const row = file.rows[idx];
    const rawOb = obCol ? String(row[obCol] || '').trim() : '';

    if (isTotalizadorRow(row, favCol ? row[favCol] : undefined, rawOb)) {
      continue;
    }

    const cleanOb = rawOb.toUpperCase().replace(/[\s\.\-\/]/g, '');
    if (!cleanOb || cleanOb === '-' || cleanOb === '0') continue;

    let rawVal = valCol ? row[valCol] : 0;
    if (
      isFonteString(rawVal) ||
      (fonteCol && rawVal === row[fonteCol]) ||
      (typeof rawVal === 'number' && (rawVal === 600 || rawVal === 500 || rawVal === 1500 || rawVal === 1600 || rawVal === 0.6 || rawVal === 0.5))
    ) {
      rawVal = 0;
    }
    if (typeof rawVal === 'string' && (/\b\d{3}\s+\d{4,5}[-\s]|\b\d{3}\s+\d{5}/.test(rawVal) || /^\d{4,}[-\s]\w$/.test(rawVal))) {
      rawVal = 0;
    }
    let valor = typeof rawVal === 'number' ? rawVal : parseCurrencyValue(rawVal);
    if (valor >= 1e11 || isFonteString(valor) || valor === 600 || valor === 500 || valor === 1500 || valor === 1600 || valor === 0.6 || valor === 0.5) {
      // Re-scan other columns to find true monetary value if rawVal was a fonte code or empty
      let realVal = 0;
      for (const [k, v] of Object.entries(row)) {
        if (!v || isFonteHeader(k) || k === fonteCol || k === valCol || isFonteString(v)) continue;
        const lk = k.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
        if (lk.includes('valor') || lk.includes('pago') || lk.includes('liquid') || lk.includes('total')) {
          const pv = parseCurrencyValue(v);
          if (pv > 0 && pv < 1e11 && !isFonteString(v) && pv !== 600 && pv !== 500 && pv !== 1500 && pv !== 1600 && pv !== 0.6 && pv !== 0.5) {
            realVal = pv;
            break;
          }
        }
      }
      valor = realVal;
    }

    const rawSit = sitCol ? String(row[sitCol] || '').trim().toUpperCase() : 'CB';

    // Prioritize payment date column, then general date column, then find the most recent valid date in row
    let rawData = paymentDateCol ? row[paymentDateCol] : (generalDateCol ? row[generalDateCol] : '');
    let dateInfo = parseDateToMonthInfo(rawData);

    for (const [k, v] of Object.entries(row)) {
      if (!v) continue;
      const lk = k.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
      if (lk.includes('data') || lk.includes('dt') || lk.includes('emiss') || lk.includes('pagamento') || lk.includes('ob')) {
        const cand = parseDateToMonthInfo(v);
        if (cand.dateIso) {
          if (!dateInfo.dateIso || cand.dateIso > dateInfo.dateIso) {
            dateInfo = cand;
            rawData = v;
          }
        }
      }
    }

    let rawFonte = fonteCol ? String(row[fonteCol] || '').trim() : '';
    let rawFav = favCol ? String(row[favCol] || '').trim() : '';
    let finalCnpj = cnpjCol ? String(row[cnpjCol] || '').trim() : '';
    let ppVal = ppCol ? String(row[ppCol] || '').trim() : '';

    // Only if 2-row block structure and current row lacks Favorecido or Fonte, check next row
    const isSingleRowComplete = (rawFav && rawFav.length >= 3) && (rawFonte && rawFonte !== '-' && rawFonte !== '0');
    if (!isSingleRowComplete) {
      const nextRow = idx + 1 < file.rows.length ? file.rows[idx + 1] : null;
      const nextRowHasOb = nextRow && obCol && isObNumber(nextRow[obCol]);
      if (!nextRowHasOb && nextRow) {
        if (!rawFonte || rawFonte === '-' || rawFonte === '0' || rawFonte === 'OUTRAS') {
          const nFonte = fonteCol ? String(nextRow[fonteCol] || '').trim() : '';
          if (nFonte && nFonte !== '-' && nFonte !== '0') {
            rawFonte = nFonte;
          }
        }
        if (!rawFav || rawFav.length < 3) {
          const nFav = favCol ? String(nextRow[favCol] || '').trim() : '';
          if (nFav && nFav.length >= 3) {
            rawFav = nFav;
          }
        }
        if (!ppVal && ppCol && nextRow[ppCol]) {
          ppVal = String(nextRow[ppCol]).trim();
        }
        idx++; // Consume row 2 of the block
      }
    }

    const { cnpj: parsedCnpj, name: parsedName } = parseFavorecidoField(rawFav);
    finalCnpj = finalCnpj || parsedCnpj || 'N/I';
    const finalName = parsedName || rawFav || 'NÃO INFORMADO';

    items.push({
      id: `${file.id}_ob_${idx}_${cleanOb}`,
      ob: cleanOb,
      dataReferencia: dateInfo.dateIso || String(rawData || '').trim(),
      dataIso: dateInfo.dateIso,
      mesAnoKey: dateInfo.mesAnoKey,
      mesAnoLabel: dateInfo.mesAnoLabel,
      valor,
      situacao: rawSit || 'CB',
      pp: ppVal,
      fonteRecurso: (rawFonte || 'OUTRAS').trim(),
      favorecidoRaw: rawFav,
      favorecidoCnpj: finalCnpj,
      favorecidoName: finalName,
      linhaOrigem: idx + 1,
    });
  }

  return items;
}

/**
 * Universal parser for "Listar Ordem Bancária" using strict Regex text patterns and 2-row blocks.
 * Implements exact line 1 (OB, Data, Situação, Valor) and line 2 (PP, Fonte, Favorecido) reading.
 */
export function parseListarOrdemBancariaMatrix(matrix: any[][]): OrdemBancariaItem[] {
  if (!matrix || matrix.length < 1) return [];

  const items: OrdemBancariaItem[] = [];

  for (let i = 0; i < matrix.length; i++) {
    const row1 = matrix[i];
    if (!Array.isArray(row1)) continue;

    // Verifique se row1 contém uma string que case com o Regex de OB (/\b\d{4}OB\d+\b/i). Se NÃO contiver, faça continue (pule a linha).
    let ob = '';
    for (let c = 0; c < row1.length; c++) {
      const cell = row1[c];
      if (cell !== null && cell !== undefined) {
        const s = String(cell).trim();
        const m = s.match(/\b\d{4}OB\d+\b/i);
        if (m) {
          ob = m[0].toUpperCase();
          break;
        }
      }
    }
    if (!ob) continue;

    // Se achou a OB, sabemos que estamos na "Linha 1". Extraia da row1:
    // OB: A string que deu match no regex \d{4}OB\d+. (já em `ob`)

    // Data: A string que der match no regex \d{2}\/\d{2}\/\d{4}.
    let dataReferencia = '';
    for (let c = 0; c < row1.length; c++) {
      const cell = row1[c];
      if (cell !== null && cell !== undefined) {
        const s = String(cell).trim();
        const m = s.match(/\b\d{2}\/\d{2}\/\d{4}\b/);
        if (m) {
          dataReferencia = m[0];
          break;
        }
      }
    }
    const dateInfo = parseDateToMonthInfo(dataReferencia);

    // Situação: A string que der match em \b(CB|RJ|PPCB|PPRJ|AO|C|PAGA)\b. Se não achar, assuma 'CB'.
    let situacao = 'CB';
    for (let c = 0; c < row1.length; c++) {
      const cell = row1[c];
      if (cell !== null && cell !== undefined) {
        const s = String(cell).trim().toUpperCase();
        const m = s.match(/\b(CB|RJ|PPCB|PPNC|PPRJ|AO|C|PAGA|LIQUIDADA)\b/);
        if (m) {
          situacao = m[1];
          break;
        }
      }
    }

    // Valor: Filtre a row1 buscando o maior valor numérico válido com parseCurrencyValue().
    let valor = 0;
    for (let c = 0; c < row1.length; c++) {
      const cell = row1[c];
      if (cell === null || cell === undefined) continue;
      const s = String(cell).trim();
      if (!s) continue;
      // Rejeita células que sejam a própria OB ou Data
      if (/\b\d{4}OB\d+\b/i.test(s)) continue;
      if (/\b\d{2}\/\d{2}\/\d{4}\b/.test(s)) continue;
      // Rejeita domicílio bancário (ex: 001 03795-8 000011655-6) para evitar parsing indevido
      if (/\b\d{3}\s+\d{4,5}[-\s]|\b\d{3}\s+\d{5}/.test(s) || /^\d{4,}[-\s]\w$/.test(s)) continue;
      // Rejeita expressamente códigos de Fonte de Recursos para não confundir com o valor monetário da despesa
      if (isFonteString(s) || isFonteString(cell) || isFonte600OuDerivada(s)) continue;

      if (typeof cell === 'number') {
        if (cell === 600 || cell === 500 || cell === 1500 || cell === 1600 || cell === 0.6 || cell === 0.5) continue;
        if (cell > 0 && cell < 1e11 && cell > valor) {
          valor = cell;
        }
        continue;
      }
      const valParsed = parseCurrencyValue(cell);
      if (valParsed === 600 || valParsed === 500 || valParsed === 1500 || valParsed === 1600 || valParsed === 0.6 || valParsed === 0.5) continue;
      if (valParsed > 0 && valParsed < 1e11 && valParsed > valor) {
        valor = valParsed;
      }
    }

    // Verifique fonteRecurso, favorecidoRaw e pp já na row1
    let pp = '';
    for (let c = 0; c < row1.length; c++) {
      const cell = row1[c];
      if (cell !== null && cell !== undefined) {
        const s = String(cell).trim();
        const m = s.match(/\b\d{4}PP\d+\b/i);
        if (m) {
          pp = m[0].toUpperCase();
          break;
        }
      }
    }

    let fonteRecurso = findFonteRecursoInRow(row1);
    let favorecidoRaw = '';

    for (let c = 0; c < row1.length; c++) {
      const cell = row1[c];
      if (cell === null || cell === undefined) continue;
      const s = String(cell).trim();
      if (/\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2}/.test(s)) {
        favorecidoRaw = s;
        break;
      }
    }
    if (!favorecidoRaw) {
      for (let c = 0; c < row1.length; c++) {
        const cell = row1[c];
        if (cell === null || cell === undefined) continue;
        const s = String(cell).trim();
        if (/\b\d{4}OB\d+\b/i.test(s) || /\b\d{4}PP\d+\b/i.test(s)) continue;
        if (/\b\d{2}\/\d{2}\/\d{4}\b/.test(s)) continue;
        if (/\b(CB|RJ|PPCB|PPNC|PPRJ|AO|C|PAGA|LIQUIDADA)\b/i.test(s) && s.length <= 10) continue;
        if (/\b\d{3}\s+\d{4,5}[-\s]\w\b/i.test(s) || /\b\d{3}\s\d{5}\-\d\b/i.test(s)) continue;
        if (isValidFonteRecursoCode(s)) continue;
        if (s.length > 10 && /[A-Za-z]/.test(s)) {
          favorecidoRaw = s;
          break;
        }
      }
    }

    // Só avança para row2 (i++) se a row1 não tiver fonte/favorecido E a row2 não contiver outra OB
    const row2 = i + 1 < matrix.length ? matrix[i + 1] : null;
    const row2HasOB =
      Array.isArray(row2) &&
      row2.some((cell) => cell !== null && cell !== undefined && /\b\d{4}OB\d+\b/i.test(String(cell)));

    if (row2 && Array.isArray(row2) && (!fonteRecurso || !favorecidoRaw) && !row2HasOB) {
      if (!pp) {
        for (let c = 0; c < row2.length; c++) {
          const cell = row2[c];
          if (cell !== null && cell !== undefined) {
            const s = String(cell).trim();
            const m = s.match(/\b\d{4}PP\d+\b/i);
            if (m) {
              pp = m[0].toUpperCase();
              break;
            }
          }
        }
      }

      if (!fonteRecurso) {
        fonteRecurso = findFonteRecursoInRow(row2);
      }

      if (!favorecidoRaw) {
        for (let c = 0; c < row2.length; c++) {
          const cell = row2[c];
          if (cell === null || cell === undefined) continue;
          const s = String(cell).trim();
          if (/\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2}/.test(s)) {
            favorecidoRaw = s;
            break;
          }
        }
        if (!favorecidoRaw) {
          for (let c = 0; c < row2.length; c++) {
            const cell = row2[c];
            if (cell === null || cell === undefined) continue;
            const s = String(cell).trim();
            if (/\b\d{3}\s+\d{4,5}[-\s]\w\b/i.test(s) || /\b\d{3}\s\d{5}\-\d\b/i.test(s)) continue;
            if (/\b\d{4}PP\d+\b/i.test(s)) continue;
            if (isValidFonteRecursoCode(s)) continue;
            if (s.length > 10 && /[A-Za-z]/.test(s)) {
              favorecidoRaw = s;
              break;
            }
          }
        }
      }

      // Se valor não foi encontrado em row1, confere row2
      if (valor <= 0) {
        for (let c = 0; c < row2.length; c++) {
          const cell = row2[c];
          if (cell === null || cell === undefined) continue;
          const s = String(cell).trim();
          if (!s) continue;
          if (/\b\d{4}PP\d+\b/i.test(s)) continue;
          if (/\b\d{3}\s+\d{4,5}/.test(s) || /^\d{4,}[-\s]\w$/.test(s)) continue;
          if (/\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2}/.test(s)) continue;
          if (isValidFonteRecursoCode(s)) continue;
          if (typeof cell === 'number') {
            if (cell > 0 && cell < 1e11 && cell > valor) valor = cell;
            continue;
          }
          const valParsed = parseCurrencyValue(cell);
          if (valParsed > 0 && valParsed < 1e11 && valParsed > valor) valor = valParsed;
        }
      }

      // Avança i++ para pular a row2 consumida no bloco
      i++;
    }

    // Limpe as strings, valide se o valor > 0, e faça o .push() no array items.
    if (valor > 0) {
      const { cnpj: parsedCnpj, name: parsedName } = parseFavorecidoField(favorecidoRaw);
      const cnpjMatch = favorecidoRaw.match(/\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2}/);
      const finalCnpj = parsedCnpj || cnpjMatch?.[0] || 'N/I';
      const finalName = (parsedName || (cnpjMatch ? favorecidoRaw.replace(cnpjMatch[0], '').trim() : favorecidoRaw) || 'NÃO INFORMADO').trim();

      items.push({
        id: `ob_${ob}_${i}`,
        ob,
        dataReferencia: dateInfo.dateIso || dataReferencia,
        dataIso: dateInfo.dateIso,
        mesAnoKey: dateInfo.mesAnoKey,
        mesAnoLabel: dateInfo.mesAnoLabel,
        valor,
        situacao,
        pp,
        fonteRecurso: (fonteRecurso || 'Fonte Não Identificada').trim(),
        favorecidoRaw: favorecidoRaw.trim(),
        favorecidoCnpj: finalCnpj,
        favorecidoName: finalName,
        linhaOrigem: i + 1,
      });
    }
  }

  return items;
}

/**
 * Asynchronous chunked parser for large spreadsheets using the exact same strict regex logic.
 * Iterates in chunks of 1000 rows and pauses with setTimeout(0) between chunks
 * so the browser main thread breathes and never freezes or crashes.
 */
export async function parseListarOrdemBancariaMatrixAsync(
  matrix: any[][],
  onProgress?: (percent: number) => void
): Promise<OrdemBancariaItem[]> {
  if (!matrix || matrix.length < 1) return [];

  const items: OrdemBancariaItem[] = [];
  const total = matrix.length;
  const CHUNK_SIZE = 1000;

  for (let i = 0; i < total; i++) {
    if (i > 0 && i % CHUNK_SIZE === 0) {
      if (onProgress) onProgress(Math.min(99, Math.round((i / total) * 100)));
      await new Promise<void>((resolve) => setTimeout(resolve, 0));
    }

    const row1 = matrix[i];
    if (!Array.isArray(row1)) continue;

    let ob = '';
    for (let c = 0; c < row1.length; c++) {
      const cell = row1[c];
      if (cell !== null && cell !== undefined) {
        const s = String(cell).trim();
        const m = s.match(/\b\d{4}OB\d+\b/i);
        if (m) {
          ob = m[0].toUpperCase();
          break;
        }
      }
    }
    if (!ob) continue;

    let dataReferencia = '';
    for (let c = 0; c < row1.length; c++) {
      const cell = row1[c];
      if (cell !== null && cell !== undefined) {
        const s = String(cell).trim();
        const m = s.match(/\b\d{2}\/\d{2}\/\d{4}\b/);
        if (m) {
          dataReferencia = m[0];
          break;
        }
      }
    }
    const dateInfo = parseDateToMonthInfo(dataReferencia);

    let situacao = 'CB';
    for (let c = 0; c < row1.length; c++) {
      const cell = row1[c];
      if (cell !== null && cell !== undefined) {
        const s = String(cell).trim().toUpperCase();
        const m = s.match(/\b(CB|RJ|PPCB|PPNC|PPRJ|AO|C|PAGA|LIQUIDADA)\b/);
        if (m) {
          situacao = m[1];
          break;
        }
      }
    }

    let valor = 0;
    for (let c = 0; c < row1.length; c++) {
      const cell = row1[c];
      if (cell === null || cell === undefined) continue;
      const s = String(cell).trim();
      if (!s) continue;
      if (/\b\d{4}OB\d+\b/i.test(s)) continue;
      if (/\b\d{2}\/\d{2}\/\d{4}\b/.test(s)) continue;
      if (/\b\d{3}\s+\d{4,5}[-\s]|\b\d{3}\s+\d{5}/.test(s) || /^\d{4,}[-\s]\w$/.test(s)) continue;
      // Rejeita expressamente códigos de Fonte de Recursos para não confundir com o valor monetário da despesa
      if (isFonteString(s) || isFonteString(cell) || isFonte600OuDerivada(s)) continue;

      if (typeof cell === 'number') {
        if (cell === 600 || cell === 500 || cell === 1500 || cell === 1600 || cell === 0.6 || cell === 0.5) continue;
        if (cell > 0 && cell < 1e11 && cell > valor) {
          valor = cell;
        }
        continue;
      }
      const valParsed = parseCurrencyValue(cell);
      if (valParsed === 600 || valParsed === 500 || valParsed === 1500 || valParsed === 1600 || valParsed === 0.6 || valParsed === 0.5) continue;
      if (valParsed > 0 && valParsed < 1e11 && valParsed > valor) {
        valor = valParsed;
      }
    }

    // Verifique fonteRecurso, favorecidoRaw e pp já na row1
    let pp = '';
    for (let c = 0; c < row1.length; c++) {
      const cell = row1[c];
      if (cell !== null && cell !== undefined) {
        const s = String(cell).trim();
        const m = s.match(/\b\d{4}PP\d+\b/i);
        if (m) {
          pp = m[0].toUpperCase();
          break;
        }
      }
    }

    let fonteRecurso = findFonteRecursoInRow(row1);
    let favorecidoRaw = '';

    for (let c = 0; c < row1.length; c++) {
      const cell = row1[c];
      if (cell !== null && cell !== undefined) {
        const s = String(cell).trim();
        if (/\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2}/.test(s)) {
          favorecidoRaw = s;
          break;
        }
      }
    }
    if (!favorecidoRaw) {
      for (let c = 0; c < row1.length; c++) {
        const cell = row1[c];
        if (cell === null || cell === undefined) continue;
        const s = String(cell).trim();
        if (/\b\d{4}OB\d+\b/i.test(s) || /\b\d{4}PP\d+\b/i.test(s)) continue;
        if (/\b\d{2}\/\d{2}\/\d{4}\b/.test(s)) continue;
        if (/\b(CB|RJ|PPCB|PPNC|PPRJ|AO|C|PAGA|LIQUIDADA)\b/i.test(s) && s.length <= 10) continue;
        if (/\b\d{3}\s+\d{4,5}[-\s]\w\b/i.test(s) || /\b\d{3}\s\d{5}\-\d\b/i.test(s)) continue;
        if (isValidFonteRecursoCode(s)) continue;
        if (s.length > 10 && /[A-Za-z]/.test(s)) {
          favorecidoRaw = s;
          break;
        }
      }
    }

    // Só avança para row2 (i++) se a row1 não tiver fonte/favorecido E a row2 não contiver outra OB
    const row2 = i + 1 < total ? matrix[i + 1] : null;
    const row2HasOB =
      Array.isArray(row2) &&
      row2.some((cell) => cell !== null && cell !== undefined && /\b\d{4}OB\d+\b/i.test(String(cell)));

    if (row2 && Array.isArray(row2) && (!fonteRecurso || !favorecidoRaw) && !row2HasOB) {
      if (!pp) {
        for (let c = 0; c < row2.length; c++) {
          const cell = row2[c];
          if (cell !== null && cell !== undefined) {
            const s = String(cell).trim();
            const m = s.match(/\b\d{4}PP\d+\b/i);
            if (m) {
              pp = m[0].toUpperCase();
              break;
            }
          }
        }
      }

      if (!fonteRecurso) {
        fonteRecurso = findFonteRecursoInRow(row2);
      }

      if (!favorecidoRaw) {
        for (let c = 0; c < row2.length; c++) {
          const cell = row2[c];
          if (cell !== null && cell !== undefined) {
            const s = String(cell).trim();
            if (/\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2}/.test(s)) {
              favorecidoRaw = s;
              break;
            }
          }
        }
        if (!favorecidoRaw) {
          for (let c = 0; c < row2.length; c++) {
            const cell = row2[c];
            if (cell === null || cell === undefined) continue;
            const s = String(cell).trim();
            if (/\b\d{3}\s+\d{4,5}[-\s]\w\b/i.test(s) || /\b\d{3}\s\d{5}\-\d\b/i.test(s)) continue;
            if (/\b\d{4}PP\d+\b/i.test(s)) continue;
            if (isValidFonteRecursoCode(s)) continue;
            if (s.length > 10 && /[A-Za-z]/.test(s)) {
              favorecidoRaw = s;
              break;
            }
          }
        }
      }

      if (valor <= 0) {
        for (let c = 0; c < row2.length; c++) {
          const cell = row2[c];
          if (cell === null || cell === undefined) continue;
          const s = String(cell).trim();
          if (!s) continue;
          if (/\b\d{4}PP\d+\b/i.test(s)) continue;
          if (/\b\d{3}\s+\d{4,5}/.test(s) || /^\d{4,}[-\s]\w$/.test(s)) continue;
          if (/\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2}/.test(s)) continue;
          if (isValidFonteRecursoCode(s)) continue;
          if (typeof cell === 'number') {
            if (cell > 0 && cell < 1e11 && cell > valor) valor = cell;
            continue;
          }
          const valParsed = parseCurrencyValue(cell);
          if (valParsed > 0 && valParsed < 1e11 && valParsed > valor) valor = valParsed;
        }
      }

      i++;
    }

    if (valor > 0) {
      const { cnpj: parsedCnpj, name: parsedName } = parseFavorecidoField(favorecidoRaw);
      const cnpjMatch = favorecidoRaw.match(/\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2}/);
      const finalCnpj = parsedCnpj || cnpjMatch?.[0] || 'N/I';
      const finalName = (parsedName || (cnpjMatch ? favorecidoRaw.replace(cnpjMatch[0], '').trim() : favorecidoRaw) || 'NÃO INFORMADO').trim();

      items.push({
        id: `ob_${ob}_${i}`,
        ob,
        dataReferencia: dateInfo.dateIso || dataReferencia,
        dataIso: dateInfo.dateIso,
        mesAnoKey: dateInfo.mesAnoKey,
        mesAnoLabel: dateInfo.mesAnoLabel,
        valor,
        situacao,
        pp,
        fonteRecurso: (fonteRecurso || 'Fonte Não Identificada').trim(),
        favorecidoRaw: favorecidoRaw.trim(),
        favorecidoCnpj: finalCnpj,
        favorecidoName: finalName,
        linhaOrigem: i + 1,
      });
    }
  }

  if (onProgress) onProgress(100);
  return items;
}

/**
 * Storage helpers for Ordens Bancárias
 * Uses IndexedDB as primary store to support large datasets without localStorage quota limits.
 * Synchronous in-memory caching and localStorage fallback for small metadata.
 */
const DB_NAME_OB = 'FINANCE_OB_STORE';
const DB_VERSION_OB = 1;
const STORE_NAME_OB = 'ordens_bancarias';
const STORAGE_KEY_OB_FILENAME = 'painel_financeiro_ordens_bancarias_filename_v1';
const STORAGE_KEY_OB_FALLBACK = 'painel_financeiro_ordens_bancarias_v1';

let memoryCachedOBItems: OrdemBancariaItem[] | null = null;
let memoryCachedOBFileName: string = '';

function openOBDatabase(): Promise<IDBDatabase> {
  const dbPromise = new Promise<IDBDatabase>((resolve, reject) => {
    try {
      if (typeof window === 'undefined' || !window.indexedDB) {
        reject(new Error('IndexedDB indisponível.'));
        return;
      }
      const request = indexedDB.open(DB_NAME_OB, DB_VERSION_OB);
      request.onupgradeneeded = (e) => {
        try {
          const db = (e.target as IDBOpenDBRequest).result;
          if (!db.objectStoreNames.contains(STORE_NAME_OB)) {
            db.createObjectStore(STORE_NAME_OB);
          }
        } catch (err) {
          reject(err);
        }
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error || new Error('Falha ao abrir IndexedDB OB'));
      request.onblocked = () => reject(new Error('IndexedDB OB bloqueado'));
    } catch (err) {
      reject(err);
    }
  });

  const timeoutPromise = new Promise<IDBDatabase>((_, reject) => {
    setTimeout(() => reject(new Error('Timeout ao conectar com IndexedDB OB')), 7000);
  });

  return Promise.race([dbPromise, timeoutPromise]);
}

async function saveOBToIndexedDB(items: OrdemBancariaItem[] | null, fileName: string): Promise<void> {
  try {
    const db = await openOBDatabase();
    const tx = db.transaction(STORE_NAME_OB, 'readwrite');
    const store = tx.objectStore(STORE_NAME_OB);
    if (!items || items.length === 0) {
      store.delete('current_dataset');
    } else {
      store.put({ items, fileName, updatedAt: new Date().toISOString() }, 'current_dataset');
    }
    await new Promise<void>((resolve, reject) => {
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch (err) {
    console.warn('Erro ao salvar Ordens Bancárias no IndexedDB:', err);
  }
}

async function loadOBFromIndexedDB(): Promise<{ items: OrdemBancariaItem[] | null; fileName: string } | null> {
  try {
    const db = await openOBDatabase();
    if (!db.objectStoreNames.contains(STORE_NAME_OB)) return null;
    const tx = db.transaction(STORE_NAME_OB, 'readonly');
    const store = tx.objectStore(STORE_NAME_OB);
    const result = await new Promise<any>((resolve, reject) => {
      const req = store.get('current_dataset');
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
    if (result && Array.isArray(result.items)) {
      return { items: result.items, fileName: result.fileName || '' };
    }
    return null;
  } catch (err) {
    console.warn('Erro ao carregar Ordens Bancárias do IndexedDB:', err);
    return null;
  }
}

export function saveOrdensBancariasToStorage(items: OrdemBancariaItem[] | null, fileName?: string): void {
  const finalFileName = fileName !== undefined ? fileName : memoryCachedOBFileName;
  memoryCachedOBItems = items;
  memoryCachedOBFileName = finalFileName;

  // Persist to IndexedDB asynchronously (handles large 10MB+ datasets easily)
  saveOBToIndexedDB(items, finalFileName);

  // Store file name in localStorage (small footprint)
  try {
    if (finalFileName) {
      localStorage.setItem(STORAGE_KEY_OB_FILENAME, finalFileName);
    } else {
      localStorage.removeItem(STORAGE_KEY_OB_FILENAME);
    }
  } catch (e) {
    // Ignore
  }

  // Only mirror items to localStorage if very small (< 200 items) to prevent quota exceeded error
  try {
    if (!items || items.length === 0) {
      localStorage.removeItem(STORAGE_KEY_OB_FALLBACK);
    } else if (items.length < 200) {
      const json = JSON.stringify(items);
      if (json.length < 500 * 1024) {
        localStorage.setItem(STORAGE_KEY_OB_FALLBACK, json);
      }
    } else {
      // Remove stale large data from localStorage so it doesn't take quota
      localStorage.removeItem(STORAGE_KEY_OB_FALLBACK);
    }
  } catch (err) {
    // localStorage quota reached - safely ignored since IndexedDB and memory already stored it
    try {
      localStorage.removeItem(STORAGE_KEY_OB_FALLBACK);
    } catch (_) {
      // Ignore
    }
  }

  if (typeof window !== 'undefined') {
    window.dispatchEvent(
      new CustomEvent('finance_ordens_bancarias_changed', { detail: { items, fileName: finalFileName } })
    );
  }
}

export function loadOrdensBancariasFromStorage(): { items: OrdemBancariaItem[] | null; fileName: string } {
  // If already in memory cache, return immediately
  if (memoryCachedOBItems) {
    return { items: memoryCachedOBItems, fileName: memoryCachedOBFileName };
  }

  // Otherwise check localStorage fallback (for small cached sets or file name)
  try {
    const fileName = localStorage.getItem(STORAGE_KEY_OB_FILENAME) || '';
    const raw = localStorage.getItem(STORAGE_KEY_OB_FALLBACK);
    if (raw) {
      const items = JSON.parse(raw);
      if (Array.isArray(items)) {
        memoryCachedOBItems = items;
        memoryCachedOBFileName = fileName;
        return { items, fileName };
      }
    }
    return { items: null, fileName };
  } catch (err) {
    return { items: null, fileName: '' };
  }
}

export async function loadOrdensBancariasAsync(): Promise<{ items: OrdemBancariaItem[] | null; fileName: string }> {
  try {
    // 1. Try IndexedDB first (most complete source)
    const idbData = await loadOBFromIndexedDB();
    if (idbData && idbData.items && idbData.items.length > 0) {
      memoryCachedOBItems = idbData.items;
      memoryCachedOBFileName = idbData.fileName;
      return idbData;
    }
  } catch (err) {
    console.warn('Erro ao carregar OBs do IndexedDB, usando fallback:', err);
  }

  // 2. Fallback to synchronous load
  return loadOrdensBancariasFromStorage();
}

// Aliases for compatibility
export const getStoredOrdensBancarias = () => {
  const res = loadOrdensBancariasFromStorage();
  return { items: res.items || [], fileName: res.fileName };
};
export const saveStoredOrdensBancarias = (items: OrdemBancariaItem[], fileName: string) =>
  saveOrdensBancariasToStorage(items, fileName);
export const clearStoredOrdensBancarias = () =>
  saveOrdensBancariasToStorage(null, '');

/**
 * Normalizes text for robust matching between OB record and ConsolidatedSupplier
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
 * Extracts digits from CNPJ / CPF
 */
function extractDigits(val: string): string {
  return (val || '').replace(/\D/g, '');
}

/**
 * Matches an OrdemBancariaItem to a ConsolidatedSupplier
 */
export function matchesSupplierOB(
  supplier: { name: string; cnpj: string },
  item: OrdemBancariaItem
): boolean {
  const suppDigits = extractDigits(supplier.cnpj);
  const obDigits = extractDigits(item.favorecidoCnpj);

  // 1. Direct CNPJ match if >= 11 digits
  if (suppDigits.length >= 11 && obDigits.length >= 11) {
    if (suppDigits === obDigits) return true;
  }

  // 2. CNPJ root match (first 8 digits) if both are company CNPJs
  if (suppDigits.length === 14 && obDigits.length === 14) {
    if (suppDigits.substring(0, 8) === obDigits.substring(0, 8)) {
      return true;
    }
  }

  // 3. Name match
  const suppNorm = normalizeText(supplier.name);
  const obNorm = normalizeText(item.favorecidoName);

  if (suppNorm && obNorm) {
    if (suppNorm === obNorm) return true;
    if (suppNorm.length >= 6 && obNorm.length >= 6) {
      if (suppNorm.includes(obNorm) || obNorm.includes(suppNorm)) return true;
    }
  }

  return false;
}

/**
 * Attaches or recalculates payment breakdown by Fonte de Recurso and by Mês
 * on ConsolidatedSupplier objects based on the uploaded Ordens Bancárias list.
 * Highly optimized with pre-indexed hash maps (O(N + M)) so it executes in ~2ms even for thousands of records.
 */
export function enrichSuppliersWithOrdensBancarias(
  suppliers: any[],
  obItems: OrdemBancariaItem[]
): void {
  if (!suppliers || suppliers.length === 0 || !obItems || obItems.length === 0) {
    return;
  }

  // 1. Build fast lookup maps for OB items:
  // - by exact CNPJ / CPF (digits only)
  // - by root CNPJ (8 digits)
  // - by normalized company name
  const obByCnpj = new Map<string, OrdemBancariaItem[]>();
  const obByCnpjRoot = new Map<string, OrdemBancariaItem[]>();
  const obByName = new Map<string, OrdemBancariaItem[]>();
  const obByNamePrefix = new Map<string, OrdemBancariaItem[]>();

  for (let i = 0; i < obItems.length; i++) {
    const item = obItems[i];
    const cnpjDigits = extractDigits(item.favorecidoCnpj);
    const cnpjRoot = cnpjDigits.length === 14 ? cnpjDigits.substring(0, 8) : '';
    const normName = normalizeText(item.favorecidoName);

    if (cnpjDigits.length >= 11) {
      const arr = obByCnpj.get(cnpjDigits);
      if (arr) arr.push(item);
      else obByCnpj.set(cnpjDigits, [item]);
    }

    if (cnpjRoot) {
      const arr = obByCnpjRoot.get(cnpjRoot);
      if (arr) arr.push(item);
      else obByCnpjRoot.set(cnpjRoot, [item]);
    }

    if (normName && normName.length >= 4) {
      const arr = obByName.get(normName);
      if (arr) arr.push(item);
      else obByName.set(normName, [item]);

      if (normName.length >= 8) {
        const prefix = normName.substring(0, 14);
        const pArr = obByNamePrefix.get(prefix);
        if (pArr) pArr.push(item);
        else obByNamePrefix.set(prefix, [item]);
      }
    }
  }

  // 2. Enrich each supplier in O(1) average lookup
  for (let sIdx = 0; sIdx < suppliers.length; sIdx++) {
    const s = suppliers[sIdx];
    const suppDigits = extractDigits(s.cnpj);
    const suppRoot = suppDigits.length === 14 ? suppDigits.substring(0, 8) : '';
    const suppNorm = normalizeText(s.name);

    let matchedItems: OrdemBancariaItem[] | undefined;

    // A. Direct full CNPJ match
    if (suppDigits.length >= 11 && obByCnpj.has(suppDigits)) {
      matchedItems = obByCnpj.get(suppDigits);
    }
    // B. Root CNPJ match (first 8 digits)
    else if (suppRoot && obByCnpjRoot.has(suppRoot)) {
      matchedItems = obByCnpjRoot.get(suppRoot);
    }
    // C. Exact normalized name match
    else if (suppNorm && obByName.has(suppNorm)) {
      matchedItems = obByName.get(suppNorm);
    }
    // D. Prefix name match O(1)
    else if (suppNorm && suppNorm.length >= 8) {
      const prefix = suppNorm.substring(0, 14);
      if (obByNamePrefix.has(prefix)) {
        matchedItems = obByNamePrefix.get(prefix);
      }
    }

    if (!matchedItems || matchedItems.length === 0) continue;

    let totalFromOB = 0;
    const porFonte: Record<string, number> = {};
    const porMes: Record<string, { total: number; count: number; porFonte: Record<string, number> }> = {};

    for (let m = 0; m < matchedItems.length; m++) {
      const it = matchedItems[m];
      totalFromOB += it.valor;
      const fonteKey = it.fonteRecurso || 'OUTRAS';
      porFonte[fonteKey] = (porFonte[fonteKey] || 0) + it.valor;

      const mesKey = it.mesAnoKey || 'sem_mes';
      let mesData = porMes[mesKey];
      if (!mesData) {
        mesData = { total: 0, count: 0, porFonte: {} };
        porMes[mesKey] = mesData;
      }
      mesData.total += it.valor;
      mesData.count += 1;
      mesData.porFonte[fonteKey] = (mesData.porFonte[fonteKey] || 0) + it.valor;
    }

    s.pagamentosDetalhePorFonte = porFonte;
    s.pagamentosPorMes = porMes;

    // Set or adjust valorPagoAno
    if (!s.valorPagoAno || s.valorPagoAno === 0) {
      s.valorPagoAno = totalFromOB;
      s.countPagoAno = matchedItems.length;
    } else if (totalFromOB > s.valorPagoAno) {
      s.valorPagoAno = totalFromOB;
      s.countPagoAno = Math.max(s.countPagoAno || 0, matchedItems.length);
    }
  }
}

