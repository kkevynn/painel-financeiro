import * as XLSX from 'xlsx';
import { LiberacaoItem, LiberacaoCrossResult, SituacaoCategory } from '../types';
import { parseCurrencyValue, readSpreadsheetAsMatrix } from './excelParser';
import { isSituacaoAO, isSituacaoEnviadaBanco } from './quebraOrdemParser';

export { isSituacaoAO, isSituacaoEnviadaBanco, readSpreadsheetAsMatrix };

// Official mapping of situation acronyms
export const SITUACAO_SIGLAS: Record<string, string> = {
  C: 'Cancelada',
  AP: 'A pagar',
  PPRJ: 'Cancelada Banco',
  PPNC: 'Enviada a Banco - Não Confirmada Banco (PPNC)',
  REPP: 'Resolvida por outra PP',
  RENE: 'Resolvida por NE',
  PPCB: 'Enviada a Banco - Confirmada Banco (PPCB)',
  OB: 'Com OB sem situação',
  AO: 'PP Pronta (Associada 2 Ordenadores - AO)',
  LD: 'Liberada Arquivo Diário',
  LI: 'Liberada Pag. Imediato',
  II: 'Impressa Pag. Imediato',
  EI: 'Enviada Imediato',
  EN: 'Enviada Normal',
  CM: 'Confirmada Manual',
  CB: 'Confirmada Banco',
  RJ: 'Rejeitada Banco',
  CBNC: 'Conf. Banco e PP Não Conf.',
  SD: 'Solicitada Arq. Diário',
  SI: 'Solicitada Pag. Imediato',
};

// Categorize situation for visual styling (Badge classes)
export function getSituacaoCategory(sigla: string): SituacaoCategory {
  const s = (sigla || '').trim().toUpperCase();
  if (['AO'].includes(s) || isSituacaoAO(s)) {
    return 'pp_pronta';
  }
  if (['PPCB', 'PPNC'].includes(s) || isSituacaoEnviadaBanco(s).isEnviada) {
    return 'enviada_banco';
  }
  if (['CB', 'CM', 'LI', 'LD', 'II'].includes(s)) {
    return 'confirmada';
  }
  if (['C', 'RJ', 'PPRJ', 'CBNC'].includes(s)) {
    return 'rejeitada';
  }
  if (['AP', 'EI', 'EN', 'SD', 'SI', 'OB', 'REPP', 'RENE'].includes(s)) {
    return 'pendente';
  }
  return 'pendente';
}

// Badge color definitions based on category
export function getBadgeClassForCategory(cat: SituacaoCategory): string {
  switch (cat) {
    case 'pp_pronta':
      return 'bg-emerald-100 text-emerald-950 border-emerald-400 font-extrabold';
    case 'enviada_banco':
      return 'bg-blue-100 text-blue-950 border-blue-400 font-extrabold';
    case 'confirmada':
      return 'bg-emerald-50 text-emerald-800 border-emerald-300';
    case 'rejeitada':
      return 'bg-rose-100 text-rose-800 border-rose-300';
    case 'pendente':
      return 'bg-amber-100 text-amber-800 border-amber-300';
    case 'nao_localizada':
    default:
      return 'bg-slate-100 text-slate-600 border-slate-300';
  }
}

// Normalize search key (remove leading zeros, spaces, hyphens for flexible cross-referencing)
export function normalizeKey(val: any): string {
  if (val === null || val === undefined) return '';
  const str = String(val).trim().toUpperCase();
  // Remove non-alphanumeric except letters and numbers
  const cleaned = str.replace(/[^A-Z0-9]/gi, '');
  // Also provide non-leading-zero version for numeric codes (e.g. "000123" -> "123")
  return cleaned;
}

export interface ParsedDatabaseRecord {
  pp: string;
  ob: string;
  favorecido: string;
  notaEmpenho: string;
  favorecidoNE: string;
  valor: number;
  sigla: string;
  origemLinha: number;
  hasFullSigefFields?: boolean;
}

/**
 * Parse the main system report (Base de Dados)
 * Handles files where headers start further down (e.g. line 13)
 */
export function parseDatabaseReport(sheetData: any[][]): ParsedDatabaseRecord[] {
  if (!sheetData || sheetData.length === 0) return [];

  // 1. Locate the header row by searching for key header names
  let headerRowIndex = -1;
  let ppCol = -1;
  let obCol = -1;
  let favCol = -1;
  let neCol = -1;
  let favNeCol = -1;
  let valCol = -1;
  let siglaCol = -1;

  for (let r = 0; r < Math.min(sheetData.length, 35); r++) {
    const row = sheetData[r];
    if (!Array.isArray(row)) continue;

    const rowStrings = row.map((cell) =>
      String(cell || '')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toLowerCase()
        .trim()
    );

    const hasPP = rowStrings.some(
      (c) =>
        c === 'pp' ||
        c === 'numero' ||
        c === 'número' ||
        c === 'num' ||
        c === 'nº' ||
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

    if ((hasPP && hasFav) || (hasOB && hasVal) || (hasPP && hasVal) || (hasOB && hasFav)) {
      headerRowIndex = r;
      rowStrings.forEach((c, colIdx) => {
        if (
          (c === 'pp' ||
            c === 'numero' ||
            c === 'número' ||
            c === 'num' ||
            c === 'nº' ||
            c.includes('proposta') ||
            c.includes('programacao') ||
            c.includes('numero pp') ||
            c.includes('num pp') ||
            c.includes('n pp') ||
            c.includes('nr pp') ||
            c.includes('cod pp')) &&
          ppCol === -1
        ) {
          ppCol = colIdx;
        } else if (
          (c === 'ob' ||
            c.includes('ordem banc') ||
            c.includes('ordem de pag') ||
            c.includes('numero ob') ||
            c.includes('num ob') ||
            c.includes('n ob') ||
            c.includes('nr ob') ||
            c.includes('cod ob')) &&
          obCol === -1
        ) {
          obCol = colIdx;
        } else if (
          (c.includes('favorecido') ||
            c.includes('credor') ||
            c.includes('fornecedor') ||
            c.includes('benefici') ||
            c.includes('razao')) &&
          !c.includes('empenho') &&
          favCol === -1
        ) {
          favCol = colIdx;
        } else if (c.includes('empenho') && !c.includes('favorecido') && neCol === -1) {
          neCol = colIdx;
        } else if (c.includes('favorecido') && c.includes('empenho') && favNeCol === -1) {
          favNeCol = colIdx;
        } else if (
          (c.includes('valor') ||
            c.includes('vlr') ||
            c.includes('liquido') ||
            c.includes('pago') ||
            c.includes('total')) &&
          valCol === -1
        ) {
          valCol = colIdx;
        } else if (
          (c.includes('situacao') || c.includes('sigla') || c === 'sit' || c.includes('status') || c.includes('estado')) &&
          siglaCol === -1
        ) {
          siglaCol = colIdx;
        }
      });
      break;
    }
  }

  // Fallback defaults if header was not found
  if (headerRowIndex === -1) {
    headerRowIndex = 0;
    ppCol = 0;
    obCol = 1;
    favCol = 2;
    neCol = 3;
    favNeCol = 4;
    valCol = 5;
    siglaCol = 6;
  } else {
    if (siglaCol === -1) {
      if (valCol !== -1 && valCol + 1 < (sheetData[headerRowIndex]?.length || 10)) {
        siglaCol = valCol + 1;
      } else {
        siglaCol = 6;
      }
    }
    if (ppCol === -1) ppCol = 0;
    if (obCol === -1) obCol = 1;
    if (favCol === -1) favCol = 2;
    if (valCol === -1) valCol = 5;
  }

  const records: ParsedDatabaseRecord[] = [];

  for (let r = headerRowIndex + 1; r < sheetData.length; r++) {
    const row = sheetData[r];
    if (!row || !Array.isArray(row) || row.length === 0) continue;

    const rawPP = row[ppCol] !== undefined ? String(row[ppCol]).trim() : '';
    const rawOB = row[obCol] !== undefined ? String(row[obCol]).trim() : '';
    const rawFav = row[favCol] !== undefined ? String(row[favCol]).trim() : '';
    const rawNE = neCol !== -1 && row[neCol] !== undefined ? String(row[neCol]).trim() : '';
    const rawFavNE = favNeCol !== -1 && row[favNeCol] !== undefined ? String(row[favNeCol]).trim() : '';
    const rawVal = row[valCol];

    let rawSigla = row[siglaCol] !== undefined ? String(row[siglaCol]).trim() : '';
    if (!rawSigla && row.length > 6) {
      const lastCell = String(row[row.length - 1] || '').trim();
      if (lastCell.length >= 1 && lastCell.length <= 5) {
        rawSigla = lastCell;
      }
    }

    // Skip blank or total footer lines and administrative header/footer text
    if (!rawPP && !rawOB && !rawFav) continue;
    const combinedLower = (rawPP + ' ' + rawFav + ' ' + rawOB).toLowerCase();
    if (
      combinedLower.includes('total') ||
      combinedLower.includes('subtotal') ||
      combinedLower.includes('página') ||
      combinedLower.includes('pagina') ||
      combinedLower.includes('emitido em') ||
      combinedLower.includes('relatório') ||
      combinedLower.includes('relatorio')
    ) {
      continue;
    }

    const parsedVal = parseCurrencyValue(rawVal);

    const hasAllSix = Boolean(
      rawPP &&
      rawOB &&
      rawFav &&
      rawNE &&
      rawFavNE &&
      parsedVal > 0
    );

    records.push({
      pp: rawPP,
      ob: rawOB,
      favorecido: rawFav || rawFavNE || 'NÃO INFORMADO',
      notaEmpenho: rawNE,
      favorecidoNE: rawFavNE,
      valor: parsedVal,
      sigla: rawSigla.toUpperCase(),
      origemLinha: r + 1,
      hasFullSigefFields: hasAllSix,
    });
  }

  return records;
}

/**
 * Parse the Search Spreadsheet from Leadership (Planilha de Pesquisa)
 */
export function parsePesquisaSpreadsheet(sheetData: any[][]): { key: string; rawRow: any; line: number }[] {
  if (!sheetData || sheetData.length === 0) return [];

  // Find header row or key column
  let headerRowIndex = -1;
  let keyColIndex = 0;

  for (let r = 0; r < Math.min(sheetData.length, 15); r++) {
    const row = sheetData[r];
    if (!Array.isArray(row)) continue;
    const rowLower = row.map((c) =>
      String(c || '')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toLowerCase()
        .trim()
    );

    const ppIdx = rowLower.findIndex(
      (c) =>
        c === 'pp' ||
        c.includes('proposta') ||
        c.includes('numero pp') ||
        c.includes('num pp') ||
        c.includes('n pp') ||
        c.includes('nr pp')
    );
    const obIdx = rowLower.findIndex(
      (c) =>
        c === 'ob' ||
        c.includes('ordem banc') ||
        c.includes('ordem de pag') ||
        c.includes('numero ob') ||
        c.includes('num ob') ||
        c.includes('n ob') ||
        c.includes('nr ob')
    );

    if (obIdx !== -1) {
      headerRowIndex = r;
      keyColIndex = obIdx;
      break;
    }
    if (ppIdx !== -1) {
      headerRowIndex = r;
      keyColIndex = ppIdx;
      break;
    }
  }

  // If no header found with keywords, check if row 0 has strings or numbers
  let startRow = 0;
  if (headerRowIndex === -1) {
    // Check if row 0 looks like an actual code (contains digits)
    const firstCell = String(sheetData[0]?.[0] || '').trim();
    if (/\d/.test(firstCell)) {
      startRow = 0; // row 0 is already a data item!
    } else {
      startRow = 1; // row 0 was likely a header label
    }
    keyColIndex = 0;
  } else {
    startRow = headerRowIndex + 1;
  }

  const searchItems: { key: string; rawRow: any; line: number }[] = [];

  for (let r = startRow; r < sheetData.length; r++) {
    const row = sheetData[r];
    if (!row || !Array.isArray(row) || row.length === 0) continue;

    let key = '';
    if (row[keyColIndex] !== undefined && String(row[keyColIndex]).trim()) {
      key = String(row[keyColIndex]).trim();
    } else {
      const firstValid = row.find((c) => c !== null && c !== undefined && String(c).trim().length > 0);
      if (firstValid) key = String(firstValid).trim();
    }

    if (key && !key.toLowerCase().includes('total')) {
      searchItems.push({
        key,
        rawRow: row,
        line: r + 1,
      });
    }
  }

  return searchItems;
}

/**
 * Cross-reference the Leadership search list against the System Database
 */
export function crossReferenceLiberacao(
  dbRecords: ParsedDatabaseRecord[],
  searchItems: { key: string; rawRow: any; line: number }[],
  dbFileName?: string,
  pesquisaFileName?: string
): LiberacaoCrossResult {
  // Build lookup index by normalized PP and normalized OB
  const dbByPP = new Map<string, ParsedDatabaseRecord>();
  const dbByOB = new Map<string, ParsedDatabaseRecord>();
  const dbByStrippedPP = new Map<string, ParsedDatabaseRecord>();
  const dbByStrippedOB = new Map<string, ParsedDatabaseRecord>();

  dbRecords.forEach((rec) => {
    const normPP = normalizeKey(rec.pp);
    const normOB = normalizeKey(rec.ob);

    if (normPP) {
      dbByPP.set(normPP, rec);
      // Strip leading zeros (e.g. 000456 -> 456)
      const stripped = normPP.replace(/^0+/, '');
      if (stripped) dbByStrippedPP.set(stripped, rec);
    }

    if (normOB) {
      dbByOB.set(normOB, rec);
      const stripped = normOB.replace(/^0+/, '');
      if (stripped) dbByStrippedOB.set(stripped, rec);
    }
  });

  const resultItems: LiberacaoItem[] = [];
  let valorTotalEncontrado = 0;
  let totalEncontrado = 0;
  let totalNaoEncontrado = 0;

  searchItems.forEach((item, index) => {
    const rawKey = item.key;
    const normKey = normalizeKey(rawKey);
    const strippedKey = normKey.replace(/^0+/, '');

    // Attempt to match by exact PP, exact OB, or stripped PP/OB
    let match: ParsedDatabaseRecord | undefined =
      dbByPP.get(normKey) ||
      dbByOB.get(normKey) ||
      (strippedKey ? dbByStrippedPP.get(strippedKey) : undefined) ||
      (strippedKey ? dbByStrippedOB.get(strippedKey) : undefined);

    // If still not matched, check if search key is contained in any PP/OB
    if (!match && normKey.length >= 4) {
      match = dbRecords.find((rec) => {
        const p = normalizeKey(rec.pp);
        const o = normalizeKey(rec.ob);
        return p.includes(normKey) || o.includes(normKey) || normKey.includes(p) || normKey.includes(o);
      });
    }

    if (match) {
      totalEncontrado++;
      valorTotalEncontrado += match.valor;

      const sigla = match.sigla || '';
      const situacaoTraduzida = SITUACAO_SIGLAS[sigla] || (sigla ? `Situação (${sigla})` : 'Sem Situação Informada');
      const categoria = getSituacaoCategory(sigla);

      resultItems.push({
        id: `lib_match_${index}_${match.origemLinha}`,
        pp: match.pp || rawKey,
        ob: match.ob || '-',
        fornecedor: match.favorecido,
        valor: match.valor,
        sigla: match.sigla,
        situacaoTraduzida,
        categoria,
        origemLinha: match.origemLinha,
        notaEmpenho: match.notaEmpenho,
        favorecidoNE: match.favorecidoNE,
        hasFullSigefFields: match.hasFullSigefFields,
        foundInDatabase: true,
        searchQueryKey: rawKey,
      });
    } else {
      totalNaoEncontrado++;
      resultItems.push({
        id: `lib_not_found_${index}`,
        pp: rawKey,
        ob: '-',
        fornecedor: 'NÃO LOCALIZADO NA BASE',
        valor: 0,
        sigla: 'N/L',
        situacaoTraduzida: 'Não Localizada na Base',
        categoria: 'nao_localizada',
        origemLinha: item.line,
        foundInDatabase: false,
        searchQueryKey: rawKey,
      });
    }
  });

  return {
    totalPesquisado: searchItems.length,
    totalEncontrado,
    totalNaoEncontrado,
    valorTotalEncontrado,
    items: resultItems,
    databaseFileName: dbFileName,
    pesquisaFileName: pesquisaFileName,
    processedAt: new Date(),
  };
}

/**
 * Generate Sample Demo Data for Instant Testing of Liberação Bancária
 */
export function generateSampleLiberacaoDemo(): {
  dbRecords: ParsedDatabaseRecord[];
  searchItems: { key: string; rawRow: any; line: number }[];
  result: LiberacaoCrossResult;
} {
  const sampleDatabase: ParsedDatabaseRecord[] = [
    {
      pp: '2026PP001420',
      ob: '2026OB008912',
      favorecido: 'MEDICAMENTOS & HOSPITALAR NORTE LTDA',
      notaEmpenho: '2026NE000412',
      favorecidoNE: 'MEDICAMENTOS & HOSPITALAR NORTE LTDA',
      valor: 345890.5,
      sigla: 'CB',
      origemLinha: 14,
    },
    {
      pp: '2026PP001421',
      ob: '2026OB008913',
      favorecido: 'CONSTRUTORA POTIGUAR ENGENHARIA S/A',
      notaEmpenho: '2026NE000984',
      favorecidoNE: 'CONSTRUTORA POTIGUAR ENGENHARIA S/A',
      valor: 1120450.0,
      sigla: 'LI',
      origemLinha: 15,
    },
    {
      pp: '2026PP001422',
      ob: '2026OB008914',
      favorecido: 'SERVICOS DE LIMPEZA E CONSERVACAO LTDA',
      notaEmpenho: '2026NE001102',
      favorecidoNE: 'SERVICOS DE LIMPEZA E CONSERVACAO LTDA',
      valor: 87430.2,
      sigla: 'EI',
      origemLinha: 16,
    },
    {
      pp: '2026PP001423',
      ob: '2026OB008915',
      favorecido: 'ALIMENTOS E NUTRICAO ESCOLAR LTDA',
      notaEmpenho: '2026NE000755',
      favorecidoNE: 'ALIMENTOS E NUTRICAO ESCOLAR LTDA',
      valor: 215300.0,
      sigla: 'LD',
      origemLinha: 17,
    },
    {
      pp: '2026PP001424',
      ob: '2026OB008916',
      favorecido: 'TECH SOLUCOES EM SOFTWARE PUBLICO',
      notaEmpenho: '2026NE000311',
      favorecidoNE: 'TECH SOLUCOES EM SOFTWARE PUBLICO',
      valor: 64200.0,
      sigla: 'RJ',
      origemLinha: 18,
    },
    {
      pp: '2026PP001425',
      ob: '2026OB008917',
      favorecido: 'POSTO E COMBUSTIVEIS LITORAL LTDA',
      notaEmpenho: '2026NE000620',
      favorecidoNE: 'POSTO E COMBUSTIVEIS LITORAL LTDA',
      valor: 145000.0,
      sigla: 'AP',
      origemLinha: 19,
    },
    {
      pp: '2026PP001426',
      ob: '2026OB008918',
      favorecido: 'SEGURANCA PATRIMONIAL VIGILANCIA LTDA',
      notaEmpenho: '2026NE000501',
      favorecidoNE: 'SEGURANCA PATRIMONIAL VIGILANCIA LTDA',
      valor: 198750.4,
      sigla: 'PPCB',
      origemLinha: 20,
    },
    {
      pp: '2026PP001427',
      ob: '2026OB008919',
      favorecido: 'GRAFICA E EDITORA ESTADUAL LTDA',
      notaEmpenho: '2026NE000889',
      favorecidoNE: 'GRAFICA E EDITORA ESTADUAL LTDA',
      valor: 32400.0,
      sigla: 'C',
      origemLinha: 21,
    },
  ];

  const sampleSearch: { key: string; rawRow: any; line: number }[] = [
    { key: '2026PP001420', rawRow: ['2026PP001420'], line: 2 },
    { key: '2026OB008913', rawRow: ['2026OB008913'], line: 3 },
    { key: '2026PP001422', rawRow: ['2026PP001422'], line: 4 },
    { key: '2026PP001424', rawRow: ['2026PP001424'], line: 5 },
    { key: '2026OB008918', rawRow: ['2026OB008918'], line: 6 },
    { key: '2026PP001427', rawRow: ['2026PP001427'], line: 7 },
    { key: '2026PP009999', rawRow: ['2026PP009999'], line: 8 }, // Not found demo
  ];

  const result = crossReferenceLiberacao(
    sampleDatabase,
    sampleSearch,
    'Relatorio_14082026101422.xls (Exemplo)',
    "OB's 13.08.2026_.xlsx (Exemplo)"
  );

  return {
    dbRecords: sampleDatabase,
    searchItems: sampleSearch,
    result,
  };
}
