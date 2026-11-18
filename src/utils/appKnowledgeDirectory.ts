/**
 * App Knowledge Directory
 * Central cross-dataset intelligence service that indexes and auto-resolves
 * supplier information, CNPJs, NL numbers, and Fonte de Recursos across
 * all imported spreadsheets, bank orders, and publications.
 */
import { FileData } from '../types';
import { OrdemBancariaItem } from './ordemBancariaParser';
import { parseCurrencyValue, parseFavorecidoField } from './excelParser';

export interface KnownNlInfo {
  nl: string;
  fornecedor: string;
  cnpj: string;
  valor: number;
  fonteRecurso?: string;
  processo?: string;
  data?: string;
  stage?: string;
  sourceFile?: string;
}

export interface SupplierDirectoryEntry {
  canonicalName: string;
  cnpj: string;
  cnpjDigits: string;
  totalNls: number;
  fontes: Set<string>;
}

// In-memory directory indexes
const supplierNameToCnpj = new Map<string, string>();
const cnpjDigitsToSupplierName = new Map<string, string>();
const nlDirectory = new Map<string, KnownNlInfo>();
const supplierDirectory = new Map<string, SupplierDirectoryEntry>();

/**
 * Normalizes supplier name for robust matching (removes accents, punctuation, extra spaces)
 */
export function normalizeSupplierKey(name: string | undefined): string {
  if (!name) return '';
  return name
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Formats clean CNPJ digits into standard mask XX.XXX.XXX/XXXX-XX
 */
export function formatCnpj(rawCnpj: string | undefined): string {
  if (!rawCnpj) return '';
  const digits = rawCnpj.replace(/\D/g, '');
  if (digits.length === 14) {
    return `${digits.slice(0, 2)}.${digits.slice(2, 5)}.${digits.slice(5, 8)}/${digits.slice(8, 12)}-${digits.slice(12, 14)}`;
  }
  if (digits.length === 11) {
    return `${digits.slice(0, 3)}.${digits.slice(3, 6)}.${digits.slice(6, 9)}-${digits.slice(9, 11)}`;
  }
  return rawCnpj.trim();
}

/**
 * Determines whether a string is a clean, actual NL (and not a line number)
 */
export function isValidNL(doc: string | undefined): boolean {
  if (!doc) return false;
  const s = String(doc).trim();
  if (/^linha\s*\d+$/i.test(s)) return false;
  if (/^#?\d{1,4}$/.test(s)) return false; // Pure small integers like 1, 2, 14 are row lines, not NLs
  if (/\b(?:\d{2,4})?NL[\s\-_]*\d+\b/i.test(s)) return true;
  if (/\b\d+NL\d+\b/i.test(s)) return true;
  if (/^nl\s*[:#]?\s*[A-Z0-9]+/i.test(s)) return true;
  return false;
}

/**
 * Normalizes and cleans NL document identifier without row numbers or prefix noise
 */
export function formatCleanNL(doc: string | undefined): string {
  if (!doc) return 'NL Pendente';
  const clean = String(doc).trim();
  if (/^linha\s*\d+$/i.test(clean) || /^#?\d{1,4}$/.test(clean)) return 'NL Pendente';

  const match4 = clean.match(/\b\d{4}NL\d+\b/i);
  if (match4) return match4[0].toUpperCase();

  const matchGeneric = clean.match(/\b\d+NL\d+\b/i);
  if (matchGeneric) return matchGeneric[0].toUpperCase();

  const matchPrefix = clean.match(/^nl\s*[:#]?\s*([A-Z0-9]+)/i);
  if (matchPrefix) return `NL ${matchPrefix[1].toUpperCase()}`;

  const matchWord = clean.match(/\bNL[\s\-_]*\d+\b/i);
  if (matchWord) return matchWord[0].replace(/[\s\-_]+/g, '').toUpperCase();

  return clean;
}

/**
 * Extracts the real NL from a spreadsheet row by inspecting all cells,
 * rejecting row index numbers or line numbers.
 */
export function extractRealNlFromRow(
  row: Record<string, any>,
  preferredCol?: string
): string {
  if (!row) return '';

  // 1. Check preferred column if it has a real NL
  if (preferredCol && row[preferredCol]) {
    const val = String(row[preferredCol]).trim();
    if (isValidNL(val)) {
      return formatCleanNL(val);
    }
  }

  // 2. Scan all cells for explicit NL regex
  for (const [, v] of Object.entries(row)) {
    if (!v) continue;
    const str = String(v).trim();
    if (isValidNL(str)) {
      return formatCleanNL(str);
    }
  }

  // 3. Fallback: check liquidation/document columns, but strictly reject row lines
  for (const [k, v] of Object.entries(row)) {
    if (!v) continue;
    const str = String(v).trim();
    if (!str || str === '-' || str === '0') continue;
    if (/^linha\s*\d+$/i.test(str) || /^#?\d{1,4}$/.test(str)) continue;

    const lk = k.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    if (
      lk.includes('liquidacao') ||
      lk.includes('liquidac') ||
      lk.includes('nota de liquidacao') ||
      lk.includes('nota lancamento') ||
      lk === 'documento' ||
      lk === 'doc'
    ) {
      if (!/NE/i.test(str)) {
        return formatCleanNL(str);
      }
    }
  }

  return '';
}

/**
 * Indexes all files and data to form the application's central knowledge base
 */
export function indexAppKnowledge(
  files: FileData[],
  ordensBancariasList?: OrdemBancariaItem[] | null
): void {
  if (!files || files.length === 0) return;

  // 1. Index files: Obedece, Não Obedece, Base de PPs, Publicação, Outros
  // (ordem_bancaria é indexada no passo 2 via ordensBancariasList; natureza_despesa não contém NLs)
  files.forEach((file) => {
    const stage = file.stage || 'outros';
    if (stage === 'ordem_bancaria' || stage === 'natureza_despesa') return;
    const favCol = file.detectedFavorecidoCol;
    const valCol = file.detectedValorCol;
    const numCol = file.detectedNumeroCol || file.detectedPpCol;
    const cnpjCol = file.detectedCnpjCol;
    const fonteCol = file.detectedFonteCol;
    const procCol = file.detectedProcessoCol;
    const dataCol = file.detectedDataCol;

    file.rows.forEach((row) => {
      // Extract supplier name and CNPJ
      let rawFav = favCol && row[favCol] ? String(row[favCol]).trim() : '';
      let rawCnpj = cnpjCol && row[cnpjCol] ? String(row[cnpjCol]).trim() : '';

      // Fallback search across row cells for CNPJ if not mapped
      if (!rawCnpj) {
        for (const k in row) {
          const v = row[k];
          if (!v) continue;
          const s = String(v).trim();
          const cnpjMatch = s.match(/\b\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}\b/) || s.match(/\b\d{14}\b/);
          if (cnpjMatch) {
            rawCnpj = cnpjMatch[0];
            break;
          }
        }
      }

      const parsed = parseFavorecidoField(rawFav);
      const supplierName = parsed.name || rawFav;
      const effectiveCnpj = (rawCnpj && rawCnpj !== 'N/I' && rawCnpj !== '-') ? rawCnpj : (parsed.cnpj && parsed.cnpj !== 'N/I' ? parsed.cnpj : '');
      const cnpjDigits = effectiveCnpj.replace(/\D/g, '');

      // Register Supplier <-> CNPJ mapping
      if (supplierName && cnpjDigits.length >= 11) {
        const normKey = normalizeSupplierKey(supplierName);
        const formatted = formatCnpj(effectiveCnpj);
        if (normKey && formatted) {
          supplierNameToCnpj.set(normKey, formatted);
          cnpjDigitsToSupplierName.set(cnpjDigits, supplierName);

          const cur = supplierDirectory.get(normKey) || {
            canonicalName: supplierName,
            cnpj: formatted,
            cnpjDigits,
            totalNls: 0,
            fontes: new Set<string>(),
          };
          supplierDirectory.set(normKey, cur);
        }
      }

      // Extract real NL
      const realNl = extractRealNlFromRow(row, numCol);
      const val = valCol ? parseCurrencyValue(row[valCol]) : 0;

      // Extract Fonte de Recurso
      let fonte = '';
      if (fonteCol && row[fonteCol]) {
        fonte = String(row[fonteCol]).trim();
      }
      if (!fonte) {
        for (const k in row) {
          const v = row[k];
          if (!v) continue;
          const s = String(v).trim();
          if (/\b0\s*[.,]\s*600\b/i.test(s) || /600/i.test(s)) {
            fonte = '0.600';
            break;
          }
          if (/\b0\s*[.,]\s*500\b/i.test(s) || /500/i.test(s)) {
            fonte = '0.500';
            break;
          }
        }
      }
      if (fonte) {
        if (/600/.test(fonte)) fonte = '0.600';
        else if (/500/.test(fonte)) fonte = '0.500';
      }

      if (realNl && realNl !== 'NL Pendente') {
        const cleanNlKey = realNl.toUpperCase().replace(/[^A-Z0-9]/g, '');
        const entry: KnownNlInfo = {
          nl: realNl,
          fornecedor: supplierName,
          cnpj: effectiveCnpj ? formatCnpj(effectiveCnpj) : '',
          valor: val,
          fonteRecurso: fonte || undefined,
          processo: procCol && row[procCol] ? String(row[procCol]).trim() : undefined,
          data: dataCol && row[dataCol] ? String(row[dataCol]).trim() : undefined,
          stage,
          sourceFile: file.fileName,
        };

        nlDirectory.set(cleanNlKey, entry);

        // Also update supplier directory count and fontes
        if (supplierName) {
          const normKey = normalizeSupplierKey(supplierName);
          const sEntry = supplierDirectory.get(normKey);
          if (sEntry) {
            sEntry.totalNls += 1;
            if (fonte) sEntry.fontes.add(fonte);
          }
        }
      }
    });
  });

  // 2. Index Ordens Bancárias
  if (ordensBancariasList && ordensBancariasList.length > 0) {
    ordensBancariasList.forEach((ob) => {
      const fav = ob.favorecidoName || ob.favorecidoRaw || '';
      const cnpj = ob.favorecidoCnpj || '';
      const cnpjDigits = cnpj.replace(/\D/g, '');

      if (fav && cnpjDigits.length >= 11) {
        const normKey = normalizeSupplierKey(fav);
        const formatted = formatCnpj(cnpj);
        if (normKey && formatted) {
          if (!supplierNameToCnpj.has(normKey)) {
            supplierNameToCnpj.set(normKey, formatted);
          }
          if (!cnpjDigitsToSupplierName.has(cnpjDigits)) {
            cnpjDigitsToSupplierName.set(cnpjDigits, fav);
          }
        }
      }
    });
  }
}

/**
 * Resolves CNPJ automatically:
 * 1. Returns existing if valid
 * 2. Looks up by NL number
 * 3. Looks up by supplier name (exact or substring match)
 */
export function resolveCnpj(
  supplierName?: string,
  nlNumero?: string,
  existingCnpj?: string
): string {
  // If already valid
  if (existingCnpj) {
    const digits = existingCnpj.replace(/\D/g, '');
    if (digits.length >= 11) {
      return formatCnpj(existingCnpj);
    }
  }

  // 1. Try look up by NL
  if (nlNumero) {
    const cleanNlKey = nlNumero.toUpperCase().replace(/[^A-Z0-9]/g, '');
    const nlInfo = nlDirectory.get(cleanNlKey);
    if (nlInfo && nlInfo.cnpj) {
      return nlInfo.cnpj;
    }
  }

  // 2. Try look up by Supplier Name
  if (supplierName) {
    const normKey = normalizeSupplierKey(supplierName);
    if (normKey) {
      if (supplierNameToCnpj.has(normKey)) {
        return supplierNameToCnpj.get(normKey)!;
      }

      // High-confidence partial/fuzzy match (e.g. if one has LTDA and other doesn't)
      for (const [key, cnpj] of supplierNameToCnpj.entries()) {
        if (normKey.length > 8 && key.length > 8) {
          if (normKey.includes(key) || key.includes(normKey)) {
            return cnpj;
          }
        }
      }
    }
  }

  return '';
}

/**
 * Resolves Supplier Name automatically:
 * 1. Returns existing if valid
 * 2. Looks up by CNPJ digits
 * 3. Looks up by NL number
 */
export function resolveSupplierName(
  cnpj?: string,
  nlNumero?: string,
  existingName?: string
): string {
  if (existingName && existingName.trim().length > 2 && existingName !== 'N/I' && existingName !== '-') {
    return existingName.trim();
  }

  if (cnpj) {
    const digits = cnpj.replace(/\D/g, '');
    if (cnpjDigitsToSupplierName.has(digits)) {
      return cnpjDigitsToSupplierName.get(digits)!;
    }
  }

  if (nlNumero) {
    const cleanNlKey = nlNumero.toUpperCase().replace(/[^A-Z0-9]/g, '');
    const nlInfo = nlDirectory.get(cleanNlKey);
    if (nlInfo && nlInfo.fornecedor) {
      return nlInfo.fornecedor;
    }
  }

  return existingName || '';
}

/**
 * Resolves Fonte de Recurso automatically from NL or Supplier history
 */
export function resolveFonteRecurso(
  nlNumero?: string,
  supplierName?: string,
  existingFonte?: string
): string {
  if (existingFonte && (existingFonte.includes('600') || existingFonte.includes('500') || existingFonte.length >= 3)) {
    return existingFonte.includes('600') ? '0.600' : existingFonte.includes('500') ? '0.500' : existingFonte;
  }

  if (nlNumero) {
    const cleanNlKey = nlNumero.toUpperCase().replace(/[^A-Z0-9]/g, '');
    const nlInfo = nlDirectory.get(cleanNlKey);
    if (nlInfo && nlInfo.fonteRecurso) {
      return nlInfo.fonteRecurso;
    }
  }

  if (supplierName) {
    const normKey = normalizeSupplierKey(supplierName);
    const sEntry = supplierDirectory.get(normKey);
    if (sEntry && sEntry.fontes.size > 0) {
      return Array.from(sEntry.fontes)[0];
    }
  }

  return '0.600'; // Default Federal
}

/**
 * Enriches any liquidation or publication item with missing details from the global knowledge base
 */
export function enrichItemWithKnowledge<
  T extends {
    fornecedor?: string;
    supplierName?: string;
    cnpj?: string;
    nlNumero?: string;
    documentNumber?: string;
    fonteRecurso?: string;
  }
>(item: T): T {
  const name = item.fornecedor || item.supplierName || '';
  const rawNl = item.nlNumero || item.documentNumber || '';
  const realNl = isValidNL(rawNl) ? formatCleanNL(rawNl) : (rawNl && !/^linha\s*\d+$/i.test(rawNl) && !/^#?\d{1,4}$/.test(rawNl) ? rawNl : '');
  const cleanCnpj = resolveCnpj(name, realNl, item.cnpj);
  const cleanName = resolveSupplierName(cleanCnpj, realNl, name);
  const cleanFonte = resolveFonteRecurso(realNl, cleanName, item.fonteRecurso);

  return {
    ...item,
    fornecedor: cleanName || item.fornecedor,
    supplierName: cleanName || item.supplierName,
    cnpj: cleanCnpj || item.cnpj || '',
    nlNumero: realNl || item.nlNumero,
    documentNumber: realNl || item.documentNumber,
    fonteRecurso: cleanFonte || item.fonteRecurso,
  };
}
