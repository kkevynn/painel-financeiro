import { FileData } from '../types';
import { parseCurrencyValue, parseFavorecidoField } from './excelParser';
import { OrdemBancariaItem, findFonteRecursoInRow, isFonte600OuDerivada } from './ordemBancariaParser';
import {
  extractRealNlFromRow,
  resolveCnpj,
  resolveSupplierName,
  formatCleanNL,
  isValidNL,
} from './appKnowledgeDirectory';

export interface NlSuggestionItem {
  nlNumero: string;
  cleanDoc: string;
  fornecedor: string;
  cnpj: string;
  valor: number;
  fonteRecurso: string;
  fonteTipo: 'federal' | 'estadual' | 'outro';
  fonteLabel: string;
  stage: 'obedece' | 'nao_obedece';
  fileName: string;
  suggestedPP?: string;
  suggestedOB?: string;
}

export interface PpSuggestionItem {
  ppNumero: string;
  cleanNum: string;
  obNumero?: string;
  valor?: number;
  fornecedor?: string;
  cnpj?: string;
  situacao?: string;
  fonteRecurso?: string;
  source: string;
}

export interface ObSuggestionItem {
  obNumero: string;
  cleanNum: string;
  ppNumero?: string;
  valor?: number;
  fornecedor?: string;
  cnpj?: string;
  situacao?: string;
  fonteRecurso?: string;
  source: string;
}

const STORAGE_KEY_CUSTOM_PPS = 'sesap_custom_pps_user_registry_v1';
const STORAGE_KEY_CUSTOM_OBS = 'sesap_custom_obs_user_registry_v1';

/**
 * Normalizes document number for comparison (removes non-alphanumeric, uppercase)
 */
export function cleanDoc(val?: any): string {
  if (!val) return '';
  return String(val).toUpperCase().replace(/[^A-Z0-9]/g, '');
}

/**
 * Analyzes and normalizes Fonte de Recursos to 0.600 (Federal), 0.500 (Estadual), or custom
 */
export function normalizeFonteRecurso(rawFonte: any): {
  code: string;
  tipo: 'federal' | 'estadual' | 'outro';
  label: string;
} {
  if (!rawFonte) {
    return { code: '0.600', tipo: 'federal', label: '0.600 (Federal - SUS)' };
  }
  const s = String(rawFonte).trim();
  const upper = s.toUpperCase();
  const digits = s.replace(/\D/g, '');

  if (
    isFonte600OuDerivada(s) ||
    upper.includes('0.600') ||
    upper.includes('0.6.00') ||
    upper.includes('600') ||
    digits.startsWith('0600') ||
    digits.endsWith('000600') ||
    upper.includes('SUS') ||
    upper.includes('FNS') ||
    upper.includes('FEDERAL')
  ) {
    return { code: '0.600', tipo: 'federal', label: '0.600 (Federal - SUS)' };
  }

  if (
    upper.includes('0.500') ||
    upper.includes('0.5.00') ||
    upper.includes('500') ||
    digits.startsWith('0500') ||
    digits.endsWith('000500') ||
    upper.includes('FES') ||
    upper.includes('TESOURO') ||
    upper.includes('ESTADUAL')
  ) {
    return { code: '0.500', tipo: 'estadual', label: '0.500 (Estadual - FES)' };
  }

  if (s && s !== '-' && s !== '0') {
    return { code: s, tipo: 'outro', label: `Fonte ${s}` };
  }

  return { code: '0.600', tipo: 'federal', label: '0.600 (Federal - SUS)' };
}

/**
 * Extracts all unique NL records from Obedece and Não Obedece files,
 * capturing the exact associated Fonte de Recurso (0.600 Federal / 0.500 Estadual)
 */
export function extractAvailableNls(files: FileData[]): NlSuggestionItem[] {
  if (!files || files.length === 0) return [];

  const resultsMap = new Map<string, NlSuggestionItem>();

  files.forEach((file) => {
    const stage = file.stage || 'outros';
    const isLiquidation =
      stage === 'obedece' ||
      stage === 'nao_obedece' ||
      (file.fileName && file.fileName.toLowerCase().includes('liquida')) ||
      (file.fileName && file.fileName.toLowerCase().includes('obedece'));

    if (!isLiquidation) return;

    const favCol = file.detectedFavorecidoCol;
    const valCol = file.detectedValorCol;
    const numCol = file.detectedNumeroCol || file.detectedPpCol;
    const fonteCol = file.detectedFonteCol;
    const obCol = file.detectedObCol;

    file.rows.forEach((row) => {
      const val = valCol ? parseCurrencyValue(row[valCol]) : 0;
      const rawFav = favCol ? row[favCol] : '';
      const { name, cnpj: parsedCnpj } = parseFavorecidoField(rawFav);

      // Find real NL document (rejecting row numbers)
      const docNum = extractRealNlFromRow(row, numCol);
      if (!docNum || !isValidNL(docNum)) return;

      const effectiveCnpj = resolveCnpj(name, docNum, parsedCnpj);
      const effectiveName = resolveSupplierName(name, effectiveCnpj);

      // Extract Fonte de Recursos from this row
      let rawFonte = fonteCol ? String(row[fonteCol] || '').trim() : '';
      if (!rawFonte || rawFonte === '-' || rawFonte === '0') {
        rawFonte = findFonteRecursoInRow(Object.values(row));
      }
      const normFonte = normalizeFonteRecurso(rawFonte);

      // Check if there is an OB or PP in this row
      let detectedOB = obCol && row[obCol] ? String(row[obCol]).trim() : '';
      if (!detectedOB) {
        for (const v of Object.values(row)) {
          if (!v) continue;
          const m = String(v).match(/\b\d{4}OB\d+\b/i);
          if (m) {
            detectedOB = m[0];
            break;
          }
        }
      }

      const cleanKey = cleanDoc(docNum);
      if (!cleanKey) return;

      if (!resultsMap.has(cleanKey)) {
        resultsMap.set(cleanKey, {
          nlNumero: docNum,
          cleanDoc: cleanKey,
          fornecedor: effectiveName || name || 'Fornecedor',
          cnpj: effectiveCnpj || '',
          valor: val,
          fonteRecurso: normFonte.code,
          fonteTipo: normFonte.tipo,
          fonteLabel: normFonte.label,
          stage: stage === 'nao_obedece' ? 'nao_obedece' : 'obedece',
          fileName: file.fileName,
          suggestedOB: detectedOB || undefined,
        });
      } else {
        // Enrich existing entry if new data has better Fonte or value
        const existing = resultsMap.get(cleanKey)!;
        if (existing.fonteTipo === 'outro' && normFonte.tipo !== 'outro') {
          existing.fonteRecurso = normFonte.code;
          existing.fonteTipo = normFonte.tipo;
          existing.fonteLabel = normFonte.label;
        }
        if (val > 0 && existing.valor <= 0) {
          existing.valor = val;
        }
        if (detectedOB && !existing.suggestedOB) {
          existing.suggestedOB = detectedOB;
        }
      }
    });
  });

  return Array.from(resultsMap.values());
}

/**
 * Extracts all unique PP records across files and ordensBancariasList
 */
export function extractAvailablePPs(
  files: FileData[],
  ordensBancariasList?: OrdemBancariaItem[] | null
): PpSuggestionItem[] {
  const map = new Map<string, PpSuggestionItem>();

  // 1. From ordensBancariasList
  if (ordensBancariasList && ordensBancariasList.length > 0) {
    ordensBancariasList.forEach((it) => {
      if (it.pp && it.pp.trim() && it.pp !== '-' && it.pp !== '0') {
        const clean = cleanDoc(it.pp);
        if (clean && !map.has(clean)) {
          map.set(clean, {
            ppNumero: it.pp.trim(),
            cleanNum: clean,
            obNumero: it.ob || undefined,
            valor: it.valor,
            fornecedor: it.favorecidoName,
            cnpj: it.favorecidoCnpj,
            situacao: it.situacao || 'CB',
            fonteRecurso: it.fonteRecurso,
            source: 'Ordem Bancária',
          });
        }
      }
    });
  }

  // 2. From all uploaded files (base_pps, pp_prontas, etc.)
  (files || []).forEach((file) => {
    const numCol = file.detectedPpCol || file.detectedNumeroCol;
    const obCol = file.detectedObCol;
    const valCol = file.detectedValorCol;
    const favCol = file.detectedFavorecidoCol;
    const sitCol = file.detectedSituacaoCol;
    const fonteCol = file.detectedFonteCol;

    file.rows.forEach((row) => {
      let rawPP = numCol ? String(row[numCol] || '').trim() : '';
      if (!rawPP || rawPP === '-' || rawPP === '0' || rawPP.toUpperCase().includes('NE')) {
        for (const [k, v] of Object.entries(row)) {
          if (!v) continue;
          const lk = k.toLowerCase();
          if (lk.includes('pp') || lk.includes('pagamento') || lk.includes('prep')) {
            const s = String(v).trim();
            if (s && s !== '-' && s !== '0' && !s.toUpperCase().includes('NE')) {
              rawPP = s;
              break;
            }
          }
        }
      }

      if (!rawPP || rawPP === '-' || rawPP === '0' || rawPP.toUpperCase().includes('NE')) return;

      const clean = cleanDoc(rawPP);
      if (!clean) return;

      const rawOB = obCol ? String(row[obCol] || '').trim() : '';
      const val = valCol ? parseCurrencyValue(row[valCol]) : 0;
      const { name, cnpj } = parseFavorecidoField(favCol ? row[favCol] : '');
      const effCnpj = resolveCnpj(name, rawPP, cnpj);
      const effName = resolveSupplierName(name, effCnpj);
      const sit = sitCol ? String(row[sitCol] || '').trim().toUpperCase() : 'AO';
      const fonte = fonteCol ? String(row[fonteCol] || '').trim() : '';

      if (!map.has(clean)) {
        map.set(clean, {
          ppNumero: rawPP,
          cleanNum: clean,
          obNumero: rawOB || undefined,
          valor: val || undefined,
          fornecedor: effName || name || undefined,
          cnpj: effCnpj || cnpj || undefined,
          situacao: sit,
          fonteRecurso: fonte || undefined,
          source: file.fileName,
        });
      } else {
        const existing = map.get(clean)!;
        if (rawOB && !existing.obNumero) existing.obNumero = rawOB;
        if (val > 0 && !existing.valor) existing.valor = val;
        if ((effName || name) && !existing.fornecedor) existing.fornecedor = effName || name;
        if ((effCnpj || cnpj) && !existing.cnpj) existing.cnpj = effCnpj || cnpj;
      }
    });
  });

  // 3. User-custom registered PPs from localStorage
  try {
    if (typeof localStorage !== 'undefined') {
      const raw = localStorage.getItem(STORAGE_KEY_CUSTOM_PPS);
      if (raw) {
        const customItems: PpSuggestionItem[] = JSON.parse(raw);
        customItems.forEach((ci) => {
          if (ci.cleanNum && !map.has(ci.cleanNum)) {
            map.set(ci.cleanNum, ci);
          }
        });
      }
    }
  } catch {}

  return Array.from(map.values());
}

/**
 * Extracts all unique OB records across files and ordensBancariasList
 */
export function extractAvailableOBs(
  files: FileData[],
  ordensBancariasList?: OrdemBancariaItem[] | null
): ObSuggestionItem[] {
  const map = new Map<string, ObSuggestionItem>();

  // 1. From ordensBancariasList
  if (ordensBancariasList && ordensBancariasList.length > 0) {
    ordensBancariasList.forEach((it) => {
      if (it.ob && it.ob.trim() && it.ob !== '-' && it.ob !== '0') {
        const clean = cleanDoc(it.ob);
        if (clean && !map.has(clean)) {
          map.set(clean, {
            obNumero: it.ob.trim(),
            cleanNum: clean,
            ppNumero: it.pp || undefined,
            valor: it.valor,
            fornecedor: it.favorecidoName,
            cnpj: it.favorecidoCnpj,
            situacao: it.situacao || 'CB',
            fonteRecurso: it.fonteRecurso,
            source: 'Ordem Bancária',
          });
        }
      }
    });
  }

  // 2. From all uploaded files with OB columns
  (files || []).forEach((file) => {
    const obCol = file.detectedObCol;
    const numCol = file.detectedPpCol || file.detectedNumeroCol;
    const valCol = file.detectedValorCol;
    const favCol = file.detectedFavorecidoCol;
    const sitCol = file.detectedSituacaoCol;
    const fonteCol = file.detectedFonteCol;

    file.rows.forEach((row) => {
      let rawOB = obCol ? String(row[obCol] || '').trim() : '';
      if (!rawOB || rawOB === '-' || rawOB === '0') {
        for (const [k, v] of Object.entries(row)) {
          if (!v) continue;
          const lk = k.toLowerCase();
          if (lk.includes('ob') || lk.includes('ordem bancaria') || lk.includes('banco')) {
            const s = String(v).trim();
            if (s && s !== '-' && s !== '0' && /\b\d{4}OB\d+\b/i.test(s)) {
              rawOB = s;
              break;
            }
          }
        }
      }

      if (!rawOB || rawOB === '-' || rawOB === '0') return;

      const clean = cleanDoc(rawOB);
      if (!clean) return;

      const rawPP = numCol ? String(row[numCol] || '').trim() : '';
      const val = valCol ? parseCurrencyValue(row[valCol]) : 0;
      const { name, cnpj } = parseFavorecidoField(favCol ? row[favCol] : '');
      const effCnpj = resolveCnpj(name, rawOB, cnpj);
      const effName = resolveSupplierName(name, effCnpj);
      const sit = sitCol ? String(row[sitCol] || '').trim().toUpperCase() : 'CB';
      const fonte = fonteCol ? String(row[fonteCol] || '').trim() : '';

      if (!map.has(clean)) {
        map.set(clean, {
          obNumero: rawOB,
          cleanNum: clean,
          ppNumero: rawPP || undefined,
          valor: val || undefined,
          fornecedor: effName || name || undefined,
          cnpj: effCnpj || cnpj || undefined,
          situacao: sit,
          fonteRecurso: fonte || undefined,
          source: file.fileName,
        });
      } else {
        const existing = map.get(clean)!;
        if (rawPP && !existing.ppNumero) existing.ppNumero = rawPP;
        if (val > 0 && !existing.valor) existing.valor = val;
        if ((effName || name) && !existing.fornecedor) existing.fornecedor = effName || name;
        if ((effCnpj || cnpj) && !existing.cnpj) existing.cnpj = effCnpj || cnpj;
      }
    });
  });

  // 3. User-custom registered OBs from localStorage
  try {
    if (typeof localStorage !== 'undefined') {
      const raw = localStorage.getItem(STORAGE_KEY_CUSTOM_OBS);
      if (raw) {
        const customItems: ObSuggestionItem[] = JSON.parse(raw);
        customItems.forEach((ci) => {
          if (ci.cleanNum && !map.has(ci.cleanNum)) {
            map.set(ci.cleanNum, ci);
          }
        });
      }
    }
  } catch {}

  return Array.from(map.values());
}

/**
 * Saves a user custom-entered PP so it is remembered and suggested across the app
 */
export function recordCustomPP(item: {
  ppNumero: string;
  obNumero?: string;
  valor?: number;
  fornecedor?: string;
  cnpj?: string;
  situacao?: string;
}) {
  const clean = cleanDoc(item.ppNumero);
  if (!clean) return;

  try {
    if (typeof localStorage !== 'undefined') {
      const raw = localStorage.getItem(STORAGE_KEY_CUSTOM_PPS);
      const list: PpSuggestionItem[] = raw ? JSON.parse(raw) : [];
      const idx = list.findIndex((p) => p.cleanNum === clean);
      const entry: PpSuggestionItem = {
        ppNumero: item.ppNumero.trim(),
        cleanNum: clean,
        obNumero: item.obNumero?.trim() || undefined,
        valor: item.valor,
        fornecedor: item.fornecedor,
        cnpj: item.cnpj,
        situacao: item.situacao || 'AO',
        source: 'Adicionado Manualmente',
      };

      if (idx >= 0) {
        list[idx] = { ...list[idx], ...entry };
      } else {
        list.unshift(entry);
      }
      localStorage.setItem(STORAGE_KEY_CUSTOM_PPS, JSON.stringify(list.slice(0, 200)));
    }
  } catch {}
}

/**
 * Saves a user custom-entered OB so it is remembered and suggested across the app
 */
export function recordCustomOB(item: {
  obNumero: string;
  ppNumero?: string;
  valor?: number;
  fornecedor?: string;
  cnpj?: string;
  situacao?: string;
}) {
  const clean = cleanDoc(item.obNumero);
  if (!clean) return;

  try {
    if (typeof localStorage !== 'undefined') {
      const raw = localStorage.getItem(STORAGE_KEY_CUSTOM_OBS);
      const list: ObSuggestionItem[] = raw ? JSON.parse(raw) : [];
      const idx = list.findIndex((o) => o.cleanNum === clean);
      const entry: ObSuggestionItem = {
        obNumero: item.obNumero.trim(),
        cleanNum: clean,
        ppNumero: item.ppNumero?.trim() || undefined,
        valor: item.valor,
        fornecedor: item.fornecedor,
        cnpj: item.cnpj,
        situacao: item.situacao || 'CB',
        source: 'Adicionado Manualmente',
      };

      if (idx >= 0) {
        list[idx] = { ...list[idx], ...entry };
      } else {
        list.unshift(entry);
      }
      localStorage.setItem(STORAGE_KEY_CUSTOM_OBS, JSON.stringify(list.slice(0, 200)));
    }
  } catch {}
}

/**
 * Searches and filters NL suggestions as the user types
 */
export function searchNls(query: string, items: NlSuggestionItem[], limit = 8): NlSuggestionItem[] {
  const qClean = cleanDoc(query);
  const qNorm = (query || '').toLowerCase().trim();
  if (!qNorm) return items.slice(0, limit);

  return items
    .filter((it) => {
      if (it.cleanDoc.includes(qClean)) return true;
      if (it.nlNumero.toLowerCase().includes(qNorm)) return true;
      if (it.fornecedor.toLowerCase().includes(qNorm)) return true;
      if (it.cnpj && it.cnpj.replace(/\D/g, '').includes(qClean)) return true;
      return false;
    })
    .slice(0, limit);
}

/**
 * Searches and filters PP suggestions as the user types
 */
export function searchPPs(query: string, items: PpSuggestionItem[], limit = 8): PpSuggestionItem[] {
  const qClean = cleanDoc(query);
  const qNorm = (query || '').toLowerCase().trim();
  if (!qNorm) return items.slice(0, limit);

  return items
    .filter((it) => {
      if (it.cleanNum.includes(qClean)) return true;
      if (it.ppNumero.toLowerCase().includes(qNorm)) return true;
      if (it.obNumero && it.obNumero.toLowerCase().includes(qNorm)) return true;
      if (it.fornecedor && it.fornecedor.toLowerCase().includes(qNorm)) return true;
      if (it.cnpj && it.cnpj.replace(/\D/g, '').includes(qClean)) return true;
      return false;
    })
    .slice(0, limit);
}

/**
 * Searches and filters OB suggestions as the user types
 */
export function searchOBs(query: string, items: ObSuggestionItem[], limit = 8): ObSuggestionItem[] {
  const qClean = cleanDoc(query);
  const qNorm = (query || '').toLowerCase().trim();
  if (!qNorm) return items.slice(0, limit);

  return items
    .filter((it) => {
      if (it.cleanNum.includes(qClean)) return true;
      if (it.obNumero.toLowerCase().includes(qNorm)) return true;
      if (it.ppNumero && it.ppNumero.toLowerCase().includes(qNorm)) return true;
      if (it.fornecedor && it.fornecedor.toLowerCase().includes(qNorm)) return true;
      if (it.cnpj && it.cnpj.replace(/\D/g, '').includes(qClean)) return true;
      return false;
    })
    .slice(0, limit);
}
