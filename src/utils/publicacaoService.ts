import { FileData, PublicacaoItem, PublicacaoStatus } from '../types';
import { parseCurrencyValue, parseFavorecidoField, formatBRL } from './excelParser';
import { OrdemBancariaItem } from './ordemBancariaParser';
import { mlInferenceService } from './machineLearningInferenceService';
import {
  indexAppKnowledge,
  resolveCnpj,
  resolveSupplierName,
  resolveFonteRecurso,
  extractRealNlFromRow,
  formatCleanNL,
  isValidNL,
} from './appKnowledgeDirectory';

const RESOLVED_CACHE_STORAGE_KEY = 'sesap_resolved_publicacao_nls_v2';

/**
 * Normalizes document numbers (NL, PP, OB)
 */
export function cleanDocNum(s?: any): string {
  if (!s) return '';
  return String(s).trim().toUpperCase().replace(/[\s\.\-\/]/g, '');
}

/**
 * Extracts publication date string (e.g. "05/11", "06/11") from file name or raw cell date
 */
export function extractDiaPublicacao(fileName?: string, rawDate?: string): string {
  if (fileName) {
    const normalizedName = fileName
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '');
    const m = normalizedName.match(/(?:publicacao|publicar|doe|dia)?[^\d]*(\d{1,2}[\/\-\.]\d{1,2})/i);
    if (m && m[1]) {
      return m[1].replace(/[\-\.]/g, '/');
    }
  }
  if (rawDate) {
    const m = String(rawDate).match(/(\d{1,2}\/\d{1,2})/);
    if (m && m[1]) return m[1];
  }
  return '';
}

/**
 * Extracts indexed matching keys from an NL document string
 */
export function getNlMatchKeys(raw: any): string[] {
  if (!raw) return [];
  const str = String(raw).trim().toUpperCase();
  const keys: string[] = [];

  const cleanFull = str.replace(/[^A-Z0-9]/g, '');
  if (cleanFull && cleanFull.length >= 3) {
    keys.push(cleanFull);
  }

  // Look for NL pattern: e.g. 2024NL000123, 24NL000123, NL000123, NL 123
  const nlMatch = str.match(/(?:(\d{2,4})[\s\-\.]*)?NL[\s\-\.]*(\d+)/i);
  if (nlMatch) {
    const year = nlMatch[1] || '';
    const numPart = nlMatch[2];
    const numInt = parseInt(numPart, 10);

    if (year) {
      keys.push(`${year}NL${numPart}`);
      keys.push(`${year}NL${numInt}`);
    }
    keys.push(`NL${numPart}`);
    keys.push(`NL${numInt}`);
    if (numPart.length >= 3) {
      keys.push(`NUM_${numPart}`);
      keys.push(`NUM_${numInt}`);
    }
  } else {
    // If it's a numeric code like "000123"
    const numOnly = str.replace(/\D/g, '');
    if (numOnly.length >= 3) {
      keys.push(`NUM_${numOnly}`);
      keys.push(`NUM_${parseInt(numOnly, 10)}`);
    }
  }

  return Array.from(new Set(keys));
}

interface PaymentMatch {
  ppNumero: string;
  obNumero: string;
  situacao: string;
  valor: number;
  fornecedor: string;
  cnpj: string;
  notaEmpenho?: string;
  data?: string;
  source: string;
}

// In-memory cache of resolved NLs to satisfy rule: "se uma nl ja foi resolvida, nao é necessario processar ela novamente"
interface CachedResolution {
  equivalentePP?: string;
  equivalenteOB?: string;
  valorEquivalente?: number;
  situacaoEquivalente?: string;
  diferencaValor?: number;
  percentualRetencao?: number;
  valorNL?: number;
  fonteRecurso?: string;
  statusDetalhe: string;
}

const resolvedNlsMemoryCache = new Map<string, CachedResolution>();

function loadResolvedCacheFromStorage() {
  try {
    if (typeof localStorage !== 'undefined') {
      const raw = localStorage.getItem(RESOLVED_CACHE_STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        for (const [k, v] of Object.entries(parsed)) {
          resolvedNlsMemoryCache.set(k, v as CachedResolution);
        }
      }
    }
  } catch {
    // Ignore storage parse error
  }
}
loadResolvedCacheFromStorage();

function persistResolvedCacheToStorage() {
  try {
    if (typeof localStorage !== 'undefined') {
      const obj: Record<string, CachedResolution> = {};
      for (const [k, v] of resolvedNlsMemoryCache.entries()) {
        obj[k] = v;
      }
      localStorage.setItem(RESOLVED_CACHE_STORAGE_KEY, JSON.stringify(obj));
    }
  } catch {
    // Ignore storage save error
  }
}

/**
 * Manually links PP and OB to a publication NL, marking it as resolved,
 * saving the resolution persistently, and registering the tax retention pattern in AI/ML model.
 * Supports parcel split editing between Federal (0.600) and State (0.500) resources.
 */
export function saveManualPublicacaoResolution(params: {
  nlNumero: string;
  cnpj?: string;
  fornecedor?: string;
  valorNL?: number;
  ppNumero?: string;
  obNumero?: string;
  situacao?: string;
  valorPago?: number;
  fonteRecurso?: string;
}): void {
  const nlClean = (params.nlNumero || '').trim().toUpperCase();
  if (!nlClean) return;

  const valorNL = params.valorNL !== undefined && params.valorNL > 0 ? params.valorNL : 0;
  const valorPago = params.valorPago !== undefined && params.valorPago > 0 ? params.valorPago : valorNL;
  const dif = valorNL - valorPago;
  const retencao = dif > 0 && valorNL > 0 ? dif / valorNL : 0;
  const ppClean = (params.ppNumero || '').trim().toUpperCase();
  const obClean = (params.obNumero || '').trim().toUpperCase();
  const sitClean = (params.situacao || 'AO').trim().toUpperCase();
  const fonteClean = (params.fonteRecurso || '').trim();

  const fonteInfo = fonteClean ? ` [Fonte: ${fonteClean}]` : '';

  const resolution: CachedResolution = {
    equivalentePP: ppClean || undefined,
    equivalenteOB: obClean || undefined,
    valorEquivalente: valorPago,
    situacaoEquivalente: sitClean,
    diferencaValor: dif > 0 ? dif : 0,
    percentualRetencao: retencao,
    valorNL: valorNL > 0 ? valorNL : undefined,
    fonteRecurso: fonteClean || undefined,
    statusDetalhe: `Vinculado manualmente pelo usuário: PP ${ppClean || '-'} / OB ${obClean || '-'} (${sitClean})${fonteInfo}. Regularizado com sucesso.`,
  };

  const keys = getNlMatchKeys(nlClean);
  keys.forEach((k) => {
    resolvedNlsMemoryCache.set(k, resolution);
  });
  resolvedNlsMemoryCache.set(nlClean, resolution);

  persistResolvedCacheToStorage();

  // Register in AI pattern learning model
  if (params.fornecedor || params.cnpj) {
    mlInferenceService.recordManualPattern({
      credor: params.fornecedor || 'Fornecedor',
      cnpj: params.cnpj,
      valorNL,
      valorPP: valorPago,
      nlNumero: nlClean,
      ppNumero: ppClean,
      obNumero: obClean,
    });
  }
}

/**
 * Analyzes and cross-references Publicacao files with Obedece files, General PPs (SIGEF) and OBs.
 * 
 * Strict Business Rules:
 * 1. Matching MUST be deterministic based on the NL document number ("nao deve ser feito pesquisa de valores daqueles fornecedores").
 * 2. When publication is emitted in SIGEF, a PP and OB are issued in situation "AO" (Associada aos Ordenadores / Aguardando Ordem),
 *    NOT immediately "PPCB". Both "AO" and "PPCB"/"CB" confirm that the publication was processed and the PP/OB was generated.
 * 3. Verified from the General PP spreadsheet extracted from SIGEF (base_pps / pp_prontas) and Ordens Bancárias.
 * 4. Checks with extreme precision whether the NL is STILL in Obedece (using cell-by-cell multi-layer indexing to never miss an NL).
 * 5. Caches already resolved NLs: "se uma nl ja foi resolvida, nao é necessario processar ela novamente".
 * 6. Explicitly marks `pendenteEmObedece` for NLs that are listed in Obedece AND were sent to publication to resolve them directly.
 */
export function analyzePublicacaoRecords(
  files: FileData[],
  ordensBancariasList?: OrdemBancariaItem[] | null
): PublicacaoItem[] {
  if (!files || files.length === 0) return [];

  // 0. Build central knowledge directory across all datasets
  indexAppKnowledge(files, ordensBancariasList);

  // 1. Identify files for Publicacao
  const publicacaoFiles = files.filter((f) => f.stage === 'publicacao');
  if (publicacaoFiles.length === 0) return [];

  // 2. Identify ALL active NLs currently inside Obedece files
  // Supports all variants of naming ("Obedece - DD/MM", stage === 'obedece', etc.)
  const obedeceFiles = files.filter((f) => {
    if (f.stage === 'obedece') return true;
    const norm = (f.fileName || '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toLowerCase();
    return norm.includes('obedece') && !norm.includes('nao');
  });

  const activeObedeceNls = new Set<string>();
  const activeObedeceNlDocKeys = new Set<string>();
  const activeObedeceRowTokens = new Set<string>();

  obedeceFiles.forEach((file) => {
    // Detect dedicated NL column in this Obedece file
    const sampleRow = file.rows[0] || {};
    let detectedNlCol = '';
    for (const k of Object.keys(sampleRow)) {
      const lk = k.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
      if (lk.includes('empenho') || lk.includes(' ne ') || lk.endsWith(' ne')) continue;
      if (
        lk.includes('nl') ||
        lk.includes('liquidacao') ||
        lk.includes('liquidac') ||
        lk.includes('nota de liquidacao') ||
        lk.includes('nota liquidacao') ||
        lk === 'documento' ||
        lk === 'doc'
      ) {
        detectedNlCol = k;
        break;
      }
    }

    const numCol = detectedNlCol || file.detectedNumeroCol || file.detectedPpCol;
    const favCol = file.detectedFavorecidoCol;
    const valCol = file.detectedValorCol;

    file.rows.forEach((row) => {
      // 1. Check detected column and real extracted NL
      const docsInRow: string[] = [];
      const realExtracted = extractRealNlFromRow(row, numCol);
      if (realExtracted) {
        docsInRow.push(realExtracted);
      }
      if (numCol && row[numCol]) {
        const valStr = String(row[numCol]).trim();
        if (valStr && valStr !== '-' && valStr !== '0' && !/NE/i.test(valStr)) {
          docsInRow.push(valStr);
        }
      }

      // 2. Deep scan: check ALL cells in row for NL numbers and row tokens
      for (const [k, v] of Object.entries(row)) {
        if (!v) continue;
        const str = String(v).trim();
        if (!str || str === '-' || str === '0') continue;

        const cleanVal = cleanDocNum(str);
        if (cleanVal) {
          activeObedeceRowTokens.add(cleanVal);
        }

        const lk = k.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
        if (lk.includes('empenho') || lk.includes(' ne ')) continue;

        // Check if cell contains NL pattern
        const m = str.match(/\b(?:\d{2,4})?NL[\s\-_]*\d+\b/gi);
        if (m) {
          m.forEach((val) => docsInRow.push(val));
        }

        // If header suggests liquidation/document, include value
        if (
          lk.includes('nl') ||
          lk.includes('liquidacao') ||
          lk.includes('liquidac') ||
          lk === 'documento' ||
          lk === 'doc' ||
          lk === 'numero' ||
          lk === 'número'
        ) {
          if (!/NE/i.test(str)) {
            docsInRow.push(str);
          }
        }
      }

      // Index all found NL variations for this Obedece row
      docsInRow.forEach((rawDoc) => {
        const cleanDoc = cleanDocNum(rawDoc);
        if (cleanDoc) {
          activeObedeceNls.add(cleanDoc);
          getNlMatchKeys(rawDoc).forEach((k) => activeObedeceNls.add(k));

          const rawFav = favCol ? row[favCol] : '';
          const { cnpj } = parseFavorecidoField(rawFav);
          const cnpjDigits = (cnpj || '').replace(/\D/g, '');
          const val = valCol ? parseCurrencyValue(row[valCol]) : 0;
          const valCents = Math.round(val * 100);

          if (cnpjDigits) {
            activeObedeceNlDocKeys.add(`${cleanDoc}_${cnpjDigits}`);
            activeObedeceNlDocKeys.add(`${cleanDoc}_${valCents}`);
          }
        }
      });
    });
  });

  // 3. Build index of General PPs (from SIGEF report) and Ordens Bancárias
  const paymentByNlKey = new Map<string, PaymentMatch>();

  const registerPaymentMatch = (pm: PaymentMatch, nlCandidates: string[]) => {
    nlCandidates.forEach((rawKey) => {
      const keys = getNlMatchKeys(rawKey);
      keys.forEach((k) => {
        if (!paymentByNlKey.has(k)) {
          paymentByNlKey.set(k, pm);
        } else {
          const existing = paymentByNlKey.get(k)!;
          if ((!existing.obNumero && pm.obNumero) || (!existing.ppNumero && pm.ppNumero)) {
            paymentByNlKey.set(k, pm);
          }
        }
      });
    });
  };

  // 3a. Extract from Base de PPs (SIGEF General Report)
  files.forEach((file) => {
    const isPaymentFile =
      file.stage === 'base_pps' ||
      file.stage === 'pp_prontas' ||
      (file.stage === 'outros' && file.fileName && /relatorio|sigef|pagamento/i.test(file.fileName));

    if (!isPaymentFile) return;

    const favCol = file.detectedFavorecidoCol;
    const valCol = file.detectedValorCol;
    const numCol = file.detectedNumeroCol || file.detectedPpCol;
    const obCol = file.detectedObCol;
    const sitCol = file.detectedSituacaoCol || file.detectedSituacaoObCol;
    const neCol = file.detectedNotaEmpenhoCol;

    const colKeys = Object.keys(file.rows[0] || {});
    const nlCols = colKeys.filter((c) => {
      const lc = c.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
      return (
        lc.includes('nl') ||
        lc.includes('liquidacao') ||
        lc.includes('liquidac') ||
        lc.includes('doc orig') ||
        lc.includes('documento origem') ||
        lc.includes('doc. orig') ||
        lc.includes('origem')
      );
    });

    file.rows.forEach((row) => {
      const val = valCol ? parseCurrencyValue(row[valCol]) : 0;
      const rawFav = favCol ? row[favCol] : '';
      const { name, cnpj } = parseFavorecidoField(rawFav);
      const rawPP = numCol ? String(row[numCol] || '').trim() : '';
      const rawOB = obCol ? String(row[obCol] || '').trim() : '';
      let rawSit = sitCol ? String(row[sitCol] || '').trim().toUpperCase() : '';
      if (!rawSit || rawSit === '-' || rawSit === '0') {
        rawSit = 'AO'; // Newly emitted publication payment default in SIGEF
      }

      const nlFound: string[] = [];

      nlCols.forEach((col) => {
        const v = row[col];
        if (v && String(v).trim() && String(v).trim() !== '-' && String(v).trim() !== '0') {
          nlFound.push(String(v).trim());
        }
      });

      for (const [k, v] of Object.entries(row)) {
        if (!v) continue;
        const str = String(v).trim();
        const matches = str.match(/\b(?:\d{2,4})?NL[\s\-_]*\d+\b/gi);
        if (matches) {
          matches.forEach((m) => nlFound.push(m));
        }
      }

      if (/NL/i.test(rawPP)) {
        nlFound.push(rawPP);
      }

      if (nlFound.length > 0 && (rawPP || rawOB || val > 0)) {
        registerPaymentMatch(
          {
            ppNumero: rawPP,
            obNumero: rawOB,
            situacao: rawSit,
            valor: val,
            fornecedor: name,
            cnpj: cnpj !== 'N/I' ? cnpj : '',
            notaEmpenho: neCol ? String(row[neCol] || '') : undefined,
            source: file.fileName,
          },
          nlFound
        );
      }
    });
  });

  // 3b. Extract from ordensBancariasList
  if (ordensBancariasList && ordensBancariasList.length > 0) {
    ordensBancariasList.forEach((it) => {
      const nlFound: string[] = [];
      const itemAny = it as any;
      if (itemAny.processo) {
        const matches = String(itemAny.processo).match(/\b(?:\d{2,4})?NL[\s\-_]*\d+\b/gi);
        if (matches) matches.forEach((m: string) => nlFound.push(m));
      }
      if (itemAny.historico) {
        const matches = String(itemAny.historico).match(/\b(?:\d{2,4})?NL[\s\-_]*\d+\b/gi);
        if (matches) matches.forEach((m: string) => nlFound.push(m));
      }
      if (itemAny.documento && /NL/i.test(String(itemAny.documento))) {
        nlFound.push(String(itemAny.documento));
      }
      if (it.pp && /NL/i.test(it.pp)) {
        nlFound.push(it.pp);
      }

      if (nlFound.length > 0) {
        registerPaymentMatch(
          {
            ppNumero: it.pp || '',
            obNumero: it.ob || '',
            situacao: it.situacao || 'AO',
            valor: it.valor || 0,
            fornecedor: it.favorecidoName || '',
            cnpj: it.favorecidoCnpj || '',
            data: it.dataReferencia,
            source: 'Ordens Bancárias',
          },
          nlFound
        );
      }
    });
  }

  // 4. Process all rows in Publicacao files
  const results: PublicacaoItem[] = [];
  const processedDocKeys = new Set<string>();
  let hasNewResolutions = false;

  publicacaoFiles.forEach((file) => {
    const numCol = file.detectedNumeroCol || file.detectedPpCol;
    const favCol = file.detectedFavorecidoCol;
    const valCol = file.detectedValorCol;
    const procCol = file.detectedProcessoCol;
    const dataCol = file.detectedDataCol;

    const fileDia = extractDiaPublicacao(file.fileName);

    file.rows.forEach((row, rowIdx) => {
      // Extract NL document number from publication row (strictly rejecting line numbers)
      let rawDoc = extractRealNlFromRow(row, numCol);
      if (!rawDoc && numCol && row[numCol]) {
        const valStr = String(row[numCol]).trim();
        if (isValidNL(valStr)) rawDoc = formatCleanNL(valStr);
      }
      if (!rawDoc) {
        for (const [k, v] of Object.entries(row)) {
          const lk = k.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
          if (lk.includes('empenho') || lk.includes(' ne ')) continue;
          if (
            (lk.includes('documento') ||
              lk.includes('nl') ||
              lk.includes('nota de lancamento') ||
              lk.includes('nota lancamento') ||
              lk.includes('liquidacao')) &&
            v
          ) {
            const s = String(v).trim();
            if (isValidNL(s)) {
              rawDoc = formatCleanNL(s);
              break;
            }
          }
        }
      }
      if (!rawDoc) {
        rawDoc = extractRealNlFromRow(row);
      }
      if (!rawDoc && numCol && row[numCol]) {
        const fallback = String(row[numCol]).trim();
        if (!/^linha\s*\d+$/i.test(fallback) && !/^#?\d{1,4}$/.test(fallback)) {
          rawDoc = fallback;
        }
      }

      const cleanDoc = cleanDocNum(rawDoc);
      const val = valCol ? parseCurrencyValue(row[valCol]) : 0;
      const rawFav = favCol ? row[favCol] : '';
      const { name, cnpj: parsedCnpj } = parseFavorecidoField(rawFav);

      // Auto-resolve missing CNPJ, Supplier Name and Fonte from global app knowledge base
      const effectiveCnpj = resolveCnpj(name, rawDoc, parsedCnpj !== 'N/I' ? parsedCnpj : undefined);
      const effectiveName = resolveSupplierName(effectiveCnpj, rawDoc, name);
      const effectiveFonte = resolveFonteRecurso(rawDoc, effectiveName);
      const cnpjDigits = (effectiveCnpj || '').replace(/\D/g, '');
      const rawProc = procCol ? String(row[procCol] || '').trim() : '';
      const rawData = dataCol ? String(row[dataCol] || '').trim() : '';

      const diaPublicacao = fileDia || extractDiaPublicacao(undefined, rawData) || 'Geral';

      if (val <= 0 && (!cleanDoc || cleanDoc === '-')) {
        return;
      }

      const dedupeKey = `${cleanDoc || 'DOC'}_${cnpjDigits || cleanDocNum(name)}_${Math.round(val * 100)}_${diaPublicacao}`;
      if (processedDocKeys.has(dedupeKey)) return;
      processedDocKeys.add(dedupeKey);

      // Check rule 1: Did the NL disappear from Obedece?
      // Multi-layer deep inspection so an NL that is still in Obedece is NEVER reported as having left!
      const nlKeys = getNlMatchKeys(rawDoc);
      let isStillInObedece = false;

      if (obedeceFiles.length > 0) {
        if (cleanDoc && activeObedeceNls.has(cleanDoc)) {
          isStillInObedece = true;
        } else if (nlKeys.some((k) => activeObedeceNls.has(k))) {
          isStillInObedece = true;
        } else if (cleanDoc && activeObedeceRowTokens.has(cleanDoc)) {
          isStillInObedece = true;
        } else if (cleanDoc && cnpjDigits && activeObedeceNlDocKeys.has(`${cleanDoc}_${cnpjDigits}`)) {
          isStillInObedece = true;
        } else {
          const numOnly = cleanDoc.replace(/\D/g, '');
          if (numOnly.length >= 4 && (activeObedeceNls.has(`NUM_${numOnly}`) || activeObedeceNls.has(`NUM_${parseInt(numOnly, 10)}`))) {
            isStillInObedece = true;
          }
        }
      }

      const sumiuObedece = obedeceFiles.length > 0 ? !isStillInObedece : false;

      // RULE: "se uma nl ja foi resolvida, nao é necessario processar ela novamente"
      const cacheKey = cleanDoc || dedupeKey;
      if (resolvedNlsMemoryCache.has(cacheKey)) {
        const cached = resolvedNlsMemoryCache.get(cacheKey)!;
        results.push({
          id: `pub_${file.id}_${rowIdx}_${cleanDoc || 'item'}`,
          nlNumero: rawDoc || 'S/N',
          fornecedor: effectiveName,
          cnpj: effectiveCnpj,
          valorNL: cached.valorNL !== undefined && cached.valorNL > 0 ? cached.valorNL : val,
          valorPublicacao: val,
          fonteRecurso: cached.fonteRecurso || effectiveFonte,
          processo: rawProc,
          dataNL: rawData,
          diaPublicacao,
          status: 'concluido',
          statusDetalhe: cached.statusDetalhe,
          sumiuObedece,
          pendenteEmObedece: isStillInObedece,
          equivalentePP: cached.equivalentePP,
          equivalenteOB: cached.equivalenteOB,
          valorEquivalente: cached.valorEquivalente,
          situacaoEquivalente: cached.situacaoEquivalente || 'AO',
          diferencaValor: cached.diferencaValor,
          percentualRetencao: cached.percentualRetencao,
          fileName: file.fileName,
        });
        return;
      }

      // Check rule 2: Find equivalent PP and OB verified strictly from General PP spreadsheet
      let matchedPayment: PaymentMatch | null = null;
      for (const k of nlKeys) {
        if (paymentByNlKey.has(k)) {
          matchedPayment = paymentByNlKey.get(k)!;
          break;
        }
      }

      if (!matchedPayment && cleanDoc && cleanDoc.length >= 4) {
        for (const [k, pm] of paymentByNlKey.entries()) {
          if (k.includes(cleanDoc) || cleanDoc.includes(k)) {
            matchedPayment = pm;
            break;
          }
        }
      }

      let status: PublicacaoStatus = 'pendente';
      let statusDetalhe = '';

      if (matchedPayment) {
        status = 'concluido';
        const docLabel = [
          matchedPayment.ppNumero ? `PP ${matchedPayment.ppNumero}` : '',
          matchedPayment.obNumero ? `OB ${matchedPayment.obNumero}` : '',
        ]
          .filter(Boolean)
          .join(' e ');

        const sitLabel = matchedPayment.situacao || 'AO';

        if (matchedPayment.valor > 0 && Math.abs(matchedPayment.valor - val) > 0.05) {
          const diff = val - matchedPayment.valor;
          const pct = val > 0 && diff > 0 ? (diff / val) * 100 : 0;
          statusDetalhe = `Feito: Publicação processada no SIGEF. Emitida ${docLabel || 'PP/OB'} em situação '${sitLabel}' (${formatBRL(matchedPayment.valor)}, retenção tributária de ${pct.toFixed(1).replace('.', ',')}%).`;
        } else {
          statusDetalhe = `Feito: Publicação processada no SIGEF. Emitida ${docLabel || 'PP/OB'} em situação '${sitLabel}'.`;
        }

        // Cache resolution so this NL never needs re-processing
        const resolutionData: CachedResolution = {
          equivalentePP: matchedPayment.ppNumero || undefined,
          equivalenteOB: matchedPayment.obNumero || undefined,
          valorEquivalente: matchedPayment.valor || undefined,
          situacaoEquivalente: matchedPayment.situacao || 'AO',
          diferencaValor: matchedPayment && val > matchedPayment.valor ? val - matchedPayment.valor : undefined,
          percentualRetencao: matchedPayment && val > matchedPayment.valor && val > 0 ? (val - matchedPayment.valor) / val : undefined,
          statusDetalhe,
        };

        resolvedNlsMemoryCache.set(cacheKey, resolutionData);
        if (cleanDoc) resolvedNlsMemoryCache.set(cleanDoc, resolutionData);
        hasNewResolutions = true;
      } else {
        status = 'pendente';
        if (isStillInObedece) {
          statusDetalhe = 'Pendente: A NL permanece listada na fila de Obedece. Necessário realizar a liberação no SIGEF para emissão de PP/OB em AO.';
        } else {
          statusDetalhe = 'Pendente: A NL não consta mais na fila de Obedece, aguardando registro da emissão de PP/OB em situação AO no SIGEF.';
        }
      }

      // Flag if this NL is pending action in Obedece
      const pendenteEmObedece = status === 'pendente' && isStillInObedece;

      results.push({
        id: `pub_${file.id}_${rowIdx}_${cleanDoc || 'item'}`,
        nlNumero: rawDoc || 'S/N',
        fornecedor: effectiveName,
        cnpj: effectiveCnpj,
        fonteRecurso: effectiveFonte,
        valorNL: val,
        processo: rawProc,
        dataNL: rawData,
        diaPublicacao,
        status,
        statusDetalhe,
        sumiuObedece,
        pendenteEmObedece,
        equivalentePP: matchedPayment?.ppNumero || undefined,
        equivalenteOB: matchedPayment?.obNumero || undefined,
        valorEquivalente: matchedPayment?.valor || undefined,
        situacaoEquivalente: matchedPayment?.situacao || undefined,
        diferencaValor: matchedPayment && val > matchedPayment.valor ? val - matchedPayment.valor : undefined,
        percentualRetencao: matchedPayment && val > matchedPayment.valor && val > 0 ? (val - matchedPayment.valor) / val : undefined,
        fileName: file.fileName,
      });
    });
  });

  if (hasNewResolutions) {
    persistResolvedCacheToStorage();
  }

  return results;
}
