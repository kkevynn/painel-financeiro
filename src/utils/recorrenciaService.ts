import { FileData, StageKey, RecurringSupplier, RecurringSupplierDetail, ManualRecurringSupplierEntry } from '../types';
import {
  parseFavorecidoField,
  parseCurrencyValue,
  detectStageFromFileName,
  isSituacaoAO,
  isPPProntaAOValida,
  isTotalizadorRow,
} from './excelParser';

const RECURRING_STORAGE_KEY = 'finance_manual_recurring_suppliers';

// Initial default recurring suppliers (empty by default per user request)
export const DEFAULT_MANUAL_RECURRING: ManualRecurringSupplierEntry[] = [];

let cachedRecurring: ManualRecurringSupplierEntry[] | null = null;
let lastRecurringRaw: string | null = null;

/**
 * Load manual recurring suppliers from persistent storage with in-memory caching
 */
export function getManualRecurringSuppliers(): ManualRecurringSupplierEntry[] {
  try {
    const raw = localStorage.getItem(RECURRING_STORAGE_KEY);
    if (raw === lastRecurringRaw && cachedRecurring !== null) {
      return cachedRecurring;
    }
    lastRecurringRaw = raw;
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) {
        // Filter out legacy demo entries if any
        cachedRecurring = parsed.filter((item: any) => !item.id?.startsWith('rec_init_'));
        return cachedRecurring;
      }
    }
  } catch (err) {
    console.error('Erro ao ler fornecedores recorrentes salvos:', err);
  }
  cachedRecurring = DEFAULT_MANUAL_RECURRING;
  return DEFAULT_MANUAL_RECURRING;
}

/**
 * Save manual recurring suppliers to persistent storage
 */
export function saveManualRecurringSuppliers(list: ManualRecurringSupplierEntry[]): void {
  try {
    cachedRecurring = list;
    const str = JSON.stringify(list);
    lastRecurringRaw = str;
    localStorage.setItem(RECURRING_STORAGE_KEY, str);
  } catch (err) {
    console.error('Erro ao salvar fornecedores recorrentes:', err);
  }
}

/**
 * Add a new manual recurring supplier
 */
export function addManualRecurringSupplier(
  entry: Omit<ManualRecurringSupplierEntry, 'id' | 'createdAt'>
): ManualRecurringSupplierEntry[] {
  const current = getManualRecurringSuppliers();
  
  // Normalize CNPJ
  const rawCnpj = entry.cnpj ? entry.cnpj.trim() : '';
  const digitsOnly = rawCnpj.replace(/\D/g, '');
  let formattedCnpj = rawCnpj || 'N/I';
  if (digitsOnly.length === 14) {
    formattedCnpj = digitsOnly.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/, '$1.$2.$3/$4-$5');
  }

  const cleanName = entry.name.trim().toUpperCase();

  // Avoid duplicates by CNPJ or exact Name
  const exists = current.find(
    (item) =>
      (formattedCnpj !== 'N/I' && item.cnpj === formattedCnpj) ||
      item.name.toUpperCase() === cleanName
  );

  if (exists) {
    // Update notes if provided
    const updated = current.map((item) =>
      item.id === exists.id ? { ...item, notes: entry.notes || item.notes } : item
    );
    saveManualRecurringSuppliers(updated);
    return updated;
  }

  const newEntry: ManualRecurringSupplierEntry = {
    id: `rec_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`,
    name: cleanName,
    cnpj: formattedCnpj,
    notes: entry.notes || '',
    categoryTag: entry.categoryTag || '',
    createdAt: new Date().toISOString(),
  };

  const updated = [newEntry, ...current];
  saveManualRecurringSuppliers(updated);
  return updated;
}

/**
 * Add multiple manual recurring suppliers at once (batch)
 */
export function addMultipleManualRecurringSuppliers(
  entries: Array<Omit<ManualRecurringSupplierEntry, 'id' | 'createdAt'>>
): { updatedList: ManualRecurringSupplierEntry[]; addedCount: number } {
  let current = getManualRecurringSuppliers();
  let addedCount = 0;
  const newItemsToAdd: ManualRecurringSupplierEntry[] = [];

  for (const entry of entries) {
    const rawCnpj = entry.cnpj ? entry.cnpj.trim() : '';
    const digitsOnly = rawCnpj.replace(/\D/g, '');
    let formattedCnpj = rawCnpj || 'N/I';
    if (digitsOnly.length === 14) {
      formattedCnpj = digitsOnly.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/, '$1.$2.$3/$4-$5');
    }

    const cleanName = entry.name.trim().toUpperCase();

    // Check if already in current or in newItemsToAdd
    const alreadyExists =
      current.some(
        (item) =>
          (formattedCnpj !== 'N/I' && item.cnpj === formattedCnpj) ||
          item.name.toUpperCase() === cleanName
      ) ||
      newItemsToAdd.some(
        (item) =>
          (formattedCnpj !== 'N/I' && item.cnpj === formattedCnpj) ||
          item.name.toUpperCase() === cleanName
      );

    if (!alreadyExists) {
      newItemsToAdd.push({
        id: `rec_${Date.now()}_${Math.random().toString(36).substring(2, 7)}_${addedCount}`,
        name: cleanName,
        cnpj: formattedCnpj,
        notes: entry.notes || '',
        categoryTag: entry.categoryTag || '',
        createdAt: new Date().toISOString(),
      });
      addedCount++;
    }
  }

  const updatedList = [...newItemsToAdd, ...current];
  saveManualRecurringSuppliers(updatedList);
  return { updatedList, addedCount };
}

/**
 * Update an existing manual recurring supplier
 */
export function updateManualRecurringSupplier(
  id: string,
  updates: Partial<Omit<ManualRecurringSupplierEntry, 'id' | 'createdAt'>>
): ManualRecurringSupplierEntry[] {
  const current = getManualRecurringSuppliers();
  const updated = current.map((item) => {
    if (item.id !== id) return item;

    let formattedCnpj = item.cnpj;
    if (updates.cnpj !== undefined) {
      const rawCnpj = updates.cnpj.trim();
      const digitsOnly = rawCnpj.replace(/\D/g, '');
      formattedCnpj = rawCnpj || 'N/I';
      if (digitsOnly.length === 14) {
        formattedCnpj = digitsOnly.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/, '$1.$2.$3/$4-$5');
      }
    }

    return {
      ...item,
      name: updates.name ? updates.name.trim().toUpperCase() : item.name,
      cnpj: formattedCnpj,
      notes: updates.notes !== undefined ? updates.notes : item.notes,
      categoryTag: updates.categoryTag !== undefined ? updates.categoryTag : item.categoryTag,
    };
  });

  saveManualRecurringSuppliers(updated);
  return updated;
}

/**
 * Remove a supplier from recurring list by ID, CNPJ or Name
 */
export function removeManualRecurringSupplier(idOrIdentifier: string): ManualRecurringSupplierEntry[] {
  const current = getManualRecurringSuppliers();
  const cleanKey = (idOrIdentifier || '').trim();
  const digitsKey = cleanKey.replace(/\D/g, '');

  const updated = current.filter((item) => {
    if (item.id && item.id === cleanKey) return false;
    if (digitsKey.length === 14 && item.cnpj) {
      const itemDigits = item.cnpj.replace(/\D/g, '');
      if (itemDigits === digitsKey) return false;
    }
    if (item.name && cleanKey && item.name.trim().toUpperCase() === cleanKey.toUpperCase()) {
      return false;
    }
    return true;
  });

  saveManualRecurringSuppliers(updated);
  return updated;
}

/**
 * Clear all manual recurring suppliers or restore default
 */
export function clearAllManualRecurringSuppliers(restoreDefaults = false): ManualRecurringSupplierEntry[] {
  const nextList = restoreDefaults ? DEFAULT_MANUAL_RECURRING : [];
  saveManualRecurringSuppliers(nextList);
  return nextList;
}

export interface ManualRecorrenciaSummary {
  totalRegistered: number;
  totalWithTransactions: number;
  totalRecurringValue: number;
  totalOperations: number;
  averageOperationsPerActiveSupplier: number;
  suppliers: RecurringSupplier[];
}

/**
 * Cross-match manually registered recurring suppliers against all uploaded spreadsheets
 */
export function calculateManualRecurringMetrics(
  files: FileData[],
  manualList: ManualRecurringSupplierEntry[]
): ManualRecorrenciaSummary {
  const list = manualList.length > 0 ? manualList : getManualRecurringSuppliers();

  if (list.length === 0 || files.length === 0) {
    return {
      totalRegistered: list.length,
      totalWithTransactions: 0,
      totalRecurringValue: 0,
      totalOperations: 0,
      averageOperationsPerActiveSupplier: 0,
      suppliers: list.map((entry) => ({
        id: entry.id,
        name: entry.name,
        cnpj: entry.cnpj,
        notes: entry.notes || '',
        categoryTag: entry.categoryTag,
        createdAt: entry.createdAt,
        ppProntasValue: 0,
        liqObedeceValue: 0,
        liqNaoObedeceValue: 0,
        outrosValue: 0,
        totalValue: 0,
        totalCount: 0,
        averageValue: 0,
        stagesCount: 0,
        isMultiStage: false,
        recurrenceLevel: 'moderada' as const,
        details: [],
      })),
    };
  }
  const allRows: {
    fileName: string;
    fileId: string;
    stage: StageKey;
    rawFav: string;
    name: string;
    cnpj: string;
    cleanCnpjDigits: string;
    valor: number;
    rowIndex: number;
    rowData: Record<string, any>;
  }[] = [];

  files.forEach((file) => {
    const favCol = file.detectedFavorecidoCol;
    const valCol = file.detectedValorCol;
    const sitCol = file.detectedSituacaoCol;
    const numCol = file.detectedNumeroCol;
    const obCol = file.detectedObCol;
    const neCol = file.detectedNotaEmpenhoCol;
    const stage: StageKey = file.stage || detectStageFromFileName(file.fileName);

    file.rows.forEach((row, rowIndex) => {
      let rawFav = row[favCol];

      if (isTotalizadorRow(row, rawFav)) {
        return;
      }

      if (rawFav === null || rawFav === undefined || String(rawFav).trim() === '') {
        const keys = Object.keys(row);
        for (const k of keys) {
          const lk = k.toLowerCase();
          if (
            (lk.includes('favorecido') ||
              lk.includes('credor') ||
              lk.includes('fornecedor') ||
              lk.includes('nome')) &&
            row[k]
          ) {
            rawFav = row[k];
            break;
          }
        }
      }

      const val = parseCurrencyValue(row[valCol]);
      const rawNumero = numCol ? String(row[numCol] || '').trim() : '';
      const rawOB = obCol ? String(row[obCol] || '').trim() : '';
      const rawNE = neCol ? String(row[neCol] || '').trim() : '';
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

      const isPPPronta = isPPProntaAOValida({
        situacaoRaw: rawSit,
        numeroRaw: rawNumero,
        obRaw: rawOB,
        favorecidoRaw: rawFav,
        notaEmpenhoRaw: rawNE,
        valor: val,
      });

      const effectiveStage: StageKey =
        stage === 'base_pps' || stage === 'pp_prontas'
          ? (isAO || !sitCol ? 'base_pps' : 'outros')
          : (isPPPronta ? 'base_pps' : stage);

      const { cnpj, name } = parseFavorecidoField(rawFav);
      const cleanCnpjDigits = cnpj.replace(/\D/g, '');

      allRows.push({
        fileName: file.fileName,
        fileId: file.id,
        stage: effectiveStage,
        rawFav: String(rawFav || name),
        name: name.toUpperCase(),
        cnpj,
        cleanCnpjDigits,
        valor: val,
        rowIndex,
        rowData: row,
      });
    });
  });

  // Normalize text helper for flexible matching
  const norm = (s: string) =>
    (s || '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toLowerCase()
      .replace(/[^a-z0-9]/g, ' ')
      .trim();

  // Calculate metrics for each registered recurring supplier
  const calculatedSuppliers: RecurringSupplier[] = list.map((entry) => {
    const cleanEntryDigits = entry.cnpj ? entry.cnpj.replace(/\D/g, '') : '';
    const padEntryCnpj = cleanEntryDigits.length >= 11 ? cleanEntryDigits.padStart(14, '0') : '';
    const normEntryName = norm(entry.name);

    // Match rows
    const matchedRows = allRows.filter((r) => {
      // 1. If entry has a valid CNPJ / CPF, match by padded digits
      if (padEntryCnpj && r.cleanCnpjDigits.length >= 11) {
        const padRowCnpj = r.cleanCnpjDigits.padStart(14, '0');
        if (padRowCnpj === padEntryCnpj) return true;
      }
      
      // 2. Otherwise or in addition, match by normalized Name
      if (normEntryName && normEntryName !== 'nao informado') {
        const normRowName = norm(r.name);
        if (normRowName === normEntryName) return true;
        if (normRowName.includes(normEntryName) || normEntryName.includes(normRowName)) return true;
      }

      return false;
    });

    let basePpsValue = 0;
    let ppProntasValue = 0;
    let liqObedeceValue = 0;
    let liqNaoObedeceValue = 0;
    let outrosValue = 0;
    let totalValue = 0;
    const stagesSet = new Set<StageKey>();
    const details: RecurringSupplierDetail[] = [];

    matchedRows.forEach((item) => {
      const val = item.valor;
      totalValue += val;
      stagesSet.add(item.stage);

      if (item.stage === 'base_pps' || item.stage === 'pp_prontas') {
        basePpsValue += val;
        ppProntasValue += val;
      } else if (item.stage === 'obedece') {
        liqObedeceValue += val;
      } else if (item.stage === 'nao_obedece') {
        liqNaoObedeceValue += val;
      } else {
        outrosValue += val;
      }

      details.push({
        id: `det_${item.fileId}_${item.rowIndex}`,
        fileName: item.fileName,
        stage: item.stage,
        rawFavorecido: item.rawFav,
        valor: val,
        rowNumber: item.rowIndex + 2,
        rowData: item.rowData,
      });
    });

    const totalCount = matchedRows.length;
    const stagesCount = stagesSet.size;
    const isMultiStage = stagesCount > 1;

    let recurrenceLevel: 'alta' | 'media' | 'moderada' = 'moderada';
    if (totalCount >= 5) {
      recurrenceLevel = 'alta';
    } else if (totalCount >= 3) {
      recurrenceLevel = 'media';
    }

    return {
      id: entry.id,
      name: entry.name,
      cnpj: entry.cnpj,
      notes: entry.notes,
      categoryTag: entry.categoryTag,
      createdAt: entry.createdAt,
      basePpsValue,
      ppProntasValue,
      liqObedeceValue,
      liqNaoObedeceValue,
      outrosValue,
      totalValue,
      totalCount,
      averageValue: totalCount > 0 ? totalValue / totalCount : 0,
      stagesCount,
      isMultiStage,
      recurrenceLevel,
      details,
    };
  });

  // Sort: active suppliers with records first by value descending, then suppliers with 0 records
  calculatedSuppliers.sort((a, b) => {
    if (b.totalCount !== a.totalCount) {
      return b.totalCount - a.totalCount;
    }
    return b.totalValue - a.totalValue;
  });

  const totalRegistered = calculatedSuppliers.length;
  const activeSuppliers = calculatedSuppliers.filter((s) => s.totalCount > 0);
  const totalWithTransactions = activeSuppliers.length;
  const totalRecurringValue = calculatedSuppliers.reduce((acc, curr) => acc + curr.totalValue, 0);
  const totalOperations = calculatedSuppliers.reduce((acc, curr) => acc + curr.totalCount, 0);
  const averageOperationsPerActiveSupplier =
    totalWithTransactions > 0 ? totalOperations / totalWithTransactions : 0;

  return {
    totalRegistered,
    totalWithTransactions,
    totalRecurringValue,
    totalOperations,
    averageOperationsPerActiveSupplier,
    suppliers: calculatedSuppliers,
  };
}
