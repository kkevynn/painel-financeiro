/**
 * Service for managing monthly supplier attendance ("Atendimento de Fornecedores por Mês")
 * Persists data to localStorage with live sync across tabs and views.
 */

export interface AtendimentoMensalRecord {
  id: string; // Key: `${monthKey}_${normalizedKey}`
  monthKey: string; // Format: "YYYY-MM", e.g., "2026-09"
  supplierName: string;
  cnpj: string;
  isRecurring: boolean;
  atendido: boolean;
  dataAtendimento?: string; // ISO string when marked
  observacao?: string;
  updatedAt: string;
}

const STORAGE_KEY = 'finance_fornecedores_atendimento_mensal_v1';
const BACKUP_STORAGE_KEY = 'finance_fornecedores_atendimento_mensal_backup';
const SYNC_EVENT = 'finance_atendimento_mensal_changed';

// Potential legacy or alternate keys to automatically scan and restore from if needed
const FALLBACK_KEYS = [
  'finance_fornecedores_atendimento_mensal_v1',
  'finance_fornecedores_atendimento_mensal_backup',
  'finance_fornecedores_atendimento_mensal',
  'sigecon_fornecedores_atendimento_mensal',
  'finance_atendimento_mensal',
  'fornecedores_atendimento_mensal',
];

/**
 * Normalizes a key for a supplier based on CNPJ digits or uppercase name
 */
export function getSupplierStorageKey(name: string, cnpj?: string): string {
  const digits = (cnpj || '').replace(/\D/g, '');
  if (digits.length === 14) return `cnpj_${digits}`;
  return `name_${(name || '').trim().toUpperCase()}`;
}

/**
 * Returns current month key in "YYYY-MM" format
 */
export function getCurrentMonthKey(): string {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  return `${year}-${month}`;
}

const MONTH_NAMES = [
  'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
  'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
];

const MONTH_NAMES_SHORT = [
  'Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun',
  'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'
];

/**
 * Formats "YYYY-MM" to "Setembro de 2026"
 */
export function formatMonthFull(monthKey: string): string {
  if (!monthKey) return '';
  const [yearStr, monthStr] = monthKey.split('-');
  const monthIdx = parseInt(monthStr, 10) - 1;
  if (monthIdx >= 0 && monthIdx < 12) {
    return `${MONTH_NAMES[monthIdx]} de ${yearStr}`;
  }
  return monthKey;
}

/**
 * Formats "YYYY-MM" to "Set/26"
 */
export function formatMonthShort(monthKey: string): string {
  if (!monthKey) return '';
  const [yearStr, monthStr] = monthKey.split('-');
  const monthIdx = parseInt(monthStr, 10) - 1;
  const shortYear = (yearStr || '').slice(-2);
  if (monthIdx >= 0 && monthIdx < 12) {
    return `${MONTH_NAMES_SHORT[monthIdx]}/${shortYear}`;
  }
  return monthKey;
}

/**
 * Generates a list of recent months for selection
 */
export function getAvailableMonthKeys(rangeMonths: number = 14): { key: string; label: string; isCurrent: boolean }[] {
  const currentKey = getCurrentMonthKey();
  const [currYear, currMonth] = currentKey.split('-').map(Number);
  const result: { key: string; label: string; isCurrent: boolean }[] = [];

  // Generate 8 months prior to current and 3 months after
  for (let i = -8; i <= 3; i++) {
    const d = new Date(currYear, currMonth - 1 + i, 1);
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const key = `${y}-${m}`;
    result.push({
      key,
      label: formatMonthFull(key),
      isCurrent: key === currentKey,
    });
  }

  return result;
}

let cachedAtendimentos: Record<string, AtendimentoMensalRecord> | null = null;
let lastRawAtendimentos: string | null = null;

/**
 * Invalidate in-memory cache if needed
 */
export function invalidateAtendimentosCache(): void {
  cachedAtendimentos = null;
  lastRawAtendimentos = null;
}

function getRecordTimestampMs(rec?: AtendimentoMensalRecord | null): number {
  if (!rec) return 0;
  const raw = rec.updatedAt || rec.dataAtendimento || '';
  const ts = Date.parse(raw);
  return Number.isFinite(ts) ? ts : 0;
}

/**
 * Faz o merge seguro de dois mapas de atendimento mensal registro a registro,
 * mantendo sempre a alteração mais recente (updatedAt) por fornecedor/mês.
 */
export function mergeAtendimentosMaps(
  localMap: Record<string, AtendimentoMensalRecord> | null | undefined,
  incomingMap: Record<string, AtendimentoMensalRecord> | null | undefined
): { merged: Record<string, AtendimentoMensalRecord>; localHadNewer: boolean; changedFromLocal: boolean } {
  const merged: Record<string, AtendimentoMensalRecord> = {};
  const safeLocal = localMap && typeof localMap === 'object' ? localMap : {};
  const safeIncoming = incomingMap && typeof incomingMap === 'object' ? incomingMap : {};

  let localHadNewer = false;
  let changedFromLocal = false;

  for (const [id, rec] of Object.entries(safeLocal)) {
    if (rec && typeof rec === 'object') {
      merged[id] = rec;
    }
  }

  for (const id of Object.keys(safeLocal)) {
    if (!(id in safeIncoming)) {
      localHadNewer = true;
    }
  }

  for (const [id, incRec] of Object.entries(safeIncoming)) {
    if (!incRec || typeof incRec !== 'object') continue;
    const locRec = merged[id];
    if (!locRec) {
      merged[id] = incRec;
      changedFromLocal = true;
      continue;
    }

    const locTs = getRecordTimestampMs(locRec);
    const incTs = getRecordTimestampMs(incRec);

    if (incTs > locTs) {
      merged[id] = {
        ...incRec,
        observacao: incRec.observacao !== undefined ? incRec.observacao : locRec.observacao,
      };
      changedFromLocal = true;
    } else if (locTs > incTs) {
      localHadNewer = true;
    } else if (!locRec.atendido && incRec.atendido) {
      merged[id] = incRec;
      changedFromLocal = true;
    } else if (locRec.atendido && !incRec.atendido) {
      localHadNewer = true;
    }
  }

  return { merged, localHadNewer, changedFromLocal };
}

function parseAtendimentosJson(raw: string | null): Record<string, AtendimentoMensalRecord> | null {
  if (!raw) return null;
  try {
    const candidate = JSON.parse(raw);
    if (candidate && typeof candidate === 'object' && !Array.isArray(candidate) && Object.keys(candidate).length > 0) {
      return candidate;
    }
  } catch (e) {
    console.warn('Erro ao decodificar JSON de atendimentos:', e);
  }
  return null;
}

/**
 * Load all records from localStorage with ultra fast in-memory cache.
 * Automatically merges STORAGE_KEY with BACKUP_STORAGE_KEY by updatedAt and searches fallback keys to guarantee ZERO DATA LOSS.
 */
export function getAllAtendimentos(): Record<string, AtendimentoMensalRecord> {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw === lastRawAtendimentos && cachedAtendimentos !== null) {
      return cachedAtendimentos;
    }
    lastRawAtendimentos = raw;

    let parsed = parseAtendimentosJson(raw);
    const backupParsed = parseAtendimentosJson(localStorage.getItem(BACKUP_STORAGE_KEY));

    // Se o backup possuir marcações mais recentes ou ausentes na chave primária, recupera automaticamente
    if (parsed && backupParsed) {
      const { merged, changedFromLocal } = mergeAtendimentosMaps(parsed, backupParsed);
      if (changedFromLocal) {
        parsed = merged;
        const mergedStr = JSON.stringify(merged);
        lastRawAtendimentos = mergedStr;
        try {
          localStorage.setItem(STORAGE_KEY, mergedStr);
          localStorage.setItem(BACKUP_STORAGE_KEY, mergedStr);
        } catch (_) {}
      }
    } else if (!parsed && backupParsed) {
      parsed = backupParsed;
      const mergedStr = JSON.stringify(backupParsed);
      lastRawAtendimentos = mergedStr;
      try {
        localStorage.setItem(STORAGE_KEY, mergedStr);
      } catch (_) {}
    }

    // If primary & backup storage were empty, actively search other fallback/legacy keys
    if (!parsed || Object.keys(parsed).length === 0) {
      for (const fallbackKey of FALLBACK_KEYS) {
        try {
          const candidate = parseAtendimentosJson(localStorage.getItem(fallbackKey));
          if (candidate) {
            parsed = candidate;
            saveAll(candidate);
            break;
          }
        } catch (_) {}
      }

      // Also search any other keys in localStorage containing 'atendimento'
      if (!parsed || Object.keys(parsed).length === 0) {
        for (let i = 0; i < localStorage.length; i++) {
          const key = localStorage.key(i);
          if (key && key.includes('atendimento') && !FALLBACK_KEYS.includes(key)) {
            try {
              const candidate = parseAtendimentosJson(localStorage.getItem(key));
              if (candidate) {
                parsed = candidate;
                saveAll(candidate);
                break;
              }
            } catch (_) {}
          }
        }
      }
    }

    if (parsed) {
      cachedAtendimentos = parsed;
      return parsed;
    }
  } catch (err) {
    console.error('Erro ao ler atendimentos mensais:', err);
  }
  cachedAtendimentos = {};
  return cachedAtendimentos;
}

/**
 * Mescla atendimentos vindos da nuvem com os registros locais (e backup) usando updatedAt,
 * evitando que um pull atrasado apague fornecedores recém-marcados localmente.
 */
export function mergeAtendimentosFromCloud(remoteRawJson: string): {
  mergedJson: string;
  localHadNewer: boolean;
} {
  const remoteMap = parseAtendimentosJson(remoteRawJson) || {};
  invalidateAtendimentosCache();
  const localMap = getAllAtendimentos();
  const { merged, localHadNewer } = mergeAtendimentosMaps(localMap, remoteMap);
  const mergedJson = JSON.stringify(merged);
  cachedAtendimentos = merged;
  lastRawAtendimentos = mergedJson;
  try {
    localStorage.setItem(STORAGE_KEY, mergedJson);
    localStorage.setItem(BACKUP_STORAGE_KEY, mergedJson);
  } catch (err) {
    console.error('Erro ao salvar merge de atendimentos da nuvem:', err);
  }
  return { mergedJson, localHadNewer };
}

/**
 * Save all records to localStorage and dispatch sync event
 * Keeps double backup to prevent any accidental clearing or loss.
 */
function saveAll(records: Record<string, AtendimentoMensalRecord>): void {
  try {
    cachedAtendimentos = records;
    const str = JSON.stringify(records);
    lastRawAtendimentos = str;
    localStorage.setItem(STORAGE_KEY, str);
    localStorage.setItem(BACKUP_STORAGE_KEY, str);
    window.dispatchEvent(new CustomEvent(SYNC_EVENT, { detail: records }));
  } catch (err) {
    console.error('Erro ao salvar atendimentos mensais:', err);
  }
}

/**
 * Check if a supplier was attended in a specific month
 */
export function isSupplierAtendidoNoMes(
  name: string,
  cnpj: string = '',
  monthKey: string = getCurrentMonthKey()
): boolean {
  const all = getAllAtendimentos();
  const supKey = getSupplierStorageKey(name, cnpj);
  const recordId = `${monthKey}_${supKey}`;
  return !!all[recordId]?.atendido;
}

/**
 * Get the full record for a supplier in a specific month
 */
export function getSupplierAtendimentoRecord(
  name: string,
  cnpj: string = '',
  monthKey: string = getCurrentMonthKey()
): AtendimentoMensalRecord | null {
  const all = getAllAtendimentos();
  const supKey = getSupplierStorageKey(name, cnpj);
  const recordId = `${monthKey}_${supKey}`;
  return all[recordId] || null;
}

/**
 * Toggle or set supplier attendance in a given month
 */
export function toggleSupplierAtendimento(
  name: string,
  cnpj: string = '',
  isRecurring: boolean = false,
  monthKey: string = getCurrentMonthKey(),
  observacao?: string
): AtendimentoMensalRecord {
  const all = getAllAtendimentos();
  const supKey = getSupplierStorageKey(name, cnpj);
  const recordId = `${monthKey}_${supKey}`;

  const current = all[recordId];
  const nextAtendido = current ? !current.atendido : true;

  const updated: AtendimentoMensalRecord = {
    id: recordId,
    monthKey,
    supplierName: name,
    cnpj: cnpj || 'N/I',
    isRecurring,
    atendido: nextAtendido,
    dataAtendimento: nextAtendido ? new Date().toISOString() : undefined,
    observacao: observacao !== undefined ? observacao : current?.observacao,
    updatedAt: new Date().toISOString(),
  };

  all[recordId] = updated;
  saveAll(all);
  return updated;
}

/**
 * Set an observation/note for a supplier's attendance in a given month
 */
export function updateSupplierAtendimentoObs(
  name: string,
  cnpj: string = '',
  monthKey: string = getCurrentMonthKey(),
  observacao: string
): void {
  const all = getAllAtendimentos();
  const supKey = getSupplierStorageKey(name, cnpj);
  const recordId = `${monthKey}_${supKey}`;

  if (!all[recordId]) {
    all[recordId] = {
      id: recordId,
      monthKey,
      supplierName: name,
      cnpj: cnpj || 'N/I',
      isRecurring: false,
      atendido: false,
      observacao,
      updatedAt: new Date().toISOString(),
    };
  } else {
    all[recordId] = {
      ...all[recordId],
      observacao,
      updatedAt: new Date().toISOString(),
    };
  }
  saveAll(all);
}

/**
 * Hook or helper to subscribe to changes across components
 */
export function subscribeToAtendimentos(callback: () => void): () => void {
  const handler = (e?: Event) => {
    if (e && 'key' in e) {
      const storageEvent = e as StorageEvent;
      // Strictly ignore storage events from other unrelated keys (e.g. app state, toast alerts, etc.)
      if (storageEvent.key !== STORAGE_KEY) {
        return;
      }
    }
    invalidateAtendimentosCache();
    callback();
  };
  window.addEventListener(SYNC_EVENT, handler);
  window.addEventListener('storage', handler);
  return () => {
    window.removeEventListener(SYNC_EVENT, handler);
    window.removeEventListener('storage', handler);
  };
}
