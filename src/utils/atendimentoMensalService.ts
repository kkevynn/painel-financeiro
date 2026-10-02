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

/**
 * Load all records from localStorage with ultra fast in-memory cache
 * Automatically searches and restores from backup and legacy keys to guarantee ZERO DATA LOSS.
 */
export function getAllAtendimentos(): Record<string, AtendimentoMensalRecord> {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw === lastRawAtendimentos && cachedAtendimentos !== null) {
      return cachedAtendimentos;
    }
    lastRawAtendimentos = raw;

    let parsed: Record<string, AtendimentoMensalRecord> | null = null;

    if (raw) {
      try {
        const candidate = JSON.parse(raw);
        if (candidate && typeof candidate === 'object' && Object.keys(candidate).length > 0) {
          parsed = candidate;
        }
      } catch (e) {
        console.warn('Erro ao decodificar JSON primário de atendimentos:', e);
      }
    }

    // If primary storage was empty or had 0 records, actively search fallback and backup keys!
    if (!parsed || Object.keys(parsed).length === 0) {
      for (const fallbackKey of FALLBACK_KEYS) {
        try {
          const fallbackRaw = localStorage.getItem(fallbackKey);
          if (fallbackRaw) {
            const candidate = JSON.parse(fallbackRaw);
            if (candidate && typeof candidate === 'object' && Object.keys(candidate).length > 0) {
              parsed = candidate;
              // Restore immediately to primary and backup
              saveAll(candidate);
              break;
            }
          }
        } catch (_) {}
      }

      // Also search any other keys in localStorage containing 'atendimento'
      if (!parsed || Object.keys(parsed).length === 0) {
        for (let i = 0; i < localStorage.length; i++) {
          const key = localStorage.key(i);
          if (key && key.includes('atendimento') && !FALLBACK_KEYS.includes(key)) {
            try {
              const val = localStorage.getItem(key);
              if (val) {
                const candidate = JSON.parse(val);
                if (candidate && typeof candidate === 'object' && Object.keys(candidate).length > 0) {
                  parsed = candidate;
                  saveAll(candidate);
                  break;
                }
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
      // Ignore storage events from other unrelated keys (e.g. app state, toast alerts, etc.)
      if (storageEvent.key && storageEvent.key !== STORAGE_KEY) {
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
