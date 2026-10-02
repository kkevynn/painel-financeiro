import { ToastAlert, ToastAlertType, FileData } from '../types';
import { isSituacaoAO, parseCurrencyValue, parseFavorecidoField, isSituacaoConclusivaPagamento } from './excelParser';
import { OrdemBancariaItem, parseDateToMonthInfo } from './ordemBancariaParser';

const TOAST_STORAGE_KEY = 'painel_pp_status_history';
const ALERTS_HISTORY_KEY = 'painel_all_alerts_history';
const BASELINE_SNAPSHOT_KEY = 'painel_base_files_snapshot';
const ACTIVE_TOASTS_MAX = 5;

export function getAlertUniqueKey(a: {
  pp?: string;
  ob?: string;
  ne?: string;
  cnpj?: string;
  fornecedor?: string;
  situacaoNova?: string;
  tipo?: string;
  valor?: number;
}): string {
  const ppClean = (a.pp || '').trim().toUpperCase().replace(/[\s\.\-\/]/g, '');
  const obClean = (a.ob || '').trim().toUpperCase().replace(/[\s\.\-\/]/g, '');
  const neClean = (a.ne || '').trim().toUpperCase().replace(/[\s\.\-\/]/g, '');
  const cnpjClean = (a.cnpj || '').replace(/\D/g, '');
  const valKey = a.valor !== undefined && a.valor > 0 ? `_${a.valor.toFixed(0)}` : '';
  const doc = obClean || ppClean || neClean || (cnpjClean ? `${cnpjClean}${valKey}` : '') || (a.fornecedor ? `${a.fornecedor.trim().toUpperCase()}${valKey}` : 'DOC');
  const sit = (a.situacaoNova || '').trim().toUpperCase();
  const tipo = (a.tipo || '').trim().toLowerCase();
  return `${doc}__${obClean}__${sit}__${tipo}`;
}

function isValidOBString(s?: string): boolean {
  if (!s) return false;
  const c = s.trim().toUpperCase();
  if (
    !c ||
    c === '-' ||
    c === '0' ||
    c === 'N/I' ||
    c === 'NULL' ||
    c === 'UNDEFINED' ||
    c === 'NÃO INFORMADO' ||
    c === 'NAO INFORMADO'
  ) {
    return false;
  }
  if (/^(OB|ORDEM|TOTAL|VALOR|DOCUMENTO)$/i.test(c)) return false;
  return c.length >= 3;
}

function cleanDocStr(s?: any): string {
  if (!s) return '';
  return String(s).trim().toUpperCase().replace(/[\s\.\-\/]/g, '');
}

/**
 * Extracts a Date object and formatted DD/MM/YYYY date from any spreadsheet row
 */
function extractRowDate(row: Record<string, any>): { dateObj: Date; dateStr: string } | null {
  for (const [k, v] of Object.entries(row)) {
    if (v === null || v === undefined || v === '') continue;
    const lk = k.toLowerCase();
    if (
      lk.includes('data') ||
      lk.includes('dt_') ||
      lk === 'dt' ||
      lk.includes('emiss') ||
      lk.includes('pagamento') ||
      lk.includes('refer') ||
      lk.includes('liquid')
    ) {
      const dateInfo = parseDateToMonthInfo(v);
      if (dateInfo.dateIso) {
        const d = new Date(dateInfo.dateIso + 'T12:00:00');
        if (!isNaN(d.getTime())) {
          const [y, m, day] = dateInfo.dateIso.split('-');
          return { dateObj: d, dateStr: `${day}/${m}/${y}` };
        }
      }
    }
  }
  return null;
}

export function isSituacaoRejeitada(situacaoRaw: any): boolean {
  if (situacaoRaw === null || situacaoRaw === undefined) return false;
  const str = String(situacaoRaw)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toUpperCase();
  if (!str) return false;
  return (
    str === 'RJ' ||
    str.startsWith('RJ') ||
    str === 'PPRJ' ||
    str.startsWith('PPRJ') ||
    str === 'CBNC' ||
    str.startsWith('CBNC') ||
    str.includes('REJEITAD') ||
    str.includes('CANCELAD') ||
    str.includes('DEVOLVID') ||
    str.includes('ESTORNAD') ||
    str.includes('ANULAD')
  );
}

type ToastListener = (toasts: ToastAlert[]) => void;

class ToastAlertManager {
  private activeToasts: ToastAlert[] = [];
  private alertsHistory: ToastAlert[] = [];
  private listeners: Set<ToastListener> = new Set();

  constructor() {
    this.activeToasts = [];
    this.loadAlertsHistory();
  }

  private loadAlertsHistory() {
    try {
      const raw = localStorage.getItem(ALERTS_HISTORY_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        const seen = new Set<string>();
        const unique: ToastAlert[] = [];

        for (const item of parsed) {
          const key = getAlertUniqueKey(item);
          if (!seen.has(key)) {
            seen.add(key);
            unique.push({
              ...item,
              dataHora: new Date(item.dataHora),
            });
          }
        }
        unique.sort((a, b) => new Date(b.dataHora).getTime() - new Date(a.dataHora).getTime());
        this.alertsHistory = unique;
        this.saveAlertsHistory();
      }
    } catch (e) {
      this.alertsHistory = [];
    }
  }

  private saveAlertsHistory() {
    try {
      localStorage.setItem(ALERTS_HISTORY_KEY, JSON.stringify(this.alertsHistory.slice(0, 500)));
    } catch (e) {
      // safe quota catch
    }
  }

  public subscribe(listener: ToastListener): () => void {
    this.listeners.add(listener);
    return () => {
      this.listeners.delete(listener);
    };
  }

  private notify() {
    setTimeout(() => {
      this.listeners.forEach((fn) => {
        try {
          fn(this.activeToasts);
        } catch (e) {
          console.error('Toast listener notification error:', e);
        }
      });
    }, 0);
  }

  public addToast(
    toast: Omit<ToastAlert, 'id' | 'dataHora' | 'resolvido'> & {
      id?: string;
      dataHora?: Date;
      resolvido?: boolean;
      showPopup?: boolean;
    }
  ) {
    const uniqueKey = getAlertUniqueKey(toast);
    const existingIndex = this.alertsHistory.findIndex((t) => getAlertUniqueKey(t) === uniqueKey);

    const fullToast: ToastAlert = {
      ...toast,
      id: toast.id || `toast_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`,
      dataHora: toast.dataHora || new Date(),
      resolvido: toast.resolvido || false,
    };

    if (existingIndex !== -1) {
      // Update existing item with fresh timestamp
      this.alertsHistory[existingIndex] = {
        ...this.alertsHistory[existingIndex],
        ...fullToast,
        dataHora: fullToast.dataHora,
      };
      this.saveAlertsHistory();
      this.notify();
      return;
    }

    this.alertsHistory.unshift(fullToast);
    this.saveAlertsHistory();

    const showInToastPopup = toast.showPopup ?? (toast.tipo === 'critico');
    if (showInToastPopup) {
      const alreadyActive = this.activeToasts.some((t) => getAlertUniqueKey(t) === uniqueKey);
      if (!alreadyActive) {
        this.activeToasts = [fullToast, ...this.activeToasts.slice(0, ACTIVE_TOASTS_MAX - 1)];
      }

      const timeoutMs = toast.tipo === 'critico' ? 12000 : 7000;
      setTimeout(() => {
        this.dismissToast(fullToast.id);
      }, timeoutMs);
    }

    this.notify();
  }

  public batchAddToasts(
    toasts: Array<
      Omit<ToastAlert, 'id' | 'dataHora' | 'resolvido'> & {
        id?: string;
        dataHora?: Date;
        resolvido?: boolean;
        showPopup?: boolean;
      }
    >
  ) {
    if (!toasts || toasts.length === 0) return;
    const existingKeys = new Set(this.alertsHistory.map((t) => getAlertUniqueKey(t)));
    const newItems: ToastAlert[] = [];

    for (const toast of toasts) {
      const uniqueKey = getAlertUniqueKey(toast);
      if (!existingKeys.has(uniqueKey)) {
        existingKeys.add(uniqueKey);
        const fullToast: ToastAlert = {
          ...toast,
          id: toast.id || `toast_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`,
          dataHora: toast.dataHora || new Date(),
          resolvido: toast.resolvido || false,
        };
        newItems.push(fullToast);
      }
    }

    if (newItems.length > 0) {
      this.alertsHistory = [...newItems, ...this.alertsHistory].slice(0, 300);
      const toPopup = newItems.filter((t) => (t as any).showPopup).slice(0, ACTIVE_TOASTS_MAX);
      if (toPopup.length > 0) {
        this.activeToasts = [...toPopup, ...this.activeToasts].slice(0, ACTIVE_TOASTS_MAX);
        toPopup.forEach((p) => {
          setTimeout(() => this.dismissToast(p.id), 7000);
        });
      }
      this.saveAlertsHistory();
      this.notify();
    }
  }

  public dismissToast(id: string) {
    this.activeToasts = this.activeToasts.filter((t) => t.id !== id);
    this.notify();
  }

  public clearActiveToasts() {
    this.activeToasts = [];
    this.notify();
  }

  public clearAlertsHistory() {
    this.alertsHistory = [];
    this.activeToasts = [];
    this.saveAlertsHistory();
    this.notify();
  }

  public clearAll() {
    this.activeToasts = [];
    this.notify();
  }

  public getActiveToasts(): ToastAlert[] {
    return [...this.activeToasts];
  }

  public getAlertsHistory(): ToastAlert[] {
    return [...this.alertsHistory];
  }

  public updateToast(id: string, updates: Partial<ToastAlert>) {
    this.activeToasts = this.activeToasts.map((t) => (t.id === id ? { ...t, ...updates } : t));
    this.alertsHistory = this.alertsHistory.map((t) => (t.id === id ? { ...t, ...updates } : t));
    this.saveAlertsHistory();
    this.notify();
  }

  public resolveToast(id: string, observacao?: string) {
    this.updateToast(id, { resolvido: true, observacao });
    this.dismissToast(id);
  }

  /**
   * Resets baseline snapshot so the next upload will establish a fresh baseline
   */
  public resetBaseline() {
    try {
      localStorage.removeItem(BASELINE_SNAPSHOT_KEY);
      this.alertsHistory = [];
      this.activeToasts = [];
      this.saveAlertsHistory();
      this.notify();
    } catch (e) {
      // safe
    }
  }

  public hasBaseline(): boolean {
    try {
      return !!localStorage.getItem(BASELINE_SNAPSHOT_KEY);
    } catch (e) {
      return false;
    }
  }

  private loadBaselineSnapshot(): Record<string, any> | null {
    try {
      const raw = localStorage.getItem(BASELINE_SNAPSHOT_KEY);
      if (!raw) return null;
      const parsed = JSON.parse(raw);
      const res: Record<string, any> = {};
      for (const [k, v] of Object.entries(parsed as Record<string, any>)) {
        res[k] = {
          situacao: v.s || v.situacao || '',
          ob: v.ob || '',
          isConclusivo: v.c !== undefined ? v.c : !!v.isConclusivo,
          isRejected: v.r !== undefined ? v.r : !!v.isRejected,
        };
      }
      return res;
    } catch (e) {
      return null;
    }
  }

  private saveBaselineSnapshot(records: Record<string, any>) {
    try {
      // Save compact representation of up to 3500 keys to avoid quota/memory issues
      const compact: Record<string, { s: string; ob?: string; c: boolean; r: boolean }> = {};
      const keys = Object.keys(records).slice(0, 3500);
      for (const k of keys) {
        const item = records[k];
        compact[k] = {
          s: item.situacao || '',
          ob: item.ob || '',
          c: !!item.isConclusivo,
          r: !!item.isRejected,
        };
      }
      localStorage.setItem(BASELINE_SNAPSHOT_KEY, JSON.stringify(compact));
    } catch (e) {
      // safe quota catch
    }
  }

  /**
   * Extracts normalized records from current files and ordens bancarias
   */
  private extractAllCurrentRecords(
    files: FileData[],
    ordensBancariasList?: OrdemBancariaItem[] | null
  ): Record<
    string,
    {
      key: string;
      ob: string;
      pp: string;
      ne: string;
      cnpj: string;
      fornecedor: string;
      valor: number;
      situacao: string;
      isConclusivo: boolean;
      isRejected: boolean;
      dataReferencia?: string;
      fileName?: string;
    }
  > {
    const records: Record<string, any> = {};

    // 1. Extract from Ordens Bancárias
    if (ordensBancariasList && ordensBancariasList.length > 0) {
      ordensBancariasList.forEach((obItem) => {
        const obNum = isValidOBString(obItem.ob) ? obItem.ob.trim().toUpperCase() : '';
        const ppNum = cleanDocStr(obItem.pp);
        const cnpjNum = (obItem.favorecidoCnpj || '').replace(/\D/g, '');
        const valNum = typeof obItem.valor === 'number' ? obItem.valor : 0;
        const sitStr = (obItem.situacao || 'CB').trim().toUpperCase();
        const isConclusivo = isSituacaoConclusivaPagamento(sitStr) || sitStr === 'CB';
        const isRejected = isSituacaoRejeitada(sitStr);

        const key = obNum
          ? `OB_${obNum}`
          : ppNum
          ? `PP_${ppNum}`
          : cnpjNum
          ? `CNPJ_${cnpjNum}_${Math.round(valNum)}`
          : `FORN_${cleanDocStr(obItem.favorecidoName)}_${Math.round(valNum)}`;

        records[key] = {
          key,
          ob: obNum,
          pp: ppNum,
          ne: '',
          cnpj: obItem.favorecidoCnpj || '',
          fornecedor: obItem.favorecidoName || '',
          valor: valNum,
          situacao: sitStr,
          isConclusivo,
          isRejected,
          dataReferencia: obItem.dataReferencia,
          fileName: 'Listar Ordem Bancária',
        };
      });
    }

    // 2. Extract from Files
    files.forEach((file) => {
      const sitCol = file.detectedSituacaoCol;
      const numCol = file.detectedNumeroCol || file.detectedPpCol;
      const valCol = file.detectedValorCol;
      const favCol = file.detectedFavorecidoCol;
      const obCol = file.detectedObCol;
      const neCol = file.detectedNotaEmpenhoCol;

      file.rows.forEach((row) => {
        let rawSit = sitCol ? String(row[sitCol] || '').trim().toUpperCase() : '';
        if (!rawSit) {
          for (const [k, v] of Object.entries(row)) {
            const lk = k.toLowerCase();
            if (lk.includes('situa') || lk.includes('status') || lk === 'sigla' || lk === 'sit') {
              rawSit = String(v || '').trim().toUpperCase();
              if (rawSit) break;
            }
          }
        }

        let rawOB = obCol ? String(row[obCol] || '').trim() : '';
        if (!isValidOBString(rawOB)) {
          for (const [k, v] of Object.entries(row)) {
            const lk = k.toLowerCase();
            if ((lk.includes('ob') || lk.includes('ordem bancaria') || lk.includes('ordem bancária')) && v) {
              const valStr = String(v).trim();
              if (isValidOBString(valStr)) {
                rawOB = valStr;
                break;
              }
            }
          }
        }

        let rawNum = numCol ? String(row[numCol] || '').trim() : '';
        if (!rawNum) {
          for (const [k, v] of Object.entries(row)) {
            const lk = k.toLowerCase();
            if (
              (lk.includes('documento') ||
                lk.includes('número') ||
                lk.includes('numero') ||
                lk.includes('pp') ||
                lk.includes('processo') ||
                lk.includes('nl')) &&
              v
            ) {
              const valStr = String(v).trim();
              if (valStr && valStr !== '-' && valStr !== '0') {
                rawNum = valStr;
                break;
              }
            }
          }
        }

        let rawNE = neCol ? String(row[neCol] || '').trim() : '';
        if (!rawNE) {
          for (const [k, v] of Object.entries(row)) {
            const lk = k.toLowerCase();
            if ((lk.includes('empenho') || lk.includes('ne')) && v) {
              const valStr = String(v).trim();
              if (valStr && valStr !== '-' && valStr !== '0') {
                rawNE = valStr;
                break;
              }
            }
          }
        }

        const rawVal = valCol ? row[valCol] : 0;
        const val = typeof rawVal === 'number' ? rawVal : parseCurrencyValue(rawVal);
        if (val <= 0) return;

        const rawFav = favCol ? row[favCol] : '';
        const { name: favName, cnpj: favCnpj } = parseFavorecidoField(rawFav);
        const cnpjNum = (favCnpj || '').replace(/\D/g, '');
        const obNum = isValidOBString(rawOB) ? rawOB.trim().toUpperCase() : '';
        const ppNum = cleanDocStr(rawNum);
        const neNum = cleanDocStr(rawNE);

        const isConclusivo = isSituacaoConclusivaPagamento(rawSit);
        const isRejected = isSituacaoRejeitada(rawSit);

        const rowDateInfo = extractRowDate(row);

        const key = obNum
          ? `OB_${obNum}`
          : ppNum
          ? `PP_${ppNum}`
          : neNum
          ? `NE_${neNum}`
          : cnpjNum
          ? `CNPJ_${cnpjNum}_${Math.round(val)}`
          : `FORN_${cleanDocStr(favName)}_${Math.round(val)}`;

        if (!records[key]) {
          records[key] = {
            key,
            ob: obNum,
            pp: ppNum,
            ne: neNum,
            cnpj: favCnpj,
            fornecedor: favName,
            valor: val,
            situacao: rawSit,
            isConclusivo,
            isRejected,
            dataReferencia: rowDateInfo?.dateStr,
            fileName: file.fileName,
          };
        }
      });
    });

    return records;
  }

  /**
   * Refreshes confirmations from current files and ordensBancariasList.
   */
  public syncConfirmations(
    files: FileData[],
    ordensBancariasList?: OrdemBancariaItem[] | null,
    forceRefresh: boolean = false
  ): number {
    this.analyzeFileTransitions(files, ordensBancariasList, forceRefresh);
    return this.alertsHistory.filter((a) => a.tipo === 'sucesso').length;
  }

  /**
   * Snapshot-based comparison:
   * When user uploads base files, stores snapshot.
   * On subsequent uploads, compares against previous snapshot to report only updates
   * that occurred in that lapse of time (confirmed, rejected, virou OB, changed situation).
   */
  public analyzeFileTransitions(
    files: FileData[],
    ordensBancariasList?: OrdemBancariaItem[] | null,
    forceRefresh: boolean = false
  ) {
    if ((!files || files.length === 0) && (!ordensBancariasList || ordensBancariasList.length === 0)) {
      return;
    }

    const currentRecords = this.extractAllCurrentRecords(files, ordensBancariasList);
    const previousSnapshot = this.loadBaselineSnapshot();

    // 1. Initial baseline upload: store snapshot, do not spam with ancient historic payments
    if (!previousSnapshot || Object.keys(previousSnapshot).length === 0) {
      this.saveBaselineSnapshot(currentRecords);
      // Keep only manual/resolved alerts
      this.alertsHistory = this.alertsHistory.filter((a) => a.resolvido || !!a.observacao);
      this.activeToasts = [];
      this.saveAlertsHistory();
      this.notify();
      return;
    }

    // 2. Subsequent upload: compare current records with previous snapshot
    const toastsToBatch: any[] = [];
    let countUpdates = 0;
    const currentKeys = Object.keys(currentRecords);

    for (const key of currentKeys) {
      if (countUpdates >= 25) break; // Cap at 25 updates per upload to keep UI instant and clean
      const curr = currentRecords[key];
      // Match by exact key, or secondary check by PP / OB
      let prev = previousSnapshot[key];
      if (!prev && curr.pp) {
        prev = previousSnapshot[`PP_${curr.pp}`];
      }
      if (!prev && curr.ob) {
        prev = previousSnapshot[`OB_${curr.ob}`];
      }

      if (prev) {
        // Newly Concluded / Paid in this lapse of time
        if (!prev.isConclusivo && curr.isConclusivo) {
          toastsToBatch.push({
            tipo: 'sucesso',
            titulo: 'Pagamento Confirmado no Banco',
            mensagem: `${curr.ob ? `OB ${curr.ob}` : ''}${curr.pp ? ` | PP ${curr.pp}` : ''} - Confirmado no Banco (${curr.situacao || 'CB'})`,
            fornecedor: curr.fornecedor,
            cnpj: curr.cnpj,
            ob: curr.ob,
            pp: curr.pp,
            valor: curr.valor,
            situacaoAnterior: prev.situacao,
            situacaoNova: curr.situacao || 'CB',
            dataReferencia: curr.dataReferencia,
            dataHora: new Date(),
            showPopup: countUpdates < 3,
          });
          countUpdates++;
        }
        // Newly Rejected in this lapse of time
        else if (!prev.isRejected && curr.isRejected) {
          toastsToBatch.push({
            tipo: 'critico',
            titulo: 'Pagamento Rejeitado / Estornado',
            mensagem: `${curr.ob ? `OB ${curr.ob}` : ''}${curr.pp ? ` | PP ${curr.pp}` : ''} - ${curr.situacao}`,
            fornecedor: curr.fornecedor,
            cnpj: curr.cnpj,
            ob: curr.ob,
            pp: curr.pp,
            valor: curr.valor,
            situacaoAnterior: prev.situacao,
            situacaoNova: curr.situacao,
            dataReferencia: curr.dataReferencia,
            dataHora: new Date(),
            showPopup: countUpdates < 3,
          });
          countUpdates++;
        }
        // Newly Generated OB (Virou OB) in this lapse of time
        else if (!isValidOBString(prev.ob) && isValidOBString(curr.ob)) {
          toastsToBatch.push({
            tipo: 'sucesso',
            titulo: 'Ordem Bancária Gerada',
            mensagem: `PP ${curr.pp || ''} virou OB ${curr.ob} (${curr.situacao})`,
            fornecedor: curr.fornecedor,
            cnpj: curr.cnpj,
            ob: curr.ob,
            pp: curr.pp,
            valor: curr.valor,
            situacaoAnterior: prev.situacao,
            situacaoNova: 'Virou OB',
            dataReferencia: curr.dataReferencia,
            dataHora: new Date(),
            showPopup: countUpdates < 3,
          });
          countUpdates++;
        }
        // Situation updated between uploads
        else if (prev.situacao !== curr.situacao && !curr.isConclusivo && !curr.isRejected) {
          toastsToBatch.push({
            tipo: 'sucesso',
            titulo: `Transição: ${prev.situacao} → ${curr.situacao}`,
            mensagem: `${curr.ob ? `OB ${curr.ob}` : ''}${curr.pp ? ` | PP ${curr.pp}` : ''} - Situação atualizada`,
            fornecedor: curr.fornecedor,
            cnpj: curr.cnpj,
            ob: curr.ob,
            pp: curr.pp,
            valor: curr.valor,
            situacaoAnterior: prev.situacao,
            situacaoNova: curr.situacao,
            dataReferencia: curr.dataReferencia,
            dataHora: new Date(),
            showPopup: countUpdates < 3,
          });
          countUpdates++;
        }
      }
    }

    if (toastsToBatch.length > 0) {
      this.batchAddToasts(toastsToBatch);
    }

    // 3. Update baseline snapshot for the next upload
    this.saveBaselineSnapshot(currentRecords);
    this.notify();
  }
}

export const toastAlertService = new ToastAlertManager();
