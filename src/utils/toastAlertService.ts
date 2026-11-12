import { ToastAlert, ToastAlertType, FileData } from '../types';
import {
  formatBRL,
  isSituacaoAO,
  parseCurrencyValue,
  parseFavorecidoField,
  isSituacaoConclusivaPagamento,
  isFonteHeader,
  isFonteString,
} from './excelParser';
import {
  OrdemBancariaItem,
  parseDateToMonthInfo,
  isFonte600OuDerivada,
  findFonteRecursoInRow,
} from './ordemBancariaParser';
import { mlInferenceService } from './machineLearningInferenceService';

const TOAST_STORAGE_KEY = 'painel_pp_status_history';
const ALERTS_HISTORY_KEY = 'painel_all_alerts_history';
const BASELINE_SNAPSHOT_KEY = 'painel_base_files_snapshot';
const ACTIVE_TOASTS_MAX = 5;

export function getAlertUniqueKey(a: {
  pp?: string;
  ob?: string;
  ne?: string;
  nl?: string;
  cnpj?: string;
  fornecedor?: string;
  situacaoNova?: string;
  tipo?: string;
  valor?: number;
}): string {
  const ppClean = (a.pp || '').trim().toUpperCase().replace(/[\s\.\-\/]/g, '');
  const obClean = (a.ob || '').trim().toUpperCase().replace(/[\s\.\-\/]/g, '');
  const neClean = (a.ne || '').trim().toUpperCase().replace(/[\s\.\-\/]/g, '');
  const nlClean = (a.nl || '').trim().toUpperCase().replace(/[\s\.\-\/]/g, '');
  const cnpjClean = (a.cnpj || '').replace(/\D/g, '');
  const valKey = a.valor !== undefined && a.valor > 0 ? `_${a.valor.toFixed(0)}` : '';

  const doc = ppClean || obClean || nlClean || neClean || (cnpjClean ? `${cnpjClean}${valKey}` : 'DOC');
  const rawSit = (a.situacaoNova || '').trim().toUpperCase();

  // Agrupa CB, PPCB, CM e PAGO sob a mesma chave lógica para nunca duplicar alerta entre planilhas diferentes
  const sitGroup = isSituacaoConfirmadaBanco(rawSit).isMatch
    ? 'CONF_BANCO'
    : isSituacaoRejeitadaBanco(rawSit).isMatch
    ? 'REJ_BANCO'
    : rawSit;

  return `${doc}__${sitGroup}`;
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
 * Extracts a Date object and formatted DD/MM/YYYY date from any spreadsheet row,
 * prioritizing payment/OB/confirmation resolution dates over older emission dates.
 */
function extractRowDate(row: Record<string, any>): { dateObj: Date; dateStr: string } | null {
  let bestPaymentDate: { dateObj: Date; dateStr: string } | null = null;
  let latestAnyDate: { dateObj: Date; dateStr: string } | null = null;

  for (const [k, v] of Object.entries(row)) {
    if (v === null || v === undefined || v === '') continue;
    const lk = k.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');

    const isPaymentDateCol =
      lk.includes('pagamento') ||
      lk.includes('pgto') ||
      lk.includes('ob') ||
      lk.includes('confirm') ||
      lk.includes('situac') ||
      lk.includes('baixa') ||
      lk.includes('efetiv');

    const isDateCol =
      isPaymentDateCol ||
      lk.includes('data') ||
      lk.includes('dt') ||
      lk.includes('emiss') ||
      lk.includes('refer') ||
      lk.includes('liquid');

    if (isDateCol) {
      const dateInfo = parseDateToMonthInfo(v);
      if (dateInfo.dateIso) {
        const d = new Date(dateInfo.dateIso + 'T12:00:00');
        if (!isNaN(d.getTime())) {
          const [y, m, day] = dateInfo.dateIso.split('-');
          const cand = { dateObj: d, dateStr: `${day}/${m}/${y}` };

          if (isPaymentDateCol) {
            if (!bestPaymentDate || d.getTime() > bestPaymentDate.dateObj.getTime()) {
              bestPaymentDate = cand;
            }
          }
          if (!latestAnyDate || d.getTime() > latestAnyDate.dateObj.getTime()) {
            latestAnyDate = cand;
          }
        }
      }
    }
  }

  return bestPaymentDate || latestAnyDate;
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

/**
 * Checks if a PP or OB is currently in an unpaid situation (open, pending, authorising, etc.)
 */
export function isSituacaoNaoPaga(situacaoRaw: any): boolean {
  if (situacaoRaw === null || situacaoRaw === undefined) return true;
  const str = String(situacaoRaw)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toUpperCase();
  if (!str) return true;
  return (
    str === 'AO' ||
    str.startsWith('AO') ||
    str === 'AP' ||
    str.startsWith('AP') ||
    str === 'PENDENTE' ||
    str.includes('PENDENTE') ||
    str.includes('LIQUIDAD') ||
    str.includes('ABERTO') ||
    str.includes('AGUARDANDO') ||
    str.includes('PROCESSAMENTO') ||
    str.includes('ANDAMENTO') ||
    (!isSituacaoConclusivaPagamento(str) && !isSituacaoRejeitada(str))
  );
}

/**
 * Checks if a situation represents "AO" (Associada 2 Ordenadores) or is empty/unspecified.
 */
export function isSituacaoAOOuVazia(situacaoRaw: any): boolean {
  if (situacaoRaw === null || situacaoRaw === undefined) return true;
  const str = String(situacaoRaw)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toUpperCase();
  if (
    !str ||
    str === '-' ||
    str === '0' ||
    str === 'N/I' ||
    str === 'NULL' ||
    str === 'UNDEFINED' ||
    str === 'NAO INFORMADO' ||
    str === 'NÃO INFORMADO'
  ) {
    return true;
  }
  return isSituacaoAO(str) || str === 'AO' || str.startsWith('AO ') || str.startsWith('AO-') || str.startsWith('AO/');
}

/**
 * Checks if a situation represents a confirmed payment: "PPCB", "CB", or "PAGO"
 */
export function isSituacaoConfirmadaBanco(situacaoRaw: any): { isMatch: boolean; normalized: string } {
  if (situacaoRaw === null || situacaoRaw === undefined) return { isMatch: false, normalized: '' };
  const str = String(situacaoRaw)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toUpperCase();
  if (!str) return { isMatch: false, normalized: '' };

  if (str === 'PPCB' || str.startsWith('PPCB')) {
    return { isMatch: true, normalized: 'PPCB' };
  }
  if (str === 'CB' || str.startsWith('CB ') || str.startsWith('CB-') || str.startsWith('CB/') || str.startsWith('CB_')) {
    return { isMatch: true, normalized: 'CB' };
  }
  if (
    str === 'CONFIRMADA' ||
    str === 'CONFIRMADO' ||
    str.startsWith('CONFIRMADA') ||
    str.startsWith('CONFIRMADO') ||
    str.includes('CONFIRMADA') ||
    str.includes('CONFIRMADO') ||
    str.includes('CONFIRM')
  ) {
    return { isMatch: true, normalized: 'CONFIRMADA' };
  }
  if (
    str === 'PAGO' ||
    str === 'PAGA' ||
    str.startsWith('PAGO') ||
    str.startsWith('PAGA') ||
    str.startsWith('PAGTO') ||
    str.includes('PAGO NO BANCO') ||
    str.includes('PAGA NO BANCO') ||
    str.includes('PAGAMENTO CONFIRMADO') ||
    str.includes('PAGAMENTO EFETIVADO') ||
    str.includes('EFETIVADO NO BANCO') ||
    str.includes('EFETIVADA NO BANCO') ||
    str.includes('COMPENSAD') ||
    str.includes('CREDITAD')
  ) {
    return { isMatch: true, normalized: 'PAGO' };
  }

  if (str === 'CM' || str.startsWith('CM ') || str.startsWith('CM-') || str.includes('CONFIRMADA MANUAL')) {
    return { isMatch: true, normalized: 'CM' };
  }

  return { isMatch: false, normalized: '' };
}

/**
 * Checks if a situation represents a rejected payment: "PPRJ", "RJ", or "C" (Cancelada)
 */
export function isSituacaoRejeitadaBanco(situacaoRaw: any): { isMatch: boolean; normalized: string } {
  if (situacaoRaw === null || situacaoRaw === undefined) return { isMatch: false, normalized: '' };
  const str = String(situacaoRaw)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toUpperCase();
  if (!str) return { isMatch: false, normalized: '' };

  if (str === 'PPRJ' || str.startsWith('PPRJ')) {
    return { isMatch: true, normalized: 'PPRJ' };
  }
  if (str === 'RJ' || str.startsWith('RJ ') || str.startsWith('RJ-') || str.startsWith('RJ/') || str.startsWith('RJ_')) {
    return { isMatch: true, normalized: 'RJ' };
  }
  if (
    str === 'C' ||
    str.startsWith('C ') ||
    str.startsWith('C-') ||
    str.startsWith('C/') ||
    str.startsWith('C_') ||
    str === 'CANCELADA' ||
    str === 'CANCELADO' ||
    str.startsWith('CANCELAD') ||
    str.includes('CANCELAD') ||
    str.includes('REJEITAD') ||
    str.includes('DEVOLVID') ||
    str.includes('ESTORNAD') ||
    str.includes('ANULAD') ||
    str.includes('CBNC')
  ) {
    return { isMatch: true, normalized: 'C' };
  }

  return { isMatch: false, normalized: '' };
}

/**
 * Determines whether all 4 base files have been uploaded into the system:
 * 1. Base de PPs (stage 'base_pps' / 'pp_prontas')
 * 2. Liquidações - Obedece (stage 'obedece')
 * 3. Liquidações - Não Obedece (stage 'nao_obedece')
 * 4. Relatório de Ordens Bancárias / Pagamentos (ordensBancariasList)
 */
export function hasAll4BaseUploads(
  files: FileData[],
  ordensBancariasList?: OrdemBancariaItem[] | null
): boolean {
  if (!files || files.length === 0) return false;
  const hasBasePPs = files.some(
    (f) => f.stage === 'base_pps' || f.stage === 'pp_prontas' || Boolean(f.detectedPpCol)
  );
  const hasObedece = files.some((f) => f.stage === 'obedece');
  const hasNaoObedece = files.some((f) => f.stage === 'nao_obedece');
  const hasOBs = Boolean(
    (ordensBancariasList && ordensBancariasList.length > 0) ||
    files.some(
      (f) =>
        f.stage === 'ordem_bancaria' ||
        f.stage === 'outros' ||
        Boolean(f.detectedObCol && f.detectedSituacaoCol)
    )
  );

  return hasBasePPs && hasObedece && hasNaoObedece && hasOBs;
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
        if (Array.isArray(parsed)) {
          const seen = new Set<string>();
          const unique: ToastAlert[] = [];
          const nowMs = Date.now();
          const MAX_AGE_MS = 180 * 24 * 60 * 60 * 1000; // 180 days max age

          for (const item of parsed) {
            if (!item || typeof item !== 'object') continue;
            // Filter out legacy "Virou OB" alerts from previous versions
            if (
              item.situacaoNova === 'Virou OB' ||
              (item.titulo && item.titulo.toLowerCase().includes('virou ob')) ||
              (item.titulo && item.titulo.toLowerCase().includes('ordem bancária gerada')) ||
              (item.titulo && item.titulo.toLowerCase().includes('ordem bancaria gerada'))
            ) {
              continue;
            }

            // Filter out old historical events from beginning of the year or older than 45 days
            if (item.dataReferencia) {
              const dInfo = parseDateToMonthInfo(item.dataReferencia);
              if (dInfo.dateIso) {
                const itemTime = new Date(dInfo.dateIso + 'T12:00:00').getTime();
                if (!isNaN(itemTime) && nowMs - itemTime > MAX_AGE_MS) {
                  continue;
                }
              }
            }

            const key = getAlertUniqueKey(item);
            if (!seen.has(key)) {
              seen.add(key);
              const d = item.dataHora ? new Date(item.dataHora) : new Date();

              // Sanitize item.valor and item.mensagem if it was corrupted by a Fonte code (e.g. 600, 500, 0.60, etc.)
              let itemVal = typeof item.valor === 'number' ? item.valor : 0;
              if (
                isFonteString(itemVal) ||
                itemVal === 600 ||
                itemVal === 500 ||
                itemVal === 1500 ||
                itemVal === 1600 ||
                itemVal === 0.6 ||
                itemVal === 0.5
              ) {
                itemVal = 0;
              }
              let itemMsg = item.mensagem || '';
              if (itemVal === 0 && itemMsg) {
                // Strip corrupted (R$ 600,00), (R$ 500,00), (R$ 0,60), etc. from alert message text
                itemMsg = itemMsg
                  .replace(/\s*\(\s*R\$\s*600(?:,00)?\s*\)/gi, '')
                  .replace(/\s*\(\s*R\$\s*500(?:,00)?\s*\)/gi, '')
                  .replace(/\s*\(\s*R\$\s*1\.?500(?:,00)?\s*\)/gi, '')
                  .replace(/\s*\(\s*R\$\s*1\.?600(?:,00)?\s*\)/gi, '')
                  .replace(/\s*\(\s*R\$\s*0,[56]0?\s*\)/gi, '');
              }

              unique.push({
                ...item,
                valor: itemVal,
                mensagem: itemMsg,
                dataHora: !isNaN(d.getTime()) ? d : new Date(),
              });
            }
          }
          unique.sort((a, b) => new Date(b.dataHora).getTime() - new Date(a.dataHora).getTime());
          // Cap history to 1000 items to prevent UI and storage flood
          this.alertsHistory = unique.slice(0, 1000);
          this.saveAlertsHistory();
        }
      }
    } catch (e) {
      this.alertsHistory = [];
    }
  }

  private saveAlertsHistory() {
    try {
      localStorage.setItem(ALERTS_HISTORY_KEY, JSON.stringify(this.alertsHistory.slice(0, 1000)));
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
      this.alertsHistory = [...newItems, ...this.alertsHistory].slice(0, 2500);
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

  public restoreAlertsHistory(alerts: ToastAlert[]) {
    if (!Array.isArray(alerts) || alerts.length === 0) return;
    const existingKeys = new Set(this.alertsHistory.map((t) => getAlertUniqueKey(t)));
    const newItems = alerts.filter((a) => !existingKeys.has(getAlertUniqueKey(a)));
    this.alertsHistory = [...newItems, ...this.alertsHistory].slice(0, 1000);
    this.saveAlertsHistory();
    this.notify();
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

  public confirmTransition(id: string) {
    const toast = this.alertsHistory.find((t) => t.id === id);
    if (!toast) return;

    toast.resolvido = true;
    toast.requerValidacao = false;
    toast.situacaoNova = 'Confirmado pelo Usuário';
    toast.tipo = 'sucesso';
    toast.titulo = `Transição Confirmada: NL ${toast.nl || ''} → PP ${toast.pp || ''}`;

    if (toast.inferenceMatchId) {
      mlInferenceService.confirmMatch(toast.inferenceMatchId);
    }

    this.saveAlertsHistory();
    this.dismissToast(id);
    this.notify();
  }

  public rejectTransition(id: string) {
    const toast = this.alertsHistory.find((t) => t.id === id);
    if (!toast) return;

    toast.resolvido = true;
    toast.requerValidacao = false;
    toast.situacaoNova = 'Descartado pelo Usuário';
    toast.tipo = 'info';
    toast.titulo = `Transição Descartada: NL ${toast.nl || ''}`;

    if (toast.inferenceMatchId) {
      mlInferenceService.rejectMatch(toast.inferenceMatchId);
    }

    this.saveAlertsHistory();
    this.dismissToast(id);
    this.notify();
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

  public ensureBaseline(files: FileData[], ordensBancariasList?: OrdemBancariaItem[] | null) {
    if (!this.hasBaseline() && files && files.length > 0) {
      const records = this.extractAllCurrentRecords(files, ordensBancariasList);
      this.saveBaselineSnapshot(records);
    }
  }

  private loadBaselineSnapshot(): Record<string, any> | null {
    try {
      const raw = localStorage.getItem(BASELINE_SNAPSHOT_KEY);
      if (!raw) return null;
      const parsed = JSON.parse(raw);
      if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) return null;
      const res: Record<string, any> = {};
      for (const [k, v] of Object.entries(parsed as Record<string, any>)) {
        if (!v || typeof v !== 'object') continue;
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
      // Save compact representation of up to 50000 keys to avoid quota/memory issues
      const compact: Record<string, { s: string; ob?: string; c: boolean; r: boolean }> = {};
      const keys = Object.keys(records).slice(0, 50000);
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
      fonteRecurso: string;
      isFonte600: boolean;
      stage?: string;
      etapaOrigem?: string;
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
        let valNum = typeof obItem.valor === 'number' ? obItem.valor : 0;
        if (
          isFonteString(valNum) ||
          valNum === 600 ||
          valNum === 500 ||
          valNum === 1500 ||
          valNum === 1600 ||
          valNum === 0.6 ||
          valNum === 0.5
        ) {
          valNum = 0;
        }
        const sitStr = (obItem.situacao || 'CB').trim().toUpperCase();
        const isRejected = isSituacaoRejeitada(sitStr);
        const isConclusivo = isSituacaoConclusivaPagamento(sitStr) || sitStr === 'CB' || !isRejected;

        const fonteRecurso = (obItem.fonteRecurso || '').trim() || 'OUTRAS';
        const isFonte600 = isFonte600OuDerivada(fonteRecurso);

        const key = ppNum
          ? `PP_${ppNum}`
          : obNum
          ? `OB_${obNum}`
          : cnpjNum
          ? `CNPJ_${cnpjNum}_${Math.round(valNum)}`
          : `FORN_${cleanDocStr(obItem.favorecidoName)}_${Math.round(valNum)}`;

        const itemRecord = {
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
          fonteRecurso,
          isFonte600,
          stage: 'pagamentos_emitidos',
          dataReferencia: obItem.dataReferencia,
          fileName: 'Listar Ordem Bancária',
        };

        records[key] = itemRecord;
      });
    }

    // 2. Extract from Files
    files.forEach((file) => {
      const sitCol = file.detectedSituacaoCol;
      const sitObCol = file.detectedSituacaoObCol;
      const numCol = file.detectedNumeroCol || file.detectedPpCol;
      const fonteCol = file.detectedFonteCol || file.headers?.find((k) => isFonteHeader(k));
      let valCol = file.detectedValorCol;
      if (valCol && (isFonteHeader(valCol) || valCol === fonteCol)) {
        valCol = '';
      }
      if (!valCol && file.headers) {
        valCol = file.headers.find((k) => {
          if (isFonteHeader(k) || k === fonteCol) return false;
          const lk = k.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
          return (
            lk.includes('valor liquidado') ||
            lk.includes('valor da liquidacao') ||
            lk.includes('valor pago') ||
            lk.includes('valor doc') ||
            lk.includes('valor') ||
            lk.includes('vlr')
          );
        });
      }
      const favCol = file.detectedFavorecidoCol;
      const obCol = file.detectedObCol;
      const neCol = file.detectedNotaEmpenhoCol;

      file.rows.forEach((row) => {
        let rawSit = '';
        if (sitObCol && row[sitObCol] !== undefined && row[sitObCol] !== null && String(row[sitObCol]).trim() !== '') {
          rawSit = String(row[sitObCol]).trim().toUpperCase();
        }
        if (!rawSit && sitCol && row[sitCol] !== undefined && row[sitCol] !== null) {
          rawSit = String(row[sitCol]).trim().toUpperCase();
        }
        if (!rawSit) {
          for (const [k, v] of Object.entries(row)) {
            const lk = k.toLowerCase();
            if (lk.includes('situa') || lk.includes('status') || lk === 'sigla' || lk === 'sit') {
              rawSit = String(v || '').trim().toUpperCase();
              if (rawSit) break;
            }
          }
        }

        // Check if any cell in the row explicitly indicates OB status (CONFIRMADA, CB, PPCB, RJ, etc.)
        if (!rawSit || isSituacaoAOOuVazia(rawSit)) {
          for (const [k, v] of Object.entries(row)) {
            if (!v) continue;
            const lk = k.toLowerCase();
            if (lk.includes('ob') || lk.includes('pagamento') || lk.includes('status') || lk.includes('situac')) {
              const vStr = String(v).trim().toUpperCase();
              if (isSituacaoConfirmadaBanco(vStr).isMatch || isSituacaoRejeitadaBanco(vStr).isMatch) {
                rawSit = vStr;
                break;
              }
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

        // Fonte extraction
        let rawFonte = fonteCol ? String(row[fonteCol] || '').trim() : '';
        if (!rawFonte || rawFonte === '-' || rawFonte === '0') {
          rawFonte = findFonteRecursoInRow(Object.values(row));
        }

        const isFonte600 = isFonte600OuDerivada(rawFonte);

        let rawVal = valCol ? row[valCol] : 0;
        if (
          isFonteString(rawVal) ||
          rawVal === rawFonte ||
          (typeof rawVal === 'number' && (rawVal === 600 || rawVal === 500 || rawVal === 1500 || rawVal === 1600 || rawVal === 0.6 || rawVal === 0.5))
        ) {
          rawVal = 0;
        }
        let val = typeof rawVal === 'number' ? rawVal : parseCurrencyValue(rawVal);
        if (
          val >= 1e11 ||
          isFonteString(val) ||
          val === 600 ||
          val === 500 ||
          val === 1500 ||
          val === 1600 ||
          val === 0.6 ||
          val === 0.5
        ) {
          val = 0;
        }

        // If val is 0, find genuine monetary value from other columns
        if (val <= 0) {
          for (const [k, v] of Object.entries(row)) {
            if (!v || isFonteHeader(k) || k === fonteCol || k === valCol || isFonteString(v)) continue;
            const lk = k.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
            if (lk.includes('valor') || lk.includes('liquid') || lk.includes('pago') || lk.includes('total')) {
              const pv = parseCurrencyValue(v);
              if (pv > 0 && pv < 1e11 && !isFonteString(v) && pv !== 600 && pv !== 500 && pv !== 1500 && pv !== 1600 && pv !== 0.6 && pv !== 0.5) {
                val = pv;
                break;
              }
            }
          }
        }

        const rawFav = favCol ? row[favCol] : '';
        const { name: favName, cnpj: favCnpj } = parseFavorecidoField(rawFav);
        const cnpjNum = (favCnpj || '').replace(/\D/g, '');
        const obNum = isValidOBString(rawOB) ? rawOB.trim().toUpperCase() : '';
        const ppNum = cleanDocStr(rawNum);
        const neNum = cleanDocStr(rawNE);

        const isRejected = isSituacaoRejeitada(rawSit) || isSituacaoRejeitadaBanco(rawSit).isMatch;
        const isConclusivo = isSituacaoConclusivaPagamento(rawSit) || isSituacaoConfirmadaBanco(rawSit).isMatch || (!!obNum && !isRejected);

        const rowDateInfo = extractRowDate(row);

        const key = ppNum
          ? `PP_${ppNum}`
          : obNum
          ? `OB_${obNum}`
          : neNum
          ? `NE_${neNum}`
          : cnpjNum
          ? `CNPJ_${cnpjNum}_${Math.round(val)}`
          : `FORN_${cleanDocStr(favName)}_${Math.round(val)}`;

        const fileRecord = {
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
          fonteRecurso: rawFonte || 'OUTRAS',
          isFonte600,
          stage: file.stage,
          dataReferencia: rowDateInfo?.dateStr,
          fileName: file.fileName,
        };

        const mergeRecord = (existing?: any) => {
          if (!existing) return fileRecord;
          // Never overwrite real expense value with 0 or a fonte code
          const trueVal = val > 0 ? val : (existing.valor > 0 ? existing.valor : 0);
          const trueOb = obNum || existing.ob || '';
          const truePP = ppNum || existing.pp || '';
          const trueNE = neNum || existing.ne || '';
          const trueFav = favName || existing.fornecedor || '';
          const trueCnpj = favCnpj || existing.cnpj || '';
          const trueFonte = (rawFonte && rawFonte !== 'OUTRAS') ? rawFonte : (existing.fonteRecurso || 'OUTRAS');
          const trueIsF600 = isFonte600 || existing.isFonte600;

          const isBetterSit =
            (isSituacaoAOOuVazia(existing.situacao) && !isSituacaoAOOuVazia(rawSit)) ||
            (isConclusivo && !existing.isConclusivo) ||
            (!existing.isConclusivo && Boolean(rawSit && rawSit !== existing.situacao));

          return {
            ...existing,
            ...fileRecord,
            valor: trueVal,
            ob: trueOb,
            pp: truePP,
            ne: trueNE,
            fornecedor: trueFav,
            cnpj: trueCnpj,
            fonteRecurso: trueFonte,
            isFonte600: trueIsF600,
            situacao: isBetterSit ? rawSit : (existing.situacao || rawSit),
            isConclusivo: existing.isConclusivo || isConclusivo,
            isRejected: existing.isRejected || isRejected,
          };
        };

        records[key] = mergeRecord(records[key]);
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
   * Monitora estritamente a resolução bancária de Ordens Bancárias / PPs preexistentes:
   * 1. Alerta verde (Sucesso): transição de situação pendente (AO ou vazia) para Confirmada (PPCB, CB, PAGO).
   * 2. Alerta vermelho (Crítico): transição de situação pendente (AO ou vazia) para Rejeitada (PPRJ, RJ, C).
   * Transições heurísticas de NL -> PP/OB (dedução de impostos) pertencem exclusivamente à Auditoria IA e NÃO disparam Toasts.
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

    // Se ainda não existia um snapshot prévio gravado, salva o estado atual como ponto de partida
    // e não dispara notificações retroativas/em massa no primeiro carregamento
    if (!previousSnapshot && !forceRefresh) {
      this.saveBaselineSnapshot(currentRecords);
      return;
    }

    const toastsToBatch: Array<
      Omit<ToastAlert, 'id' | 'dataHora' | 'resolvido'> & {
        id?: string;
        dataHora?: Date;
        resolvido?: boolean;
        showPopup?: boolean;
      }
    > = [];
    const batchNotifiedKeys = new Set<string>();
    let countUpdates = 0;

    // Ordenar as chaves para que os registros com data mais recente (ex: Março/Fevereiro) sejam processados e alertados primeiro
    const currentKeys = Object.keys(currentRecords);
    currentKeys.sort((kA, kB) => {
      const rA = currentRecords[kA];
      const rB = currentRecords[kB];
      const isoA = rA?.dataReferencia ? parseDateToMonthInfo(rA.dataReferencia).dateIso || '' : '';
      const isoB = rB?.dataReferencia ? parseDateToMonthInfo(rB.dataReferencia).dateIso || '' : '';
      if (isoA && isoB) {
        return isoB.localeCompare(isoA); // Data mais recente primeiro
      }
      if (isoB) return 1;
      if (isoA) return -1;
      return 0;
    });

    // O(1) Pre-index of alerted history documents to eliminate quadratic overhead on large imports
    const historyConfDocs = new Set<string>();
    const historyRejDocs = new Set<string>();
    const historyDocSituations = new Set<string>();

    for (const hist of this.alertsHistory) {
      const hPP = cleanDocStr(hist.pp);
      const hOB = cleanDocStr(hist.ob);
      const hSit = (hist.situacaoNova || '').trim().toUpperCase();
      const isConf = hist.tipo === 'sucesso' || isSituacaoConfirmadaBanco(hSit).isMatch;
      const isRej = hist.tipo === 'critico' || isSituacaoRejeitadaBanco(hSit).isMatch;

      if (hOB) {
        if (hSit) historyDocSituations.add(`OB_${hOB}_${hSit}`);
        if (isConf) historyConfDocs.add(`OB_${hOB}`);
        if (isRej) historyRejDocs.add(`OB_${hOB}`);
      }
      if (hPP) {
        if (hSit) historyDocSituations.add(`PP_${hPP}_${hSit}`);
        if (isConf) historyConfDocs.add(`PP_${hPP}`);
        if (isRej) historyRejDocs.add(`PP_${hPP}`);
      }
    }

    for (const key of currentKeys) {
      const curr = currentRecords[key];
      if (!curr || typeof curr !== 'object') continue;

      // Ignora eventos históricos ocorridos há mais de 180 dias ou no início do ano
      if (curr.dataReferencia) {
        const dInfo = parseDateToMonthInfo(curr.dataReferencia);
        if (dInfo.dateIso) {
          const itemTime = new Date(dInfo.dateIso + 'T12:00:00').getTime();
          if (!isNaN(itemTime) && Date.now() - itemTime > 180 * 24 * 60 * 60 * 1000) {
            continue;
          }
        }
      }

      const currPP = cleanDocStr(curr.pp);
      const currOB = cleanDocStr(curr.ob);

      // É necessário ter ao menos identificação de OB ou PP para monitorar a resolução bancária
      if (!currOB && !currPP) continue;

      let prev = previousSnapshot ? previousSnapshot[key] : undefined;
      if (!prev && previousSnapshot && currPP) prev = previousSnapshot[`PP_${currPP}`];
      if (!prev && previousSnapshot && currOB) prev = previousSnapshot[`OB_${currOB}`];

      const prevSituacao = prev ? (prev.situacao || '').trim().toUpperCase() : '';
      const currSit = (curr.situacao || '').trim().toUpperCase();
      if (!currSit) continue;

      const confInfo = isSituacaoConfirmadaBanco(currSit);
      const rejInfo = isSituacaoRejeitadaBanco(currSit);
      const prevConf = isSituacaoConfirmadaBanco(prevSituacao);
      const prevRej = isSituacaoRejeitadaBanco(prevSituacao);

      // Se já estava confirmada no banco antes e continua confirmada (ex: CB -> PPCB), não alertar novamente
      if (prev && prevConf.isMatch && confInfo.isMatch) continue;
      if (prev && prevRej.isMatch && rejInfo.isMatch) continue;

      const hasSituationChanged = prev ? prevSituacao !== currSit : false;
      if (prev && !hasSituationChanged && !forceRefresh) continue;
      if (!prev && !forceRefresh) continue;

      const situacaoNova = confInfo.isMatch
        ? currSit || confInfo.normalized
        : rejInfo.isMatch
        ? currSit || rejInfo.normalized
        : currSit;
      const situacaoNovaNorm = situacaoNova.trim().toUpperCase();

      // Bloqueio estrito de duplicidade no histórico (mesmo em forceRefresh)
      let isAlreadyInHistory = false;
      if (currOB) {
        if (historyDocSituations.has(`OB_${currOB}_${situacaoNovaNorm}`)) isAlreadyInHistory = true;
        if (confInfo.isMatch && historyConfDocs.has(`OB_${currOB}`)) isAlreadyInHistory = true;
        if (rejInfo.isMatch && historyRejDocs.has(`OB_${currOB}`)) isAlreadyInHistory = true;
      }
      if (!isAlreadyInHistory && currPP) {
        if (historyDocSituations.has(`PP_${currPP}_${situacaoNovaNorm}`)) isAlreadyInHistory = true;
        if (confInfo.isMatch && historyConfDocs.has(`PP_${currPP}`)) isAlreadyInHistory = true;
        if (rejInfo.isMatch && historyRejDocs.has(`PP_${currPP}`)) isAlreadyInHistory = true;
      }
      if (isAlreadyInHistory) continue;

      const batchGroup = confInfo.isMatch ? 'CONF_BANCO' : rejInfo.isMatch ? 'REJ_BANCO' : situacaoNovaNorm;
      const bKeyOB = currOB ? `${currOB}__${batchGroup}` : '';
      const bKeyPP = currPP ? `${currPP}__${batchGroup}` : '';

      if ((bKeyOB && batchNotifiedKeys.has(bKeyOB)) || (bKeyPP && batchNotifiedKeys.has(bKeyPP))) continue;
      if (bKeyOB) batchNotifiedKeys.add(bKeyOB);
      if (bKeyPP) batchNotifiedKeys.add(bKeyPP);

      const docDesc = curr.ob && curr.pp ? `OB ${curr.ob} (PP ${curr.pp})` : curr.ob ? `OB ${curr.ob}` : `PP ${curr.pp}`;
      const isValidExpenseVal =
        typeof curr.valor === 'number' &&
        curr.valor > 0 &&
        curr.valor < 1e11 &&
        !isFonteString(curr.valor) &&
        ![600, 500, 1500, 1600, 0.6, 0.5].includes(curr.valor);

      const valorFormatado = isValidExpenseVal ? ` (${formatBRL(curr.valor)})` : '';
      const finalValor = isValidExpenseVal ? curr.valor : 0;
      const sitAnteriorLabel = prevSituacao || 'AO';

      if (confInfo.isMatch) {
        toastsToBatch.push({
          id: `toast_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`,
          tipo: 'sucesso' as const,
          titulo: 'Pagamento Confirmado no Banco',
          mensagem: `${docDesc} de ${curr.fornecedor || 'Credor'}${valorFormatado} mudou de ${sitAnteriorLabel} para ${situacaoNova}.`,
          fornecedor: curr.fornecedor,
          cnpj: curr.cnpj,
          ob: curr.ob,
          pp: curr.pp,
          ne: curr.ne,
          valor: finalValor,
          fonteRecurso: curr.fonteRecurso,
          isFonte600: curr.isFonte600,
          etapaOrigem: curr.etapaOrigem,
          situacaoAnterior: sitAnteriorLabel,
          situacaoNova,
          dataReferencia: curr.dataReferencia,
          dataHora: new Date(),
          showPopup: countUpdates < 5,
        });
        countUpdates++;
        continue;
      }

      if (rejInfo.isMatch) {
        toastsToBatch.push({
          id: `toast_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`,
          tipo: 'critico' as const,
          titulo: 'Pagamento Rejeitado / Cancelado',
          mensagem: `Atenção: ${docDesc} de ${curr.fornecedor || 'Credor'}${valorFormatado} mudou de ${sitAnteriorLabel} para ${situacaoNova}.`,
          fornecedor: curr.fornecedor,
          cnpj: curr.cnpj,
          ob: curr.ob,
          pp: curr.pp,
          ne: curr.ne,
          valor: finalValor,
          fonteRecurso: curr.fonteRecurso,
          isFonte600: curr.isFonte600,
          etapaOrigem: curr.etapaOrigem,
          situacaoAnterior: sitAnteriorLabel,
          situacaoNova,
          dataReferencia: curr.dataReferencia,
          dataHora: new Date(),
          showPopup: countUpdates < 5,
        });
        countUpdates++;
        continue;
      }

      // Qualquer outra alteração de situação (AO -> LI, EI, LD, EN, II, SD, SI, PPNC, AP -> AO, REPP, RENE, etc.)
      toastsToBatch.push({
        id: `toast_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`,
        tipo: 'info' as const,
        titulo: 'Alteração de Situação da PP / OB',
        mensagem: `${docDesc} de ${curr.fornecedor || 'Credor'}${valorFormatado} mudou de situação: ${sitAnteriorLabel} → ${situacaoNova}.`,
        fornecedor: curr.fornecedor,
        cnpj: curr.cnpj,
        ob: curr.ob,
        pp: curr.pp,
        ne: curr.ne,
        valor: finalValor,
        fonteRecurso: curr.fonteRecurso,
        isFonte600: curr.isFonte600,
        etapaOrigem: curr.etapaOrigem,
        situacaoAnterior: sitAnteriorLabel,
        situacaoNova,
        dataReferencia: curr.dataReferencia,
        dataHora: new Date(),
        showPopup: countUpdates < 5,
      });
      countUpdates++;
    }

    if (toastsToBatch.length > 0) {
      this.batchAddToasts(toastsToBatch);
    }

    // Salva o snapshot atual para a próxima verificação
    this.saveBaselineSnapshot(currentRecords);
    this.notify();
  }
}

export const toastAlertService = new ToastAlertManager();
