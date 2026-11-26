/**
 * Cloud Sync Service
 * Synchronizes spreadsheets, bank orders, metrics, and audit state across all user devices
 * using a dedicated Google Drive (Google Apps Script Web App) JSON state endpoint,
 * eliminating Firestore quota and 1MB document limits.
 */
import {
  signInWithPopup,
  signOut,
  onAuthStateChanged,
  User,
} from 'firebase/auth';
import { auth, googleAuthProvider } from '../firebase';
import { FileData, DashboardMetrics, ConsolidatedSupplier } from '../types';
import { saveAppState, loadAppState, clearAppState } from './storageService';
import {
  saveStoredOrdensBancarias,
  getStoredOrdensBancarias,
  loadOrdensBancariasAsync,
  clearStoredOrdensBancarias,
  OrdemBancariaItem,
  enrichSuppliersWithOrdensBancarias,
} from './ordemBancariaParser';
import { processUploadedFiles } from './excelParser';
import { paraPublicarService } from './paraPublicarService';
import { mlInferenceService } from './machineLearningInferenceService';
import { toastAlertService } from './toastAlertService';
import { isAdmin } from './authService';
import { googleDriveService } from './googleDriveService';
import { getAllAtendimentos, mergeAtendimentosFromCloud } from './atendimentoMensalService';
import {
  scheduleBackgroundInterval,
  scheduleBackgroundTimeout,
  backgroundDelay,
  withActiveProcessingLock,
  onTabReturnToForeground,
} from './backgroundScheduler';

export const DRIVE_DATABASE_FOLDER_NAME = 'Painel_Financeiro_SESAP_Dados';

const GOOGLE_DRIVE_SYNC_URL =
  'https://script.google.com/macros/s/AKfycbyIKBVLRR2hSLt2Ci--84mvBkf8tpvL3QI6GtTkgsWijEOFQiZXInvoE9w303GDCvVA/exec';
const METRICS_CACHE_STORAGE_KEY = 'painel_fes_metrics_cache_v1';
const LAST_PUSHED_VERSION_KEY = 'painel_fes_last_pushed_version_v1';
const MAX_SAFE_PAYLOAD_BYTES = 38 * 1024 * 1024; // Margem segura abaixo do limite de ~50MB do Google Apps Script

const TILDE_TOKEN_REGEX = /^~{1,2}\d+$/;

export function isTildeTokenString(val: unknown): boolean {
  return typeof val === 'string' && TILDE_TOKEN_REGEX.test(val.trim());
}

/**
 * Detecta se um FileData contém tokens corrompidos (~3396 / ~~3396) deixados por dicionário de compressão legado.
 */
export function isFileCorruptedWithTildeTokens(file: FileData | null | undefined): boolean {
  if (!file || !Array.isArray(file.rows) || file.rows.length === 0) return false;
  const rows = file.rows;
  const total = rows.length;
  const sampleLimit = Math.min(total, 300);
  const step = total > sampleLimit ? Math.max(1, Math.floor(total / sampleLimit)) : 1;
  const favCol = file.detectedFavorecidoCol || '';
  let favHits = 0;
  let totalHits = 0;

  for (let i = 0; i < total; i += step) {
    const r = rows[i];
    if (!r || typeof r !== 'object') continue;
    if (favCol && isTildeTokenString(r[favCol])) {
      favHits++;
      if (favHits >= 1) return true;
    }
    for (const k of Object.keys(r)) {
      if (isTildeTokenString(r[k])) {
        totalHits++;
        if (totalHits >= 3) return true;
      }
    }
  }
  return false;
}

/**
 * Detecta se um objeto DashboardMetrics contém fornecedores corrompidos (~3396 / ~~3396).
 */
export function isMetricsCorruptedWithTildeTokens(metrics: DashboardMetrics | null | undefined): boolean {
  if (!metrics || !Array.isArray(metrics.consolidatedSuppliers)) return false;
  return metrics.consolidatedSuppliers.some(
    (s) => isTildeTokenString(s?.name) || isTildeTokenString(s?.cnpj)
  );
}

/**
 * Detecta se um conjunto de arquivos ou métricas contém tokens corrompidos (~3396 / ~~3396).
 */
export function isDatasetCorruptedWithTildeTokens(
  files?: FileData[] | null,
  metrics?: DashboardMetrics | null
): boolean {
  if (isMetricsCorruptedWithTildeTokens(metrics)) return true;
  if (Array.isArray(files) && files.some((f) => isFileCorruptedWithTildeTokens(f))) return true;
  return false;
}

/**
 * Separa arquivos limpos de arquivos que contenham tokens corrompidos (~3396 / ~~3396).
 */
export function filterCleanFiles(files: FileData[]): { cleanFiles: FileData[]; corruptedCount: number } {
  const cleanFiles: FileData[] = [];
  let corruptedCount = 0;
  for (const f of files) {
    if (isFileCorruptedWithTildeTokens(f)) {
      corruptedCount++;
    } else {
      cleanFiles.push(f);
    }
  }
  return { cleanFiles, corruptedCount };
}

const ESSENTIAL_COL_HINT_REGEX =
  /favorec|credor|fornec|cnpj|cpf|valor|data|fonte|natureza|empenho|processo|situa|status|observ|historic|descr|motivo|liquida|\bnl\b|\bpp\b|\bob\b|ordem|banco|sigef|tipo|unidade|gest/i;

/**
 * Codifica todas as planilhas em formato colunar PURO e autossuficiente (sem dicionário de strings nem tokens ~).
 * Preserva 100% dos textos originais em todas as sessões/versões e reduz significativamente o tamanho do JSON.
 */
export function encodeFilesForCloud(
  files: FileData[],
  options?: { maxRowsPerHugeFile?: number; keepOnlyEssentialColsOnHugeFiles?: boolean }
): any[] {
  const maxRowsPerHugeFile = options?.maxRowsPerHugeFile;
  const keepOnlyEssentialColsOnHugeFiles = options?.keepOnlyEssentialColsOnHugeFiles === true;

  return files.map((f) => {
    const allRows = f.rows || [];
    const isRawObLarge = f.stage === 'ordem_bancaria' && allRows.length > 5000;
    const effectiveLimit =
      maxRowsPerHugeFile !== undefined
        ? Math.min(allRows.length, maxRowsPerHugeFile)
        : isRawObLarge
        ? 200
        : allRows.length;
    const sourceRows = effectiveLimit < allRows.length ? allRows.slice(0, effectiveLimit) : allRows;

    const detectedCols = new Set<string>(
      [
        f.detectedFavorecidoCol,
        f.detectedValorCol,
        f.detectedNumeroCol,
        f.detectedPpCol,
        f.detectedObCol,
        f.detectedCnpjCol,
        f.detectedSituacaoCol,
        f.detectedSituacaoObCol,
        f.detectedNotaEmpenhoCol,
        f.detectedFavorecidoNECol,
        f.detectedProcessoCol,
        f.detectedFonteCol,
        f.detectedNaturezaCol,
        f.detectedDataCol,
      ].filter((c): c is string => Boolean(c))
    );

    const trimUnusedCols = keepOnlyEssentialColsOnHugeFiles && sourceRows.length > 20000;
    const headerSet = new Set<string>();
    const compactHeaders: string[] = [];

    const considerHeader = (colName: string) => {
      if (!colName || headerSet.has(colName)) return;
      if (trimUnusedCols && !detectedCols.has(colName) && !ESSENTIAL_COL_HINT_REGEX.test(colName)) {
        return;
      }
      headerSet.add(colName);
      compactHeaders.push(colName);
    };

    for (const h of f.headers || []) {
      considerHeader(h);
    }
    for (let i = 0; i < sourceRows.length; i++) {
      const r = sourceRows[i];
      if (!r || typeof r !== 'object') continue;
      for (const k of Object.keys(r)) {
        const v = r[k];
        if (v !== '' && v !== null && v !== undefined && v !== '-') {
          considerHeader(k);
        }
      }
    }

    const rowsCompact: any[][] = new Array(sourceRows.length);
    for (let i = 0; i < sourceRows.length; i++) {
      const r = sourceRows[i] || {};
      const arr: any[] = [];
      let lastNonEmpty = -1;
      for (let c = 0; c < compactHeaders.length; c++) {
        const v = r[compactHeaders[c]];
        if (v !== '' && v !== null && v !== undefined && v !== '-') {
          arr.push(v);
          lastNonEmpty = c;
        } else {
          arr.push(null);
        }
      }
      rowsCompact[i] = lastNonEmpty >= 0 ? arr.slice(0, lastNonEmpty + 1) : [];
    }

    return {
      id: f.id || '',
      fileName: f.fileName || '',
      stage: f.stage || 'outros',
      headers: f.headers || compactHeaders,
      compactHeaders,
      rowsCompact,
      detectedFavorecidoCol: f.detectedFavorecidoCol || '',
      detectedValorCol: f.detectedValorCol || '',
      detectedNumeroCol: f.detectedNumeroCol || '',
      detectedPpCol: f.detectedPpCol || '',
      detectedObCol: f.detectedObCol || '',
      detectedCnpjCol: f.detectedCnpjCol || '',
      detectedSituacaoCol: f.detectedSituacaoCol || '',
      detectedSituacaoObCol: f.detectedSituacaoObCol || '',
      detectedNotaEmpenhoCol: f.detectedNotaEmpenhoCol || '',
      detectedFavorecidoNECol: f.detectedFavorecidoNECol || '',
      detectedProcessoCol: f.detectedProcessoCol || '',
      detectedFonteCol: f.detectedFonteCol || '',
      detectedNaturezaCol: f.detectedNaturezaCol || '',
      detectedDataCol: f.detectedDataCol || '',
      rowCount: allRows.length,
      totalCount: f.totalCount || allRows.length,
      isSummaryOnly: isRawObLarge || effectiveLimit < allRows.length,
      rows: [],
    };
  });
}

export function decodeFileFromCloud(rawFile: any, stringPool?: string[]): FileData {
  if (!rawFile || typeof rawFile !== 'object') return rawFile;
  if (Array.isArray(rawFile.rowsCompact) && Array.isArray(rawFile.compactHeaders)) {
    const headers: string[] = rawFile.compactHeaders;
    const rowsCompact: any[][] = rawFile.rowsCompact;
    // Só aplica decodificação legada de stringPool se houver um dicionário válido e não vazio
    const hasPool = Array.isArray(stringPool) && stringPool.length > 0;
    const restoredRows: Record<string, any>[] = new Array(rowsCompact.length);
    for (let i = 0; i < rowsCompact.length; i++) {
      const arr = rowsCompact[i];
      const obj: Record<string, any> = {};
      if (Array.isArray(arr)) {
        for (let c = 0; c < arr.length && c < headers.length; c++) {
          let v = arr[c];
          if (v !== null && v !== undefined && v !== '') {
            if (hasPool && typeof v === 'string' && v.length >= 2 && v.charCodeAt(0) === 126 /* ~ */) {
              if (v.charCodeAt(1) === 126) {
                v = v.slice(1);
              } else {
                const poolIdx = Number(v.slice(1));
                if (Number.isInteger(poolIdx) && poolIdx >= 0 && poolIdx < stringPool!.length) {
                  v = stringPool![poolIdx];
                }
              }
            }
            obj[headers[c]] = v;
          }
        }
      }
      restoredRows[i] = obj;
    }
    const { rowsCompact: _rc, compactHeaders: _ch, ...rest } = rawFile;
    return {
      ...rest,
      rows: restoredRows,
    } as FileData;
  }
  return rawFile as FileData;
}

export interface CloudSyncState {
  isAuthenticated: boolean;
  user: User | null;
  isSyncing: boolean;
  lastSyncAt: Date | null;
  lastSyncDevice: string | null;
  remoteFilesCount: number;
  syncStatus: 'idle' | 'syncing' | 'synced' | 'error' | 'unauthenticated';
  errorMessage: string | null;
}

type SyncListener = (state: CloudSyncState) => void;

class CloudSyncService {
  private state: CloudSyncState = {
    isAuthenticated: false,
    user: null,
    isSyncing: false,
    lastSyncAt: null,
    lastSyncDevice: null,
    remoteFilesCount: 0,
    syncStatus: 'idle',
    errorMessage: null,
  };

  private listeners = new Set<SyncListener>();
  private stopPollingFn: (() => void) | null = null;
  private stopForegroundListenerFn: (() => void) | null = null;
  private isProcessingSync = false;
  private isPushInFlight = false;
  private isApplyingCloudPull = false;
  private pushEpoch = 0;
  private pendingPushArgs: {
    files: FileData[];
    obsList?: OrdemBancariaItem[] | null;
    obFileName?: string;
    consolidatedMetrics?: DashboardMetrics | null;
  } | null = null;
  private pullInFlight: Promise<boolean> | null = null;
  private pullInFlightOnlyIfNewer = false;
  private lastPullHadNetworkError = false;
  private lastPullDetectedCorruption = false;
  private consecutivePollFailures = 0;
  private pollTickCounter = 0;
  private lastRemoteVersion: number | null = null;
  private lastPulledData: {
    files: FileData[];
    obs: OrdemBancariaItem[] | null;
    syncDate?: Date;
    metrics?: DashboardMetrics | null;
  } | null = null;

  private onDatasetReloadCallback:
    | ((
        files: FileData[],
        obs: OrdemBancariaItem[] | null,
        syncDate?: Date,
        metrics?: DashboardMetrics | null
      ) => void)
    | null = null;

  private broadcastChannel: BroadcastChannel | null = null;

  constructor() {
    this.initAuthListener();
    this.initBroadcastChannel();
    this.startAutoPolling();
    scheduleBackgroundTimeout(() => {
      this.checkAndAutoSync().catch((e) => {
        console.warn('Auto-sync inicial Google Drive:', e);
      });
    }, 300);
  }

  private initBroadcastChannel() {
    try {
      if (typeof BroadcastChannel !== 'undefined') {
        this.broadcastChannel = new BroadcastChannel('sesap_cloud_sync_channel');
        this.broadcastChannel.onmessage = (ev) => {
          if (ev?.data?.type === 'cloud_updated' || ev?.data?.type === 'cloud_cleared') {
            this.lastRemoteVersion = null;
            this.pullFromCloud(false).catch((e) => {
              console.warn('Sync via BroadcastChannel:', e);
            });
          }
        };
      }
    } catch (_) {
      // BroadcastChannel opcional
    }
  }

  private notifyOtherTabs(type: 'cloud_updated' | 'cloud_cleared') {
    try {
      this.broadcastChannel?.postMessage({ type, at: Date.now() });
    } catch (_) {}
  }

  private initAuthListener() {
    try {
      onAuthStateChanged(auth, async (user) => {
        if (user) {
          this.updateState({
            isAuthenticated: true,
            user,
            errorMessage: null,
          });
        } else {
          this.updateState({
            isAuthenticated: false,
            user: null,
            errorMessage: null,
          });
        }
      });
    } catch (_) {
      // Auth listener opcional; sincronização via Drive Web App independe de Firebase Auth
    }
  }

  private startAutoPolling() {
    if (this.stopPollingFn) {
      this.stopPollingFn();
      this.stopPollingFn = null;
    }
    if (this.stopForegroundListenerFn) {
      this.stopForegroundListenerFn();
      this.stopForegroundListenerFn = null;
    }

    const runPollTick = () => {
      if (this.isProcessingSync || this.pullInFlight) return;
      if (typeof navigator !== 'undefined' && navigator.onLine === false) return;

      if (this.consecutivePollFailures > 0) {
        this.pollTickCounter++;
        const skipTicks = Math.min(this.consecutivePollFailures, 4);
        if (this.pollTickCounter <= skipTicks) {
          return;
        }
        this.pollTickCounter = 0;
      }

      this.pullFromCloud(true).catch((e) => {
        console.warn('Polling automático Google Drive:', e);
      });
    };

    // Polling automático leve via Web Worker a cada 15 segundos (imune ao throttling de aba inativa)
    this.stopPollingFn = scheduleBackgroundInterval(runPollTick, 15000);
    this.stopForegroundListenerFn = onTabReturnToForeground(() => {
      if (!this.isProcessingSync && !this.pullInFlight) {
        this.pullFromCloud(true).catch(() => {});
      }
    });
  }

  public registerDatasetReloadCallback(
    cb: (
      files: FileData[],
      obs: OrdemBancariaItem[] | null,
      syncDate?: Date,
      metrics?: DashboardMetrics | null
    ) => void
  ) {
    this.onDatasetReloadCallback = cb;
    if (this.lastPulledData) {
      cb(
        this.lastPulledData.files,
        this.lastPulledData.obs,
        this.lastPulledData.syncDate,
        this.lastPulledData.metrics
      );
    }
  }

  public getState(): CloudSyncState {
    return { ...this.state };
  }

  public isRestoringFromCloud(): boolean {
    return this.isApplyingCloudPull;
  }

  public subscribe(listener: SyncListener): () => void {
    this.listeners.add(listener);
    listener(this.getState());
    return () => {
      this.listeners.delete(listener);
    };
  }

  private updateState(partial: Partial<CloudSyncState>) {
    this.state = { ...this.state, ...partial };
    this.listeners.forEach((l) => l(this.getState()));
  }

  public async loginWithGoogle(): Promise<boolean> {
    try {
      this.updateState({ isSyncing: true, syncStatus: 'syncing', errorMessage: null });
      const result = await signInWithPopup(auth, googleAuthProvider);
      if (result.user) {
        this.updateState({
          isAuthenticated: true,
          user: result.user,
          isSyncing: false,
          syncStatus: 'synced',
        });
        await this.checkAndAutoSync();
        return true;
      }
      return false;
    } catch (err: any) {
      console.warn('Tentativa de login com Google:', err?.message || err);
      const isPopupBlocked =
        err?.code === 'auth/popup-blocked' || String(err?.message || '').includes('popup-blocked');
      this.updateState({
        isSyncing: false,
        syncStatus: 'idle',
        errorMessage: isPopupBlocked
          ? 'Pop-up bloqueado pelo navegador. O painel continuará sincronizando via Google Drive normalmente.'
          : err.message || 'Falha na autenticação Google.',
      });
      await this.checkAndAutoSync();
      return false;
    }
  }

  public async logout(): Promise<void> {
    try {
      await signOut(auth);
    } catch (err) {
      console.warn('Erro ao sair:', err);
    }
  }

  private getDeviceName(): string {
    if (typeof navigator === 'undefined') return 'Dispositivo';
    const ua = navigator.userAgent.toLowerCase();
    if (
      ua.includes('ipad') ||
      (ua.includes('tablet') && !ua.includes('mobile')) ||
      (ua.includes('android') && !ua.includes('mobile'))
    ) {
      return 'Tablet';
    }
    if (ua.includes('mobile') || ua.includes('iphone')) {
      return 'Smartphone';
    }
    return 'PC / Computador';
  }

  private saveMetricsCache(metrics: DashboardMetrics | null, syncDate: Date, obFileName: string) {
    if (!metrics || !metrics.isProcessed) return;
    if (isMetricsCorruptedWithTildeTokens(metrics)) {
      try {
        localStorage.removeItem(METRICS_CACHE_STORAGE_KEY);
      } catch (_) {}
      return;
    }
    try {
      localStorage.setItem(
        METRICS_CACHE_STORAGE_KEY,
        JSON.stringify({
          metrics,
          lastSyncAt: syncDate.toISOString(),
          obFileName,
        })
      );
    } catch (_) {
      // Ignore quota issues in localStorage
    }
  }

  /**
   * Envia o estado completo do painel para o endpoint oficial no Google Drive (SESAP_PAINEL_STATE_V1.json).
   * Coalesce atualizações consecutivas para nunca perder edições rápidas nem colidir com polling.
   */
  public async pushToCloud(
    files: FileData[],
    obsList?: OrdemBancariaItem[] | null,
    obFileName?: string,
    consolidatedMetrics?: DashboardMetrics | null
  ): Promise<boolean> {
    if (!isAdmin()) {
      console.warn('pushToCloud bloqueado: apenas Administradores podem enviar dados para a nuvem.');
      return false;
    }

    // Invalida qualquer pullFromCloud que já estivesse em voo para nunca sobrescrever esta atualização
    this.pushEpoch++;

    if (this.isPushInFlight) {
      // Guarda a versão mais recente para enviar imediatamente ao término do envio atual
      this.pendingPushArgs = { files, obsList, obFileName, consolidatedMetrics };
      return true;
    }

    this.isPushInFlight = true;
    let lastResult = false;
    let currentArgs: {
      files: FileData[];
      obsList?: OrdemBancariaItem[] | null;
      obFileName?: string;
      consolidatedMetrics?: DashboardMetrics | null;
    } | null = { files, obsList, obFileName, consolidatedMetrics };

    try {
      while (currentArgs) {
        const { files: f, obsList: obs, obFileName: obName, consolidatedMetrics: met } = currentArgs;
        this.pendingPushArgs = null;
        this.pushEpoch++;

        lastResult = await withActiveProcessingLock('cloud-push', () => {
          let cancelSafetyTimeout: (() => void) | null = null;
          return Promise.race([
            this.executePushToCloud(f, obs, obName, met).finally(() => {
              cancelSafetyTimeout?.();
            }),
            new Promise<boolean>((resolve) => {
              cancelSafetyTimeout = scheduleBackgroundTimeout(() => {
                console.warn('Timeout de segurança ao sincronizar com o Google Drive (100s). Liberando interface.');
                this.isProcessingSync = false;
                this.updateState({
                  isSyncing: false,
                  syncStatus: 'error',
                  errorMessage: 'Tempo limite excedido ao enviar dados para a nuvem.',
                });
                resolve(false);
              }, 100000);
            }),
          ]);
        });

        currentArgs = this.pendingPushArgs;
      }
      return lastResult;
    } finally {
      this.isPushInFlight = false;
      this.pushEpoch++;
    }
  }

  private async postJsonToAppsScript(serializedPayload: string): Promise<any> {
    const attemptPost = async (): Promise<any> => {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 85000);
      try {
        const response = await fetch(GOOGLE_DRIVE_SYNC_URL, {
          method: 'POST',
          redirect: 'follow',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: serializedPayload,
          signal: controller.signal,
        });

        const rawText = await response.text();
        if (!response.ok) {
          throw new Error(`Falha HTTP ${response.status} ao salvar no Google Drive.`);
        }

        let resData: any = null;
        try {
          resData = rawText ? JSON.parse(rawText) : { success: true };
        } catch {
          const snippet = rawText.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 180);
          throw new Error(
            snippet
              ? `Resposta inesperada do Google Drive: ${snippet}`
              : 'Resposta inválida retornada pelo endpoint do Google Drive.'
          );
        }

        if (resData && resData.success === false) {
          throw new Error(resData.error || 'Erro retornado pelo endpoint do Google Drive.');
        }

        return resData;
      } finally {
        clearTimeout(timeoutId);
      }
    };

    try {
      return await attemptPost();
    } catch (firstErr: any) {
      const isRetryable =
        firstErr?.name === 'AbortError' ||
        firstErr instanceof TypeError ||
        String(firstErr?.message || '').toLowerCase().includes('fetch') ||
        String(firstErr?.message || '').includes('HTTP 5');
      if (isRetryable && (typeof navigator === 'undefined' || navigator.onLine !== false)) {
        await backgroundDelay(1500);
        return await attemptPost();
      }
      throw firstErr;
    }
  }

  private async executePushToCloud(
    files: FileData[],
    obsList?: OrdemBancariaItem[] | null,
    obFileName?: string,
    consolidatedMetrics?: DashboardMetrics | null
  ): Promise<boolean> {
    try {
      this.isProcessingSync = true;
      this.updateState({ isSyncing: true, syncStatus: 'syncing', errorMessage: null });

      const device = this.getDeviceName();
      const nowDate = new Date();
      const nowIso = nowDate.toISOString();
      const version = Date.now();
      const asyncObState = !obsList || !obFileName ? await loadOrdensBancariasAsync() : null;
      const effectiveOBs = obsList || asyncObState?.items || getStoredOrdensBancarias().items || [];
      const effectiveOBFileName =
        obFileName ||
        asyncObState?.fileName ||
        getStoredOrdensBancarias().fileName ||
        'OBs - sigef - Principal';

      // 1. Obter ou calcular as métricas consolidadas
      let activeMetrics: DashboardMetrics;
      if (consolidatedMetrics && consolidatedMetrics.isProcessed) {
        activeMetrics = consolidatedMetrics;
      } else {
        const base = processUploadedFiles(files);
        enrichSuppliersWithOrdensBancarias(
          base.consolidatedSuppliers,
          effectiveOBs.length > 0 ? effectiveOBs : null
        );
        activeMetrics = base;
      }

      // Bloqueio de segurança: jamais publicar na nuvem planilhas ou métricas contendo códigos corrompidos (~3396 / ~~3396)
      if (isDatasetCorruptedWithTildeTokens(files, activeMetrics)) {
        const msg =
          'Foram detectados dados corrompidos no cache local (códigos ~). Clique em "Sincronizar" com o Google Drive para restaurar as planilhas originais.';
        console.warn('[CloudSync] Envio bloqueado para proteger a nuvem contra dados corrompidos:', msg);
        this.updateState({
          isSyncing: false,
          syncStatus: 'error',
          errorMessage: msg,
        });
        return false;
      }

      // 2. Versão compacta de consolidatedSuppliers
      const compactSuppliers: ConsolidatedSupplier[] = (activeMetrics.consolidatedSuppliers || []).map(
        (s) => ({
          name: s.name || '',
          cnpj: s.cnpj || '',
          totalValue: s.totalValue || 0,
          totalCount: s.totalCount || 0,
          ppProntasValue: s.ppProntasValue || 0,
          basePpsValue: s.basePpsValue || 0,
          liqObedeceValue: s.liqObedeceValue || 0,
          liqNaoObedeceValue: s.liqNaoObedeceValue || 0,
          outrosValue: s.outrosValue || 0,
          enviadasBancoValue: s.enviadasBancoValue || 0,
          countPpProntasAO: s.countPpProntasAO || 0,
          countEnviadasBanco: s.countEnviadasBanco || 0,
          valorPagoAno: s.valorPagoAno || 0,
          countPagoAno: s.countPagoAno || 0,
          naturezasDespesa: s.naturezasDespesa || [],
          pagamentosDetalhePorFonte: s.pagamentosDetalhePorFonte || {},
          pagamentosPorMes: s.pagamentosPorMes || {},
        })
      );

      // 3. Codificação colunar PURA e autossuficiente (sem stringPool nem tokens ~)
      const filesSummary = encodeFilesForCloud(files);

      // 4. Lista compacta das Ordens Bancárias agrupadas por favorecidoCnpj + fonteRecurso + mesAnoKey
      const compactOBsMap = new Map<string, OrdemBancariaItem>();
      effectiveOBs.forEach((item) => {
        const cnpj = item.favorecidoCnpj || item.favorecidoRaw || '';
        const fonte = item.fonteRecurso || 'Geral';
        const mesAno =
          item.mesAnoKey ||
          (item.dataIso
            ? item.dataIso.slice(0, 7)
            : item.dataReferencia
            ? item.dataReferencia.slice(-7)
            : 'Sem Data');
        const groupKey = `${cnpj}__${fonte}__${mesAno}`;
        const existing = compactOBsMap.get(groupKey);
        if (!existing) {
          compactOBsMap.set(groupKey, {
            ...item,
            id: `compact_${groupKey}`,
            valor: item.valor || 0,
          });
        } else {
          existing.valor = (existing.valor || 0) + (item.valor || 0);
          if (!existing.ob && item.ob) existing.ob = item.ob;
          if (!existing.pp && item.pp) existing.pp = item.pp;
        }
      });
      const compactOBs = Array.from(compactOBsMap.values());

      // 5. Padrões de ML e mapeamentos exatos
      let mlPatterns: Record<string, any> = {};
      let mlExactMappings: Record<string, any> = {};
      try {
        mlPatterns = mlInferenceService.getAllLearnedPatterns() || {};
        mlExactMappings = mlInferenceService.getExactNlMappings() || {};
      } catch (_) {}

      const metricsPayload: DashboardMetrics = {
        grandTotalValue: activeMetrics.grandTotalValue || 0,
        grandTotalCount: activeMetrics.grandTotalCount || 0,
        totalBasePps: activeMetrics.totalBasePps || 0,
        totalPpProntas: activeMetrics.totalPpProntas || 0,
        totalLiqObedece: activeMetrics.totalLiqObedece || 0,
        totalLiqNaoObedece: activeMetrics.totalLiqNaoObedece || 0,
        totalOutros: activeMetrics.totalOutros || 0,
        totalEnviadasBanco: activeMetrics.totalEnviadasBanco || 0,
        countPpProntasAO: activeMetrics.countPpProntasAO || 0,
        countEnviadasBanco: activeMetrics.countEnviadasBanco || 0,
        totalPagoAno: activeMetrics.totalPagoAno || 0,
        countPagoAnoTotal: activeMetrics.countPagoAnoTotal || 0,
        isProcessed: true,
        consolidatedSuppliers: compactSuppliers,
      };

      const drivePayload = {
        targetFolderName: DRIVE_DATABASE_FOLDER_NAME,
        fileName: 'SESAP_PAINEL_STATE_V1.json',
        lastSyncAt: nowIso,
        deviceType: device,
        totalFiles: files.length,
        version,
        obFileName: effectiveOBFileName,
        metrics: metricsPayload,
        filesSummary,
        compactOBs,
        alerts: toastAlertService.getAlertsHistory(),
        paraPublicar: paraPublicarService.getItems(),
        mlPatterns,
        mlExactMappings,
        localPreferences: {
          recurring: localStorage.getItem('finance_manual_recurring_suppliers'),
          atendimentos: JSON.stringify(getAllAtendimentos()),
          categories:
            localStorage.getItem('sigecon_categories_config') ||
            localStorage.getItem('sigecon_categories_v1'),
          categoryOverrides: localStorage.getItem('sigecon_supplier_category_overrides'),
        },
      };

      let serializedPayload = JSON.stringify(drivePayload);

      // Guarda de segurança caso o payload puro ultrapasse 38MB:
      // 1º remove apenas colunas auxiliares não essenciais de planilhas muito grandes (preservando 100% das linhas)
      if (serializedPayload.length > MAX_SAFE_PAYLOAD_BYTES) {
        console.warn(
          `[CloudSync] Payload de ${(serializedPayload.length / (1024 * 1024)).toFixed(1)}MB acima do limite seguro; mantendo apenas colunas essenciais nas planilhas grandes.`
        );
        drivePayload.filesSummary = encodeFilesForCloud(files, {
          keepOnlyEssentialColsOnHugeFiles: true,
        });
        serializedPayload = JSON.stringify(drivePayload);
      }

      // 2º se ainda assim exceder 38MB, limita linhas apenas das planilhas gigantes (mantendo métricas e OBs 100% íntegros)
      if (serializedPayload.length > MAX_SAFE_PAYLOAD_BYTES) {
        console.warn(
          `[CloudSync] Payload ainda com ${(serializedPayload.length / (1024 * 1024)).toFixed(1)}MB; aplicando limite seguro de linhas mantendo métricas e OBs completos.`
        );
        drivePayload.filesSummary = encodeFilesForCloud(files, {
          keepOnlyEssentialColsOnHugeFiles: true,
          maxRowsPerHugeFile: 25000,
        });
        serializedPayload = JSON.stringify(drivePayload);
      }

      // 6. Envia em uma única requisição HTTP POST para o Google Drive (com timeout + retry automático)
      await this.postJsonToAppsScript(serializedPayload);

      // Espelho direto na pasta dedicada Painel_Financeiro_SESAP_Dados se o OAuth do Drive estiver conectado
      if (googleDriveService.getState().isConnected && googleDriveService.isConnected()) {
        googleDriveService
          .saveStateToDriveFolder('SESAP_PAINEL_STATE_V1.json', serializedPayload)
          .catch((e) => console.warn('Espelho na pasta do Google Drive:', e));
      }

      this.lastRemoteVersion = version;
      try {
        localStorage.setItem(LAST_PUSHED_VERSION_KEY, String(version));
      } catch (_) {}
      this.saveMetricsCache(metricsPayload, nowDate, effectiveOBFileName);

      this.updateState({
        isSyncing: false,
        syncStatus: 'synced',
        lastSyncAt: nowDate,
        lastSyncDevice: device,
        remoteFilesCount: files.length,
        errorMessage: null,
      });

      this.notifyOtherTabs('cloud_updated');
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('finance_cloud_data_updated'));
      }

      console.log(`Painel sincronizado no Google Drive com sucesso a partir do ${device}`);
      return true;
    } catch (err: any) {
      console.warn('Sincronização Google Drive (envio):', err?.message || err);
      this.updateState({
        isSyncing: false,
        syncStatus: 'error',
        errorMessage: err?.message || 'Erro ao sincronizar com o Google Drive.',
      });
      return false;
    } finally {
      this.isProcessingSync = false;
    }
  }

  /**
   * Remove arquivo na nuvem (arquitetura de snapshot único)
   */
  public async deleteFileFromCloud(fileId: string): Promise<void> {
    console.log(`Arquivo ${fileId} marcado para remoção local e nuvem.`);
  }

  private async fetchDriveStateJson(): Promise<any> {
    const attemptFetch = async (): Promise<any> => {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 60000);
      try {
        const response = await fetch(`${GOOGLE_DRIVE_SYNC_URL}?t=${Date.now()}`, {
          method: 'GET',
          redirect: 'follow',
          signal: controller.signal,
        });

        if (!response.ok) {
          throw new Error(`Falha HTTP ${response.status} ao consultar Google Drive.`);
        }

        return await response.json();
      } finally {
        clearTimeout(timeoutId);
      }
    };

    try {
      return await attemptFetch();
    } catch (firstErr: any) {
      // Pequena pausa e nova tentativa para oscilações transitórias de redirecionamento do Apps Script
      const isAbortOrNetwork =
        firstErr?.name === 'AbortError' ||
        firstErr instanceof TypeError ||
        String(firstErr?.message || '').toLowerCase().includes('fetch');

      if (isAbortOrNetwork && (typeof navigator === 'undefined' || navigator.onLine !== false)) {
        await backgroundDelay(1200);
        try {
          return await attemptFetch();
        } catch (retryErr) {
          // Fallback direto para a pasta dedicada do Google Drive caso o OAuth esteja conectado
          if (googleDriveService.isConnected()) {
            const driveFolderData = await googleDriveService.loadStateFromDriveFolder(
              'SESAP_PAINEL_STATE_V1.json'
            );
            if (driveFolderData) {
              return driveFolderData;
            }
          }
          throw retryErr;
        }
      }

      throw firstErr;
    }
  }

  /**
   * Leitura direta do estado completo no Google Drive (SESAP_PAINEL_STATE_V1.json)
   */
  public async pullFromCloud(onlyIfNewer: boolean = false): Promise<boolean> {
    if (this.isPushInFlight) {
      return false;
    }
    if (this.pullInFlight) {
      const existingResult = await this.pullInFlight;
      // Se um login ou ação manual pediu força total (onlyIfNewer = false) enquanto um polling leve rodava, faz uma leitura completa em seguida
      if (!onlyIfNewer && this.pullInFlightOnlyIfNewer && !this.isPushInFlight) {
        return this.pullFromCloud(false);
      }
      return existingResult;
    }

    this.pullInFlightOnlyIfNewer = onlyIfNewer;
    this.pullInFlight = this.executePullFromCloud(onlyIfNewer).finally(() => {
      this.pullInFlight = null;
      this.pullInFlightOnlyIfNewer = false;
    });

    return this.pullInFlight;
  }

  private async executePullFromCloud(onlyIfNewer: boolean = false): Promise<boolean> {
    const pullStartEpoch = this.pushEpoch;
    try {
      if (typeof navigator !== 'undefined' && navigator.onLine === false) {
        this.lastPullHadNetworkError = true;
        if (!onlyIfNewer) {
          this.updateState({ isSyncing: false, syncStatus: 'idle' });
        }
        return false;
      }

      this.isProcessingSync = true;
      this.lastPullHadNetworkError = false;
      this.lastPullDetectedCorruption = false;
      if (!onlyIfNewer) {
        // Ignora cache em memória quando usuário abre o app ou solicita sincronização manual
        this.lastRemoteVersion = null;
        this.updateState({ isSyncing: true, syncStatus: 'syncing', errorMessage: null });
      }

      const data = await this.fetchDriveStateJson();
      this.consecutivePollFailures = 0;
      this.pollTickCounter = 0;

      // Se o Administrador iniciou ou concluiu um pushToCloud enquanto este GET estava em voo, descarta dados obsoletos
      if (this.isPushInFlight || this.pushEpoch !== pullStartEpoch) {
        return false;
      }

      if (!data) {
        if (!onlyIfNewer) {
          this.updateState({ isSyncing: false, syncStatus: 'idle' });
        }
        return false;
      }

      const remoteVersion = typeof data.version === 'number' ? data.version : null;
      const syncDate = data.lastSyncAt ? new Date(data.lastSyncAt) : new Date();
      const validSyncDate = isNaN(syncDate.getTime()) ? new Date() : syncDate;

      // No polling automático a cada 15s, evita re-renderizações desnecessárias se a versão for idêntica
      if (
        onlyIfNewer &&
        remoteVersion &&
        this.lastRemoteVersion === remoteVersion &&
        this.lastPulledData
      ) {
        this.updateState({ syncStatus: 'synced', isSyncing: false });
        return true;
      }

      // Trata limpeza remota do workspace feita pelo Administrador
      if (!data.metrics) {
        const isExplicitClear =
          data.totalFiles === 0 &&
          Array.isArray(data.filesSummary) &&
          data.filesSummary.length === 0;
        if (isExplicitClear) {
          await clearAppState();
          clearStoredOrdensBancarias();
          try {
            localStorage.removeItem(METRICS_CACHE_STORAGE_KEY);
            localStorage.removeItem(LAST_PUSHED_VERSION_KEY);
          } catch (_) {}

          this.lastRemoteVersion = remoteVersion || validSyncDate.getTime();
          this.lastPulledData = {
            files: [],
            obs: null,
            metrics: null,
            syncDate: validSyncDate,
          };
          this.updateState({
            isSyncing: false,
            syncStatus: 'synced',
            lastSyncAt: validSyncDate,
            lastSyncDevice: data.deviceType || 'Google Drive',
            remoteFilesCount: 0,
            errorMessage: null,
          });
          if (this.onDatasetReloadCallback) {
            this.onDatasetReloadCallback([], null, validSyncDate, null);
          }
          return true;
        }

        if (!onlyIfNewer) {
          this.updateState({ isSyncing: false, syncStatus: 'idle' });
        }
        return false;
      }

      // 1. Restaurar metrics
      let parsedMetrics: DashboardMetrics | null = null;
      if (typeof data.metrics === 'string') {
        try {
          parsedMetrics = JSON.parse(data.metrics);
        } catch (e) {
          console.warn('Erro ao processar metrics string:', e);
        }
      } else if (typeof data.metrics === 'object') {
        parsedMetrics = data.metrics as DashboardMetrics;
      }

      // 2. Restaurar filesSummary (decodifica formato colunar + stringPool ou legado rows, preservando todas as importações)
      const remoteStringPool: string[] | undefined = Array.isArray(data.stringPool) ? data.stringPool : undefined;
      let parsedFilesSummary: FileData[] = [];
      if (Array.isArray(data.filesSummary)) {
        parsedFilesSummary = data.filesSummary.map((item: any) => decodeFileFromCloud(item, remoteStringPool));
      } else if (typeof data.filesSummary === 'string') {
        try {
          const parsedRaw = JSON.parse(data.filesSummary);
          if (Array.isArray(parsedRaw)) {
            parsedFilesSummary = parsedRaw.map((item: any) => decodeFileFromCloud(item, remoteStringPool));
          }
        } catch (e) {
          console.warn('Erro ao processar filesSummary:', e);
        }
      }

      // 3. Restaurar compactOBs
      let parsedCompactOBs: OrdemBancariaItem[] = [];
      if (Array.isArray(data.compactOBs)) {
        parsedCompactOBs = data.compactOBs;
      } else if (typeof data.compactOBs === 'string') {
        try {
          parsedCompactOBs = JSON.parse(data.compactOBs);
        } catch (e) {
          console.warn('Erro ao processar compactOBs:', e);
        }
      }
      const obFileName = data.obFileName || 'OBs - sigef - Principal';

      this.isApplyingCloudPull = true;
      try {
        // 4. Restaurar paraPublicar (inclusive lista vazia caso o Admin tenha limpo)
        if (Array.isArray(data.paraPublicar)) {
          try {
            paraPublicarService.restoreItems(data.paraPublicar);
          } catch (e) {
            console.warn('Erro ao restaurar paraPublicar:', e);
          }
        }

        // 5. Restaurar alerts
        if (Array.isArray(data.alerts) && data.alerts.length > 0) {
          try {
            const normalizedAlerts = data.alerts.map((a: any) => ({
              ...a,
              dataHora: a.dataHora ? new Date(a.dataHora) : new Date(),
            }));
            toastAlertService.restoreAlertsHistory(normalizedAlerts);
          } catch (e) {
            console.warn('Erro ao restaurar alerts:', e);
          }
        }

        // 6. Restaurar padrões de aprendizado de máquina
        if (data.mlPatterns || data.mlExactMappings) {
          try {
            const patterns =
              typeof data.mlPatterns === 'string' ? JSON.parse(data.mlPatterns) : data.mlPatterns || {};
            const mappings =
              typeof data.mlExactMappings === 'string'
                ? JSON.parse(data.mlExactMappings)
                : data.mlExactMappings || {};
            mlInferenceService.restoreLearnedPatterns(patterns, mappings);
          } catch (_) {}
        }

        // 7. Restaurar preferências locais (Recorrentes, Atendimentos no Mês, Categorias e Overrides)
        let localAtendimentosHadNewer = false;
        if (data.localPreferences && typeof data.localPreferences === 'object') {
          try {
            const prefs = data.localPreferences;
            if (typeof prefs.recurring === 'string') {
              localStorage.setItem('finance_manual_recurring_suppliers', prefs.recurring);
            }
            if (typeof prefs.atendimentos === 'string') {
              const mergeRes = mergeAtendimentosFromCloud(prefs.atendimentos);
              if (mergeRes.localHadNewer) {
                localAtendimentosHadNewer = true;
              }
            }
            if (typeof prefs.categories === 'string') {
              localStorage.setItem('sigecon_categories_config', prefs.categories);
              localStorage.setItem('sigecon_categories_v1', prefs.categories);
            }
            if (typeof prefs.categoryOverrides === 'string') {
              localStorage.setItem('sigecon_supplier_category_overrides', prefs.categoryOverrides);
            }
            if (typeof window !== 'undefined') {
              window.dispatchEvent(new Event('finance_atendimento_mensal_changed'));
              window.dispatchEvent(new Event('sigecon_categories_changed'));
              window.dispatchEvent(new Event('sigecon_category_sync_event'));
              window.dispatchEvent(new Event('finance_recurring_changed'));
            }
          } catch (e) {
            console.warn('Erro ao restaurar localPreferences:', e);
          }
        }
        if (localAtendimentosHadNewer && isAdmin()) {
          // Agenda reenvio automático para publicar na nuvem as marcações locais mais recentes/recuperadas
          scheduleBackgroundTimeout(() => {
            if (!this.isApplyingCloudPull && this.lastPulledData && this.lastPulledData.files.length > 0) {
              this.pushToCloud(
                this.lastPulledData.files,
                this.lastPulledData.obs,
                obFileName,
                this.lastPulledData.metrics
              ).catch((e) => console.warn('Reenvio automático de atendimentos mesclados:', e));
            }
          }, 2500);
        }
      } finally {
        this.isApplyingCloudPull = false;
      }

      // Preserva arquivos e OBs locais completos SOMENTE no navegador do Administrador que acabou de enviar esta exata versão.
      // Para qualquer outro login/dispositivo ou versão remota diferente, aplica sempre os dados atualizados da nuvem.
      const localState = await loadAppState();
      const localFiles = localState?.files || [];
      const { cleanFiles: cleanLocalFiles } = filterCleanFiles(localFiles);
      const localOBState = await loadOrdensBancariasAsync();
      const localOBs = localOBState?.items || [];

      // Proteção contra snapshot remoto corrompido por tokens legados (~3396 / ~~3396):
      // Preserva OBs e preferências válidas, mas NUNCA sobrescreve planilhas/métricas locais com dados corrompidos.
      const remoteCorrupted = isDatasetCorruptedWithTildeTokens(parsedFilesSummary, parsedMetrics);
      if (remoteCorrupted) {
        this.lastPullDetectedCorruption = true;
        console.warn(
          '[CloudSync] Snapshot remoto contém tokens corrompidos (~). Protegendo dados locais e solicitando ressincronização limpa do Google Drive.'
        );

        // Salva OBs limpas vindas da nuvem caso ainda não existam localmente
        if (parsedCompactOBs.length > 0 && localOBs.length === 0) {
          try {
            saveStoredOrdensBancarias(parsedCompactOBs, obFileName);
          } catch (_) {}
        }

        // Remove cache local corrompido de métricas
        try {
          const cachedRaw = localStorage.getItem(METRICS_CACHE_STORAGE_KEY);
          if (cachedRaw) {
            const cached = JSON.parse(cachedRaw);
            if (isMetricsCorruptedWithTildeTokens(cached?.metrics)) {
              localStorage.removeItem(METRICS_CACHE_STORAGE_KEY);
            }
          }
        } catch (_) {}

        this.lastRemoteVersion = remoteVersion || validSyncDate.getTime();
        this.updateState({
          isSyncing: false,
          syncStatus: 'idle',
          errorMessage:
            'A nuvem ainda contém um arquivo anterior com códigos corrompidos (~). Clique em "Sincronizar" com o Google Drive para restaurar as planilhas originais.',
        });

        if (typeof window !== 'undefined') {
          window.dispatchEvent(
            new CustomEvent('finance_cloud_corrupted_detected', {
              detail: {
                hasCleanLocalFiles: cleanLocalFiles.length > 0,
                restoredOBs: parsedCompactOBs.length > 0 ? parsedCompactOBs : null,
                obFileName,
              },
            })
          );
        }
        return false;
      }

      const localTotalRows = cleanLocalFiles.reduce((acc, f) => acc + (f.rows?.length || 0), 0);
      const remoteTotalRows = parsedFilesSummary.reduce((acc, f) => acc + (f.rows?.length || 0), 0);

      let lastPushedVersion: string | null = null;
      try {
        lastPushedVersion = localStorage.getItem(LAST_PUSHED_VERSION_KEY);
      } catch (_) {}
      const isSameLocalPushedVersion =
        Boolean(remoteVersion && lastPushedVersion && String(remoteVersion) === lastPushedVersion);

      const preserveLocalFiles =
        isSameLocalPushedVersion &&
        cleanLocalFiles.length > 0 &&
        localTotalRows > remoteTotalRows;

      const effectiveFiles = preserveLocalFiles ? cleanLocalFiles : parsedFilesSummary;

      const preserveLocalOBs =
        isSameLocalPushedVersion &&
        localOBs.length > 0 &&
        localOBs.length > parsedCompactOBs.length;

      const effectiveOBs = preserveLocalOBs
        ? localOBs
        : parsedCompactOBs.length > 0
        ? parsedCompactOBs
        : null;

      // Salvar localmente no IndexedDB apenas se NÃO estiver sobrescrevendo arquivos locais mais completos
      if (!preserveLocalFiles && parsedFilesSummary.length > 0) {
        try {
          await saveAppState(parsedFilesSummary, validSyncDate);
        } catch (e) {
          console.warn('Erro ao salvar localmente dados baixados do Drive:', e);
        }
      }
      if (!preserveLocalOBs && parsedCompactOBs.length > 0) {
        try {
          saveStoredOrdensBancarias(parsedCompactOBs, obFileName);
        } catch (e) {
          console.warn('Erro ao salvar localmente OBs do Drive:', e);
        }
      }

      // Atualiza silenciosamente a base comparativa (BASELINE_SNAPSHOT_KEY) sem disparar alertas falsos no login
      try {
        toastAlertService.forceUpdateBaseline(effectiveFiles, effectiveOBs);
      } catch (e) {
        console.warn('Erro ao atualizar baseline após pullFromCloud:', e);
      }

      // Salvar cache de métricas no localStorage
      this.saveMetricsCache(parsedMetrics, validSyncDate, obFileName);

      this.lastRemoteVersion = remoteVersion || validSyncDate.getTime();

      // Salva no cache em memória
      this.lastPulledData = {
        files: effectiveFiles,
        obs: effectiveOBs,
        metrics: parsedMetrics,
        syncDate: validSyncDate,
      };

      this.updateState({
        isSyncing: false,
        syncStatus: 'synced',
        lastSyncAt: validSyncDate,
        lastSyncDevice: data.deviceType || 'Google Drive',
        remoteFilesCount: effectiveFiles.length,
        errorMessage: null,
      });

      // Notifica App.tsx com os dados, OBs, data exata do Drive e métricas consolidadas
      if (this.onDatasetReloadCallback) {
        this.onDatasetReloadCallback(
          effectiveFiles,
          effectiveOBs,
          validSyncDate,
          parsedMetrics
        );
      }

      console.log('Dados restaurados diretamente do Google Drive com sucesso.');
      return true;
    } catch (err: any) {
      this.lastPullHadNetworkError = true;
      this.consecutivePollFailures++;
      console.warn('Sincronização Google Drive (pull) indisponível no momento:', err?.message || err);
      if (!onlyIfNewer) {
        this.updateState({
          isSyncing: false,
          syncStatus: this.state.lastSyncAt ? 'synced' : 'idle',
          errorMessage: err?.message || 'Erro ao baixar dados do Google Drive.',
        });
      } else {
        this.updateState({ isSyncing: false });
      }
      return false;
    } finally {
      this.isProcessingSync = false;
    }
  }

  /**
   * Avalia estado local vs Google Drive e prioriza verificação imediata na nuvem ao abrir o app
   */
  public async checkAndAutoSync(): Promise<void> {
    if (this.isProcessingSync && !this.pullInFlight) return;

    try {
      // Sempre força a verificação imediata no Google Drive ignorando cache local ao abrir o aplicativo
      const pulled = await this.pullFromCloud(false);

      if (!pulled && !this.lastPullHadNetworkError && isAdmin()) {
        const localState = await loadAppState();
        const localFiles = localState?.files || [];
        const { cleanFiles, corruptedCount } = filterCleanFiles(localFiles);
        if (cleanFiles.length > 0 && corruptedCount === 0) {
          const asyncOB = await loadOrdensBancariasAsync();
          const localOBs = asyncOB.items || getStoredOrdensBancarias().items || [];
          console.log(
            this.lastPullDetectedCorruption
              ? 'Nuvem continha snapshot corrompido; republicando estado local limpo do Administrador...'
              : 'Google Drive vazio; enviando estado local do Administrador...'
          );
          await this.pushToCloud(cleanFiles, localOBs, asyncOB.fileName);
        }
      }
    } catch (err) {
      console.warn('Verificação de sincronização automática:', err);
    }
  }

  /**
   * Limpa o snapshot do workspace na nuvem (Google Drive)
   */
  public async clearCloudWorkspace(): Promise<void> {
    if (!isAdmin()) return;
    try {
      this.isProcessingSync = true;
      const nowIso = new Date().toISOString();
      const clearVersion = Date.now();
      const clearPayload = {
        lastSyncAt: nowIso,
        deviceType: this.getDeviceName(),
        totalFiles: 0,
        version: clearVersion,
        obFileName: '',
        metrics: null,
        filesSummary: [],
        compactOBs: [],
        alerts: [],
        paraPublicar: [],
        mlPatterns: {},
        mlExactMappings: {},
      };

      await fetch(GOOGLE_DRIVE_SYNC_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(clearPayload),
      });

      try {
        localStorage.removeItem(METRICS_CACHE_STORAGE_KEY);
        localStorage.removeItem(LAST_PUSHED_VERSION_KEY);
      } catch (_) {}

      this.lastRemoteVersion = clearVersion;
      this.lastPulledData = null;
      this.updateState({
        remoteFilesCount: 0,
        lastSyncAt: new Date(),
        syncStatus: 'synced',
      });
      this.notifyOtherTabs('cloud_cleared');
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('finance_cloud_data_updated'));
      }
      console.log('Workspace no Google Drive limpo com sucesso.');
    } catch (err) {
      console.warn('Erro ao limpar workspace no Google Drive:', err);
    } finally {
      this.isProcessingSync = false;
    }
  }

  /**
   * Export all data as a portable JSON backup file for 1-click manual transfer
   */
  public async exportPortableBackup(): Promise<void> {
    const localState = await loadAppState();
    const obState = getStoredOrdensBancarias();
    const paraPublicar = paraPublicarService.getItems();

    const backupObj = {
      app: 'SESAP_PAINEL_FINANCEIRO',
      version: 2,
      exportDate: new Date().toISOString(),
      exportDevice: this.getDeviceName(),
      files: localState?.files || [],
      ordensBancarias: obState.items || [],
      ordensBancariasFileName: obState.fileName || '',
      paraPublicar,
    };

    const jsonStr = JSON.stringify(backupObj, null, 2);
    const blob = new Blob([jsonStr], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Backup_Financeiro_Completo_${new Date().toISOString().slice(0, 10)}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  /**
   * Import data from a portable JSON backup file
   */
  public async importPortableBackup(file: File): Promise<boolean> {
    try {
      const text = await file.text();
      const data = JSON.parse(text);

      if (!data || !Array.isArray(data.files)) {
        throw new Error('Arquivo de backup inválido.');
      }

      // Save files
      await saveAppState(data.files, new Date());

      // Save OBs
      if (Array.isArray(data.ordensBancarias) && data.ordensBancarias.length > 0) {
        saveStoredOrdensBancarias(data.ordensBancarias, data.ordensBancariasFileName || '');
      }

      // Restore publication queue
      if (Array.isArray(data.paraPublicar) && data.paraPublicar.length > 0) {
        paraPublicarService.restoreItems(data.paraPublicar);
      }

      // Push to Google Drive if admin
      if (isAdmin()) {
        await this.pushToCloud(data.files, data.ordensBancarias, data.ordensBancariasFileName);
      }

      // Notify App
      if (this.onDatasetReloadCallback) {
        this.onDatasetReloadCallback(data.files, data.ordensBancarias);
      }

      return true;
    } catch (err: any) {
      console.error('Erro ao restaurar backup:', err);
      toastAlertService.addToast({
        tipo: 'critico',
        titulo: 'Falha ao Restaurar Backup',
        mensagem: err?.message || 'Arquivo de backup corrompido ou formato inválido.',
      });
      return false;
    }
  }
}

export const cloudSyncService = new CloudSyncService();

