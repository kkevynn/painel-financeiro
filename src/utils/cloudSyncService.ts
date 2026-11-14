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
import { saveAppState, loadAppState } from './storageService';
import {
  saveStoredOrdensBancarias,
  getStoredOrdensBancarias,
  OrdemBancariaItem,
  enrichSuppliersWithOrdensBancarias,
} from './ordemBancariaParser';
import { processUploadedFiles } from './excelParser';
import { paraPublicarService } from './paraPublicarService';
import { mlInferenceService } from './machineLearningInferenceService';
import { toastAlertService } from './toastAlertService';
import { isAdmin } from './authService';
import { googleDriveService } from './googleDriveService';

export const DRIVE_DATABASE_FOLDER_NAME = 'Painel_Financeiro_SESAP_Dados';

const GOOGLE_DRIVE_SYNC_URL =
  'https://script.google.com/macros/s/AKfycbyIKBVLRR2hSLt2Ci--84mvBkf8tpvL3QI6GtTkgsWijEOFQiZXInvoE9w303GDCvVA/exec';
const METRICS_CACHE_STORAGE_KEY = 'painel_fes_metrics_cache_v1';

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
  private pollingIntervalId: ReturnType<typeof setInterval> | null = null;
  private isProcessingSync = false;
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

  constructor() {
    this.initAuthListener();
    this.startAutoPolling();
    setTimeout(() => {
      this.checkAndAutoSync().catch((e) => {
        console.warn('Auto-sync inicial Google Drive:', e);
      });
    }, 300);
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
    if (this.pollingIntervalId) {
      clearInterval(this.pollingIntervalId);
    }
    // Polling automático leve a cada 60 segundos para manter todos os dispositivos sincronizados
    this.pollingIntervalId = setInterval(() => {
      if (!this.isProcessingSync) {
        this.pullFromCloud(true).catch((e) => {
          console.warn('Polling automático Google Drive:', e);
        });
      }
    }, 60000);
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
   * Envia o estado completo do painel para o endpoint oficial no Google Drive (SESAP_PAINEL_STATE_V1.json)
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

    // Evita colisões de concorrência: espera até 2s se outra sincronização estiver em curso
    let waitCycles = 0;
    while (this.isProcessingSync && waitCycles < 10) {
      await new Promise((r) => setTimeout(r, 200));
      waitCycles++;
    }

    return Promise.race([
      this.executePushToCloud(files, obsList, obFileName, consolidatedMetrics),
      new Promise<boolean>((resolve) => {
        setTimeout(() => {
          console.warn('Timeout de segurança ao sincronizar com o Google Drive (90s). Liberando interface.');
          this.isProcessingSync = false;
          this.updateState({ isSyncing: false, syncStatus: 'idle' });
          resolve(false);
        }, 90000);
      }),
    ]);
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
      const effectiveOBs = obsList || getStoredOrdensBancarias().items || [];
      const effectiveOBFileName =
        obFileName || getStoredOrdensBancarias().fileName || 'OBs - sigef - Principal';

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

      // 3. Resumo completo dos arquivos carregados:
      // Nunca trunca planilhas financeiras (base_pps, pp_prontas, obedece, nao_obedece, natureza_despesa, publicacao, outros),
      // apenas remove propriedades vazias para manter o JSON enxuto sem perder nenhum registro.
      const compactRow = (r: Record<string, any>) => {
        const out: Record<string, any> = {};
        for (const [k, v] of Object.entries(r)) {
          if (v !== '' && v !== null && v !== undefined && v !== '-') {
            out[k] = v;
          }
        }
        return out;
      };

      const filesSummary = files.map((f) => {
        const allRows = f.rows || [];
        const isRawObLarge = f.stage === 'ordem_bancaria' && allRows.length > 5000;
        const keptRows = isRawObLarge
          ? allRows.slice(0, 200).map(compactRow)
          : allRows.map(compactRow);
        return {
          id: f.id || '',
          fileName: f.fileName || '',
          stage: f.stage || 'outros',
          headers: f.headers || [],
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
          isSummaryOnly: isRawObLarge,
          rows: keptRows,
        };
      });

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
        alerts: toastAlertService.getAlertsHistory().slice(0, 100),
        paraPublicar: paraPublicarService.getItems(),
        mlPatterns,
        mlExactMappings,
        localPreferences: {
          recurring: localStorage.getItem('finance_manual_recurring_suppliers'),
          atendimentos: localStorage.getItem('finance_fornecedores_atendimento_mensal_v1'),
          categories:
            localStorage.getItem('sigecon_categories_config') ||
            localStorage.getItem('sigecon_categories_v1'),
          categoryOverrides: localStorage.getItem('sigecon_supplier_category_overrides'),
        },
      };

      const serializedPayload = JSON.stringify(drivePayload);

      // 6. Envia em uma única requisição HTTP POST para o Google Drive (sem CORS preflight)
      const response = await fetch(GOOGLE_DRIVE_SYNC_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: serializedPayload,
      });

      if (!response.ok) {
        throw new Error(`Falha HTTP ${response.status} ao salvar no Google Drive.`);
      }

      const resData = await response.json();
      if (resData && resData.success === false) {
        throw new Error(resData.error || 'Erro retornado pelo endpoint do Google Drive.');
      }

      // Espelho direto na pasta dedicada Painel_Financeiro_SESAP_Dados se o OAuth do Drive estiver conectado
      if (googleDriveService.getState().isConnected && googleDriveService.isConnected()) {
        googleDriveService
          .saveStateToDriveFolder('SESAP_PAINEL_STATE_V1.json', serializedPayload)
          .catch((e) => console.warn('Espelho na pasta do Google Drive:', e));
      }

      this.lastRemoteVersion = version;
      this.saveMetricsCache(metricsPayload, nowDate, effectiveOBFileName);

      this.updateState({
        isSyncing: false,
        syncStatus: 'synced',
        lastSyncAt: nowDate,
        lastSyncDevice: device,
        remoteFilesCount: files.length,
        errorMessage: null,
      });

      console.log(`Painel sincronizado no Google Drive com sucesso a partir do ${device}`);
      return true;
    } catch (err: any) {
      console.error('ERRO GOOGLE DRIVE PUSH:', err);
      this.updateState({
        isSyncing: false,
        syncStatus: 'error',
        errorMessage: err.message || 'Erro ao sincronizar com o Google Drive.',
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

  /**
   * Leitura direta do estado completo no Google Drive (SESAP_PAINEL_STATE_V1.json)
   */
  public async pullFromCloud(onlyIfNewer: boolean = false): Promise<boolean> {
    try {
      this.isProcessingSync = true;
      if (!onlyIfNewer) {
        this.updateState({ isSyncing: true, syncStatus: 'syncing', errorMessage: null });
      }

      const response = await fetch(`${GOOGLE_DRIVE_SYNC_URL}?t=${Date.now()}`, {
        method: 'GET',
        redirect: 'follow',
      });

      if (!response.ok) {
        throw new Error(`Falha HTTP ${response.status} ao consultar Google Drive.`);
      }

      const data = await response.json();
      if (!data || !data.metrics) {
        if (!onlyIfNewer) {
          this.updateState({ isSyncing: false, syncStatus: 'idle' });
        }
        return false;
      }

      const remoteVersion = typeof data.version === 'number' ? data.version : null;
      const syncDate = data.lastSyncAt ? new Date(data.lastSyncAt) : new Date();
      const validSyncDate = isNaN(syncDate.getTime()) ? new Date() : syncDate;

      // No polling automático a cada 60s, evita re-renderizações desnecessárias se a versão for idêntica
      if (
        onlyIfNewer &&
        remoteVersion &&
        this.lastRemoteVersion === remoteVersion &&
        this.lastPulledData
      ) {
        this.updateState({ syncStatus: 'synced', isSyncing: false });
        return true;
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

      // 2. Restaurar filesSummary
      let parsedFilesSummary: FileData[] = [];
      if (Array.isArray(data.filesSummary)) {
        parsedFilesSummary = data.filesSummary;
      } else if (typeof data.filesSummary === 'string') {
        try {
          parsedFilesSummary = JSON.parse(data.filesSummary);
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

      // 4. Restaurar paraPublicar
      if (Array.isArray(data.paraPublicar) && data.paraPublicar.length > 0) {
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
      if (data.localPreferences && typeof data.localPreferences === 'object') {
        try {
          const prefs = data.localPreferences;
          if (typeof prefs.recurring === 'string') {
            localStorage.setItem('finance_manual_recurring_suppliers', prefs.recurring);
          }
          if (typeof prefs.atendimentos === 'string') {
            localStorage.setItem('finance_fornecedores_atendimento_mensal_v1', prefs.atendimentos);
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

      // Verificar se o computador atual já possui arquivos locais completos (com mais linhas e versão igual/mais recente)
      const localState = await loadAppState();
      const localFiles = localState?.files || [];
      const localTotalRows = localFiles.reduce((acc, f) => acc + (f.rows?.length || 0), 0);
      const remoteTotalRows = parsedFilesSummary.reduce((acc, f) => acc + (f.rows?.length || 0), 0);
      const localSavedTime = localState?.savedAt ? new Date(localState.savedAt).getTime() : 0;

      const preserveLocalFiles =
        localFiles.length > 0 &&
        localTotalRows > remoteTotalRows &&
        localSavedTime >= validSyncDate.getTime();

      const effectiveFiles = preserveLocalFiles ? localFiles : parsedFilesSummary;

      // Salvar localmente no IndexedDB apenas se NÃO estiver sobrescrevendo arquivos locais mais completos
      if (!preserveLocalFiles && parsedFilesSummary.length > 0) {
        try {
          await saveAppState(parsedFilesSummary, validSyncDate);
        } catch (e) {
          console.warn('Erro ao salvar localmente dados baixados do Drive:', e);
        }
      }
      if (parsedCompactOBs.length > 0) {
        try {
          saveStoredOrdensBancarias(parsedCompactOBs, obFileName);
        } catch (e) {
          console.warn('Erro ao salvar localmente OBs do Drive:', e);
        }
      }

      // Atualiza silenciosamente a base comparativa (BASELINE_SNAPSHOT_KEY) sem disparar alertas falsos no login
      try {
        toastAlertService.forceUpdateBaseline(
          effectiveFiles,
          parsedCompactOBs.length > 0 ? parsedCompactOBs : null
        );
      } catch (e) {
        console.warn('Erro ao atualizar baseline após pullFromCloud:', e);
      }

      // Salvar cache de métricas no localStorage
      this.saveMetricsCache(parsedMetrics, validSyncDate, obFileName);

      this.lastRemoteVersion = remoteVersion || validSyncDate.getTime();

      // Salva no cache em memória
      this.lastPulledData = {
        files: effectiveFiles,
        obs: parsedCompactOBs.length > 0 ? parsedCompactOBs : null,
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
          parsedCompactOBs.length > 0 ? parsedCompactOBs : null,
          validSyncDate,
          parsedMetrics
        );
      }

      console.log('Dados restaurados diretamente do Google Drive com sucesso.');
      return true;
    } catch (err: any) {
      console.error('ERRO GOOGLE DRIVE PULL:', err);
      this.updateState({
        isSyncing: false,
        syncStatus: 'error',
        errorMessage: err.message || 'Erro ao baixar dados do Google Drive.',
      });
      return false;
    } finally {
      this.isProcessingSync = false;
    }
  }

  /**
   * Avalia estado local vs Google Drive e prioriza recuperação automática
   */
  public async checkAndAutoSync(): Promise<void> {
    if (this.isProcessingSync) return;

    try {
      const localState = await loadAppState();
      const localFiles = localState?.files || [];

      // Se o navegador local estiver vazio (novo dispositivo / Vercel / Modo Consulta), puxa imediatamente
      if (localFiles.length === 0) {
        console.log('Nenhum dado local encontrado. Baixando estado do Google Drive...');
        await this.pullFromCloud(false);
        return;
      }

      // Mesmo com dados locais, verifica se existe atualização mais recente no Google Drive
      const response = await fetch(`${GOOGLE_DRIVE_SYNC_URL}?t=${Date.now()}`, {
        method: 'GET',
        redirect: 'follow',
      });
      if (!response.ok) return;

      const data = await response.json();
      if (data && data.metrics) {
        const remoteTime = data.lastSyncAt ? new Date(data.lastSyncAt).getTime() : 0;
        const localSavedTime = localState?.savedAt ? new Date(localState.savedAt).getTime() : 0;

        if (remoteTime > localSavedTime + 2000 || !this.lastPulledData) {
          console.log('Google Drive possui atualização mais recente. Baixando...');
          await this.pullFromCloud(false);
        }
      } else if (localFiles.length > 0 && isAdmin()) {
        const localOBs = getStoredOrdensBancarias().items || [];
        console.log('Google Drive vazio; enviando estado local do Administrador...');
        await this.pushToCloud(localFiles, localOBs);
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
      const clearPayload = {
        lastSyncAt: nowIso,
        deviceType: this.getDeviceName(),
        totalFiles: 0,
        version: Date.now(),
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
      } catch (_) {}

      this.lastRemoteVersion = null;
      this.lastPulledData = null;
      this.updateState({
        remoteFilesCount: 0,
        lastSyncAt: new Date(),
        syncStatus: 'synced',
      });
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
      alert('Falha ao restaurar backup: ' + (err.message || 'Arquivo corrompido'));
      return false;
    }
  }
}

export const cloudSyncService = new CloudSyncService();

