/**
 * Cloud Sync Service
 * Synchronizes spreadsheets, bank orders, and audit state across all user devices (Tablet, PC, Mobile)
 * using Firebase Firestore with automatic real-time updates and offline resilience.
 */
import {
  doc,
  getDoc,
  setDoc,
  collection,
  getDocs,
  onSnapshot,
  deleteDoc,
  writeBatch,
} from 'firebase/firestore';
import {
  signInWithPopup,
  signOut,
  onAuthStateChanged,
  User,
} from 'firebase/auth';
import { db, auth, googleAuthProvider } from '../firebase';
import { FileData, DashboardMetrics, ConsolidatedSupplier } from '../types';
import { saveAppState, loadAppState } from './storageService';
import {
  saveStoredOrdensBancarias,
  getStoredOrdensBancarias,
  OrdemBancariaItem,
  extractOrdensBancariasFromFileData,
  enrichSuppliersWithOrdensBancarias,
} from './ordemBancariaParser';
import { processUploadedFiles } from './excelParser';
import { paraPublicarService, ParaPublicarItem } from './paraPublicarService';
import { mlInferenceService } from './machineLearningInferenceService';
import { toastAlertService } from './toastAlertService';

enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}

interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: {
    userId?: string | null;
    email?: string | null;
    emailVerified?: boolean | null;
    isAnonymous?: boolean | null;
    providerInfo?: {
      providerId?: string | null;
      email?: string | null;
    }[];
  };
}

function handleFirestoreError(error: unknown, operationType: OperationType, path: string | null) {
  const errInfo: FirestoreErrorInfo = {
    error: error instanceof Error ? error.message : String(error),
    authInfo: {
      userId: auth.currentUser?.uid,
      email: auth.currentUser?.email,
      emailVerified: auth.currentUser?.emailVerified,
      isAnonymous: auth.currentUser?.isAnonymous,
      providerInfo:
        auth.currentUser?.providerData?.map((p) => ({
          providerId: p.providerId,
          email: p.email,
        })) || [],
    },
    operationType,
    path,
  };
  console.warn('Firestore Cloud Sync Info: ', JSON.stringify(errInfo));
  return errInfo;
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
    syncStatus: 'unauthenticated',
    errorMessage: null,
  };

  private listeners = new Set<SyncListener>();
  private unsubscribeSnapshot: (() => void) | null = null;
  private isProcessingSync = false;
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
    this.startRealtimeListener();
    setTimeout(() => {
      this.checkAndAutoSync().catch((e) => {
        console.warn('Auto-sync inicial da nuvem:', e);
      });
    }, 500);
  }

  private initAuthListener() {
    onAuthStateChanged(auth, async (user) => {
      if (user) {
        this.updateState({
          isAuthenticated: true,
          user,
          syncStatus: 'synced',
          errorMessage: null,
        });
        this.checkAndAutoSync();
      } else {
        this.updateState({
          isAuthenticated: false,
          user: null,
          syncStatus: 'idle',
          errorMessage: null,
        });
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
      const isPopupBlocked = err?.code === 'auth/popup-blocked' || String(err?.message || '').includes('popup-blocked');
      this.updateState({
        isSyncing: false,
        syncStatus: 'idle',
        errorMessage: isPopupBlocked
          ? 'Pop-up bloqueado pelo navegador. O app continuará sincronizando na nuvem pelo Workspace normalmente.'
          : (err.message || 'Falha na autenticação Google.'),
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
    if (ua.includes('ipad') || (ua.includes('tablet') && !ua.includes('mobile')) || (ua.includes('android') && !ua.includes('mobile'))) {
      return 'Tablet';
    }
    if (ua.includes('mobile') || ua.includes('iphone')) {
      return 'Smartphone';
    }
    return 'PC / Computador';
  }

  private startRealtimeListener() {
    if (this.unsubscribeSnapshot) {
      this.unsubscribeSnapshot();
    }

    const metaDocRef = doc(db, 'workspaces', 'default');
    this.unsubscribeSnapshot = onSnapshot(
      metaDocRef,
      async (snapshot) => {
        if (!snapshot.exists()) return;
        const data = snapshot.data();
        if (!data) return;

        const remoteTime = data.lastSyncAt ? new Date(data.lastSyncAt).getTime() : 0;
        const localTime = this.state.lastSyncAt ? this.state.lastSyncAt.getTime() : 0;

        // If remote has a newer timestamp from another device, pull automatically!
        if (remoteTime > localTime + 2000 && !this.isProcessingSync) {
          console.log('Detectada nova sincronização na nuvem proveniente de:', data.deviceType);
          await this.pullFromCloud();
        }
      },
      (error) => {
        handleFirestoreError(error, OperationType.GET, 'workspaces/default');
      }
    );
  }

  /**
   * Sincronização Ultra-Leve em 1 Documento Único (Single-Doc Dashboard Snapshot)
   * Salva todo o painel consolidado em workspaces/default com apenas 1 escrita (< 300KB)
   */
  public async pushToCloud(
    files: FileData[],
    obsList?: OrdemBancariaItem[] | null,
    obFileName?: string,
    consolidatedMetrics?: DashboardMetrics | null
  ): Promise<boolean> {
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
          console.warn('Timeout de segurança ao sincronizar com a nuvem (90s). Liberando interface.');
          this.isProcessingSync = false;
          this.updateState({ isSyncing: false, syncStatus: 'idle' });
          resolve(false);
        }, 90000);
      }),
    ]);
  }

  /**
   * Grava o painel consolidado em 1 único documento no Firestore sem estourar cotas
   */
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
      const nowIso = new Date().toISOString();
      const effectiveOBs = obsList || getStoredOrdensBancarias().items || [];
      const effectiveOBFileName = obFileName || getStoredOrdensBancarias().fileName || 'OBs - sigef - Principal';

      // 1. Obter ou calcular as métricas consolidadas
      let activeMetrics: DashboardMetrics;
      if (consolidatedMetrics && consolidatedMetrics.isProcessed) {
        activeMetrics = consolidatedMetrics;
      } else {
        const base = processUploadedFiles(files);
        enrichSuppliersWithOrdensBancarias(base.consolidatedSuppliers, effectiveOBs.length > 0 ? effectiveOBs : null);
        activeMetrics = base;
      }

      // 2. Versão compacta de consolidatedSuppliers
      const compactSuppliers: ConsolidatedSupplier[] = (activeMetrics.consolidatedSuppliers || []).map((s) => ({
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
      }));

      // 3. Resumo leve dos arquivos importados (máximo primeiras 30 linhas para pré-visualização)
      const filesSummary = files.map((f) => ({
        id: f.id || '',
        fileName: f.fileName || '',
        stage: f.stage || 'outros',
        headers: f.headers || [],
        detectedFavorecidoCol: f.detectedFavorecidoCol || '',
        detectedValorCol: f.detectedValorCol || '',
        detectedNumeroCol: f.detectedNumeroCol || '',
        detectedPpCol: f.detectedPpCol || '',
        detectedCnpjCol: f.detectedCnpjCol || '',
        detectedSituacaoCol: f.detectedSituacaoCol || '',
        detectedFonteCol: f.detectedFonteCol || '',
        rowCount: (f.rows || []).length,
        totalCount: f.totalCount || (f.rows || []).length,
        rows: (f.rows || []).slice(0, 30),
      }));

      // 4. Lista compacta das Ordens Bancárias agrupadas por fornecedor+fonte+mês
      const compactOBsMap = new Map<string, OrdemBancariaItem>();
      effectiveOBs.forEach((item) => {
        const cnpj = item.favorecidoCnpj || item.favorecidoRaw || '';
        const fonte = item.fonteRecurso || 'Geral';
        const mesAno = item.mesAnoKey || (item.dataIso ? item.dataIso.slice(0, 7) : (item.dataReferencia ? item.dataReferencia.slice(-7) : 'Sem Data'));
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
      let mlPatternsJson = '{}';
      let mlExactMappingsJson = '{}';
      try {
        const learnedPatterns = mlInferenceService.getAllLearnedPatterns();
        const exactMappings = mlInferenceService.getExactNlMappings();
        mlPatternsJson = JSON.stringify(learnedPatterns || {});
        mlExactMappingsJson = JSON.stringify(exactMappings || {});
      } catch (_) {}

      // 6. Grava TUDO em apenas 1 ÚNICO setDoc no documento doc(db, 'workspaces', 'default')
      const metaRef = doc(db, 'workspaces', 'default');
      await setDoc(metaRef, {
        lastSyncAt: nowIso,
        deviceType: device,
        totalFiles: files.length,
        version: Date.now(),
        metricsJson: JSON.stringify({
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
        }),
        filesSummaryJson: JSON.stringify(filesSummary),
        compactOBsJson: JSON.stringify(compactOBs),
        obFileName: effectiveOBFileName,
        alertsJson: JSON.stringify(toastAlertService.getAlertsHistory().slice(0, 100)),
        paraPublicarJson: JSON.stringify(paraPublicarService.getItems()),
        mlPatternsJson,
        mlExactMappingsJson,
      });

      this.updateState({
        isSyncing: false,
        syncStatus: 'synced',
        lastSyncAt: new Date(),
        lastSyncDevice: device,
        remoteFilesCount: files.length,
      });

      console.log(`Painel consolidado sincronizado em 1 documento com sucesso a partir do ${device}`);
      return true;
    } catch (err: any) {
      console.error('ERRO FIRESTORE:', err);
      handleFirestoreError(err, OperationType.WRITE, 'workspaces/default');
      this.updateState({
        isSyncing: false,
        syncStatus: 'error',
        errorMessage: err.message || 'Erro ao sincronizar com a nuvem.',
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
   * Leitura instantânea de 1 documento único do Firestore (Single-Doc Snapshot)
   */
  public async pullFromCloud(): Promise<boolean> {
    try {
      this.isProcessingSync = true;
      this.updateState({ isSyncing: true, syncStatus: 'syncing', errorMessage: null });

      const metaRef = doc(db, 'workspaces', 'default');
      const metaSnap = await getDoc(metaRef);
      if (!metaSnap.exists()) {
        this.updateState({ isSyncing: false, syncStatus: 'idle' });
        return false;
      }

      const metaData = metaSnap.data();
      if (!metaData) {
        this.updateState({ isSyncing: false, syncStatus: 'idle' });
        return false;
      }

      // 1. Restaurar metricsJson
      let parsedMetrics: DashboardMetrics | null = null;
      if (metaData.metricsJson) {
        try {
          parsedMetrics = JSON.parse(metaData.metricsJson);
        } catch (e) {
          console.warn('Erro ao processar metricsJson:', e);
        }
      }

      // 2. Restaurar filesSummaryJson
      let parsedFilesSummary: FileData[] = [];
      if (metaData.filesSummaryJson) {
        try {
          parsedFilesSummary = JSON.parse(metaData.filesSummaryJson);
        } catch (e) {
          console.warn('Erro ao processar filesSummaryJson:', e);
        }
      }

      // 3. Restaurar compactOBsJson
      let parsedCompactOBs: OrdemBancariaItem[] = [];
      if (metaData.compactOBsJson) {
        try {
          parsedCompactOBs = JSON.parse(metaData.compactOBsJson);
        } catch (e) {
          console.warn('Erro ao processar compactOBsJson:', e);
        }
      }
      const obFileName = metaData.obFileName || 'OBs - sigef - Principal';

      // 4. Restaurar paraPublicarJson
      if (metaData.paraPublicarJson) {
        try {
          const pubItems = JSON.parse(metaData.paraPublicarJson);
          if (Array.isArray(pubItems) && pubItems.length > 0) {
            paraPublicarService.restoreItems(pubItems);
          }
        } catch (e) {
          console.warn('Erro ao restaurar paraPublicarJson:', e);
        }
      }

      // 5. Restaurar alertsJson
      if (metaData.alertsJson) {
        try {
          const alertItems = JSON.parse(metaData.alertsJson);
          if (Array.isArray(alertItems) && alertItems.length > 0) {
            toastAlertService.restoreAlertsHistory(alertItems);
          }
        } catch (e) {
          console.warn('Erro ao restaurar alertsJson:', e);
        }
      }

      // 6. Restaurar padrões de aprendizado de máquina
      if (metaData.mlPatternsJson && metaData.mlExactMappingsJson) {
        try {
          const patterns = JSON.parse(metaData.mlPatternsJson);
          const mappings = JSON.parse(metaData.mlExactMappingsJson);
          mlInferenceService.restoreLearnedPatterns(patterns, mappings);
        } catch (_) {}
      }

      const syncDate = metaData.lastSyncAt ? new Date(metaData.lastSyncAt) : new Date();

      // Salvar localmente no IndexedDB
      if (parsedFilesSummary.length > 0) {
        try {
          await saveAppState(parsedFilesSummary, syncDate);
        } catch (e) {
          console.warn('Erro ao salvar localmente dados baixados da nuvem:', e);
        }
      }
      if (parsedCompactOBs.length > 0) {
        try {
          saveStoredOrdensBancarias(parsedCompactOBs, obFileName);
        } catch (e) {
          console.warn('Erro ao salvar localmente OBs da nuvem:', e);
        }
      }

      // Salva no cache em memória
      this.lastPulledData = {
        files: parsedFilesSummary,
        obs: parsedCompactOBs.length > 0 ? parsedCompactOBs : null,
        metrics: parsedMetrics,
        syncDate,
      };

      this.updateState({
        isSyncing: false,
        syncStatus: 'synced',
        lastSyncAt: syncDate,
        lastSyncDevice: metaData.deviceType || 'Nuvem',
        remoteFilesCount: parsedFilesSummary.length,
      });

      // Notifica App.tsx com os dados e as métricas prontas
      if (this.onDatasetReloadCallback) {
        this.onDatasetReloadCallback(
          parsedFilesSummary,
          parsedCompactOBs.length > 0 ? parsedCompactOBs : null,
          syncDate,
          parsedMetrics
        );
      }

      console.log(`Dados da nuvem restaurados instantaneamente (1 documento lido)`);
      return true;
    } catch (err: any) {
      console.error('ERRO FIRESTORE PULL:', err);
      handleFirestoreError(err, OperationType.GET, 'workspaces/default');
      this.updateState({
        isSyncing: false,
        syncStatus: 'error',
        errorMessage: err.message || 'Erro ao baixar dados da nuvem.',
      });
      return false;
    } finally {
      this.isProcessingSync = false;
    }
  }

  /**
   * Avalia estado local vs nuvem e prioriza recuperação automática
   */
  public async checkAndAutoSync(): Promise<void> {
    if (this.isProcessingSync) return;

    try {
      const localState = await loadAppState();
      const localFiles = localState?.files || [];
      const localOBs = getStoredOrdensBancarias().items || [];

      // Sempre que o navegador local não tiver arquivos (como na Vercel), puxa a nuvem imediatamente!
      if (localFiles.length === 0) {
        console.log('Nenhum dado local encontrado. Baixando snapshot da nuvem imediatamente...');
        await this.pullFromCloud();
        return;
      }

      const metaRef = doc(db, 'workspaces', 'default');
      const metaSnap = await getDoc(metaRef);

      if (metaSnap.exists()) {
        const remoteData = metaSnap.data();
        const remoteTime = remoteData.lastSyncAt ? new Date(remoteData.lastSyncAt).getTime() : 0;
        const localSavedTime = localState?.savedAt ? new Date(localState.savedAt).getTime() : 0;

        if (remoteTime !== localSavedTime || remoteTime > localSavedTime) {
          console.log('Nuvem possui versão mais recente. Sincronizando para este dispositivo...');
          await this.pullFromCloud();
        } else if (localFiles.length > 0 && localSavedTime > remoteTime + 3000) {
          console.log('Este dispositivo possui atualizações locais mais recentes. Enviando para a nuvem...');
          await this.pushToCloud(localFiles, localOBs);
        }
      } else {
        if (localFiles.length > 0 || localOBs.length > 0) {
          console.log('Iniciando primeiro upload de sincronização na nuvem para este workspace...');
          await this.pushToCloud(localFiles, localOBs);
        }
      }
    } catch (err) {
      console.warn('Verificação de sincronização automática:', err);
    }
  }

  /**
   * Limpa o snapshot do workspace na nuvem
   */
  public async clearCloudWorkspace(): Promise<void> {
    try {
      this.isProcessingSync = true;
      const metaRef = doc(db, 'workspaces', 'default');
      const nowIso = new Date().toISOString();
      await setDoc(
        metaRef,
        {
          lastSyncAt: nowIso,
          totalFiles: 0,
          version: Date.now(),
          metricsJson: null,
          filesSummaryJson: null,
          compactOBsJson: null,
          alertsJson: null,
          paraPublicarJson: null,
        },
        { merge: true }
      );

      this.lastPulledData = null;
      this.updateState({
        remoteFilesCount: 0,
        lastSyncAt: new Date(),
        syncStatus: 'synced',
      });
      console.log('Workspace na nuvem limpo com sucesso.');
    } catch (err) {
      console.warn('Erro ao limpar workspace na nuvem:', err);
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

      // Push to cloud if connected
      if (auth.currentUser) {
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
