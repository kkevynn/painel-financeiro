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
} from 'firebase/firestore';
import {
  signInWithPopup,
  signOut,
  onAuthStateChanged,
  User,
} from 'firebase/auth';
import { db, auth, googleAuthProvider } from '../firebase';
import { FileData } from '../types';
import { saveAppState, loadAppState } from './storageService';
import {
  saveStoredOrdensBancarias,
  getStoredOrdensBancarias,
  OrdemBancariaItem,
} from './ordemBancariaParser';
import { paraPublicarService, ParaPublicarItem } from './paraPublicarService';
import { mlInferenceService } from './machineLearningInferenceService';

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
  private onDatasetReloadCallback: ((files: FileData[], obs: OrdemBancariaItem[] | null) => void) | null = null;

  constructor() {
    this.initAuthListener();
    setTimeout(() => {
      if (auth.currentUser) {
        this.checkAndAutoSync().catch((e) => {
          console.warn('Auto-sync inicial:', e);
        });
      }
    }, 1200);
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
        this.startRealtimeListener();
        this.checkAndAutoSync();
      } else {
        if (this.unsubscribeSnapshot) {
          this.unsubscribeSnapshot();
          this.unsubscribeSnapshot = null;
        }
        this.updateState({
          isAuthenticated: false,
          user: null,
          syncStatus: 'idle',
          errorMessage: null,
        });
      }
    });
  }

  public registerDatasetReloadCallback(cb: (files: FileData[], obs: OrdemBancariaItem[] | null) => void) {
    this.onDatasetReloadCallback = cb;
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
   * Pushes current local dataset (files, OBs, publication queue, ML patterns) to Firestore
   */
  public async pushToCloud(
    files: FileData[],
    obsList?: OrdemBancariaItem[] | null,
    obFileName?: string
  ): Promise<boolean> {
    try {
      this.isProcessingSync = true;
      this.updateState({ isSyncing: true, syncStatus: 'syncing', errorMessage: null });

      const device = this.getDeviceName();
      const nowIso = new Date().toISOString();
      const userEmail = auth.currentUser?.email || '';
      const userId = auth.currentUser?.uid || 'workspace_user';

      // 1. Update metadata
      const metaRef = doc(db, 'workspaces', 'default');
      await setDoc(
        metaRef,
        {
          lastSyncAt: nowIso,
          lastUpdatedBy: userEmail || userId,
          userEmail: userEmail,
          deviceType: device,
          totalFiles: files.length,
          version: Date.now(),
        },
        { merge: true }
      );

      // 2. Fetch existing file docs from Firestore and delete files that were removed locally
      const currentFileIds = new Set(files.map((f) => f.id));
      try {
        const remoteFilesSnap = await getDocs(collection(db, 'workspaces', 'default', 'imported_files'));
        for (const docSnap of remoteFilesSnap.docs) {
          if (!currentFileIds.has(docSnap.id)) {
            // Delete subcollection chunks first
            try {
              const oldChunksSnap = await getDocs(collection(db, 'workspaces', 'default', 'imported_files', docSnap.id, 'chunks'));
              for (const cSnap of oldChunksSnap.docs) {
                await deleteDoc(cSnap.ref);
              }
            } catch (chunkCleanErr) {
              console.warn('Erro ao limpar chunks do arquivo removido:', chunkCleanErr);
            }
            await deleteDoc(doc(db, 'workspaces', 'default', 'imported_files', docSnap.id));
          }
        }
      } catch (cleanErr) {
        console.warn('Limpeza de arquivos remotos deletados:', cleanErr);
      }

      // 3. Save each active file into /workspaces/default/imported_files/{id} with subcollection chunking (max 1000 rows each)
      const CHUNK_SIZE = 1000;
      for (const file of files) {
        const fileRef = doc(db, 'workspaces', 'default', 'imported_files', file.id);
        const headersStr = JSON.stringify(file.headers || []);
        const fileRows = file.rows || [];
        const totalChunks = Math.ceil(fileRows.length / CHUNK_SIZE);

        // Save main file document with metadata only (bypassing Firestore 1MB document limit)
        await setDoc(fileRef, {
          id: file.id,
          fileName: file.fileName,
          stage: file.stage || 'outros',
          headersJson: headersStr,
          detectedFavorecidoCol: file.detectedFavorecidoCol || '',
          detectedValorCol: file.detectedValorCol || '',
          detectedNumeroCol: file.detectedNumeroCol || '',
          detectedPpCol: file.detectedPpCol || '',
          detectedCnpjCol: file.detectedCnpjCol || '',
          detectedSituacaoCol: file.detectedSituacaoCol || '',
          detectedFonteCol: file.detectedFonteCol || '',
          rowCount: fileRows.length,
          totalChunks,
          chunkSize: CHUNK_SIZE,
          updatedAt: nowIso,
        });

        // Clean up any surplus chunks if row count was reduced
        try {
          const existingChunksSnap = await getDocs(collection(db, 'workspaces', 'default', 'imported_files', file.id, 'chunks'));
          for (const cSnap of existingChunksSnap.docs) {
            const idx = parseInt(cSnap.id.replace('chunk_', ''), 10);
            if (!isNaN(idx) && idx >= totalChunks) {
              await deleteDoc(cSnap.ref);
            }
          }
        } catch (_) {}

        // Save row blocks in subcollection /workspaces/default/imported_files/{file.id}/chunks/chunk_{i}
        for (let i = 0; i < totalChunks; i++) {
          const chunk = fileRows.slice(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE);
          const chunkRef = doc(db, 'workspaces', 'default', 'imported_files', file.id, 'chunks', `chunk_${i}`);
          await setDoc(chunkRef, {
            chunkId: `chunk_${i}`,
            chunkIndex: i,
            rowCount: chunk.length,
            rowsJson: JSON.stringify(chunk),
            updatedAt: nowIso,
          });
        }
      }

      // 4. Save OBs dataset in chunks if present
      const effectiveOBs = obsList || getStoredOrdensBancarias().items || [];
      const effectiveOBFileName = obFileName || getStoredOrdensBancarias().fileName || '';

      if (effectiveOBs.length > 0) {
        const chunkSize = 400;
        const totalChunks = Math.ceil(effectiveOBs.length / chunkSize);

        for (let i = 0; i < totalChunks; i++) {
          const chunk = effectiveOBs.slice(i * chunkSize, (i + 1) * chunkSize);
          const chunkRef = doc(db, 'workspaces', 'default', 'ordens_bancarias_store', `chunk_${i}`);

          await setDoc(chunkRef, {
            chunkId: `chunk_${i}`,
            chunkIndex: i,
            fileName: effectiveOBFileName,
            totalItems: effectiveOBs.length,
            itemsJson: JSON.stringify(chunk),
            updatedAt: nowIso,
          });
        }
      }

      // 5. Save "Para Publicar" queue items
      const paraPublicarItems = paraPublicarService.getItems();
      for (const item of paraPublicarItems) {
        const itemDocId = item.id.replace(/[^a-zA-Z0-9_-]/g, '_');
        const itemRef = doc(db, 'workspaces', 'default', 'para_publicar', itemDocId);
        await setDoc(itemRef, {
          id: item.id,
          nlNumero: item.nlNumero || item.documentNumber || '',
          documentNumber: item.documentNumber || item.nlNumero || '',
          fornecedor: item.fornecedor || item.supplierName || '',
          supplierName: item.supplierName || item.fornecedor || '',
          cnpj: item.cnpj,
          valor: typeof item.valor === 'number' ? item.valor : (item.value || 0),
          value: typeof item.value === 'number' ? item.value : (item.valor || 0),
          fonteRecurso: item.fonteRecurso || '0.6.00',
          stageOrigem: item.stageOrigem || 'obedece',
          status: item.status || '',
          dataSelecao: item.dataSelecao || nowIso,
          updatedAt: nowIso,
        });
      }

      // 6. Save learned ML patterns and exact NL mappings across all devices
      try {
        const learnedPatterns = mlInferenceService.getAllLearnedPatterns();
        const exactMappings = mlInferenceService.getExactNlMappings();
        const mlDocRef = doc(db, 'workspaces', 'default', 'ml_patterns', 'patterns');
        await setDoc(mlDocRef, {
          patternsJson: JSON.stringify(learnedPatterns || {}),
          exactMappingsJson: JSON.stringify(exactMappings || {}),
          updatedAt: nowIso,
        });
      } catch (mlSaveErr) {
        console.warn('Erro ao salvar padrões de ML na nuvem:', mlSaveErr);
      }

      this.updateState({
        isSyncing: false,
        syncStatus: 'synced',
        lastSyncAt: new Date(),
        lastSyncDevice: device,
        remoteFilesCount: files.length,
      });

      console.log(`Dados sincronizados com a nuvem com sucesso a partir do ${device}`);
      return true;
    } catch (err: any) {
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
   * Directly removes a file document and its chunks from Firestore
   */
  public async deleteFileFromCloud(fileId: string): Promise<void> {
    try {
      try {
        const chunksSnap = await getDocs(collection(db, 'workspaces', 'default', 'imported_files', fileId, 'chunks'));
        for (const cSnap of chunksSnap.docs) {
          await deleteDoc(cSnap.ref);
        }
      } catch (err) {
        console.warn(`Erro ao limpar chunks do arquivo ${fileId}:`, err);
      }
      const fileRef = doc(db, 'workspaces', 'default', 'imported_files', fileId);
      await deleteDoc(fileRef);
      console.log(`Arquivo ${fileId} removido da nuvem com sucesso.`);
    } catch (err) {
      console.warn(`Erro ao remover arquivo ${fileId} da nuvem:`, err);
    }
  }

  /**
   * Pulls the latest cloud dataset into local state
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

      // 1. Fetch all imported files and reassemble their chunks from subcollections
      const filesColRef = collection(db, 'workspaces', 'default', 'imported_files');
      const filesSnap = await getDocs(filesColRef);

      const loadedFiles: FileData[] = [];
      for (const docSnap of filesSnap.docs) {
        const d = docSnap.data();

        let headers: string[] = [];
        try {
          headers = JSON.parse(d.headersJson || '[]');
        } catch {
          headers = [];
        }

        // Fetch all documents from the 'chunks' subcollection
        let rows: any[] = [];
        try {
          const chunksColRef = collection(db, 'workspaces', 'default', 'imported_files', docSnap.id, 'chunks');
          const chunksSnap = await getDocs(chunksColRef);

          if (!chunksSnap.empty) {
            const chunkList: { index: number; rows: any[] }[] = [];
            chunksSnap.forEach((cDoc) => {
              const cd = cDoc.data();
              let chunkRows: any[] = [];
              try {
                chunkRows = JSON.parse(cd.rowsJson || '[]');
              } catch {
                chunkRows = [];
              }
              chunkList.push({
                index: typeof cd.chunkIndex === 'number' ? cd.chunkIndex : parseInt(cDoc.id.replace('chunk_', ''), 10) || 0,
                rows: chunkRows,
              });
            });
            chunkList.sort((a, b) => a.index - b.index);
            rows = chunkList.flatMap((c) => c.rows);
          } else if (d.rowsJson) {
            // Backward-compatibility fallback for older unchunked records
            rows = JSON.parse(d.rowsJson || '[]');
          }
        } catch (readChunkErr) {
          console.warn(`Erro ao ler chunks do arquivo ${docSnap.id}:`, readChunkErr);
          if (d.rowsJson) {
            try {
              rows = JSON.parse(d.rowsJson || '[]');
            } catch {
              rows = [];
            }
          }
        }

        loadedFiles.push({
          id: d.id,
          fileName: d.fileName,
          headers,
          stage: d.stage || 'outros',
          detectedFavorecidoCol: d.detectedFavorecidoCol || '',
          detectedValorCol: d.detectedValorCol || '',
          detectedNumeroCol: d.detectedNumeroCol,
          detectedPpCol: d.detectedPpCol,
          detectedCnpjCol: d.detectedCnpjCol,
          detectedSituacaoCol: d.detectedSituacaoCol,
          detectedFonteCol: d.detectedFonteCol,
          rows,
        });
      }

      // 2. Fetch OBs chunks
      const obColRef = collection(db, 'workspaces', 'default', 'ordens_bancarias_store');
      const obSnap = await getDocs(obColRef);

      const obChunks: { index: number; fileName: string; items: OrdemBancariaItem[] }[] = [];
      obSnap.forEach((docSnap) => {
        const d = docSnap.data();
        let items: OrdemBancariaItem[] = [];
        try {
          items = JSON.parse(d.itemsJson || '[]');
        } catch {
          items = [];
        }
        obChunks.push({
          index: d.chunkIndex || 0,
          fileName: d.fileName || '',
          items,
        });
      });

      obChunks.sort((a, b) => a.index - b.index);
      const allOBItems: OrdemBancariaItem[] = obChunks.flatMap((c) => c.items);
      const obFileName = obChunks[0]?.fileName || 'OBs - sigef - Principal';

      // 3. Fetch "Para Publicar" items
      const pubColRef = collection(db, 'workspaces', 'default', 'para_publicar');
      const pubSnap = await getDocs(pubColRef);
      const pubItems: ParaPublicarItem[] = [];

      pubSnap.forEach((docSnap) => {
        const d = docSnap.data();
        pubItems.push({
          id: docSnap.id,
          nlNumero: d.nlNumero || d.documentNumber || '',
          documentNumber: d.documentNumber || d.nlNumero || '',
          fornecedor: d.fornecedor || d.supplierName || '',
          supplierName: d.supplierName || d.fornecedor || '',
          cnpj: d.cnpj,
          valor: typeof d.valor === 'number' ? d.valor : (d.value || 0),
          value: typeof d.value === 'number' ? d.value : (d.valor || 0),
          fonteRecurso: d.fonteRecurso || '0.6.00',
          stageOrigem: d.stageOrigem || 'obedece',
          status: d.status || '',
          dataSelecao: d.dataSelecao || new Date().toISOString(),
        });
      });

      if (pubItems.length > 0) {
        paraPublicarService.restoreItems(pubItems);
      }

      // 4. Fetch learned ML patterns and restore them
      try {
        const mlDocRef = doc(db, 'workspaces', 'default', 'ml_patterns', 'patterns');
        const mlSnap = await getDoc(mlDocRef);
        if (mlSnap.exists()) {
          const mlData = mlSnap.data();
          const patterns = JSON.parse(mlData.patternsJson || '{}');
          const mappings = JSON.parse(mlData.exactMappingsJson || '{}');
          mlInferenceService.restoreLearnedPatterns(patterns, mappings);
        }
      } catch (mlErr) {
        console.warn('Erro ao restaurar padrões de aprendizado da nuvem:', mlErr);
      }

      const syncDate = metaData.lastSyncAt ? new Date(metaData.lastSyncAt) : new Date();

      // Persist downloaded cloud state immediately into local storage
      if (loadedFiles.length > 0) {
        try {
          await saveAppState(loadedFiles, syncDate);
        } catch (e) {
          console.warn('Erro ao salvar localmente dados baixados da nuvem:', e);
        }
      }
      if (allOBItems.length > 0) {
        try {
          saveStoredOrdensBancarias(allOBItems, obFileName);
        } catch (e) {
          console.warn('Erro ao salvar localmente OBs da nuvem:', e);
        }
      }

      this.updateState({
        isSyncing: false,
        syncStatus: 'synced',
        lastSyncAt: syncDate,
        lastSyncDevice: metaData.deviceType || 'Nuvem',
        remoteFilesCount: loadedFiles.length,
      });

      // Notify App to reload UI with updated files
      if (this.onDatasetReloadCallback) {
        this.onDatasetReloadCallback(loadedFiles, allOBItems.length > 0 ? allOBItems : null);
      }

      console.log(`Dados da nuvem restaurados localmente (${loadedFiles.length} planilhas sincronizadas)`);
      return true;
    } catch (err: any) {
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
   * Evaluates cloud vs local state and triggers automatic sync
   */
  public async checkAndAutoSync(): Promise<void> {
    if (this.isProcessingSync) return;

    try {
      const metaRef = doc(db, 'workspaces', 'default');
      const metaSnap = await getDoc(metaRef);

      const localState = await loadAppState();
      const localFiles = localState?.files || [];
      const localOBs = getStoredOrdensBancarias().items || [];

      if (metaSnap.exists()) {
        const remoteData = metaSnap.data();
        const remoteTime = remoteData.lastSyncAt ? new Date(remoteData.lastSyncAt).getTime() : 0;
        const localSavedTime = localState?.savedAt ? new Date(localState.savedAt).getTime() : 0;
        const remoteFilesCount = remoteData.totalFiles || 0;

        // If cloud is newer OR local is empty while cloud has files, pull!
        if (remoteTime > localSavedTime + 2000 || (localFiles.length === 0 && remoteFilesCount > 0)) {
          console.log('Nuvem possui versão mais recente ou dados remotos. Sincronizando para este dispositivo...');
          await this.pullFromCloud();
        } else if (localFiles.length > 0 && localSavedTime > remoteTime + 3000) {
          // If local has newer files/adjustments, push!
          console.log('Este dispositivo possui atualizações mais recentes. Enviando para a nuvem...');
          await this.pushToCloud(localFiles, localOBs);
        }
      } else {
        // Cloud is empty: if local has data, seed cloud!
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
