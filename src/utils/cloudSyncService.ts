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
import { FileData } from '../types';
import { saveAppState, loadAppState } from './storageService';
import {
  saveStoredOrdensBancarias,
  getStoredOrdensBancarias,
  OrdemBancariaItem,
  extractOrdensBancariasFromFileData,
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
  private onDatasetReloadCallback: ((files: FileData[], obs: OrdemBancariaItem[] | null, syncDate?: Date) => void) | null = null;

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

  public registerDatasetReloadCallback(cb: (files: FileData[], obs: OrdemBancariaItem[] | null, syncDate?: Date) => void) {
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
   * Pushes the entire local dataset (spreadsheets in chunks,
   * OBs, publication queue, ML patterns) to Firestore with collision avoidance and safety timeout
   */
  public async pushToCloud(
    files: FileData[],
    obsList?: OrdemBancariaItem[] | null,
    obFileName?: string
  ): Promise<boolean> {
    // Avoid concurrency collisions: wait up to 2s if a sync is currently in progress
    let waitCycles = 0;
    while (this.isProcessingSync && waitCycles < 10) {
      await new Promise((r) => setTimeout(r, 200));
      waitCycles++;
    }

    return Promise.race([
      this.executePushToCloud(files, obsList, obFileName),
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
   * Internal execution of cloud synchronization using parallel batch writes and deduplication
   */
  private async executePushToCloud(
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

      // 1. Grava workspaces/default IMEDIATAMENTE no Passo 1 para garantir que o documento
      // sempre exista no Firestore mesmo enquanto os chunks estão sendo enviados
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
              await Promise.all(oldChunksSnap.docs.map((c) => deleteDoc(c.ref)));
            } catch (chunkCleanErr) {
              console.warn('Erro ao limpar chunks do arquivo removido:', chunkCleanErr);
            }
            await deleteDoc(doc(db, 'workspaces', 'default', 'imported_files', docSnap.id));
          }
        }
      } catch (cleanErr) {
        console.warn('Limpeza de arquivos remotos deletados:', cleanErr);
      }

      // 3. Prepare atomic write operations (up to 10 chunks/docs per parallel batch)
      interface PendingWrite {
        ref: any;
        data: any;
      }
      const pendingWrites: PendingWrite[] = [];

      // A) Save each active file and its chunks (subcollection chunks, max 500 rows each to guarantee < 300KB)
      // Removed slow getDocs on subcollections before saving to prevent network stalls
      const CHUNK_SIZE = 500;
      for (const file of files) {
        const fileRef = doc(db, 'workspaces', 'default', 'imported_files', file.id);
        const headersStr = JSON.stringify(file.headers || []);
        const fileRows = file.rows || [];
        const totalChunks = Math.ceil(fileRows.length / CHUNK_SIZE);

        pendingWrites.push({
          ref: fileRef,
          data: {
            id: file.id || '',
            fileName: file.fileName || '',
            stage: file.stage || 'outros',
            headersJson: headersStr || '[]',
            detectedFavorecidoCol: file.detectedFavorecidoCol || '',
            detectedValorCol: file.detectedValorCol || '',
            detectedNumeroCol: file.detectedNumeroCol || '',
            detectedPpCol: file.detectedPpCol || '',
            detectedCnpjCol: file.detectedCnpjCol || '',
            detectedSituacaoCol: file.detectedSituacaoCol || '',
            detectedFonteCol: file.detectedFonteCol || '',
            rowCount: fileRows.length || 0,
            totalChunks: totalChunks || 0,
            chunkSize: CHUNK_SIZE,
            updatedAt: nowIso,
          },
        });

        for (let i = 0; i < totalChunks; i++) {
          const chunk = fileRows.slice(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE);
          const chunkRef = doc(db, 'workspaces', 'default', 'imported_files', file.id, 'chunks', `chunk_${i}`);
          pendingWrites.push({
            ref: chunkRef,
            data: {
              chunkId: `chunk_${i}`,
              chunkIndex: i,
              rowCount: chunk.length || 0,
              rowsJson: JSON.stringify(chunk) || '[]',
              updatedAt: nowIso,
            },
          });
        }
      }

      // B) Eliminate duplication: if files.length > 0, OBs are already inside imported_files.
      // Only write ordens_bancarias_store when files.length === 0 and standalone OBs exist.
      const effectiveOBs = obsList || getStoredOrdensBancarias().items || [];
      const effectiveOBFileName = obFileName || getStoredOrdensBancarias().fileName || '';

      if (files.length === 0 && effectiveOBs.length > 0) {
        const obChunkSize = 500;
        const totalChunks = Math.ceil(effectiveOBs.length / obChunkSize);

        for (let i = 0; i < totalChunks; i++) {
          const chunk = effectiveOBs.slice(i * obChunkSize, (i + 1) * obChunkSize);
          const chunkRef = doc(db, 'workspaces', 'default', 'ordens_bancarias_store', `chunk_${i}`);
          pendingWrites.push({
            ref: chunkRef,
            data: {
              chunkId: `chunk_${i}`,
              chunkIndex: i,
              fileName: effectiveOBFileName || '',
              totalItems: effectiveOBs.length || 0,
              itemsJson: JSON.stringify(chunk) || '[]',
              updatedAt: nowIso,
            },
          });
        }
      }

      // C) Save "Para Publicar" queue items (zero undefined)
      const paraPublicarItems = paraPublicarService.getItems();
      for (const item of paraPublicarItems) {
        const itemDocId = (item.id || '').replace(/[^a-zA-Z0-9_-]/g, '_');
        const itemRef = doc(db, 'workspaces', 'default', 'para_publicar', itemDocId);
        pendingWrites.push({
          ref: itemRef,
          data: {
            id: item.id || '',
            nlNumero: item.nlNumero || item.documentNumber || '',
            documentNumber: item.documentNumber || item.nlNumero || '',
            fornecedor: item.fornecedor || item.supplierName || '',
            supplierName: item.supplierName || item.fornecedor || '',
            cnpj: item.cnpj || '',
            valor: typeof item.valor === 'number' ? item.valor : (item.value || 0),
            value: typeof item.value === 'number' ? item.value : (item.valor || 0),
            fonteRecurso: item.fonteRecurso || '0.6.00',
            stageOrigem: item.stageOrigem || 'obedece',
            status: item.status || '',
            dataSelecao: item.dataSelecao || nowIso,
            updatedAt: nowIso,
          },
        });
      }

      // D) Save learned ML patterns and exact NL mappings across all devices
      try {
        const learnedPatterns = mlInferenceService.getAllLearnedPatterns();
        const exactMappings = mlInferenceService.getExactNlMappings();
        const mlDocRef = doc(db, 'workspaces', 'default', 'ml_patterns', 'patterns');
        pendingWrites.push({
          ref: mlDocRef,
          data: {
            patternsJson: JSON.stringify(learnedPatterns || {}),
            exactMappingsJson: JSON.stringify(exactMappings || {}),
            updatedAt: nowIso,
          },
        });
      } catch (mlSaveErr) {
        console.warn('Erro ao salvar padrões de ML na nuvem:', mlSaveErr);
      }

      // E) Gravar os documentos em lotes paralelos de 10 com await Promise.all(...)
      const PARALLEL_BATCH_SIZE = 10;
      for (let i = 0; i < pendingWrites.length; i += PARALLEL_BATCH_SIZE) {
        const batchSlice = pendingWrites.slice(i, i + PARALLEL_BATCH_SIZE);
        await Promise.all(batchSlice.map((op) => setDoc(op.ref, op.data)));
      }

      // Repete a gravação de workspaces/default no final com o timestamp e versão finais
      const finishIso = new Date().toISOString();
      await setDoc(
        metaRef,
        {
          lastSyncAt: finishIso,
          lastUpdatedBy: userEmail || userId,
          userEmail: userEmail,
          deviceType: device,
          totalFiles: files.length,
          version: Date.now(),
        },
        { merge: true }
      );

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
      let metaSnap: any = null;
      try {
        metaSnap = await getDoc(metaRef);
      } catch (metaErr) {
        console.warn('Aviso ao consultar metadados do workspace:', metaErr);
      }

      const metaExists = metaSnap && typeof metaSnap.exists === 'function' ? metaSnap.exists() : false;
      const metaData = metaExists ? metaSnap.data() : {};

      // 1. Fetch all imported files and reassemble their chunks from subcollections
      const filesColRef = collection(db, 'workspaces', 'default', 'imported_files');
      const filesSnap = await getDocs(filesColRef);

      // Fallback: se não há metadados E a coleção de arquivos estiver vazia, encerra idle
      if (!metaExists && filesSnap.empty) {
        this.updateState({ isSyncing: false, syncStatus: 'idle' });
        return false;
      }

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
            let chunkList: { index: number; rows: any[] }[] = [];
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
            chunkList = chunkList.filter((c) => !d.totalChunks || c.index < d.totalChunks);
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
      let obFileName = obChunks[0]?.fileName || 'OBs - sigef - Principal';

      // If allOBItems is empty from ordens_bancarias_store, auto-reconstruct from loadedFiles
      if (allOBItems.length === 0 && loadedFiles.length > 0) {
        const obMap = new Map<string, OrdemBancariaItem>();
        loadedFiles.forEach((f) => {
          const lower = (f.fileName || '').toLowerCase();
          const isObFile = lower.includes('ordem') || lower.includes('ob') || lower.includes('banco') || f.stage === 'ordem_bancaria' || f.detectedObCol;
          if (isObFile) {
            const extracted = extractOrdensBancariasFromFileData(f);
            extracted.forEach((item) => {
              const k = item.ob ? ('OB_' + item.ob.trim().toUpperCase()) : (item.pp ? ('PP_' + item.pp.trim().toUpperCase()) : item.id);
              obMap.set(k, item);
            });
          }
        });
        if (obMap.size > 0) {
          allOBItems.push(...Array.from(obMap.values()));
          obFileName = 'Consolidado de Arquivos';
        }
      }

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
        this.onDatasetReloadCallback(loadedFiles, allOBItems.length > 0 ? allOBItems : null, syncDate);
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
   * Evaluates cloud vs local state and triggers automatic sync with cloud priority
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

        // Se metaSnap.exists() e remoteData.totalFiles > 0, compare remoteTime com o tempo local.
        // Se remoteTime !== localSavedTime (ou se a nuvem tiver dados/versão diferente), execute pullFromCloud() imediatamente!
        if (remoteFilesCount > 0 && (remoteTime !== localSavedTime || remoteTime > localSavedTime || localFiles.length === 0)) {
          console.log('Nuvem possui versão mais recente ou diferente. Sincronizando para este dispositivo...');
          await this.pullFromCloud();
        } else if (localFiles.length > 0 && localSavedTime > remoteTime + 3000) {
          // If local has newer files/adjustments, push!
          console.log('Este dispositivo possui atualizações locais mais recentes. Enviando para a nuvem...');
          await this.pushToCloud(localFiles, localOBs);
        }
      } else {
        // Fallback: se metaSnap ainda não existir, verifica se imported_files já possui dados
        try {
          const filesSnap = await getDocs(collection(db, 'workspaces', 'default', 'imported_files'));
          if (!filesSnap.empty) {
            console.log('Arquivos encontrados na nuvem sem metadados. Sincronizando para este dispositivo...');
            await this.pullFromCloud();
            return;
          }
        } catch (_) {}

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
   * Cleans all cloud documents from workspaces/default when user clicks Clear in the panel
   */
  public async clearCloudWorkspace(): Promise<void> {
    try {
      this.isProcessingSync = true;
      // 1. Delete imported files and chunks
      const filesSnap = await getDocs(collection(db, 'workspaces', 'default', 'imported_files'));
      for (const docSnap of filesSnap.docs) {
        try {
          const chunksSnap = await getDocs(collection(db, 'workspaces', 'default', 'imported_files', docSnap.id, 'chunks'));
          await Promise.all(chunksSnap.docs.map((c) => deleteDoc(c.ref)));
        } catch (_) {}
        await deleteDoc(docSnap.ref);
      }

      // 2. Delete OB chunks
      const obSnap = await getDocs(collection(db, 'workspaces', 'default', 'ordens_bancarias_store'));
      await Promise.all(obSnap.docs.map((c) => deleteDoc(c.ref)));

      // 3. Reset workspace metadata to totalFiles: 0
      const metaRef = doc(db, 'workspaces', 'default');
      const nowIso = new Date().toISOString();
      await setDoc(
        metaRef,
        {
          lastSyncAt: nowIso,
          totalFiles: 0,
          version: Date.now(),
        },
        { merge: true }
      );

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
