/**
 * Google Drive & Sheets Integration Service
 * 
 * Provides OAuth2 authentication via Google Identity Services (GIS),
 * searches the user's Google Drive for the most recently modified SIGEF spreadsheet report,
 * downloads it in-memory as a File object, and delivers it directly to the existing parser.
 * Supports silent 20-minute auto-polling and on-demand manual synchronization.
 */

// Production Client ID configured in Google Cloud Console
export const GOOGLE_CLIENT_ID = '866738622063-trtjudvl95g2c3lu0fhga9q79nn4eol6.apps.googleusercontent.com';

export const DRIVE_DATABASE_FOLDER_NAME = 'Painel_Financeiro_SESAP_Dados';

// Scopes supporting spreadsheet reading and dedicated app database folder storage
export const GOOGLE_DRIVE_SCOPES = [
  'https://www.googleapis.com/auth/drive.readonly',
  'https://www.googleapis.com/auth/drive.file',
  'https://www.googleapis.com/auth/spreadsheets.readonly',
].join(' ');

export interface DriveFileMetadata {
  id: string;
  name: string;
  mimeType: string;
  modifiedTime: string;
  size?: string;
}

export type DriveSyncStatus = 'idle' | 'connecting' | 'connected' | 'syncing' | 'error';

export interface GoogleDriveState {
  isConnected: boolean;
  status: DriveSyncStatus;
  userEmail: string | null;
  lastSyncAt: Date | null;
  lastFileName: string | null;
  lastFileModifiedAt: string | null;
  errorMessage: string | null;
  autoPollingEnabled: boolean;
}

type Subscriber = (state: GoogleDriveState) => void;

interface TokenClient {
  requestAccessToken: (options?: { prompt?: string }) => void;
}

interface TokenResponse {
  access_token: string;
  error?: string;
  error_description?: string;
  expires_in?: number;
}

declare global {
  interface Window {
    google?: {
      accounts?: {
        oauth2?: {
          initTokenClient: (config: {
            client_id: string;
            scope: string;
            callback: (response: TokenResponse) => void;
            error_callback?: (err: any) => void;
          }) => TokenClient;
          revoke: (accessToken: string, done?: () => void) => void;
        };
      };
    };
  }
}

class GoogleDriveService {
  private accessToken: string | null = null;
  private tokenExpiresAt: number = 0;
  private tokenClient: TokenClient | null = null;
  private subscribers: Set<Subscriber> = new Set();
  private pollingTimer: any = null;
  private cachedDatabaseFolderId: string | null = null;
  private onNewFileCallback: ((file: File, isAuto: boolean) => Promise<void> | void) | null = null;
  private existingFilesProvider: (() => { names: string[]; lastSyncAt: Date | null }) | null = null;

  // 20 minutes polling interval
  public readonly POLLING_INTERVAL_MS = 20 * 60 * 1000;

  private state: GoogleDriveState = {
    isConnected: false,
    status: 'idle',
    userEmail: null,
    lastSyncAt: null,
    lastFileName: null,
    lastFileModifiedAt: null,
    errorMessage: null,
    autoPollingEnabled: false,
  };

  constructor() {
    this.loadPersistedMetadata();
  }

  private loadPersistedMetadata() {
    try {
      const savedMeta = localStorage.getItem('sesap_drive_meta');
      if (savedMeta) {
        const parsed = JSON.parse(savedMeta);
        this.state = {
          ...this.state,
          lastFileName: parsed.lastFileName || null,
          lastFileModifiedAt: parsed.lastFileModifiedAt || null,
          userEmail: parsed.userEmail || null,
          lastSyncAt: parsed.lastSyncAt ? new Date(parsed.lastSyncAt) : null,
        };
      }
    } catch (e) {
      // Ignore localStorage errors
    }
  }

  private persistMetadata() {
    try {
      localStorage.setItem(
        'sesap_drive_meta',
        JSON.stringify({
          lastFileName: this.state.lastFileName,
          lastFileModifiedAt: this.state.lastFileModifiedAt,
          userEmail: this.state.userEmail,
          lastSyncAt: this.state.lastSyncAt ? this.state.lastSyncAt.toISOString() : null,
        })
      );
    } catch (e) {
      // Ignore localStorage errors
    }
  }

  public getState(): GoogleDriveState {
    return { ...this.state };
  }

  public subscribe(fn: Subscriber): () => void {
    this.subscribers.add(fn);
    fn(this.getState());
    return () => {
      this.subscribers.delete(fn);
    };
  }

  private notify() {
    const currentState = this.getState();
    this.subscribers.forEach((fn) => {
      try {
        fn(currentState);
      } catch (err) {
        console.error('Erro no subscriber do Google Drive:', err);
      }
    });
  }

  private setState(partial: Partial<GoogleDriveState>) {
    this.state = { ...this.state, ...partial };
    this.notify();
    this.persistMetadata();
  }

  /**
   * Dynamically loads Google Identity Services script if not already present
   */
  public async loadGisScript(): Promise<void> {
    if (window.google?.accounts?.oauth2) {
      return;
    }

    return new Promise((resolve, reject) => {
      const existingScript = document.getElementById('google-gis-script');
      if (existingScript) {
        let checks = 0;
        const interval = setInterval(() => {
          checks++;
          if (window.google?.accounts?.oauth2) {
            clearInterval(interval);
            resolve();
          } else if (checks > 50) {
            clearInterval(interval);
            reject(new Error('Tempo limite para carregamento do Google Identity Services'));
          }
        }, 100);
        return;
      }

      const script = document.createElement('script');
      script.id = 'google-gis-script';
      script.src = 'https://accounts.google.com/gsi/client';
      script.async = true;
      script.defer = true;
      script.onload = () => {
        let checks = 0;
        const interval = setInterval(() => {
          checks++;
          if (window.google?.accounts?.oauth2) {
            clearInterval(interval);
            resolve();
          } else if (checks > 30) {
            clearInterval(interval);
            resolve(); // Try anyway
          }
        }, 50);
      };
      script.onerror = () => reject(new Error('Falha ao carregar script do Google Identity Services'));
      document.head.appendChild(script);
    });
  }

  /**
   * Connect to Google Drive using OAuth2 (Popup Flow via Google Identity Services)
   */
  public async connectGoogleDrive(): Promise<boolean> {
    this.setState({ status: 'connecting', errorMessage: null });

    try {
      await this.loadGisScript();

      if (!window.google?.accounts?.oauth2) {
        throw new Error('Google Identity Services não inicializado.');
      }

      return new Promise<boolean>((resolve, reject) => {
        try {
          this.tokenClient = window.google!.accounts!.oauth2!.initTokenClient({
            client_id: GOOGLE_CLIENT_ID,
            scope: GOOGLE_DRIVE_SCOPES,
            callback: async (tokenResponse: TokenResponse) => {
              if (tokenResponse.error) {
                this.setState({
                  status: 'error',
                  errorMessage: tokenResponse.error_description || tokenResponse.error,
                  isConnected: false,
                });
                reject(new Error(tokenResponse.error_description || tokenResponse.error));
                return;
              }

              if (tokenResponse.access_token) {
                this.accessToken = tokenResponse.access_token;
                const expiresInSec = tokenResponse.expires_in || 3600;
                this.tokenExpiresAt = Date.now() + (expiresInSec - 60) * 1000;

                // Attempt to fetch user identity (email) silently
                let userEmail: string | null = null;
                try {
                  const userInfoRes = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
                    headers: { Authorization: `Bearer ${this.accessToken}` },
                  });
                  if (userInfoRes.ok) {
                    const uData = await userInfoRes.json();
                    userEmail = uData.email || null;
                  }
                } catch {
                  // userinfo might not be granted, keep email as null or previously known
                }

                this.setState({
                  isConnected: true,
                  status: 'connected',
                  userEmail: userEmail || this.state.userEmail || 'Google Drive Conectado',
                  errorMessage: null,
                });

                // Start background 15-minute polling if callback is registered
                if (this.onNewFileCallback) {
                  this.startAutoPolling();
                }

                resolve(true);
              } else {
                this.setState({
                  status: 'error',
                  errorMessage: 'Nenhum token de acesso retornado pelo Google.',
                  isConnected: false,
                });
                reject(new Error('Nenhum token retornado.'));
              }
            },
            error_callback: (err: any) => {
              const msg = err?.message || 'Falha na autenticação do Google';
              this.setState({
                status: 'error',
                errorMessage: msg,
                isConnected: false,
              });
              reject(new Error(msg));
            },
          });

          // Open Google OAuth consent popup
          this.tokenClient.requestAccessToken({ prompt: '' });
        } catch (err: any) {
          this.setState({
            status: 'error',
            errorMessage: err?.message || 'Falha ao solicitar autorização.',
            isConnected: false,
          });
          reject(err);
        }
      });
    } catch (err: any) {
      this.setState({
        status: 'error',
        errorMessage: err?.message || 'Erro ao conectar ao Google Drive.',
        isConnected: false,
      });
      return false;
    }
  }

  /**
   * Disconnect from Google Drive & clear in-memory tokens
   */
  public disconnectGoogleDrive(): void {
    if (this.accessToken && window.google?.accounts?.oauth2?.revoke) {
      try {
        window.google.accounts.oauth2.revoke(this.accessToken, () => {});
      } catch (e) {
        // Ignore revocation error
      }
    }

    this.stopAutoPolling();
    this.accessToken = null;
    this.tokenExpiresAt = 0;
    this.cachedDatabaseFolderId = null;
    this.setState({
      isConnected: false,
      status: 'idle',
      errorMessage: null,
      autoPollingEnabled: false,
    });
  }

  public isConnected(): boolean {
    return this.state.isConnected && !!this.accessToken && Date.now() < this.tokenExpiresAt;
  }

  /**
   * Ensures the dedicated database folder 'Painel_Financeiro_SESAP_Dados' exists in Google Drive,
   * creating it automatically if not found, and returns its folder ID.
   */
  public async ensureDatabaseFolder(): Promise<string> {
    if (!this.accessToken) {
      throw new Error('Usuário não autenticado no Google Drive.');
    }

    if (Date.now() >= this.tokenExpiresAt) {
      await this.connectGoogleDrive();
      if (!this.accessToken) {
        throw new Error('Sessão do Google Drive expirada.');
      }
    }

    if (this.cachedDatabaseFolderId) {
      return this.cachedDatabaseFolderId;
    }

    const searchUrl = new URL('https://www.googleapis.com/drive/v3/files');
    searchUrl.searchParams.set(
      'q',
      `mimeType='application/vnd.google-apps.folder' and name='${DRIVE_DATABASE_FOLDER_NAME}' and trashed=false`
    );
    searchUrl.searchParams.set('spaces', 'drive');
    searchUrl.searchParams.set('fields', 'files(id, name)');

    const searchRes = await fetch(searchUrl.toString(), {
      headers: { Authorization: `Bearer ${this.accessToken}` },
    });

    if (!searchRes.ok) {
      const errJson = await searchRes.json().catch(() => ({}));
      throw new Error(errJson.error?.message || `Erro ao buscar pasta no Drive (${searchRes.status})`);
    }

    const searchData = await searchRes.json();
    if (Array.isArray(searchData.files) && searchData.files.length > 0) {
      this.cachedDatabaseFolderId = searchData.files[0].id;
      return this.cachedDatabaseFolderId!;
    }

    // Create dedicated database folder
    const createRes = await fetch('https://www.googleapis.com/drive/v3/files?fields=id,name', {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${this.accessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        name: DRIVE_DATABASE_FOLDER_NAME,
        mimeType: 'application/vnd.google-apps.folder',
      }),
    });

    if (!createRes.ok) {
      const errJson = await createRes.json().catch(() => ({}));
      throw new Error(errJson.error?.message || `Erro ao criar pasta "${DRIVE_DATABASE_FOLDER_NAME}" (${createRes.status})`);
    }

    const folderData = await createRes.json();
    this.cachedDatabaseFolderId = folderData.id;
    return this.cachedDatabaseFolderId!;
  }

  /**
   * Saves (creates or updates) a JSON state file directly inside the dedicated Painel_Financeiro_SESAP_Dados folder.
   */
  public async saveStateToDriveFolder(fileName: string, jsonContent: string): Promise<boolean> {
    try {
      const folderId = await this.ensureDatabaseFolder();
      const escapedName = fileName.replace(/'/g, "\\'");

      const findUrl = new URL('https://www.googleapis.com/drive/v3/files');
      findUrl.searchParams.set('q', `name='${escapedName}' and '${folderId}' in parents and trashed=false`);
      findUrl.searchParams.set('fields', 'files(id, name)');

      const findRes = await fetch(findUrl.toString(), {
        headers: { Authorization: `Bearer ${this.accessToken}` },
      });

      if (!findRes.ok) {
        throw new Error(`Erro ao verificar arquivo "${fileName}" na pasta (${findRes.status})`);
      }

      const findData = await findRes.json();
      const existingFile = Array.isArray(findData.files) && findData.files.length > 0 ? findData.files[0] : null;

      if (existingFile?.id) {
        // Update existing file content inside dedicated folder
        const patchRes = await fetch(
          `https://www.googleapis.com/upload/drive/v3/files/${existingFile.id}?uploadType=media`,
          {
            method: 'PATCH',
            headers: {
              Authorization: `Bearer ${this.accessToken}`,
              'Content-Type': 'application/json',
            },
            body: jsonContent,
          }
        );
        if (!patchRes.ok) {
          throw new Error(`Erro ao atualizar "${fileName}" no Google Drive (${patchRes.status})`);
        }
        return true;
      }

      // Create new file strictly inside parents: [folderId]
      const metadata = {
        name: fileName,
        mimeType: 'application/json',
        parents: [folderId],
      };

      const boundary = `-------sesap_drive_boundary_${Date.now()}`;
      const multipartBody =
        `--${boundary}\r\n` +
        `Content-Type: application/json; charset=UTF-8\r\n\r\n` +
        `${JSON.stringify(metadata)}\r\n` +
        `--${boundary}\r\n` +
        `Content-Type: application/json; charset=UTF-8\r\n\r\n` +
        `${jsonContent}\r\n` +
        `--${boundary}--`;

      const createRes = await fetch(
        'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name',
        {
          method: 'POST',
          headers: {
            Authorization: `Bearer ${this.accessToken}`,
            'Content-Type': `multipart/related; boundary="${boundary}"`,
          },
          body: multipartBody,
        }
      );

      if (!createRes.ok) {
        throw new Error(`Erro ao criar "${fileName}" na pasta do Drive (${createRes.status})`);
      }

      return true;
    } catch (err) {
      console.warn('[GoogleDriveService] saveStateToDriveFolder falhou:', err);
      return false;
    }
  }

  /**
   * Loads a JSON state file from inside the dedicated Painel_Financeiro_SESAP_Dados folder.
   */
  public async loadStateFromDriveFolder(fileName: string): Promise<any | null> {
    try {
      const folderId = await this.ensureDatabaseFolder();
      const escapedName = fileName.replace(/'/g, "\\'");

      const findUrl = new URL('https://www.googleapis.com/drive/v3/files');
      findUrl.searchParams.set('q', `name='${escapedName}' and '${folderId}' in parents and trashed=false`);
      findUrl.searchParams.set('orderBy', 'modifiedTime desc');
      findUrl.searchParams.set('fields', 'files(id, name, modifiedTime)');

      const findRes = await fetch(findUrl.toString(), {
        headers: { Authorization: `Bearer ${this.accessToken}` },
      });

      if (!findRes.ok) return null;
      const findData = await findRes.json();
      const targetFile = Array.isArray(findData.files) && findData.files.length > 0 ? findData.files[0] : null;
      if (!targetFile?.id) return null;

      const mediaRes = await fetch(`https://www.googleapis.com/drive/v3/files/${targetFile.id}?alt=media`, {
        headers: { Authorization: `Bearer ${this.accessToken}` },
      });
      if (!mediaRes.ok) return null;

      return await mediaRes.json();
    } catch (err) {
      console.warn('[GoogleDriveService] loadStateFromDriveFolder falhou:', err);
      return null;
    }
  }

  /**
   * Searches user's Google Drive for the most recently modified SIGEF / financial spreadsheet.
   * Priority:
   * 1. Spreadsheets or files with SIGEF / PP / OB / Liq in their name.
   * 2. Any recent spreadsheet file (.xlsx, .xls, .csv, Google Sheet).
   */
  /**
   * Lists spreadsheets and CSVs in the user's Google Drive using a broad MIME query.
   * Returns metadata only (no file downloads), delegating filtering to syncNow.
   */
  public async fetchRelevantFinancialMetadataList(): Promise<DriveFileMetadata[]> {
    if (!this.accessToken) {
      throw new Error('Usuário não autenticado no Google Drive.');
    }

    // Check token expiration
    if (Date.now() >= this.tokenExpiresAt) {
      await this.connectGoogleDrive();
      if (!this.accessToken) {
        throw new Error('Sessão expirada. Por favor, reconecte ao Google Drive.');
      }
    }

    // Busca ampla: apenas tipos MIME de planilha/csv e trashed = false
    const query =
      "trashed = false and (" +
      "mimeType = 'application/vnd.google-apps.spreadsheet' or " +
      "mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or " +
      "mimeType = 'application/vnd.ms-excel' or " +
      "mimeType = 'text/csv'" +
      ")";

    const url = new URL('https://www.googleapis.com/drive/v3/files');
    url.searchParams.set('q', query);
    url.searchParams.set('orderBy', 'modifiedTime desc');
    url.searchParams.set('pageSize', '100');
    url.searchParams.set('fields', 'files(id, name, mimeType, modifiedTime, size)');

    const res = await fetch(url.toString(), {
      headers: {
        Authorization: `Bearer ${this.accessToken}`,
      },
    });

    if (!res.ok) {
      const errJson = await res.json().catch(() => ({}));
      throw new Error(errJson.error?.message || `Erro da API Google Drive (${res.status})`);
    }

    const data = await res.json();
    const files: DriveFileMetadata[] = data.files || [];
    return files;
  }

  /**
   * Searches user's Google Drive for the primary financial spreadsheet.
   */
  public async fetchLatestFinancialMetadata(): Promise<DriveFileMetadata | null> {
    const list = await this.fetchRelevantFinancialMetadataList();
    return list[0] || null;
  }

  /**
   * Downloads a specific file from Google Drive into memory as a File object.
   * If it's a native Google Sheet, exports it as an XLSX file.
   * If it's an existing Excel or CSV file, downloads it via alt=media.
   */
  public async downloadFileAsBrowserFile(meta: DriveFileMetadata): Promise<File> {
    if (!this.accessToken) {
      throw new Error('Usuário não autenticado.');
    }

    let downloadUrl: string;
    let fileName = meta.name;

    if (meta.mimeType === 'application/vnd.google-apps.spreadsheet') {
      // Export native Google Sheet to XLSX
      downloadUrl = `https://www.googleapis.com/drive/v3/files/${meta.id}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`;
      if (!fileName.toLowerCase().endsWith('.xlsx')) {
        fileName = `${fileName}.xlsx`;
      }
    } else {
      // Direct binary download for .xlsx, .xls, .csv, etc.
      downloadUrl = `https://www.googleapis.com/drive/v3/files/${meta.id}?alt=media`;
    }

    const res = await fetch(downloadUrl, {
      headers: {
        Authorization: `Bearer ${this.accessToken}`,
      },
    });

    if (!res.ok) {
      throw new Error(`Falha ao baixar arquivo "${meta.name}" do Google Drive (${res.status})`);
    }

    const blob = await res.blob();
    const mimeType = blob.type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

    // Construct a standard HTML5 File object
    const file = new File([blob], fileName, {
      type: mimeType,
      lastModified: new Date(meta.modifiedTime).getTime(),
    });

    return file;
  }

  /**
   * Fetches the latest financial spreadsheet from user's Google Drive,
   * converts it to a browser File object, and delivers it directly to the caller.
   */
  public async fetchLatestFinancialSheet(): Promise<{ file: File; meta: DriveFileMetadata } | null> {
    const meta = await this.fetchLatestFinancialMetadata();
    if (!meta) {
      return null;
    }

    const file = await this.downloadFileAsBrowserFile(meta);
    return { file, meta };
  }

  /**
   * Registers callback to deliver new/updated files to the application parser
   */
  public registerNewFileCallback(callback: (file: File, isAuto: boolean) => Promise<void> | void) {
    this.onNewFileCallback = callback;
  }

  /**
   * Registers provider of currently loaded file names and last sync date for safe background polling
   */
  public registerExistingFilesProvider(provider: () => { names: string[]; lastSyncAt: Date | null }) {
    this.existingFilesProvider = provider;
  }

  /**
   * Incremental smart synchronization:
   * Only downloads files already loaded in the dashboard (`existingFileNames`)
   * when forced or modified after `lastSyncAt`.
   */
  public async syncNow(
    existingFileNames: string[] = [],
    lastSyncAt: Date | null = null,
    force: boolean = false
  ): Promise<{
    success: boolean;
    hasUpdate: boolean;
    fileName?: string;
    message: string;
  }> {
    if (!this.isConnected()) {
      const connected = await this.connectGoogleDrive();
      if (!connected) {
        return {
          success: false,
          hasUpdate: false,
          message: 'Não conectado ao Google Drive.',
        };
      }
    }

    this.setState({ status: 'syncing', errorMessage: null });

    try {
      const relevantFiles = await this.fetchRelevantFinancialMetadataList();
      if (relevantFiles.length === 0) {
        this.setState({
          status: 'connected',
          errorMessage: null,
        });
        return {
          success: true,
          hasUpdate: false,
          message: 'Nenhuma planilha encontrada no Google Drive.',
        };
      }

      let downloadedCount = 0;
      const fileNames: string[] = [];
      let latestDownloadedModifiedAt: string | null = null;

      for (const meta of relevantFiles) {
        if (!existingFileNames.includes(meta.name)) {
          // Ignora silenciosamente arquivos do Drive que não estão carregados no painel
          continue;
        }

        const shouldDownload =
          force ||
          !lastSyncAt ||
          new Date(meta.modifiedTime).getTime() > lastSyncAt.getTime();

        if (!shouldDownload) {
          continue;
        }

        // Download file into memory one at a time
        const file = await this.downloadFileAsBrowserFile(meta);

        // Deliver to parser callback
        if (this.onNewFileCallback) {
          await this.onNewFileCallback(file, false);
          downloadedCount++;
          fileNames.push(meta.name);
          if (!latestDownloadedModifiedAt) {
            latestDownloadedModifiedAt = meta.modifiedTime;
          }
        }
      }

      this.setState({
        status: 'connected',
        lastSyncAt: new Date(),
        lastFileName: fileNames.length > 0 ? fileNames.join(', ') : this.state.lastFileName,
        lastFileModifiedAt: latestDownloadedModifiedAt || this.state.lastFileModifiedAt,
        errorMessage: null,
      });

      return {
        success: true,
        hasUpdate: downloadedCount > 0,
        fileName: fileNames.join(', '),
        message:
          downloadedCount === 0
            ? 'Nenhuma planilha carregada no painel precisou ser atualizada no Google Drive.'
            : downloadedCount === 1
            ? `Planilha "${fileNames[0]}" sincronizada com sucesso!`
            : `Sincronizadas ${downloadedCount} planilhas: ${fileNames.join(', ')}`,
      };
    } catch (err: any) {
      const errorMsg = err?.message || 'Erro durante a sincronização com o Google Drive.';
      this.setState({
        status: 'error',
        errorMessage: errorMsg,
      });
      return {
        success: false,
        hasUpdate: false,
        message: errorMsg,
      };
    }
  }

  /**
   * Starts the 20-minute background polling loop.
   * Checks silently if newer versions of currently loaded files exist in Drive.
   */
  public startAutoPolling(): void {
    this.stopAutoPolling();
    this.setState({ autoPollingEnabled: true });

    this.pollingTimer = setInterval(async () => {
      if (!this.isConnected()) return;

      try {
        const provided = this.existingFilesProvider ? this.existingFilesProvider() : null;
        const existingFileNames = provided?.names || [];
        if (existingFileNames.length === 0) return;

        const relevantFiles = await this.fetchRelevantFinancialMetadataList();
        if (relevantFiles.length === 0) return;

        const stateMod = this.state.lastFileModifiedAt
          ? new Date(this.state.lastFileModifiedAt).getTime()
          : 0;
        const providedMod = provided?.lastSyncAt ? provided.lastSyncAt.getTime() : 0;
        const lastMod = Math.max(stateMod, providedMod);

        let anyUpdated = false;
        const updatedNames: string[] = [];
        let newestModTime: string | null = null;

        for (const meta of relevantFiles) {
          if (!existingFileNames.includes(meta.name)) {
            continue;
          }

          const currentMod = new Date(meta.modifiedTime).getTime();
          if (currentMod > lastMod) {
            console.log('[GoogleDriveService] Nova versão detectada no Drive:', meta.name, meta.modifiedTime);
            const file = await this.downloadFileAsBrowserFile(meta);

            if (this.onNewFileCallback) {
              await this.onNewFileCallback(file, true);
              anyUpdated = true;
              updatedNames.push(meta.name);
              if (!newestModTime) {
                newestModTime = meta.modifiedTime;
              }
            }
          }
        }

        if (anyUpdated) {
          this.setState({
            lastSyncAt: new Date(),
            lastFileName: updatedNames.join(', '),
            lastFileModifiedAt: newestModTime || this.state.lastFileModifiedAt,
          });
        }
      } catch (pollErr) {
        console.warn('[GoogleDriveService] Erro durante verificação em segundo plano:', pollErr);
      }
    }, this.POLLING_INTERVAL_MS);
  }

  public stopAutoPolling(): void {
    if (this.pollingTimer) {
      clearInterval(this.pollingTimer);
      this.pollingTimer = null;
    }
    this.setState({ autoPollingEnabled: false });
  }
}

export const googleDriveService = new GoogleDriveService();
