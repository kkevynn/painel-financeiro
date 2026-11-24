import { FileData } from '../types';

const DB_NAME_PRIMARY = 'FINANCE_STORE';
const DB_VERSION = 1;
const STORE_NAME = 'app_state';

let cachedDBPromise: Promise<IDBDatabase> | null = null;

function openDB(dbName: string = DB_NAME_PRIMARY): Promise<IDBDatabase> {
  if (cachedDBPromise) return cachedDBPromise;

  const dbPromise = new Promise<IDBDatabase>((resolve, reject) => {
    try {
      if (typeof window === 'undefined' || !window.indexedDB) {
        reject(new Error('IndexedDB não suportado neste navegador.'));
        return;
      }

      const request = indexedDB.open(dbName, DB_VERSION);

      request.onupgradeneeded = (event) => {
        try {
          const db = (event.target as IDBOpenDBRequest).result;
          if (!db.objectStoreNames.contains(STORE_NAME)) {
            db.createObjectStore(STORE_NAME);
          }
        } catch (e) {
          reject(e);
        }
      };

      request.onsuccess = () => {
        const db = request.result;
        db.onclose = () => {
          cachedDBPromise = null;
        };
        resolve(db);
      };

      request.onerror = () => {
        cachedDBPromise = null;
        reject(request.error || new Error('Falha ao abrir IndexedDB'));
      };

      request.onblocked = () => {
        cachedDBPromise = null;
        reject(new Error('IndexedDB bloqueado'));
      };
    } catch (err) {
      cachedDBPromise = null;
      reject(err);
    }
  });

  // Timeout guard (7000ms max) to ensure app never hangs indefinitely during startup
  const timeoutPromise = new Promise<IDBDatabase>((_, reject) => {
    setTimeout(() => reject(new Error('Timeout ao conectar com IndexedDB')), 7000);
  });

  cachedDBPromise = Promise.race([dbPromise, timeoutPromise]).catch((err) => {
    cachedDBPromise = null;
    throw err;
  });

  return cachedDBPromise;
}

export interface StoredAppState {
  files: FileData[];
  lastProcessedAt: string | null;
  savedAt: string;
}

async function saveToDB(dbName: string, payload: StoredAppState): Promise<boolean> {
  try {
    const db = await openDB(dbName);
    const transaction = db.transaction(STORE_NAME, 'readwrite');
    const store = transaction.objectStore(STORE_NAME);

    return new Promise((resolve, reject) => {
      const request = store.put(payload, 'current_dataset');
      request.onsuccess = () => resolve(true);
      request.onerror = () => reject(request.error);
    });
  } catch (err) {
    cachedDBPromise = null;
    return false;
  }
}

async function loadFromDB(dbName: string): Promise<StoredAppState | null> {
  try {
    const db = await openDB(dbName);
    if (!db.objectStoreNames.contains(STORE_NAME)) {
      return null;
    }
    const transaction = db.transaction(STORE_NAME, 'readonly');
    const store = transaction.objectStore(STORE_NAME);

    return await new Promise<StoredAppState | null>((resolve, reject) => {
      const request = store.get('current_dataset');
      request.onsuccess = () => resolve(request.result || null);
      request.onerror = () => reject(request.error);
    });
  } catch (err) {
    cachedDBPromise = null;
    return null;
  }
}

export async function saveAppState(
  files: FileData[],
  lastProcessedAt: Date | null
): Promise<boolean> {
  const payload: StoredAppState = {
    files,
    lastProcessedAt: lastProcessedAt ? lastProcessedAt.toISOString() : null,
    savedAt: new Date().toISOString(),
  };

  let savedSuccessfully = await saveToDB(DB_NAME_PRIMARY, payload);

  try {
    const totalRows = (files || []).reduce((acc, f) => acc + (f.rows?.length || 0), 0);
    // Only mirror to localStorage if very lightweight (< 400 rows total)
    // Avoids heavy synchronous serialization and quota exceeded freeze on large datasets
    if (totalRows > 0 && totalRows < 400) {
      const jsonStr = JSON.stringify(payload);
      if (jsonStr.length < 1.5 * 1024 * 1024) {
        localStorage.setItem('finance_app_state', jsonStr);
      }
    }
    savedSuccessfully = true;
  } catch (e) {
    // IndexedDB is the primary durable store, so localStorage failure is non-blocking
  }

  return savedSuccessfully;
}

export async function loadAppState(): Promise<StoredAppState | null> {
  let data = await loadFromDB(DB_NAME_PRIMARY);
  if (data && Array.isArray(data.files) && data.files.length > 0) {
    return data;
  }

  try {
    const raw = localStorage.getItem('finance_app_state') || localStorage.getItem('sigecon_app_state');
    if (raw) {
      const parsed = JSON.parse(raw);
      if (parsed && Array.isArray(parsed.files) && parsed.files.length > 0) {
        return parsed as StoredAppState;
      }
    }
  } catch (e) {
    console.error('Erro ao ler do localStorage:', e);
  }

  return null;
}

export async function clearAppState(): Promise<boolean> {
  try {
    const db1 = await openDB(DB_NAME_PRIMARY);
    if (db1.objectStoreNames.contains(STORE_NAME)) {
      const tx1 = db1.transaction(STORE_NAME, 'readwrite');
      const req = tx1.objectStore(STORE_NAME).delete('current_dataset');
      await new Promise<void>((resolve) => {
        tx1.oncomplete = () => resolve();
        tx1.onerror = () => resolve();
        req.onerror = () => resolve();
      });
    }
  } catch (err) {
    // Ignore
  }

  try {
    localStorage.removeItem('finance_app_state');
    localStorage.removeItem('sigecon_app_state');
  } catch (e) {
    // Ignore
  }

  return true;
}

export async function saveKVToDB(key: string, value: any): Promise<boolean> {
  try {
    const db = await openDB(DB_NAME_PRIMARY);
    const transaction = db.transaction(STORE_NAME, 'readwrite');
    const store = transaction.objectStore(STORE_NAME);
    return await new Promise<boolean>((resolve, reject) => {
      const request = store.put(value, key);
      request.onsuccess = () => resolve(true);
      request.onerror = () => reject(request.error);
    });
  } catch {
    return false;
  }
}

export async function loadKVFromDB<T = any>(key: string): Promise<T | null> {
  try {
    const db = await openDB(DB_NAME_PRIMARY);
    if (!db.objectStoreNames.contains(STORE_NAME)) {
      return null;
    }
    const transaction = db.transaction(STORE_NAME, 'readonly');
    const store = transaction.objectStore(STORE_NAME);
    return await new Promise<T | null>((resolve, reject) => {
      const request = store.get(key);
      request.onsuccess = () => resolve((request.result as T) ?? null);
      request.onerror = () => reject(request.error);
    });
  } catch {
    return null;
  }
}

export async function deleteKVFromDB(key: string): Promise<boolean> {
  try {
    const db = await openDB(DB_NAME_PRIMARY);
    if (!db.objectStoreNames.contains(STORE_NAME)) {
      return true;
    }
    const transaction = db.transaction(STORE_NAME, 'readwrite');
    const store = transaction.objectStore(STORE_NAME);
    return await new Promise<boolean>((resolve) => {
      const request = store.delete(key);
      request.onsuccess = () => resolve(true);
      request.onerror = () => resolve(false);
    });
  } catch {
    return false;
  }
}

