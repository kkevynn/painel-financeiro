/**
 * Background Execution & Anti-Throttling Scheduler
 *
 * Navegadores modernos (Chrome, Edge, Safari, Firefox) aplicam throttling agressivo
 * em setTimeout/setInterval da thread principal (1.000ms a 60.000ms) e podem congelar
 * abas em segundo plano (Memory Saver / Page Lifecycle Freeze) quando o usuário muda de aba.
 *
 * Este módulo combina:
 * 1. Um Web Worker leve (inline via Blob) para temporizadores e intervalos imunes ao throttling de aba inativa.
 * 2. MessageChannel para cessão rápida de event loop (<1ms) sem bloqueio de setTimeout.
 * 3. Web Locks API (navigator.locks) + heartbeat sob demanda durante leitura/processamento ativo de planilhas,
 *    impedindo que o navegador suspenda ou congele a aba antes de concluir a tarefa.
 */

type TimerCallback = () => void;

const WORKER_SOURCE = `
  const timeouts = new Map();
  const intervals = new Map();
  let heartbeatId = null;

  self.onmessage = function(e) {
    const data = e.data || {};
    const { cmd, id, delay } = data;

    if (cmd === 'setTimeout') {
      const t = setTimeout(() => {
        timeouts.delete(id);
        self.postMessage({ type: 'timeout', id });
      }, Math.max(0, Number(delay) || 0));
      timeouts.set(id, t);
    } else if (cmd === 'clearTimeout') {
      const t = timeouts.get(id);
      if (t !== undefined) {
        clearTimeout(t);
        timeouts.delete(id);
      }
    } else if (cmd === 'setInterval') {
      const iv = setInterval(() => {
        self.postMessage({ type: 'interval', id });
      }, Math.max(10, Number(delay) || 1000));
      intervals.set(id, iv);
    } else if (cmd === 'clearInterval') {
      const iv = intervals.get(id);
      if (iv !== undefined) {
        clearInterval(iv);
        intervals.delete(id);
      }
    } else if (cmd === 'startHeartbeat') {
      if (!heartbeatId) {
        heartbeatId = setInterval(() => {
          self.postMessage({ type: 'heartbeat', ts: Date.now() });
        }, 4000);
      }
    } else if (cmd === 'stopHeartbeat') {
      if (heartbeatId) {
        clearInterval(heartbeatId);
        heartbeatId = null;
      }
    }
  };
`;

class BackgroundScheduler {
  private worker: Worker | null = null;
  private workerFailed = false;
  private nextId = 1;
  private lockSeq = 1;
  private activeProcessingTasks = 0;

  private timeoutCallbacks = new Map<number, TimerCallback>();
  private intervalCallbacks = new Map<number, TimerCallback>();

  private ensureWorker(): Worker | null {
    if (this.worker) return this.worker;
    if (this.workerFailed) return null;
    if (typeof window === 'undefined' || typeof Worker === 'undefined' || typeof Blob === 'undefined') {
      this.workerFailed = true;
      return null;
    }

    try {
      const blob = new Blob([WORKER_SOURCE], { type: 'application/javascript' });
      const url = URL.createObjectURL(blob);
      const worker = new Worker(url);
      URL.revokeObjectURL(url);

      worker.onmessage = (event: MessageEvent) => {
        const msg = event.data;
        if (!msg) return;

        if (msg.type === 'timeout') {
          const cb = this.timeoutCallbacks.get(msg.id);
          if (cb) {
            this.timeoutCallbacks.delete(msg.id);
            try {
              cb();
            } catch (err) {
              console.error('[BackgroundScheduler] Erro em timeout:', err);
            }
          }
        } else if (msg.type === 'interval') {
          const cb = this.intervalCallbacks.get(msg.id);
          if (cb) {
            try {
              cb();
            } catch (err) {
              console.error('[BackgroundScheduler] Erro em interval:', err);
            }
          }
        }
        // 'heartbeat' apenas mantém a fila de mensagens ativa durante processamento pesado
      };

      worker.onerror = () => {
        // Em caso de falha do Worker (ex: política CSP restritiva), faz fallback transparente
        this.worker = null;
        this.workerFailed = true;
      };

      this.worker = worker;
      if (this.activeProcessingTasks > 0) {
        worker.postMessage({ cmd: 'startHeartbeat' });
      }
      return worker;
    } catch {
      this.workerFailed = true;
      return null;
    }
  }

  public setTimeout(callback: TimerCallback, delayMs: number): () => void {
    const worker = this.ensureWorker();
    if (worker) {
      const id = this.nextId++;
      this.timeoutCallbacks.set(id, callback);
      worker.postMessage({ cmd: 'setTimeout', id, delay: Math.max(0, delayMs) });
      return () => {
        if (this.timeoutCallbacks.delete(id)) {
          try {
            this.worker?.postMessage({ cmd: 'clearTimeout', id });
          } catch {
            // ignore
          }
        }
      };
    }

    const nativeId = globalThis.setTimeout(callback, Math.max(0, delayMs));
    return () => globalThis.clearTimeout(nativeId);
  }

  public setInterval(callback: TimerCallback, intervalMs: number): () => void {
    const worker = this.ensureWorker();
    if (worker) {
      const id = this.nextId++;
      this.intervalCallbacks.set(id, callback);
      worker.postMessage({ cmd: 'setInterval', id, delay: Math.max(10, intervalMs) });
      return () => {
        if (this.intervalCallbacks.delete(id)) {
          try {
            this.worker?.postMessage({ cmd: 'clearInterval', id });
          } catch {
            // ignore
          }
        }
      };
    }

    const nativeId = globalThis.setInterval(callback, Math.max(10, intervalMs));
    return () => globalThis.clearInterval(nativeId);
  }

  public beginActiveWork(): void {
    this.activeProcessingTasks++;
    if (this.activeProcessingTasks === 1) {
      const worker = this.ensureWorker();
      try {
        worker?.postMessage({ cmd: 'startHeartbeat' });
      } catch {
        // ignore
      }
    }
  }

  public endActiveWork(): void {
    this.activeProcessingTasks = Math.max(0, this.activeProcessingTasks - 1);
    if (this.activeProcessingTasks === 0) {
      try {
        this.worker?.postMessage({ cmd: 'stopHeartbeat' });
      } catch {
        // ignore
      }
    }
  }

  public nextLockName(tag: string): string {
    return `painel-fes-bg-${tag}-${this.lockSeq++}`;
  }
}

const scheduler = new BackgroundScheduler();

/**
 * Agenda um callback após delayMs sem sofrer throttling de aba em segundo plano.
 * Retorna função de cancelamento.
 */
export function scheduleBackgroundTimeout(callback: TimerCallback, delayMs: number): () => void {
  return scheduler.setTimeout(callback, delayMs);
}

/**
 * Agenda um intervalo recorrente em Web Worker, imune ao throttling de aba inativa.
 * Retorna função de cancelamento.
 */
export function scheduleBackgroundInterval(callback: TimerCallback, intervalMs: number): () => void {
  return scheduler.setInterval(callback, intervalMs);
}

/**
 * Cede o controle ao event loop sem sofrer clamp de 1.000ms do setTimeout em abas inativas.
 * Usa MessageChannel (macrotask instantânea) no navegador.
 */
export function yieldWithoutThrottling(): Promise<void> {
  if (typeof MessageChannel !== 'undefined') {
    return new Promise<void>((resolve) => {
      const channel = new MessageChannel();
      channel.port1.onmessage = () => {
        channel.port1.onmessage = null;
        resolve();
      };
      channel.port2.postMessage(null);
    });
  }
  return Promise.resolve();
}

/**
 * Aguarda um pequeno intervalo sem travar em abas em segundo plano.
 * Se a aba já estiver oculta e a pausa for apenas visual (<= 60ms), cede imediatamente via MessageChannel.
 */
export function backgroundDelay(delayMs: number): Promise<void> {
  const isHidden = typeof document !== 'undefined' && document.visibilityState === 'hidden';
  if (delayMs <= 0 || (isHidden && delayMs <= 60)) {
    return yieldWithoutThrottling();
  }
  return new Promise<void>((resolve) => {
    scheduler.setTimeout(resolve, delayMs);
  });
}

/**
 * Mantém a aba ativa e protegida contra congelamento/descarte (Page Lifecycle Freeze / Memory Saver)
 * enquanto uma tarefa de leitura, processamento ou sincronização de planilhas está em execução.
 */
export async function withActiveProcessingLock<T>(
  tag: string,
  task: () => Promise<T>
): Promise<T> {
  scheduler.beginActiveWork();
  try {
    const locks = typeof navigator !== 'undefined' ? (navigator as any).locks : undefined;
    if (locks && typeof locks.request === 'function') {
      const lockName = scheduler.nextLockName(tag);
      return await locks.request(lockName, { mode: 'exclusive' }, async () => {
        return await task();
      });
    }
    return await task();
  } finally {
    scheduler.endActiveWork();
  }
}

/**
 * Registra callback para quando o usuário retorna à aba (ou rede volta),
 * útil para retomar verificações após suspensão completa do sistema operacional.
 */
export function onTabReturnToForeground(callback: () => void): () => void {
  if (typeof document === 'undefined' || typeof window === 'undefined') {
    return () => {};
  }

  const handleVisibility = () => {
    if (document.visibilityState === 'visible') {
      callback();
    }
  };
  const handleOnline = () => {
    callback();
  };

  document.addEventListener('visibilitychange', handleVisibility);
  window.addEventListener('online', handleOnline);

  return () => {
    document.removeEventListener('visibilitychange', handleVisibility);
    window.removeEventListener('online', handleOnline);
  };
}
