import { enrichItemWithKnowledge } from './appKnowledgeDirectory';

export interface ParaPublicarItem {
  id: string; // Ex: pub_NL00123_12345678000199
  nlNumero: string;
  documentNumber?: string;
  fornecedor: string;
  supplierName?: string;
  cnpj: string;
  valor: number;
  value?: number;
  fonteRecurso?: string;
  stageOrigem: 'obedece' | 'nao_obedece' | 'outros' | string;
  sourceType?: string;
  sourceLabel?: string;
  status?: string;
  dataSelecao: string;
  addedAt?: number;
  dataReferencia?: string;
  date?: string;
  observacao?: string;
}

export type ItemParaPublicar = ParaPublicarItem;

const STORAGE_KEY_PARA_PUBLICAR = 'sesap_itens_para_publicar_v1';

export type ParaPublicarListener = (items?: ParaPublicarItem[]) => void;

class ParaPublicarService {
  private items: ParaPublicarItem[] = [];
  private listeners: Set<ParaPublicarListener> = new Set();
  private onItemsChangedCloudCallback: (() => void) | null = null;

  constructor() {
    this.loadFromStorage();
  }

  public registerCloudSyncCallback(cb: () => void) {
    this.onItemsChangedCloudCallback = cb;
  }

  private loadFromStorage() {
    try {
      if (typeof localStorage !== 'undefined') {
        const raw = localStorage.getItem(STORAGE_KEY_PARA_PUBLICAR);
        if (raw) {
          const parsed = JSON.parse(raw);
          if (Array.isArray(parsed)) {
            this.items = parsed;
          }
        }
      }
    } catch (e) {
      this.items = [];
    }
  }

  private saveToStorage() {
    try {
      if (typeof localStorage !== 'undefined') {
        localStorage.setItem(STORAGE_KEY_PARA_PUBLICAR, JSON.stringify(this.items));
      }
    } catch (e) {
      // safe quota catch
    }
  }

  private notify() {
    this.saveToStorage();
    const currentItems = this.getItems();
    setTimeout(() => {
      this.listeners.forEach((fn) => {
        try {
          fn(currentItems);
        } catch (err) {
          console.error('Erro no listener do ParaPublicarService:', err);
        }
      });
      if (this.onItemsChangedCloudCallback) {
        try {
          this.onItemsChangedCloudCallback();
        } catch {
          // ignore
        }
      }
    }, 0);
  }

  public subscribe(listener: ParaPublicarListener): () => void {
    this.listeners.add(listener);
    return () => {
      this.listeners.delete(listener);
    };
  }

  public getItems(): ParaPublicarItem[] {
    return this.items.map((it) => enrichItemWithKnowledge(it));
  }

  public getCount(): number {
    return this.items.length;
  }

  public isSelected(nlNumero: string, cnpj?: string): boolean {
    const cleanNL = (nlNumero || '').trim().toUpperCase();
    const cleanCnpj = (cnpj || '').replace(/\D/g, '');
    if (!cleanNL) return false;

    return this.items.some((it) => {
      const itNL = (it.nlNumero || it.documentNumber || '').trim().toUpperCase();
      const itCnpj = (it.cnpj || '').replace(/\D/g, '');
      if (cleanCnpj && itCnpj && cleanCnpj === itCnpj) {
        return itNL === cleanNL;
      }
      return itNL === cleanNL;
    });
  }

  public addItem(item: Partial<ParaPublicarItem> & { nlNumero?: string; documentNumber?: string }): void {
    const enriched = enrichItemWithKnowledge({
      ...item,
      fornecedor: item.fornecedor || item.supplierName || 'NÃO INFORMADO',
      cnpj: item.cnpj || '',
      nlNumero: item.nlNumero || item.documentNumber || '',
    });

    const docNum = (enriched.nlNumero || enriched.documentNumber || '').trim().toUpperCase();
    const fornecedor = enriched.fornecedor || enriched.supplierName || 'NÃO INFORMADO';
    const cnpj = enriched.cnpj || '';
    const valor = typeof item.valor === 'number' ? item.valor : (typeof item.value === 'number' ? item.value : 0);
    const cleanNL = docNum;
    const cleanCnpj = cnpj.replace(/\D/g, '');
    const id = item.id || `pub_${cleanNL}_${cleanCnpj || Date.now()}`;

    // If already exists, do not duplicate
    if (this.isSelected(cleanNL, cnpj)) {
      return;
    }

    const newItem: ParaPublicarItem = {
      ...item,
      id,
      nlNumero: cleanNL,
      documentNumber: cleanNL,
      fornecedor,
      supplierName: fornecedor,
      cnpj,
      valor,
      value: valor,
      stageOrigem: item.stageOrigem || (item.sourceType as any) || 'obedece',
      sourceType: item.sourceType || item.stageOrigem,
      sourceLabel: item.sourceLabel,
      status: item.status,
      dataSelecao: item.dataSelecao || new Date().toISOString(),
      addedAt: item.addedAt || Date.now(),
      dataReferencia: item.dataReferencia || item.date,
      date: item.date || item.dataReferencia,
    };

    this.items.unshift(newItem);
    this.notify();
  }

  public removeItem(idOrNL: string): void {
    const target = idOrNL.trim().toUpperCase();
    this.items = this.items.filter(
      (it) => it.id !== idOrNL && it.nlNumero.toUpperCase() !== target && (it.documentNumber || '').toUpperCase() !== target
    );
    this.notify();
  }

  public toggleItem(item: Partial<ParaPublicarItem> & { nlNumero?: string; documentNumber?: string }): boolean {
    const cleanNL = (item.nlNumero || item.documentNumber || '').trim().toUpperCase();

    const exists = this.isSelected(cleanNL, item.cnpj);
    if (exists) {
      this.removeItem(cleanNL);
      return false; // Removed
    } else {
      this.addItem(item);
      return true; // Added
    }
  }

  public clearAll(): void {
    this.items = [];
    this.notify();
  }

  public restoreItems(items: ParaPublicarItem[]): void {
    if (Array.isArray(items)) {
      this.items = [...items];
      this.notify();
    }
  }
}

export const paraPublicarService = new ParaPublicarService();
