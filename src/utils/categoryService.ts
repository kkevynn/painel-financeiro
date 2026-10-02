import { CategoryDefinition, CategoryMetrics, ConsolidatedSupplier } from '../types';

export const COLOR_OPTIONS = [
  { id: 'emerald', name: 'Verde Esmeralda', bg: 'bg-emerald-500' },
  { id: 'blue', name: 'Azul', bg: 'bg-blue-500' },
  { id: 'purple', name: 'Roxo', bg: 'bg-purple-500' },
  { id: 'amber', name: 'Âmbar / Laranja', bg: 'bg-amber-500' },
  { id: 'teal', name: 'Verde Água', bg: 'bg-teal-500' },
  { id: 'rose', name: 'Rosa / Vermelho', bg: 'bg-rose-500' },
  { id: 'indigo', name: 'Índigo', bg: 'bg-indigo-500' },
  { id: 'slate', name: 'Cinza Neutro', bg: 'bg-slate-500' },
];

export const COLOR_MAP: Record<
  string,
  {
    badge: string;
    badgeText: string;
    bgLight: string;
    border: string;
    text: string;
    accent: string;
    borderLeft: string;
  }
> = {
  emerald: {
    badge: 'bg-emerald-500',
    badgeText: 'bg-emerald-100 text-emerald-800 border-emerald-200',
    bgLight: 'bg-emerald-50/50',
    border: 'border-emerald-200',
    text: 'text-emerald-950',
    accent: 'text-emerald-700',
    borderLeft: 'border-l-emerald-500',
  },
  blue: {
    badge: 'bg-blue-500',
    badgeText: 'bg-blue-100 text-blue-800 border-blue-200',
    bgLight: 'bg-blue-50/50',
    border: 'border-blue-200',
    text: 'text-blue-950',
    accent: 'text-blue-700',
    borderLeft: 'border-l-blue-500',
  },
  purple: {
    badge: 'bg-purple-500',
    badgeText: 'bg-purple-100 text-purple-800 border-purple-200',
    bgLight: 'bg-purple-50/50',
    border: 'border-purple-200',
    text: 'text-purple-950',
    accent: 'text-purple-700',
    borderLeft: 'border-l-purple-500',
  },
  amber: {
    badge: 'bg-amber-500',
    badgeText: 'bg-amber-100 text-amber-800 border-amber-200',
    bgLight: 'bg-amber-50/50',
    border: 'border-amber-200',
    text: 'text-amber-950',
    accent: 'text-amber-700',
    borderLeft: 'border-l-amber-500',
  },
  teal: {
    badge: 'bg-teal-500',
    badgeText: 'bg-teal-100 text-teal-800 border-teal-200',
    bgLight: 'bg-teal-50/50',
    border: 'border-teal-200',
    text: 'text-teal-950',
    accent: 'text-teal-700',
    borderLeft: 'border-l-teal-500',
  },
  rose: {
    badge: 'bg-rose-500',
    badgeText: 'bg-rose-100 text-rose-800 border-rose-200',
    bgLight: 'bg-rose-50/50',
    border: 'border-rose-200',
    text: 'text-rose-950',
    accent: 'text-rose-700',
    borderLeft: 'border-l-rose-500',
  },
  indigo: {
    badge: 'bg-indigo-500',
    badgeText: 'bg-indigo-100 text-indigo-800 border-indigo-200',
    bgLight: 'bg-indigo-50/50',
    border: 'border-indigo-200',
    text: 'text-indigo-950',
    accent: 'text-indigo-700',
    borderLeft: 'border-l-indigo-500',
  },
  slate: {
    badge: 'bg-slate-500',
    badgeText: 'bg-slate-100 text-slate-800 border-slate-200',
    bgLight: 'bg-slate-50/50',
    border: 'border-slate-200',
    text: 'text-slate-900',
    accent: 'text-slate-700',
    borderLeft: 'border-l-slate-400',
  },
};

export const CATEGORY_SYNC_EVENT = 'sigecon_category_sync_event';

export function subscribeToCategories(callback: () => void): () => void {
  const handler = () => {
    cachedCategories = null;
    cachedOverrides = null;
    callback();
  };
  window.addEventListener(CATEGORY_SYNC_EVENT, handler);
  return () => {
    window.removeEventListener(CATEGORY_SYNC_EVENT, handler);
  };
}

export function notifyCategoriesChanged() {
  if (typeof window !== 'undefined') {
    window.dispatchEvent(new CustomEvent(CATEGORY_SYNC_EVENT));
  }
}

export const DEFAULT_CATEGORIES: CategoryDefinition[] = [
  {
    id: 'cat_med',
    name: 'Medicamentos e Farmacêuticos',
    color: 'emerald',
    description: 'Insumos farmacêuticos, remédios, vacinas, soluções e produtos de farmácia.',
    isDefault: true,
    keywords: [],
  },
  {
    id: 'cat_hosp',
    name: 'Equipamentos e Materiais Hospitalares',
    color: 'blue',
    description: 'Materiais cirúrgicos, descartáveis, equipamentos médicos e hospitalares.',
    isDefault: true,
    keywords: [],
  },
  {
    id: 'cat_saude',
    name: 'Serviços Médicos e Saúde',
    color: 'purple',
    description: 'Serviços médicos especializados, clínicas, análises e diagnósticos.',
    isDefault: true,
    keywords: [],
  },
  {
    id: 'cat_eng',
    name: 'Engenharia, Obras e Manutenção',
    color: 'amber',
    description: 'Obras, reformas, manutenção predial, ar-condicionado e instalações.',
    isDefault: true,
    keywords: [],
  },
  {
    id: 'cat_serv',
    name: 'Tecnologia, Logística e Serviços Gerais',
    color: 'teal',
    description: 'Tecnologia da informação, veículos, combustível, limpeza e vigilância.',
    isDefault: true,
    keywords: [],
  },
  {
    id: 'cat_alim',
    name: 'Alimentação e Nutrição',
    color: 'rose',
    description: 'Fornecimento de refeições, gênero alimentício e nutrição hospitalar.',
    isDefault: true,
    keywords: [],
  },
  {
    id: 'cat_outros',
    name: 'Não Classificados',
    color: 'slate',
    description: 'Fornecedores pendentes de classificação manual.',
    isDefault: true,
    keywords: [],
  },
];

const CATEGORIES_STORAGE_KEY = 'sigecon_categories_config';
const OVERRIDES_STORAGE_KEY = 'sigecon_supplier_category_overrides';

export function normalizeText(text: string): string {
  if (!text) return '';
  return text
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .trim();
}

let cachedCategories: CategoryDefinition[] | null = null;
let lastCategoriesRaw: string | null = null;

export function getCategories(): CategoryDefinition[] {
  try {
    const stored = localStorage.getItem(CATEGORIES_STORAGE_KEY);
    if (stored === lastCategoriesRaw && cachedCategories !== null) {
      return cachedCategories;
    }
    lastCategoriesRaw = stored;
    if (stored) {
      const parsed = JSON.parse(stored);
      if (Array.isArray(parsed) && parsed.length > 0) {
        cachedCategories = parsed;
        return parsed;
      }
    }
  } catch (err) {
    console.error('Erro ao ler categorias do localStorage:', err);
  }
  cachedCategories = DEFAULT_CATEGORIES;
  return DEFAULT_CATEGORIES;
}

export function saveCategories(categories: CategoryDefinition[]): void {
  try {
    cachedCategories = categories;
    const str = JSON.stringify(categories);
    lastCategoriesRaw = str;
    localStorage.setItem(CATEGORIES_STORAGE_KEY, str);
    notifyCategoriesChanged();
  } catch (err) {
    console.error('Erro ao salvar categorias no localStorage:', err);
  }
}

let cachedOverrides: Record<string, string> | null = null;
let lastOverridesRaw: string | null = null;

export function getSupplierOverrides(): Record<string, string> {
  try {
    const stored = localStorage.getItem(OVERRIDES_STORAGE_KEY);
    if (stored === lastOverridesRaw && cachedOverrides !== null) {
      return cachedOverrides;
    }
    lastOverridesRaw = stored;
    if (stored) {
      cachedOverrides = JSON.parse(stored);
      return cachedOverrides || {};
    }
  } catch (err) {
    console.error('Erro ao ler overrides de fornecedores:', err);
  }
  cachedOverrides = {};
  return {};
}

export function saveSupplierOverride(cnpjOrName: string, categoryId: string): Record<string, string> {
  const current = getSupplierOverrides();
  const updated = { ...current, [cnpjOrName]: categoryId };
  try {
    cachedOverrides = updated;
    const str = JSON.stringify(updated);
    lastOverridesRaw = str;
    localStorage.setItem(OVERRIDES_STORAGE_KEY, str);
    notifyCategoriesChanged();
  } catch (err) {
    console.error('Erro ao salvar override de fornecedor:', err);
  }
  return updated;
}

export function saveMultipleSupplierOverrides(
  supplierKeys: string[],
  categoryId: string
): Record<string, string> {
  const current = getSupplierOverrides();
  const updated = { ...current };
  supplierKeys.forEach((key) => {
    if (key) {
      updated[key] = categoryId;
    }
  });
  try {
    cachedOverrides = updated;
    const str = JSON.stringify(updated);
    lastOverridesRaw = str;
    localStorage.setItem(OVERRIDES_STORAGE_KEY, str);
    notifyCategoriesChanged();
  } catch (err) {
    console.error('Erro ao salvar múltiplos overrides:', err);
  }
  return updated;
}

export function removeSupplierOverride(cnpjOrName: string): Record<string, string> {
  const current = { ...getSupplierOverrides() };
  delete current[cnpjOrName];
  try {
    cachedOverrides = current;
    const str = JSON.stringify(current);
    lastOverridesRaw = str;
    localStorage.setItem(OVERRIDES_STORAGE_KEY, str);
    notifyCategoriesChanged();
  } catch (err) {
    console.error('Erro ao remover override de fornecedor:', err);
  }
  return current;
}

export function categorizeSupplier(
  supplierName: string,
  cnpj: string,
  categories: CategoryDefinition[],
  overrides: Record<string, string>
): CategoryDefinition {
  const validCnpj = cnpj && cnpj !== 'N/I' ? cnpj : null;
  const overrideCatId = (validCnpj && overrides[validCnpj]) || (supplierName && overrides[supplierName]);
  if (overrideCatId) {
    const found = categories.find((c) => c.id === overrideCatId);
    if (found) return found;
  }

  const outros = categories.find((c) => c.id === 'cat_outros') || categories[categories.length - 1];
  return outros || DEFAULT_CATEGORIES[DEFAULT_CATEGORIES.length - 1];
}

export function groupSuppliersByCategory(
  suppliers: ConsolidatedSupplier[],
  categories: CategoryDefinition[],
  overrides: Record<string, string>
): CategoryMetrics[] {
  const categoryMap = new Map<string, CategoryMetrics>();

  categories.forEach((cat) => {
    categoryMap.set(cat.id, {
      category: cat,
      totalValue: 0,
      totalCount: 0,
      suppliers: [],
    });
  });

  suppliers.forEach((supplier) => {
    const category = categorizeSupplier(supplier.name, supplier.cnpj, categories, overrides);
    let metrics = categoryMap.get(category.id);

    if (!metrics) {
      const fallbackCat = categories.find((c) => c.id === 'cat_outros') || categories[0];
      metrics = categoryMap.get(fallbackCat.id)!;
    }

    metrics.totalValue += supplier.totalValue;
    metrics.totalCount += 1;
    metrics.suppliers.push(supplier);
  });

  const result = Array.from(categoryMap.values()).map((metrics) => ({
    ...metrics,
    suppliers: [...metrics.suppliers].sort((a, b) => b.totalValue - a.totalValue),
  }));

  return result.sort((a, b) => b.totalValue - a.totalValue);
}
