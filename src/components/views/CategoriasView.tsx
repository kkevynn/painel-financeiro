import React, { useState, useMemo, useEffect } from 'react';
import {
  FolderKanban,
  Plus,
  Trash2,
  Edit2,
  ChevronDown,
  ChevronUp,
  Tag,
  Search,
  SlidersHorizontal,
  Check,
  Layers,
  Building2,
  RotateCcw,
} from 'lucide-react';
import { DashboardMetrics, CategoryDefinition, CategoryMetrics } from '../../types';
import { formatBRL } from '../../utils/excelParser';
import {
  getCategories,
  saveCategories,
  getSupplierOverrides,
  saveSupplierOverride,
  saveMultipleSupplierOverrides,
  removeSupplierOverride,
  groupSuppliersByCategory,
  DEFAULT_CATEGORIES,
  COLOR_MAP,
  COLOR_OPTIONS,
  subscribeToCategories,
} from '../../utils/categoryService';
import { InfoHelpButton } from '../InfoHelpButton';

interface CategoriasViewProps {
  metrics: DashboardMetrics;
}

export const CategoriasView: React.FC<CategoriasViewProps> = ({ metrics }) => {
  const [categories, setCategoriesState] = useState<CategoryDefinition[]>(() => getCategories());
  const [overrides, setOverrides] = useState<Record<string, string>>(() => getSupplierOverrides());

  useEffect(() => {
    const unsub = subscribeToCategories(() => {
      setCategoriesState(getCategories());
      setOverrides(getSupplierOverrides());
    });
    return () => unsub();
  }, []);

  // Workspace subtabs: 'classificar' | 'grupos' | 'gerenciar'
  const [activeSubTab, setActiveSubTab] = useState<'classificar' | 'grupos' | 'gerenciar'>('classificar');

  // Classification Table Filters
  const [searchTerm, setSearchTerm] = useState('');
  const [categoryFilter, setCategoryFilter] = useState<string>('all');
  const [selectedSupplierKeys, setSelectedSupplierKeys] = useState<string[]>([]);
  const [bulkTargetCategory, setBulkTargetCategory] = useState<string>('');

  // Grouped metrics expansion
  const [expandedCategoryId, setExpandedCategoryId] = useState<string | null>(null);

  // Form State
  const [editingCategory, setEditingCategory] = useState<CategoryDefinition | null>(null);
  const [isAddMode, setIsAddMode] = useState(false);
  const [formName, setFormName] = useState('');
  const [formColor, setFormColor] = useState('blue');
  const [formDesc, setFormDesc] = useState('');
  const [deletingCatId, setDeletingCatId] = useState<string | null>(null);
  const [confirmResetCategories, setConfirmResetCategories] = useState(false);
  const [infoAlert, setInfoAlert] = useState<string | null>(null);

  const allSuppliers = metrics.consolidatedSuppliers || [];
  const grandTotal = metrics.grandTotalValue || 0;

  // Classified / Unclassified counters
  const classifiedCount = useMemo(() => {
    return allSuppliers.filter((s) => {
      const key = s.cnpj && s.cnpj !== 'N/I' ? s.cnpj : s.name;
      const catId = overrides[key];
      return catId && catId !== 'cat_outros';
    }).length;
  }, [allSuppliers, overrides]);

  const unclassifiedCount = allSuppliers.length - classifiedCount;

  // Group suppliers into category metrics
  const categoryMetrics = useMemo(() => {
    if (!metrics.isProcessed || !metrics.consolidatedSuppliers) return [];
    return groupSuppliersByCategory(metrics.consolidatedSuppliers, categories, overrides);
  }, [metrics, categories, overrides]);

  // Supplier category change
  const handleSupplierCategoryChange = (cnpj: string, name: string, newCategoryId: string) => {
    const key = cnpj && cnpj !== 'N/I' ? cnpj : name;
    let updated: Record<string, string>;
    if (newCategoryId === 'cat_outros' || !newCategoryId) {
      updated = removeSupplierOverride(key);
    } else {
      updated = saveSupplierOverride(key, newCategoryId);
    }
    setOverrides(updated);
  };

  // Bulk assignment
  const handleApplyBulkAssignment = () => {
    if (!bulkTargetCategory || selectedSupplierKeys.length === 0) return;

    let updated: Record<string, string>;
    if (bulkTargetCategory === 'cat_outros') {
      const current = getSupplierOverrides();
      const nextOverrides = { ...current };
      selectedSupplierKeys.forEach((key) => {
        delete nextOverrides[key];
      });
      localStorage.setItem('sigecon_supplier_category_overrides', JSON.stringify(nextOverrides));
      updated = nextOverrides;
    } else {
      updated = saveMultipleSupplierOverrides(selectedSupplierKeys, bulkTargetCategory);
    }

    setOverrides(updated);
    setSelectedSupplierKeys([]);
    setBulkTargetCategory('');
  };

  const toggleSelectSupplier = (key: string) => {
    setSelectedSupplierKeys((prev) =>
      prev.includes(key) ? prev.filter((k) => k !== key) : [...prev, key]
    );
  };

  const filteredSuppliers = useMemo(() => {
    return allSuppliers.filter((sup) => {
      const key = sup.cnpj && sup.cnpj !== 'N/I' ? sup.cnpj : sup.name;
      const currentCatId = overrides[key] || 'cat_outros';

      if (categoryFilter === 'unassigned') {
        if (currentCatId !== 'cat_outros') return false;
      } else if (categoryFilter !== 'all') {
        if (currentCatId !== categoryFilter) return false;
      }

      if (searchTerm.trim()) {
        const q = searchTerm.toLowerCase();
        const nameMatch = sup.name.toLowerCase().includes(q);
        const cnpjMatch = sup.cnpj.includes(q);
        return nameMatch || cnpjMatch;
      }

      return true;
    });
  }, [allSuppliers, overrides, categoryFilter, searchTerm]);

  const toggleSelectAllFiltered = () => {
    const keys = filteredSuppliers.map((s) => (s.cnpj && s.cnpj !== 'N/I' ? s.cnpj : s.name));
    const allSelected = keys.every((k) => selectedSupplierKeys.includes(k));

    if (allSelected) {
      setSelectedSupplierKeys((prev) => prev.filter((k) => !keys.includes(k)));
    } else {
      setSelectedSupplierKeys((prev) => Array.from(new Set([...prev, ...keys])));
    }
  };

  const handleOpenAddCategory = () => {
    setIsAddMode(true);
    setEditingCategory(null);
    setFormName('');
    setFormColor('blue');
    setFormDesc('');
  };

  const handleOpenEditCategory = (cat: CategoryDefinition) => {
    setIsAddMode(false);
    setEditingCategory(cat);
    setFormName(cat.name);
    setFormColor(cat.color);
    setFormDesc(cat.description || '');
  };

  const handleSaveCategoryForm = (e: React.FormEvent) => {
    e.preventDefault();
    if (!formName.trim()) return;

    let updatedList: CategoryDefinition[] = [];

    if (isAddMode) {
      const newCat: CategoryDefinition = {
        id: `cat_custom_${Date.now()}`,
        name: formName.trim(),
        color: formColor,
        description: formDesc.trim(),
        keywords: [],
        isDefault: false,
      };
      updatedList = [...categories, newCat];
    } else if (editingCategory) {
      updatedList = categories.map((cat) => {
        if (cat.id === editingCategory.id) {
          return {
            ...cat,
            name: formName.trim(),
            color: formColor,
            description: formDesc.trim(),
          };
        }
        return cat;
      });
    }

    setCategoriesState(updatedList);
    saveCategories(updatedList);
    setEditingCategory(null);
    setIsAddMode(false);
  };

  const executeDeleteCategory = (catId: string) => {
    if (catId === 'cat_outros') {
      setInfoAlert('A categoria "Não Classificados" é padrão do sistema e não pode ser removida.');
      setTimeout(() => setInfoAlert(null), 3500);
      return;
    }

    const updatedCategories = categories.filter((c) => c.id !== catId);
    setCategoriesState(updatedCategories);
    saveCategories(updatedCategories);

    const currentOverrides = getSupplierOverrides();
    const newOverrides = { ...currentOverrides };
    Object.keys(newOverrides).forEach((k) => {
      if (newOverrides[k] === catId) {
        delete newOverrides[k];
      }
    });
    localStorage.setItem('sigecon_supplier_category_overrides', JSON.stringify(newOverrides));
    setOverrides(newOverrides);
    setDeletingCatId(null);
  };

  const executeResetDefaults = () => {
    setCategoriesState(DEFAULT_CATEGORIES);
    saveCategories(DEFAULT_CATEGORIES);
    setConfirmResetCategories(false);
  };

  return (
    <div className="space-y-6">
      {/* Top Banner & Header */}
      <div className="bg-gradient-to-r from-slate-900 via-slate-800 to-slate-900 rounded-2xl p-6 md:p-8 text-white shadow-xl border border-slate-800 relative overflow-hidden">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-6 relative z-10">
          <div className="flex items-center gap-3">
            <h2 className="text-2xl font-black text-white tracking-tight">
              Classificação & Categorias de Fornecedores
            </h2>
            <InfoHelpButton
              title="Classificação de Fornecedores"
              variant="dark"
              content={
                <p>
                  Permite agrupar os credores importados por segmento de atuação (ex.: Serviços Médicos, Locação, Tecnologia, Medicamentos) com regras automáticas por palavras-chave ou categorização manual individual.
                </p>
              }
            />
          </div>

          {/* Quick Stats */}
          <div className="flex items-center gap-3 self-start lg:self-center bg-slate-800/80 p-3 rounded-2xl border border-slate-700/80">
            <div className="px-3 py-1.5 bg-slate-900/60 rounded-xl text-center border border-slate-700">
              <span className="text-[10px] uppercase font-bold text-slate-400 block">Total</span>
              <span className="text-sm font-black text-white">{allSuppliers.length}</span>
            </div>
            <div className="px-3 py-1.5 bg-emerald-950/60 rounded-xl text-center border border-emerald-800/50">
              <span className="text-[10px] uppercase font-bold text-emerald-400 block">Classificados</span>
              <span className="text-sm font-black text-emerald-300">{classifiedCount}</span>
            </div>
            <div className="px-3 py-1.5 bg-amber-950/60 rounded-xl text-center border border-amber-800/50">
              <span className="text-[10px] uppercase font-bold text-amber-400 block">Pendentes</span>
              <span className="text-sm font-black text-amber-300">{unclassifiedCount}</span>
            </div>
          </div>
        </div>

        {/* Navigation Tabs Bar */}
        <div className="mt-6 pt-4 border-t border-slate-800 flex items-center gap-2 overflow-x-auto">
          <button
            onClick={() => setActiveSubTab('classificar')}
            className={`px-4 py-2.5 rounded-xl font-extrabold text-xs flex items-center gap-2 transition cursor-pointer shrink-0 ${
              activeSubTab === 'classificar'
                ? 'bg-emerald-500 text-slate-950 shadow-md shadow-emerald-500/20'
                : 'bg-slate-800/70 text-slate-300 hover:bg-slate-800 hover:text-white'
            }`}
          >
            <Building2 className="w-4 h-4" />
            <span>Classificar Fornecedores</span>
            {unclassifiedCount > 0 && (
              <span
                className={`ml-1 text-[10px] px-2 py-0.5 rounded-full font-black ${
                  activeSubTab === 'classificar'
                    ? 'bg-slate-950 text-emerald-400'
                    : 'bg-amber-500 text-slate-950'
                }`}
              >
                {unclassifiedCount} pendentes
              </span>
            )}
          </button>

          <button
            onClick={() => setActiveSubTab('grupos')}
            className={`px-4 py-2.5 rounded-xl font-extrabold text-xs flex items-center gap-2 transition cursor-pointer shrink-0 ${
              activeSubTab === 'grupos'
                ? 'bg-emerald-500 text-slate-950 shadow-md shadow-emerald-500/20'
                : 'bg-slate-800/70 text-slate-300 hover:bg-slate-800 hover:text-white'
            }`}
          >
            <Layers className="w-4 h-4" />
            <span>Painel por Categoria (Métricas)</span>
          </button>

          <button
            onClick={() => setActiveSubTab('gerenciar')}
            className={`px-4 py-2.5 rounded-xl font-extrabold text-xs flex items-center gap-2 transition cursor-pointer shrink-0 ${
              activeSubTab === 'gerenciar'
                ? 'bg-emerald-500 text-slate-950 shadow-md shadow-emerald-500/20'
                : 'bg-slate-800/70 text-slate-300 hover:bg-slate-800 hover:text-white'
            }`}
          >
            <SlidersHorizontal className="w-4 h-4" />
            <span>Gerenciar Categorias ({categories.length})</span>
          </button>
        </div>
      </div>

      {/* Info Alert Banner */}
      {infoAlert && (
        <div className="bg-amber-50 border border-amber-200 text-amber-900 px-4 py-2.5 rounded-xl text-xs flex items-center justify-between animate-in fade-in duration-150">
          <span>{infoAlert}</span>
          <button
            type="button"
            onClick={() => setInfoAlert(null)}
            className="text-amber-600 hover:text-amber-800 font-bold ml-2 cursor-pointer"
          >
            ✕
          </button>
        </div>
      )}

      {/* SUB-TAB 1: CLASSIFICAR FORNECEDORES */}
      {activeSubTab === 'classificar' && (
        <div className="space-y-4">
          {!metrics.isProcessed ? (
            <div className="bg-white p-12 text-center rounded-2xl border border-slate-200 space-y-3 shadow-xs">
              <FolderKanban className="w-12 h-12 text-slate-300 mx-auto" />
              <p className="text-sm font-bold text-slate-600">Nenhum dado importado para classificação.</p>
              <p className="text-xs text-slate-400 max-w-md mx-auto">
                Faça o upload de suas planilhas na aba de Importações para carregar a lista de credores.
              </p>
            </div>
          ) : (
            <>
              {/* Controls */}
              <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-xs space-y-4">
                <div className="flex flex-col md:flex-row items-stretch md:items-center justify-between gap-4">
                  <div className="relative flex-1 max-w-md">
                    <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                    <input
                      type="text"
                      placeholder="Buscar por razão social ou CNPJ..."
                      value={searchTerm}
                      onChange={(e) => setSearchTerm(e.target.value)}
                      className="w-full pl-9 pr-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-medium text-slate-800 focus:outline-hidden focus:border-emerald-500 focus:bg-white"
                    />
                  </div>

                  <div className="flex items-center gap-2 overflow-x-auto pb-1 md:pb-0">
                    <button
                      onClick={() => setCategoryFilter('all')}
                      className={`px-3 py-1.5 rounded-xl text-xs font-bold transition cursor-pointer border shrink-0 ${
                        categoryFilter === 'all'
                          ? 'bg-slate-900 text-white border-slate-900'
                          : 'bg-slate-50 text-slate-600 border-slate-200 hover:bg-slate-100'
                      }`}
                    >
                      Todos ({allSuppliers.length})
                    </button>

                    <button
                      onClick={() => setCategoryFilter('unassigned')}
                      className={`px-3 py-1.5 rounded-xl text-xs font-bold transition cursor-pointer border shrink-0 ${
                        categoryFilter === 'unassigned'
                          ? 'bg-amber-500 text-slate-950 border-amber-500 shadow-xs'
                          : 'bg-amber-50 text-amber-800 border-amber-200 hover:bg-amber-100'
                      }`}
                    >
                      Não Classificados ({unclassifiedCount})
                    </button>

                    <select
                      value={
                        categoryFilter !== 'all' && categoryFilter !== 'unassigned'
                          ? categoryFilter
                          : ''
                      }
                      onChange={(e) => setCategoryFilter(e.target.value || 'all')}
                      className="px-3 py-1.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-700 focus:outline-hidden focus:border-emerald-500 cursor-pointer shrink-0"
                    >
                      <option value="">Filtrar Categoria...</option>
                      {categories.map((c) => (
                        <option key={c.id} value={c.id}>
                          {c.name}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>

                {/* Bulk Bar */}
                {selectedSupplierKeys.length > 0 && (
                  <div className="bg-emerald-50 border border-emerald-300 p-3.5 rounded-xl flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                    <div className="flex items-center gap-2">
                      <span className="w-6 h-6 rounded-full bg-emerald-600 text-white flex items-center justify-center text-xs font-bold">
                        {selectedSupplierKeys.length}
                      </span>
                      <span className="text-xs font-extrabold text-emerald-950">
                        fornecedor(es) selecionado(s)
                      </span>
                    </div>

                    <div className="flex items-center gap-2">
                      <select
                        value={bulkTargetCategory}
                        onChange={(e) => setBulkTargetCategory(e.target.value)}
                        className="px-3 py-1.5 bg-white border border-emerald-400 rounded-xl text-xs font-bold text-slate-800 cursor-pointer"
                      >
                        <option value="">Escolha a Categoria...</option>
                        {categories.map((c) => (
                          <option key={c.id} value={c.id}>
                            {c.name}
                          </option>
                        ))}
                      </select>

                      <button
                        onClick={handleApplyBulkAssignment}
                        disabled={!bulkTargetCategory}
                        className="px-4 py-1.5 bg-emerald-600 hover:bg-emerald-500 disabled:opacity-50 text-white font-bold text-xs rounded-xl transition cursor-pointer flex items-center gap-1.5"
                      >
                        <Check className="w-3.5 h-3.5" />
                        <span>Aplicar</span>
                      </button>

                      <button
                        onClick={() => setSelectedSupplierKeys([])}
                        className="px-3 py-1.5 text-slate-500 hover:text-slate-800 font-bold text-xs"
                      >
                        Limpar
                      </button>
                    </div>
                  </div>
                )}
              </div>

              {/* Table */}
              <div className="bg-white rounded-2xl border border-slate-200 shadow-xs overflow-hidden">
                <div className="overflow-x-auto">
                  <table className="w-full text-left text-xs border-collapse">
                    <thead className="bg-slate-100/80 text-slate-700 font-bold uppercase tracking-wider text-[10px] border-b border-slate-200">
                      <tr>
                        <th className="py-3 px-3.5 border-r border-slate-200 text-center w-10">
                          <input
                            type="checkbox"
                            checked={
                              filteredSuppliers.length > 0 &&
                              filteredSuppliers.every((s) =>
                                selectedSupplierKeys.includes(s.cnpj && s.cnpj !== 'N/I' ? s.cnpj : s.name)
                              )
                            }
                            onChange={toggleSelectAllFiltered}
                            className="rounded border-slate-300 text-emerald-600 focus:ring-emerald-500 cursor-pointer"
                          />
                        </th>
                        <th className="py-3 px-3.5 border-r border-slate-200">Favorecido / Razão Social</th>
                        <th className="py-3 px-3.5 border-r border-slate-200">CNPJ</th>
                        <th className="py-3 px-3.5 border-r border-slate-200 text-right">Valor Total (R$)</th>
                        <th className="py-3 px-4 text-center font-bold">Categoria Definida</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 font-mono">
                      {filteredSuppliers.length === 0 ? (
                        <tr>
                          <td colSpan={5} className="py-12 text-center text-slate-400 font-sans">
                            Nenhum fornecedor localizado com os filtros aplicados.
                          </td>
                        </tr>
                      ) : (
                        filteredSuppliers.map((sup, idx) => {
                          const key = sup.cnpj && sup.cnpj !== 'N/I' ? sup.cnpj : sup.name;
                          const isSelected = selectedSupplierKeys.includes(key);
                          const currentCatId = overrides[key] || 'cat_outros';
                          const isUnclassified = currentCatId === 'cat_outros';
                          const catObj = categories.find((c) => c.id === currentCatId) || categories.find((c) => c.id === 'cat_outros');
                          const theme = COLOR_MAP[catObj?.color || 'slate'] || COLOR_MAP.slate;

                          return (
                            <tr
                              key={`sup_cat_${sup.cnpj || 'NI'}_${sup.name}_${idx}`}
                              className={`transition border-b border-slate-100 ${
                                isSelected
                                  ? 'bg-emerald-50/80'
                                  : 'hover:bg-slate-50/80'
                              } ${!isUnclassified ? `border-l-4 ${theme.borderLeft}` : 'border-l-4 border-l-transparent'}`}
                            >
                              <td className="py-2.5 px-3.5 text-center border-r border-slate-100">
                                <input
                                  type="checkbox"
                                  checked={isSelected}
                                  onChange={() => toggleSelectSupplier(key)}
                                  className="rounded border-slate-300 text-emerald-600 focus:ring-emerald-500 cursor-pointer"
                                />
                              </td>
                              <td className="py-2.5 px-3.5 font-sans font-extrabold text-slate-900 border-r border-slate-100 max-w-xs truncate">
                                <div className="flex items-center gap-2">
                                  {!isUnclassified && (
                                    <span
                                      className={`w-2.5 h-2.5 rounded-full shrink-0 shadow-2xs ${theme.badge}`}
                                      title={`Categoria: ${catObj?.name}`}
                                    />
                                  )}
                                  <span className="truncate">{sup.name}</span>
                                </div>
                              </td>
                              <td className="py-2.5 px-3.5 text-slate-500 border-r border-slate-100">
                                {sup.cnpj || 'N/I'}
                              </td>
                              <td className="py-2.5 px-3.5 text-right font-black text-slate-900 border-r border-slate-100">
                                {formatBRL(sup.totalValue)}
                              </td>
                              <td className="py-2.5 px-4 text-center font-sans">
                                <div className="flex items-center justify-center">
                                  <select
                                    value={currentCatId}
                                    onChange={(e) =>
                                      handleSupplierCategoryChange(
                                        sup.cnpj,
                                        sup.name,
                                        e.target.value
                                      )
                                    }
                                    className={`w-full max-w-xs text-xs font-extrabold rounded-xl px-3 py-1.5 border transition cursor-pointer shadow-2xs ${
                                      isUnclassified
                                        ? 'bg-amber-50 text-amber-900 border-amber-300 focus:border-amber-500'
                                        : `${theme.badgeText} focus:ring-2`
                                    }`}
                                  >
                                    {categories.map((c) => (
                                      <option key={c.id} value={c.id}>
                                        {c.id === 'cat_outros' ? '⚠️ Não Classificado' : `📁 ${c.name}`}
                                      </option>
                                    ))}
                                  </select>
                                </div>
                              </td>
                            </tr>
                          );
                        })
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            </>
          )}
        </div>
      )}

      {/* SUB-TAB 2: PAINEL POR GRUPOS */}
      {activeSubTab === 'grupos' && (
        <div className="space-y-4">
          {!metrics.isProcessed ? (
            <div className="bg-white p-12 text-center rounded-2xl border border-slate-200 space-y-3 shadow-xs">
              <FolderKanban className="w-12 h-12 text-slate-300 mx-auto" />
              <p className="text-sm font-bold text-slate-600">Nenhum dado importado para o painel.</p>
            </div>
          ) : (
            <div className="space-y-4">
              {categoryMetrics.map((catMetric) => {
                const { category, totalValue, totalCount, suppliers } = catMetric;
                const theme = COLOR_MAP[category.color] || COLOR_MAP.slate;
                const isExpanded = expandedCategoryId === category.id;
                const percentOfTotal = grandTotal > 0 ? (totalValue / grandTotal) * 100 : 0;

                return (
                  <div
                    key={category.id}
                    className={`bg-white rounded-2xl border ${theme.border} shadow-xs hover:shadow-md transition-all overflow-hidden`}
                  >
                    <div className={`p-5 ${theme.bgLight} border-b ${theme.border}`}>
                      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                        <div className="flex items-start gap-3.5 min-w-0">
                          <div
                            className={`w-10 h-10 rounded-xl ${theme.badge} text-white flex items-center justify-center shrink-0 shadow-xs`}
                          >
                            <Tag className="w-5 h-5" />
                          </div>

                          <div className="min-w-0">
                            <div className="flex items-center gap-2 flex-wrap">
                              <h3 className={`text-base font-extrabold ${theme.text} tracking-tight`}>
                                {category.name}
                              </h3>
                              <span
                                className={`text-[10px] font-black px-2.5 py-0.5 rounded-full border ${theme.badgeText}`}
                              >
                                {totalCount} {totalCount === 1 ? 'credor' : 'credores'}
                              </span>
                            </div>

                            {category.description && (
                              <p className="text-xs text-slate-500 mt-0.5 line-clamp-1 font-medium">
                                {category.description}
                              </p>
                            )}
                          </div>
                        </div>

                        <div className="text-left sm:text-right shrink-0">
                          <span className="text-[10px] uppercase font-extrabold text-slate-400 tracking-wider block">
                            Valor Total
                          </span>
                          <span className="text-lg font-black text-slate-900 font-mono tracking-tight block">
                            {formatBRL(totalValue)}
                          </span>
                          <span className="text-[10px] font-bold text-slate-500">
                            {percentOfTotal.toFixed(1)}% do total
                          </span>
                        </div>
                      </div>

                      <div className="mt-4 pt-3 border-t border-slate-200/60 flex items-center justify-between gap-4 text-xs">
                        <div className="flex-1 bg-slate-200 h-2 rounded-full overflow-hidden">
                          <div
                            className={`h-full ${theme.badge}`}
                            style={{ width: `${percentOfTotal}%` }}
                          />
                        </div>

                        <button
                          onClick={() => setExpandedCategoryId(isExpanded ? null : category.id)}
                          className={`text-xs font-bold ${theme.accent} hover:underline flex items-center gap-1 cursor-pointer shrink-0`}
                        >
                          <span>{isExpanded ? 'Ocultar Lista' : `Ver ${suppliers.length} Credores`}</span>
                          {isExpanded ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
                        </button>
                      </div>
                    </div>

                    {isExpanded && (
                      <div className="p-4 bg-white">
                        {suppliers.length === 0 ? (
                          <p className="text-xs text-slate-400 py-4 text-center">
                            Nenhum fornecedor nesta categoria.
                          </p>
                        ) : (
                          <div className="overflow-x-auto rounded-xl border border-slate-200">
                            <table className="w-full text-left text-xs border-collapse">
                              <thead className="bg-slate-100 text-slate-700 font-bold uppercase tracking-wider text-[10px] border-b border-slate-200">
                                <tr>
                                  <th className="py-2.5 px-3 border-r border-slate-200 w-10 text-center">#</th>
                                  <th className="py-2.5 px-3 border-r border-slate-200">Favorecido</th>
                                  <th className="py-2.5 px-3 border-r border-slate-200">CNPJ</th>
                                  <th className="py-2.5 px-3 text-right">Valor Total</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-slate-100 font-mono">
                                {suppliers.map((sup, idx) => (
                                  <tr
                                    key={`cat_sup_row_${sup.cnpj || 'NI'}_${sup.name}_${idx}`}
                                    className={`hover:bg-slate-50 transition border-l-4 ${theme.borderLeft}`}
                                  >
                                    <td className="py-2 px-3 text-slate-400 text-center border-r border-slate-100">
                                      {idx + 1}
                                    </td>
                                    <td className="py-2 px-3 font-sans font-bold text-slate-800 border-r border-slate-100">
                                      {sup.name}
                                    </td>
                                    <td className="py-2 px-3 text-slate-500 border-r border-slate-100">
                                      {sup.cnpj || 'N/I'}
                                    </td>
                                    <td className="py-2 px-3 text-right font-black text-slate-900">
                                      {formatBRL(sup.totalValue)}
                                    </td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                        )}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          )}
        </div>
      )}

      {/* SUB-TAB 3: GERENCIAR CATEGORIAS */}
      {activeSubTab === 'gerenciar' && (
        <div className="space-y-6">
          <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-xs space-y-6">
            <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4 border-b border-slate-100 pb-4">
              <div>
                <h3 className="text-base font-extrabold text-slate-900">
                  Cadastro e Edição de Categorias
                </h3>
                <p className="text-xs text-slate-500">
                  Adicione, edite ou remova grupos de classificação.
                </p>
              </div>

              <div className="flex items-center gap-2">
                {confirmResetCategories ? (
                  <div className="flex items-center gap-1.5 bg-rose-50 border border-rose-300 rounded-xl px-2.5 py-1.5 animate-in fade-in duration-100">
                    <span className="text-[11px] font-bold text-rose-800">Restaurar categorias padrão?</span>
                    <button
                      type="button"
                      onClick={executeResetDefaults}
                      className="px-2 py-0.5 bg-rose-600 hover:bg-rose-700 text-white rounded text-[10px] font-bold cursor-pointer"
                    >
                      Sim
                    </button>
                    <button
                      type="button"
                      onClick={() => setConfirmResetCategories(false)}
                      className="px-2 py-0.5 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded text-[10px] font-bold cursor-pointer"
                    >
                      Não
                    </button>
                  </div>
                ) : (
                  <button
                    type="button"
                    onClick={() => setConfirmResetCategories(true)}
                    className="px-3.5 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs rounded-xl transition flex items-center gap-1.5 cursor-pointer"
                  >
                    <RotateCcw className="w-3.5 h-3.5" />
                    <span>Restaurar Padrão</span>
                  </button>
                )}

                <button
                  onClick={handleOpenAddCategory}
                  className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-bold text-xs rounded-xl shadow-xs transition flex items-center gap-1.5 cursor-pointer"
                >
                  <Plus className="w-4 h-4" />
                  <span>Nova Categoria</span>
                </button>
              </div>
            </div>

            {/* Form */}
            {(isAddMode || editingCategory) && (
              <form
                onSubmit={handleSaveCategoryForm}
                className="bg-emerald-50/70 border border-emerald-200 p-5 rounded-2xl space-y-4"
              >
                <div className="flex items-center justify-between border-b border-emerald-200 pb-2">
                  <h4 className="text-xs font-extrabold text-emerald-900 uppercase tracking-wider">
                    {isAddMode ? 'Adicionar Categoria' : `Editar Categoria: ${editingCategory?.name}`}
                  </h4>
                  <button
                    type="button"
                    onClick={() => {
                      setIsAddMode(false);
                      setEditingCategory(null);
                    }}
                    className="text-xs text-slate-400 hover:text-slate-600 font-bold cursor-pointer"
                  >
                    Cancelar
                  </button>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">
                      Nome da Categoria *
                    </label>
                    <input
                      type="text"
                      required
                      placeholder="Ex: Material Cirúrgico e Descartáveis"
                      value={formName}
                      onChange={(e) => setFormName(e.target.value)}
                      className="w-full px-3 py-2 bg-white border border-slate-300 rounded-xl text-xs font-semibold text-slate-800 focus:outline-hidden focus:border-emerald-500"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Cor do Grupo</label>
                    <select
                      value={formColor}
                      onChange={(e) => setFormColor(e.target.value)}
                      className="w-full px-3 py-2 bg-white border border-slate-300 rounded-xl text-xs font-semibold text-slate-800 focus:outline-hidden cursor-pointer"
                    >
                      {COLOR_OPTIONS.map((opt) => (
                        <option key={opt.id} value={opt.id}>
                          {opt.name}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Descrição</label>
                  <input
                    type="text"
                    placeholder="Descrição para identificação do grupo..."
                    value={formDesc}
                    onChange={(e) => setFormDesc(e.target.value)}
                    className="w-full px-3 py-2 bg-white border border-slate-300 rounded-xl text-xs font-semibold text-slate-800 focus:outline-hidden"
                  />
                </div>

                <div className="flex justify-end pt-2">
                  <button
                    type="submit"
                    className="px-6 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-bold text-xs rounded-xl shadow-xs transition cursor-pointer"
                  >
                    Salvar Categoria
                  </button>
                </div>
              </form>
            )}

            {/* List */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              {categories.map((cat) => {
                const theme = COLOR_MAP[cat.color] || COLOR_MAP.slate;

                return (
                  <div
                    key={cat.id}
                    className="bg-slate-50 rounded-2xl p-4 border border-slate-200 flex items-start justify-between gap-3"
                  >
                    <div className="flex items-start gap-3 min-w-0">
                      <div
                        className={`w-9 h-9 rounded-xl ${theme.badge} text-white flex items-center justify-center shrink-0 shadow-xs`}
                      >
                        <Tag className="w-4 h-4" />
                      </div>
                      <div className="min-w-0">
                        <div className="flex items-center gap-2">
                          <h4 className="text-xs font-extrabold text-slate-900 truncate">{cat.name}</h4>
                          {cat.id === 'cat_outros' && (
                            <span className="text-[9px] bg-slate-200 text-slate-600 px-1.5 py-0.5 rounded font-bold">
                              Padrão
                            </span>
                          )}
                        </div>
                        <p className="text-[11px] text-slate-500 line-clamp-2 mt-0.5">{cat.description}</p>
                      </div>
                    </div>

                    <div className="flex items-center gap-1 shrink-0">
                      <button
                        onClick={() => handleOpenEditCategory(cat)}
                        className="p-1.5 text-slate-400 hover:text-blue-600 hover:bg-blue-50 rounded-lg transition cursor-pointer"
                        title="Editar Categoria"
                      >
                        <Edit2 className="w-3.5 h-3.5" />
                      </button>

                      {cat.id !== 'cat_outros' && (
                        deletingCatId === cat.id ? (
                          <div className="flex items-center gap-1 bg-rose-50 border border-rose-300 rounded-lg px-2 py-0.5 animate-in fade-in duration-100">
                            <span className="text-[10px] text-rose-800 font-bold">Excluir?</span>
                            <button
                              type="button"
                              onClick={() => executeDeleteCategory(cat.id)}
                              className="px-1.5 py-0.2 bg-rose-600 hover:bg-rose-700 text-white rounded text-[9px] font-bold cursor-pointer"
                            >
                              Sim
                            </button>
                            <button
                              type="button"
                              onClick={() => setDeletingCatId(null)}
                              className="px-1.5 py-0.2 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded text-[9px] font-bold cursor-pointer"
                            >
                              Não
                            </button>
                          </div>
                        ) : (
                          <button
                            type="button"
                            onClick={() => setDeletingCatId(cat.id)}
                            className="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition cursor-pointer"
                            title="Excluir Categoria"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        )
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
