import React, { useState, useEffect, useMemo, useDeferredValue } from 'react';
import {
  Calendar,
  CheckCircle2,
  Clock,
  Search,
  Download,
  Building,
  Repeat,
  ExternalLink,
  ChevronLeft,
  ChevronRight,
  TrendingUp,
  FileCheck,
  Edit3,
  Check,
  X,
  ArrowRight,
  Undo2,
  Loader2,
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { DashboardMetrics, FileData, ConsolidatedSupplier, CategoryDefinition } from '../../types';
import { formatBRL } from '../../utils/excelParser';
import { getManualRecurringSuppliers } from '../../utils/recorrenciaService';
import {
  getCurrentMonthKey,
  formatMonthFull,
  formatMonthShort,
  getAvailableMonthKeys,
  getAllAtendimentos,
  getSupplierStorageKey,
  toggleSupplierAtendimento,
  updateSupplierAtendimentoObs,
  subscribeToAtendimentos,
  AtendimentoMensalRecord,
} from '../../utils/atendimentoMensalService';
import {
  getCategories,
  getSupplierOverrides,
  subscribeToCategories,
  categorizeSupplier,
  COLOR_MAP,
} from '../../utils/categoryService';
import { InfoHelpButton } from '../InfoHelpButton';

interface FornecedoresMesViewProps {
  metrics: DashboardMetrics;
  files: FileData[];
  onOpenCompanyAudit?: (supplier: { name: string; cnpj: string }) => void;
  onNavigateToTab?: (tab: string) => void;
}

export const FornecedoresMesView: React.FC<FornecedoresMesViewProps> = ({
  metrics,
  files: _files,
  onOpenCompanyAudit,
  onNavigateToTab,
}) => {
  // Current month vs selected month
  const currentMonthKey = useMemo(() => getCurrentMonthKey(), []);
  const [selectedMonthKey, setSelectedMonthKey] = useState<string>(currentMonthKey);
  const availableMonths = useMemo(() => getAvailableMonthKeys(14), []);

  // Filter & Search states
  const [searchTerm, setSearchTerm] = useState('');
  const deferredSearchTerm = useDeferredValue(searchTerm);
  const [filterType, setFilterType] = useState<'all' | 'recorrentes' | 'com_obs' | 'sem_obs'>('all');
  const [editingObsKey, setEditingObsKey] = useState<string | null>(null);
  const [obsText, setObsText] = useState('');

  // Pagination for fluid table rendering
  const [currentPage, setCurrentPage] = useState(1);
  const pageSize = 50;

  // Live attendance data
  const [allRecords, setAllRecords] = useState<Record<string, AtendimentoMensalRecord>>(() => getAllAtendimentos());

  // Recurring suppliers set
  const [recurringKeysSet, setRecurringKeysSet] = useState<Set<string>>(new Set());
  const [categoryOverrides, setCategoryOverrides] = useState<Record<string, string>>(() => getSupplierOverrides());
  const [allCategories, setAllCategories] = useState<CategoryDefinition[]>(() => getCategories());

  useEffect(() => {
    // Load recurring set
    const recList = getManualRecurringSuppliers();
    const set = new Set<string>();
    recList.forEach((r) => {
      if (r.cnpj && r.cnpj !== 'N/I') set.add(r.cnpj.replace(/\D/g, ''));
      if (r.name) set.add(r.name.trim().toUpperCase());
    });
    setRecurringKeysSet(set);

    // Subscribe to live monthly attendance changes
    const unsubscribeAtendimentos = subscribeToAtendimentos(() => {
      setAllRecords({ ...getAllAtendimentos() });
    });

    const unsubscribeCategories = subscribeToCategories(() => {
      setCategoryOverrides(getSupplierOverrides());
      setAllCategories(getCategories());
    });

    return () => {
      unsubscribeAtendimentos();
      unsubscribeCategories();
    };
  }, []);

  const isSupplierRecurring = (name: string, cnpj?: string) => {
    const cnpjDigits = (cnpj || '').replace(/\D/g, '');
    if (cnpjDigits.length === 14 && recurringKeysSet.has(cnpjDigits)) return true;
    if (name && recurringKeysSet.has(name.trim().toUpperCase())) return true;
    return false;
  };

  // Build the list of ONLY suppliers marked as attended in selectedMonthKey
  const attendedSuppliers = useMemo(() => {
    // Map of currently consolidated suppliers for quick metric lookup
    const metricsMap = new Map<string, ConsolidatedSupplier>();
    (metrics.consolidatedSuppliers || []).forEach((sup) => {
      const key = getSupplierStorageKey(sup.name, sup.cnpj);
      metricsMap.set(key, sup);
    });

    const attendedList: {
      name: string;
      cnpj: string;
      totalValue: number;
      ppProntasValue: number;
      liqObedeceValue: number;
      liqNaoObedeceValue: number;
      enviadasBancoValue: number;
      totalCount: number;
      isRec: boolean;
      recordId: string;
      record: AtendimentoMensalRecord;
      observacao: string;
      dataAtendimento: string | null;
    }[] = [];

    const seenKeys = new Set<string>();

    // 1. Iterate over attendance records stored for selectedMonthKey
    Object.values(allRecords).forEach((rec) => {
      if (rec.monthKey === selectedMonthKey && rec.atendido) {
        const supKey = getSupplierStorageKey(rec.supplierName, rec.cnpj);
        seenKeys.add(supKey);
        const metricData = metricsMap.get(supKey);
        const isRec = rec.isRecurring || (metricData ? isSupplierRecurring(metricData.name, metricData.cnpj) : isSupplierRecurring(rec.supplierName, rec.cnpj));

        attendedList.push({
          name: metricData?.name || rec.supplierName,
          cnpj: metricData?.cnpj || rec.cnpj || 'N/I',
          totalValue: metricData?.totalValue || 0,
          ppProntasValue: metricData?.ppProntasValue || 0,
          liqObedeceValue: metricData?.liqObedeceValue || 0,
          liqNaoObedeceValue: metricData?.liqNaoObedeceValue || 0,
          enviadasBancoValue: metricData?.enviadasBancoValue || 0,
          totalCount: metricData?.totalCount || 0,
          isRec,
          recordId: rec.id,
          record: rec,
          observacao: rec.observacao || '',
          dataAtendimento: rec.dataAtendimento || null,
        });
      }
    });

    // 2. Also check consolidated suppliers from active metrics if they have a matching attended record
    (metrics.consolidatedSuppliers || []).forEach((sup) => {
      const supKey = getSupplierStorageKey(sup.name, sup.cnpj);
      if (!seenKeys.has(supKey)) {
        const recordId = `${selectedMonthKey}_${supKey}`;
        const rec = allRecords[recordId];
        if (rec?.atendido) {
          seenKeys.add(supKey);
          attendedList.push({
            name: sup.name,
            cnpj: sup.cnpj || 'N/I',
            totalValue: sup.totalValue || 0,
            ppProntasValue: sup.ppProntasValue || 0,
            liqObedeceValue: sup.liqObedeceValue || 0,
            liqNaoObedeceValue: sup.liqNaoObedeceValue || 0,
            enviadasBancoValue: sup.enviadasBancoValue || 0,
            totalCount: sup.totalCount || 0,
            isRec: isSupplierRecurring(sup.name, sup.cnpj),
            recordId,
            record: rec,
            observacao: rec.observacao || '',
            dataAtendimento: rec.dataAtendimento || null,
          });
        }
      }
    });

    // Sort: highest value first, or newest attendance first
    return attendedList.sort((a, b) => b.totalValue - a.totalValue);
  }, [metrics.consolidatedSuppliers, allRecords, selectedMonthKey, recurringKeysSet]);

  // Filtered attended suppliers by search and type
  const filteredSuppliers = useMemo(() => {
    return attendedSuppliers.filter((sup) => {
      // Search term
      if (deferredSearchTerm.trim()) {
        const query = deferredSearchTerm.toLowerCase().trim();
        const matchesName = sup.name.toLowerCase().includes(query);
        const matchesCnpj = (sup.cnpj || '').includes(query);
        const matchesObs = (sup.observacao || '').toLowerCase().includes(query);
        if (!matchesName && !matchesCnpj && !matchesObs) return false;
      }

      // Filter classification
      if (filterType === 'recorrentes' && !sup.isRec) return false;
      if (filterType === 'com_obs' && !sup.observacao.trim()) return false;
      if (filterType === 'sem_obs' && sup.observacao.trim()) return false;

      return true;
    });
  }, [attendedSuppliers, deferredSearchTerm, filterType]);

  // Reset page on filter or month change
  useEffect(() => {
    setCurrentPage(1);
  }, [deferredSearchTerm, filterType, selectedMonthKey]);

  // Paginated slice
  const totalPages = Math.ceil(filteredSuppliers.length / pageSize) || 1;
  const paginatedSuppliers = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return filteredSuppliers.slice(start, start + pageSize);
  }, [filteredSuppliers, currentPage, pageSize]);

  // Statistics for selected month (exclusively for attended suppliers)
  const stats = useMemo(() => {
    const totalAtendidos = attendedSuppliers.length;
    const recorrentesAtendidos = attendedSuppliers.filter((s) => s.isRec).length;
    const comObservacao = attendedSuppliers.filter((s) => s.observacao.trim()).length;
    const valorTotalAtendido = attendedSuppliers.reduce((acc, s) => acc + s.totalValue, 0);

    return {
      totalAtendidos,
      recorrentesAtendidos,
      comObservacao,
      valorTotalAtendido,
    };
  }, [attendedSuppliers]);

  // Navigation handlers for months
  const handlePrevMonth = () => {
    const idx = availableMonths.findIndex((m) => m.key === selectedMonthKey);
    if (idx > 0) {
      setSelectedMonthKey(availableMonths[idx - 1].key);
    }
  };

  const handleNextMonth = () => {
    const idx = availableMonths.findIndex((m) => m.key === selectedMonthKey);
    if (idx !== -1 && idx < availableMonths.length - 1) {
      setSelectedMonthKey(availableMonths[idx + 1].key);
    }
  };

  // Toggle/unmark attendance
  const handleUnmarkAttendance = (supplier: typeof attendedSuppliers[0]) => {
    toggleSupplierAtendimento(
      supplier.name,
      supplier.cnpj,
      supplier.isRec,
      selectedMonthKey
    );
    setAllRecords({ ...getAllAtendimentos() });
  };

  // Observations inline editing
  const handleStartEditObs = (supplier: typeof attendedSuppliers[0]) => {
    setEditingObsKey(supplier.recordId);
    setObsText(supplier.observacao || '');
  };

  const handleSaveObs = (supplier: typeof attendedSuppliers[0]) => {
    updateSupplierAtendimentoObs(
      supplier.name,
      supplier.cnpj,
      selectedMonthKey,
      obsText.trim()
    );
    setEditingObsKey(null);
    setAllRecords({ ...getAllAtendimentos() });
  };

  const [isExporting, setIsExporting] = useState(false);

  // Export to Excel
  const handleExportExcel = () => {
    if (filteredSuppliers.length === 0 || isExporting) return;

    setIsExporting(true);
    setTimeout(() => {
      try {
        const rows = filteredSuppliers.map((s, idx) => ({
          '#': idx + 1,
          'Mês de Referência': formatMonthFull(selectedMonthKey),
          'Favorecido / Fornecedor': s.name,
          'CNPJ': s.cnpj || 'N/I',
          'Classificação': s.isRec ? 'Recorrente' : 'Comum',
          'Data de Atendimento': s.dataAtendimento ? new Date(s.dataAtendimento).toLocaleDateString('pt-BR') : '-',
          'Valor Total no Painel (R$)': s.totalValue,
          'Anotação / PP Emitida': s.observacao || '',
        }));

        const worksheet = XLSX.utils.json_to_sheet(rows);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, `Atendidos_${selectedMonthKey}`);
        XLSX.writeFile(workbook, `Fornecedores_Atendidos_${selectedMonthKey}.xlsx`);
      } finally {
        setIsExporting(false);
      }
    }, 200);
  };

  return (
    <div className="space-y-6 pb-16">
      {/* Month Navigator Header */}
      <div className="bg-white rounded-2xl p-5 md:p-6 shadow-xs border border-slate-200/90 flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <div className="flex items-center space-x-2.5">
            <div className="p-2 bg-emerald-50 rounded-xl border border-emerald-200">
              <CheckCircle2 className="w-5 h-5 text-emerald-700" />
            </div>
            <div className="flex items-center gap-2.5">
              <h2 className="text-xl font-black text-slate-900 tracking-tight">
                Fornecedores Atendidos no Mês
              </h2>
            </div>
          </div>
        </div>

        {/* Month Selector Controls */}
        <div className="flex items-center gap-1 bg-slate-50 p-1 rounded-xl border border-slate-200">
          <button
            onClick={handlePrevMonth}
            className="p-1.5 rounded-lg hover:bg-white text-slate-600 hover:text-slate-900 transition cursor-pointer"
            title="Mês anterior"
          >
            <ChevronLeft className="w-4 h-4" />
          </button>

          <select
            value={selectedMonthKey}
            onChange={(e) => setSelectedMonthKey(e.target.value)}
            className="bg-transparent font-bold text-xs text-slate-800 py-1 px-2 focus:outline-hidden cursor-pointer"
          >
            {availableMonths.map((m) => (
              <option key={m.key} value={m.key}>
                {m.label} {m.isCurrent ? '• Mês Atual' : ''}
              </option>
            ))}
          </select>

          <button
            onClick={handleNextMonth}
            className="p-1.5 rounded-lg hover:bg-white text-slate-600 hover:text-slate-900 transition cursor-pointer"
            title="Próximo mês"
          >
            <ChevronRight className="w-4 h-4" />
          </button>

          {selectedMonthKey !== currentMonthKey && (
            <button
              onClick={() => setSelectedMonthKey(currentMonthKey)}
              className="ml-1 text-[11px] font-bold text-emerald-700 hover:text-emerald-800 bg-emerald-100 hover:bg-emerald-200 px-2 py-0.5 rounded-lg transition cursor-pointer"
              title="Voltar para o mês corrente"
            >
              Mês Atual
            </button>
          )}
        </div>
      </div>

      {/* KPI Cards for the selected month */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        {/* Atendidos no Mês */}
        <div className="bg-white rounded-2xl p-5 border border-slate-200/90 shadow-xs">
          <div className="flex items-center justify-between text-xs text-slate-500 font-semibold mb-2">
            <span>Fornecedores Atendidos</span>
            <CheckCircle2 className="w-4 h-4 text-emerald-600" />
          </div>
          <div className="text-2xl font-black text-emerald-800 font-mono">
            {stats.totalAtendidos}
          </div>
        </div>

        {/* Recorrentes Atendidos */}
        <div className="bg-white rounded-2xl p-5 border border-slate-200/90 shadow-xs">
          <div className="flex items-center justify-between text-xs text-slate-500 font-semibold mb-2">
            <span>Recorrentes Atendidos</span>
            <Repeat className="w-4 h-4 text-blue-600" />
          </div>
          <div className="text-2xl font-black text-blue-800 font-mono">
            {stats.recorrentesAtendidos}
          </div>
        </div>

        {/* Valor Total Atendido */}
        <div className="bg-white rounded-2xl p-5 border border-slate-200/90 shadow-xs">
          <div className="flex items-center justify-between text-xs text-slate-500 font-semibold mb-2">
            <span>Valor Total Atendido</span>
            <TrendingUp className="w-4 h-4 text-emerald-600" />
          </div>
          <div className="text-xl font-black text-slate-900 font-mono truncate" title={formatBRL(stats.valorTotalAtendido)}>
            {formatBRL(stats.valorTotalAtendido)}
          </div>
        </div>

        {/* Anotações Registradas */}
        <div className="bg-white rounded-2xl p-5 border border-slate-200/90 shadow-xs">
          <div className="flex items-center justify-between text-xs text-slate-500 font-semibold mb-2">
            <span>Com Anotação / PP</span>
            <FileCheck className="w-4 h-4 text-emerald-600" />
          </div>
          <div className="text-2xl font-black text-slate-800 font-mono">
            {stats.comObservacao}
          </div>
        </div>
      </div>

      {/* Main Content: Table or Empty State */}
      {attendedSuppliers.length === 0 ? (
        <div className="bg-white rounded-2xl p-8 md:p-12 shadow-xs border border-slate-200/90 text-center space-y-4 max-w-2xl mx-auto my-6">
          <div className="w-14 h-14 bg-emerald-50 text-emerald-600 rounded-2xl flex items-center justify-center mx-auto border border-emerald-100">
            <CheckCircle2 className="w-7 h-7" />
          </div>
          <div className="space-y-1.5">
            <h3 className="text-lg font-black text-slate-900">
              Nenhum fornecedor marcado como atendido em {formatMonthFull(selectedMonthKey)}
            </h3>
            <p className="text-xs text-slate-500 max-w-md mx-auto leading-relaxed">
              Nesta aba são exibidos <strong>apenas</strong> os credores que você marcar como atendidos.
              Para marcar um fornecedor, vá até a <strong>Página Inicial (Painel)</strong> e clique no botão <strong>Atendido</strong> na Tabela Principal de Credores.
            </p>
          </div>
          {onNavigateToTab && (
            <div className="pt-2">
              <button
                type="button"
                onClick={() => onNavigateToTab('painel')}
                className="inline-flex items-center gap-2 px-5 py-2.5 bg-emerald-700 hover:bg-emerald-800 text-white font-bold text-xs rounded-xl shadow-xs transition cursor-pointer active:scale-98"
              >
                <span>Ir para a Página Inicial</span>
                <ArrowRight className="w-4 h-4" />
              </button>
            </div>
          )}
        </div>
      ) : (
        <div className="bg-white rounded-2xl shadow-xs border border-slate-200/90 overflow-hidden">
          {/* Table Filters & Actions Header */}
          <div className="p-5 border-b border-slate-100 flex flex-col md:flex-row md:items-center justify-between gap-4">
            {/* Search */}
            <div className="relative flex-1 max-w-md">
              <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
              <input
                type="text"
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                placeholder="Buscar por favorecido, CNPJ ou anotação..."
                className="w-full pl-9 pr-3 py-2 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:outline-hidden focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500"
              />
            </div>

            {/* Filter Pills & Export */}
            <div className="flex flex-wrap items-center gap-2">
              <div className="flex items-center bg-slate-100 p-1 rounded-xl">
                <button
                  onClick={() => setFilterType('all')}
                  className={`px-2.5 py-1 text-xs font-bold rounded-lg transition cursor-pointer ${
                    filterType === 'all'
                      ? 'bg-white text-slate-900 shadow-xs'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  Todos Atendidos ({attendedSuppliers.length})
                </button>
                <button
                  onClick={() => setFilterType('recorrentes')}
                  className={`px-2.5 py-1 text-xs font-bold rounded-lg transition cursor-pointer ${
                    filterType === 'recorrentes'
                      ? 'bg-white text-blue-800 shadow-xs'
                      : 'text-slate-600 hover:text-blue-700'
                  }`}
                >
                  Recorrentes ({stats.recorrentesAtendidos})
                </button>
                <button
                  onClick={() => setFilterType('com_obs')}
                  className={`px-2.5 py-1 text-xs font-bold rounded-lg transition cursor-pointer ${
                    filterType === 'com_obs'
                      ? 'bg-white text-emerald-800 shadow-xs'
                      : 'text-slate-600 hover:text-emerald-700'
                  }`}
                >
                  Com Anotação ({stats.comObservacao})
                </button>
                <button
                  onClick={() => setFilterType('sem_obs')}
                  className={`px-2.5 py-1 text-xs font-bold rounded-lg transition cursor-pointer ${
                    filterType === 'sem_obs'
                      ? 'bg-white text-slate-800 shadow-xs'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  Sem Anotação ({attendedSuppliers.length - stats.comObservacao})
                </button>
              </div>

              <button
                onClick={handleExportExcel}
                disabled={isExporting}
                className="px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs rounded-xl transition flex items-center gap-1.5 cursor-pointer shadow-2xs active:scale-98 disabled:opacity-75 disabled:cursor-wait"
                title="Exportar fornecedores atendidos para Excel"
              >
                {isExporting ? (
                  <>
                    <Loader2 className="w-3.5 h-3.5 animate-spin text-emerald-600" />
                    <span>Exportando...</span>
                  </>
                ) : (
                  <>
                    <Download className="w-3.5 h-3.5" />
                    <span>Exportar</span>
                  </>
                )}
              </button>
            </div>
          </div>

          {/* Table */}
          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs border-collapse">
              <thead>
                <tr className="bg-slate-50/80 border-b border-slate-200 text-slate-600 font-extrabold uppercase tracking-wider text-[10px]">
                  <th className="py-3 px-3 w-12 text-center">#</th>
                  <th className="py-3 px-4 min-w-[260px]">Favorecido</th>
                  <th className="py-3 px-3 w-36">CNPJ</th>
                  <th className="py-3 px-3 w-28 text-center">Classificação</th>
                  <th className="py-3 px-3 w-32 text-center">Data do Atendimento</th>
                  <th className="py-3 px-3 w-36 text-right">Valor no Painel</th>
                  <th className="py-3 px-3 min-w-[220px]">Anotação / PP Emitida</th>
                  <th className="py-3 px-3 w-36 text-center">Ações</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 font-medium">
                {filteredSuppliers.length === 0 ? (
                  <tr>
                    <td colSpan={8} className="py-12 text-center text-slate-400 italic">
                      Nenhum fornecedor atende aos critérios da busca.
                    </td>
                  </tr>
                ) : (
                  paginatedSuppliers.map((sup, idx) => {
                    const isEditingObs = editingObsKey === sup.recordId;
                    const realIndex = (currentPage - 1) * pageSize + idx + 1;
                    const cat = categorizeSupplier(sup.name, sup.cnpj, allCategories, categoryOverrides);
                    const isCategorized = cat && cat.id !== 'cat_outros';
                    const catTheme = isCategorized ? (COLOR_MAP[cat.color] || COLOR_MAP.slate) : null;

                    return (
                      <tr
                        key={`atendido_row_${sup.cnpj || 'NI'}_${sup.name}_${selectedMonthKey}_${idx}`}
                        className={`hover:bg-slate-50 transition-colors ${
                          catTheme ? `border-l-4 ${catTheme.borderLeft}` : 'border-l-4 border-l-transparent'
                        }`}
                      >
                        {/* # */}
                        <td className="py-3 px-3 text-center font-mono text-[11px] text-slate-400 border-r border-slate-100">
                          {realIndex}
                        </td>

                        {/* Favorecido Name */}
                        <td className="py-3 px-4 font-bold text-slate-900 border-r border-slate-100">
                          <div className="flex flex-col gap-1 items-start">
                            <button
                              type="button"
                              onClick={() => onOpenCompanyAudit && onOpenCompanyAudit({ name: sup.name, cnpj: sup.cnpj })}
                              className="flex items-center space-x-2 text-left hover:text-emerald-700 transition cursor-pointer group"
                              title="Clique para auditar todas as ocorrências desta empresa"
                            >
                              <Building className="w-4 h-4 text-slate-400 group-hover:text-emerald-600 shrink-0 transition-colors" />
                              <span className="truncate max-w-xs sm:max-w-md group-hover:underline">
                                {sup.name}
                              </span>
                            </button>
                            {catTheme && (
                              <span
                                className={`inline-flex items-center gap-1.5 px-2 py-0.5 rounded-md text-[10px] font-extrabold border shadow-2xs whitespace-nowrap ${catTheme.badgeText}`}
                                title={`Categoria: ${cat.name}`}
                              >
                                <span className={`w-1.5 h-1.5 rounded-full ${catTheme.badge} shrink-0`} />
                                <span>{cat.name}</span>
                              </span>
                            )}
                          </div>
                        </td>

                        {/* CNPJ */}
                        <td className="py-3 px-3 font-mono text-slate-700 border-r border-slate-100">
                          <span className="bg-slate-100 text-slate-800 px-2 py-0.5 rounded-md border border-slate-200/80 text-[11px] font-bold inline-block">
                            {sup.cnpj || 'N/I'}
                          </span>
                        </td>

                        {/* Classificação / Recorrente */}
                        <td className="py-3 px-3 text-center border-r border-slate-100">
                          {sup.isRec ? (
                            <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-extrabold bg-blue-100 text-blue-800 border border-blue-200">
                              <Repeat className="w-2.5 h-2.5" />
                              <span>Recorrente</span>
                            </span>
                          ) : (
                            <span className="text-slate-400 text-[11px]">Comum</span>
                          )}
                        </td>

                        {/* Data Atendimento */}
                        <td className="py-3 px-3 text-center font-mono text-[11px] text-slate-600 border-r border-slate-100">
                          {sup.dataAtendimento ? (
                            <div className="flex items-center justify-center gap-1 text-emerald-800 font-bold">
                              <CheckCircle2 className="w-3 h-3 text-emerald-600 shrink-0" />
                              <span>{new Date(sup.dataAtendimento).toLocaleDateString('pt-BR')}</span>
                            </div>
                          ) : (
                            <span className="text-slate-400">-</span>
                          )}
                        </td>

                        {/* Valor no Painel */}
                        <td className="py-3 px-3 text-right font-mono font-bold text-slate-900 border-r border-slate-100">
                          {formatBRL(sup.totalValue)}
                        </td>

                        {/* Anotação / Observação */}
                        <td className="py-2.5 px-3 border-r border-slate-100">
                          {isEditingObs ? (
                            <div className="flex items-center gap-1.5">
                              <input
                                type="text"
                                value={obsText}
                                onChange={(e) => setObsText(e.target.value)}
                                placeholder="Ex: PP nº 1234 emitida / Pago"
                                className="w-full px-2 py-1 text-xs bg-white border border-emerald-500 rounded-md focus:outline-hidden"
                                autoFocus
                                onKeyDown={(e) => {
                                  if (e.key === 'Enter') handleSaveObs(sup);
                                  if (e.key === 'Escape') setEditingObsKey(null);
                                }}
                              />
                              <button
                                onClick={() => handleSaveObs(sup)}
                                className="p-1 text-emerald-700 hover:bg-emerald-100 rounded-md transition cursor-pointer"
                                title="Salvar"
                              >
                                <Check className="w-3.5 h-3.5" />
                              </button>
                              <button
                                onClick={() => setEditingObsKey(null)}
                                className="p-1 text-slate-400 hover:bg-slate-100 rounded-md transition cursor-pointer"
                                title="Cancelar"
                              >
                                <X className="w-3.5 h-3.5" />
                              </button>
                            </div>
                          ) : (
                            <div
                              onClick={() => handleStartEditObs(sup)}
                              className="group flex items-center justify-between gap-1 text-slate-600 hover:text-slate-900 cursor-pointer p-1 rounded-md hover:bg-slate-100 transition"
                              title="Clique para editar anotação deste mês"
                            >
                              <span className="truncate text-xs text-slate-700">
                                {sup.observacao || <span className="text-slate-300 italic">+ Adicionar anotação/PP</span>}
                              </span>
                              <Edit3 className="w-3 h-3 text-slate-300 group-hover:text-slate-600 shrink-0" />
                            </div>
                          )}
                        </td>

                        {/* Ações: Desmarcar e Extrato */}
                        <td className="py-3 px-3 text-center">
                          <div className="flex items-center justify-center gap-1.5">
                            <button
                              type="button"
                              onClick={() => handleUnmarkAttendance(sup)}
                              className="p-1.5 text-slate-400 hover:text-amber-700 hover:bg-amber-50 rounded-lg transition cursor-pointer"
                              title="Desmarcar atendimento deste fornecedor"
                            >
                              <Undo2 className="w-3.5 h-3.5" />
                            </button>
                            <button
                              type="button"
                              onClick={() => onOpenCompanyAudit && onOpenCompanyAudit({ name: sup.name, cnpj: sup.cnpj })}
                              className="inline-flex items-center gap-1 px-2.5 py-1 bg-slate-100 hover:bg-emerald-100 text-slate-700 hover:text-emerald-900 rounded-lg text-[11px] font-bold border border-slate-200 hover:border-emerald-300 transition cursor-pointer"
                              title="Extrato de Auditoria da Empresa"
                            >
                              <span>Extrato</span>
                              <ExternalLink className="w-3 h-3 text-slate-400" />
                            </button>
                          </div>
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>

          {/* Pagination Controls */}
          {totalPages > 1 && (
            <div className="px-5 py-3 border-t border-slate-100 flex items-center justify-between text-xs text-slate-500 bg-slate-50/50">
              <span>
                Página {currentPage} de {totalPages} ({filteredSuppliers.length} fornecedores atendidos)
              </span>
              <div className="flex items-center space-x-2">
                <button
                  type="button"
                  disabled={currentPage === 1}
                  onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                  className="p-1.5 rounded-lg border border-slate-200 bg-white hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed transition cursor-pointer"
                >
                  <ChevronLeft className="w-4 h-4" />
                </button>
                <button
                  type="button"
                  disabled={currentPage === totalPages}
                  onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
                  className="p-1.5 rounded-lg border border-slate-200 bg-white hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed transition cursor-pointer"
                >
                  <ChevronRight className="w-4 h-4" />
                </button>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
};
