import React, { useState, useMemo, useEffect, useDeferredValue } from 'react';
import {
  Search,
  Download,
  ArrowUpDown,
  Building2,
  Building,
  Repeat,
  Check,
  Plus,
  ChevronLeft,
  ChevronRight,
  FileSpreadsheet,
  ExternalLink,
  Loader2,
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { DashboardMetrics, ConsolidatedSupplier, FileData, QuebraOrdemItem, CategoryDefinition } from '../types';
import { formatBRL } from '../utils/excelParser';
import {
  getManualRecurringSuppliers,
  addManualRecurringSupplier,
  removeManualRecurringSupplier,
} from '../utils/recorrenciaService';
import {
  getCurrentMonthKey,
  formatMonthShort,
  formatMonthFull,
  getAllAtendimentos,
  getSupplierStorageKey,
  toggleSupplierAtendimento,
  subscribeToAtendimentos,
  AtendimentoMensalRecord,
} from '../utils/atendimentoMensalService';
import {
  getCategories,
  getSupplierOverrides,
  subscribeToCategories,
  categorizeSupplier,
  COLOR_MAP,
} from '../utils/categoryService';
import { ExtratoAuditoriaEmpresaModal } from './ExtratoAuditoriaEmpresaModal';
import { BankAuditRecord } from '../utils/auditCompanyService';
import { CheckCircle2, Clock } from 'lucide-react';
import { InfoHelpButton } from './InfoHelpButton';

interface DetailedConsolidationTableProps {
  metrics: DashboardMetrics;
  files?: FileData[];
  quebraOrdemItems?: QuebraOrdemItem[];
  bankRecords?: BankAuditRecord[];
  onOpenCompanyAudit?: (supplier: { name: string; cnpj: string }) => void;
}

type SortField = 'totalValue' | 'ppProntasValue' | 'liqObedeceValue' | 'liqNaoObedeceValue' | 'name' | 'cnpj';

export const DetailedConsolidationTable: React.FC<DetailedConsolidationTableProps> = ({
  metrics,
  files = [],
  quebraOrdemItems = [],
  bankRecords = [],
  onOpenCompanyAudit,
}) => {
  const { consolidatedSuppliers, isProcessed } = metrics;

  const [searchTerm, setSearchTerm] = useState('');
  const deferredSearchTerm = useDeferredValue(searchTerm);
  const [sortField, setSortField] = useState<SortField>('totalValue');
  const [sortOrder, setSortOrder] = useState<'desc' | 'asc'>('desc');
  const [filterMode, setFilterMode] = useState<'all' | 'recurring_only' | 'attended_only' | 'pending_only'>('all');
  const [recurringKeysSet, setRecurringKeysSet] = useState<Set<string>>(new Set());
  const [attendanceTrigger, setAttendanceTrigger] = useState(0);
  const [allAtendimentosMap, setAllAtendimentosMap] = useState<Record<string, AtendimentoMensalRecord>>(() => getAllAtendimentos());

  // Month reference
  const currentMonthKey = useMemo(() => getCurrentMonthKey(), []);
  const monthShortLabel = useMemo(() => formatMonthShort(currentMonthKey), [currentMonthKey]);
  const monthFullLabel = useMemo(() => formatMonthFull(currentMonthKey), [currentMonthKey]);

  // Modal for internal company audit
  const [internalAuditedSupplier, setInternalAuditedSupplier] = useState<{ name: string; cnpj: string } | null>(null);

  // Pagination state for ultra fluid rendering
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState<number>(50);
  const [categoryOverrides, setCategoryOverrides] = useState<Record<string, string>>(() => getSupplierOverrides());
  const [allCategories, setAllCategories] = useState<CategoryDefinition[]>(() => getCategories());

  const handleAuditClick = (supplier: { name: string; cnpj: string }) => {
    if (onOpenCompanyAudit) {
      onOpenCompanyAudit(supplier);
    } else {
      setInternalAuditedSupplier(supplier);
    }
  };

  // Load recurring suppliers and subscribe to attendance and category changes
  useEffect(() => {
    const recList = getManualRecurringSuppliers();
    const set = new Set<string>();
    recList.forEach((r) => {
      if (r.cnpj && r.cnpj !== 'N/I') set.add(r.cnpj.replace(/\D/g, ''));
      if (r.name) set.add(r.name.trim().toUpperCase());
    });
    setRecurringKeysSet(set);

    const unsubscribeAtendimentos = subscribeToAtendimentos(() => {
      setAllAtendimentosMap({ ...getAllAtendimentos() });
      setAttendanceTrigger((prev) => prev + 1);
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

  const isSupplierRecurring = (supplier: ConsolidatedSupplier) => {
    const cnpjDigits = (supplier.cnpj || '').replace(/\D/g, '');
    if (cnpjDigits.length === 14 && recurringKeysSet.has(cnpjDigits)) return true;
    if (supplier.name && recurringKeysSet.has(supplier.name.trim().toUpperCase())) return true;
    return false;
  };

  const isSupplierAttendedThisMonth = (supplier: ConsolidatedSupplier) => {
    const supKey = getSupplierStorageKey(supplier.name, supplier.cnpj);
    const recordId = `${currentMonthKey}_${supKey}`;
    return !!allAtendimentosMap[recordId]?.atendido;
  };

  const toggleSupplierAttendance = (supplier: ConsolidatedSupplier) => {
    const isRec = isSupplierRecurring(supplier);
    toggleSupplierAtendimento(supplier.name, supplier.cnpj, isRec, currentMonthKey);
    setAllAtendimentosMap({ ...getAllAtendimentos() });
    setAttendanceTrigger((prev) => prev + 1);
  };

  const toggleSupplierRecurrence = (supplier: ConsolidatedSupplier) => {
    const isCurrently = isSupplierRecurring(supplier);
    const recList = getManualRecurringSuppliers();

    if (isCurrently) {
      // Find matching item and remove
      const match = recList.find(
        (r) =>
          (r.cnpj !== 'N/I' && r.cnpj.replace(/\D/g, '') === (supplier.cnpj || '').replace(/\D/g, '')) ||
          r.name.trim().toUpperCase() === supplier.name.trim().toUpperCase()
      );
      if (match) {
        const updated = removeManualRecurringSupplier(match.id);
        const set = new Set<string>();
        updated.forEach((r) => {
          if (r.cnpj && r.cnpj !== 'N/I') set.add(r.cnpj.replace(/\D/g, ''));
          if (r.name) set.add(r.name.trim().toUpperCase());
        });
        setRecurringKeysSet(set);
      }
    } else {
      // Add
      const updated = addManualRecurringSupplier({
        name: supplier.name,
        cnpj: supplier.cnpj || 'N/I',
        notes: '',
      });
      const set = new Set<string>();
      updated.forEach((r) => {
        if (r.cnpj && r.cnpj !== 'N/I') set.add(r.cnpj.replace(/\D/g, ''));
        if (r.name) set.add(r.name.trim().toUpperCase());
      });
      setRecurringKeysSet(set);
    }
  };

  // Filter and Sort Logic with deferred search
  const filteredSuppliers = useMemo(() => {
    if (!consolidatedSuppliers || consolidatedSuppliers.length === 0) return [];

    const query = deferredSearchTerm.toLowerCase().trim();

    const filtered = consolidatedSuppliers.filter((supplier) => {
      if (filterMode === 'recurring_only' && !isSupplierRecurring(supplier)) {
        return false;
      }
      if (filterMode === 'attended_only' && !isSupplierAttendedThisMonth(supplier)) {
        return false;
      }
      if (filterMode === 'pending_only' && isSupplierAttendedThisMonth(supplier)) {
        return false;
      }

      if (!query) return true;
      const cleanQuery = query.replace(/\D/g, '');
      const cleanCnpj = (supplier.cnpj || '').replace(/\D/g, '');
      return (
        supplier.name.toLowerCase().includes(query) ||
        (supplier.cnpj || '').toLowerCase().includes(query) ||
        (cleanQuery.length >= 3 && cleanCnpj.includes(cleanQuery))
      );
    });

    const isStringField = sortField === 'name' || sortField === 'cnpj';

    // Map each item to its sort key once O(N) instead of during every O(N log N) compare
    const mapped = filtered.map((s, idx) => {
      let key: any = s[sortField] ?? 0;
      if (isStringField && typeof key === 'string') {
        key = key.toLowerCase();
      }
      return { s, key, idx };
    });

    mapped.sort((a, b) => {
      if (a.key < b.key) return sortOrder === 'asc' ? -1 : 1;
      if (a.key > b.key) return sortOrder === 'asc' ? 1 : -1;
      return a.idx - b.idx;
    });

    return mapped.map((m) => m.s);
  }, [consolidatedSuppliers, filterMode, recurringKeysSet, attendanceTrigger, deferredSearchTerm, sortField, sortOrder, allAtendimentosMap]);

  // Reset page to 1 when filters or sorting change
  useEffect(() => {
    setCurrentPage(1);
  }, [deferredSearchTerm, filterMode, sortField, sortOrder, pageSize]);

  // Total pages and paginated slice
  const totalPages = pageSize === 0 ? 1 : Math.ceil(filteredSuppliers.length / pageSize) || 1;
  const paginatedSuppliers = useMemo(() => {
    if (pageSize === 0) return filteredSuppliers;
    const start = (currentPage - 1) * pageSize;
    return filteredSuppliers.slice(start, start + pageSize);
  }, [filteredSuppliers, currentPage, pageSize]);

  const toggleSort = (field: SortField) => {
    if (sortField === field) {
      setSortOrder(sortOrder === 'asc' ? 'desc' : 'asc');
    } else {
      setSortField(field);
      setSortOrder('desc');
    }
  };

  const [isExporting, setIsExporting] = useState(false);

  const handleExportConsolidatedExcel = () => {
    if (!filteredSuppliers || filteredSuppliers.length === 0 || isExporting) return;

    setIsExporting(true);
    setTimeout(() => {
      try {
        const dataToExport = filteredSuppliers.map((s) => {
          return {
            'FAVORECIDO': s.name,
            'CNPJ': s.cnpj,
            'PPS PRONTAS (R$)': s.ppProntasValue || 0,
            'LIQUIDAÇÃO OBEDECE (R$)': s.liqObedeceValue || 0,
            'LIQUIDAÇÃO NÃO OBEDECE (R$)': s.liqNaoObedeceValue || 0,
            'VALOR TOTAL (R$)': s.totalValue,
            'QTD REGISTROS': s.totalCount,
          };
        });

        const worksheet = XLSX.utils.json_to_sheet(dataToExport);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Tabela_Credores');

        XLSX.writeFile(workbook, `Tabela_Credores_${new Date().toISOString().slice(0, 10)}.xlsx`);
      } finally {
        setIsExporting(false);
      }
    }, 200);
  };

  if (!isProcessed) return null;

  return (
    <section id="tabela-principal" className="bg-white rounded-2xl p-6 shadow-xs border border-slate-200/90 space-y-5">
      
      {/* Header & Export Button */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-slate-100 pb-4">
        <div className="flex items-center gap-2.5">
          <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
            <Building2 className="w-5 h-5 text-emerald-600" />
            <span>Tabela Principal de Credores</span>
          </h3>
          <InfoHelpButton
            title="Tabela de Credores"
            variant="light"
            content={
              <p>
                Consolidação geral de credores por CNPJ/Favorecido, totalizando os pagamentos emitidos no exercício e os valores na Base de PPs e Liquidações.
              </p>
            }
          />
        </div>

        {/* Export Button */}
        <button
          onClick={handleExportConsolidatedExcel}
          disabled={isExporting}
          className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-800 border border-slate-300 text-xs font-bold rounded-xl transition flex items-center gap-2 shrink-0 shadow-2xs cursor-pointer active:scale-98 disabled:opacity-75 disabled:cursor-wait"
        >
          {isExporting ? (
            <>
              <Loader2 className="w-4 h-4 animate-spin text-emerald-600" />
              <span>Exportando Relatório...</span>
            </>
          ) : (
            <>
              <Download className="w-4 h-4 text-slate-600" />
              <span>Exportar Relatório</span>
            </>
          )}
        </button>
      </div>

      {/* Filter and Search Bar */}
      <div className="flex flex-col sm:flex-row items-center justify-between gap-3">
        {/* Search Input */}
        <div className="relative w-full sm:w-96">
          <Search className="w-4 h-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2" />
          <input
            type="text"
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            placeholder="Buscar por Favorecido ou CNPJ..."
            className="w-full pl-9 pr-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs text-slate-900 focus:bg-white focus:outline-hidden focus:border-emerald-500 transition font-medium"
          />
        </div>

        <div className="flex flex-wrap items-center gap-2 self-end sm:self-auto">
          {/* Filter by Recurrence & Attendance */}
          <div className="flex flex-wrap items-center bg-slate-100 p-1 rounded-xl border border-slate-200 text-xs font-bold gap-1">
            <button
              onClick={() => setFilterMode('all')}
              className={`px-3 py-1 rounded-lg transition cursor-pointer active:scale-98 ${
                filterMode === 'all'
                  ? 'bg-white text-slate-900 shadow-2xs font-extrabold'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              Todos ({consolidatedSuppliers.length})
            </button>
            <button
              onClick={() => setFilterMode('recurring_only')}
              className={`px-3 py-1 rounded-lg transition flex items-center gap-1.5 cursor-pointer active:scale-98 ${
                filterMode === 'recurring_only'
                  ? 'bg-blue-600 text-white shadow-2xs font-extrabold'
                  : 'text-blue-800 hover:bg-blue-50'
              }`}
            >
              <Repeat className="w-3 h-3" />
              <span>Recorrentes</span>
            </button>
            <button
              onClick={() => setFilterMode('attended_only')}
              className={`px-3 py-1 rounded-lg transition flex items-center gap-1.5 cursor-pointer active:scale-98 ${
                filterMode === 'attended_only'
                  ? 'bg-emerald-600 text-white shadow-2xs font-extrabold'
                  : 'text-emerald-800 hover:bg-emerald-50'
              }`}
            >
              <CheckCircle2 className="w-3 h-3" />
              <span>Atendidos ({monthShortLabel})</span>
            </button>
            <button
              onClick={() => setFilterMode('pending_only')}
              className={`px-3 py-1 rounded-lg transition flex items-center gap-1.5 cursor-pointer active:scale-98 ${
                filterMode === 'pending_only'
                  ? 'bg-amber-600 text-white shadow-2xs font-extrabold'
                  : 'text-amber-800 hover:bg-amber-50'
              }`}
            >
              <Clock className="w-3 h-3" />
              <span>Pendentes ({monthShortLabel})</span>
            </button>
          </div>

          {/* Page size selector */}
          <div className="flex items-center gap-1.5 text-xs text-slate-600 font-bold bg-slate-50 px-2.5 py-1 rounded-xl border border-slate-200">
            <span className="text-[11px] text-slate-400">Exibir:</span>
            <select
              value={pageSize}
              onChange={(e) => setPageSize(Number(e.target.value))}
              className="bg-transparent border-none font-bold text-slate-800 text-xs cursor-pointer focus:outline-none"
            >
              <option value={25}>25</option>
              <option value={50}>50</option>
              <option value={100}>100</option>
              <option value={0}>Todos</option>
            </select>
          </div>

          <div className="text-xs text-slate-500 font-semibold hidden md:block">
            <span><strong>{filteredSuppliers.length}</strong> credores</span>
          </div>
        </div>
      </div>

      {/* Table Content Container */}
      <div className="border border-slate-200/90 rounded-xl overflow-hidden shadow-xs">
        <div className="overflow-x-auto overflow-y-auto max-h-[560px]">
          <table className="w-full text-left border-collapse text-xs">
            <thead className="sticky top-0 z-10 bg-slate-100/95 backdrop-blur-xs shadow-xs">
              <tr className="border-b border-slate-200 text-slate-700 font-bold uppercase tracking-wider text-[11px]">
                <th className="py-3.5 px-3 w-12 text-center text-slate-400">#</th>
                <th className="py-3.5 px-4 min-w-[220px]">
                  <button
                    onClick={() => toggleSort('name')}
                    className="flex items-center gap-1.5 hover:text-emerald-700 cursor-pointer font-extrabold active:scale-98"
                  >
                    <span>FAVORECIDO / EMPRESA</span>
                    <ArrowUpDown className="w-3.5 h-3.5 text-slate-400" />
                  </button>
                </th>
                <th className="py-3.5 px-3 text-center min-w-[130px] whitespace-nowrap">
                  <span>RECORRÊNCIA</span>
                </th>
                <th className="py-3.5 px-3 text-center min-w-[140px] whitespace-nowrap">
                  <span>MÊS ATUAL ({monthShortLabel.toUpperCase()})</span>
                </th>
                <th className="py-3.5 px-3 text-center min-w-[110px] whitespace-nowrap">
                  <span>AUDITORIA</span>
                </th>
                <th className="py-3.5 px-3 text-right w-40 min-w-[140px] bg-emerald-50/40 whitespace-nowrap">
                  <button
                    onClick={() => toggleSort('ppProntasValue')}
                    className="flex items-center justify-end gap-1 w-full hover:text-emerald-700 cursor-pointer font-bold text-emerald-900 active:scale-98 whitespace-nowrap"
                  >
                    <span>BASE DE PPS (R$)</span>
                    <ArrowUpDown className="w-3 h-3 text-emerald-600" />
                  </button>
                </th>
                <th className="py-3.5 px-3 text-right w-44 min-w-[155px] bg-blue-50/40 whitespace-nowrap">
                  <button
                    onClick={() => toggleSort('liqObedeceValue')}
                    className="flex items-center justify-end gap-1 w-full hover:text-blue-700 cursor-pointer font-bold text-blue-900 active:scale-98 whitespace-nowrap"
                  >
                    <span>LIQUIDADO - OBEDECE (R$)</span>
                    <ArrowUpDown className="w-3 h-3 text-blue-600" />
                  </button>
                </th>
                <th className="py-3.5 px-3 text-right w-48 min-w-[170px] bg-amber-50/40 whitespace-nowrap">
                  <button
                    onClick={() => toggleSort('liqNaoObedeceValue')}
                    className="flex items-center justify-end gap-1 w-full hover:text-amber-700 cursor-pointer font-bold text-amber-900 active:scale-98 whitespace-nowrap"
                  >
                    <span>LIQUIDADO - NÃO OBEDECE (R$)</span>
                    <ArrowUpDown className="w-3 h-3 text-amber-600" />
                  </button>
                </th>
                <th className="py-3.5 px-4 text-right w-44 min-w-[150px] bg-slate-200/60 whitespace-nowrap">
                  <button
                    onClick={() => toggleSort('totalValue')}
                    className="flex items-center justify-end gap-1.5 w-full hover:text-emerald-700 cursor-pointer font-black text-slate-900 active:scale-98 whitespace-nowrap"
                  >
                    <span>VALOR TOTAL</span>
                    <ArrowUpDown className="w-3.5 h-3.5 text-slate-600" />
                  </button>
                </th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 font-medium">
              {paginatedSuppliers.length === 0 ? (
                <tr>
                  <td colSpan={9} className="py-12 text-center text-slate-400 italic">
                    Nenhum favorecido localizado com o termo digitado.
                  </td>
                </tr>
              ) : (
                paginatedSuppliers.map((supplier, idx) => {
                  const isRec = isSupplierRecurring(supplier);
                  const isAttended = isSupplierAttendedThisMonth(supplier);
                  const realIndex = pageSize === 0 ? idx + 1 : (currentPage - 1) * pageSize + idx + 1;
                  const cat = categorizeSupplier(supplier.name, supplier.cnpj, allCategories, categoryOverrides);
                  const isCategorized = cat && cat.id !== 'cat_outros';
                  const catTheme = isCategorized ? (COLOR_MAP[cat.color] || COLOR_MAP.slate) : null;

                  return (
                    <tr
                      key={`${supplier.cnpj}-${supplier.name}-${idx}`}
                      className={`hover:bg-slate-50 transition-colors ${
                        catTheme ? `border-l-4 ${catTheme.borderLeft}` : 'border-l-4 border-l-transparent'
                      }`}
                    >
                      {/* Index */}
                      <td className="py-3 px-3 text-center font-mono text-[11px] text-slate-400 border-r border-slate-100">
                        {realIndex}
                      </td>

                      {/* Favorecido Name - Clickable for Audit */}
                      <td className="py-3 px-4 font-bold text-slate-900 border-r border-slate-100">
                        <div className="flex flex-col gap-1 items-start">
                          <button
                            type="button"
                            onClick={() => handleAuditClick({ name: supplier.name, cnpj: supplier.cnpj })}
                            className="flex items-center space-x-2 text-left hover:text-emerald-700 transition cursor-pointer group"
                            title="Clique para auditar todas as ocorrências desta empresa"
                          >
                            <Building className="w-4 h-4 text-slate-400 group-hover:text-emerald-600 shrink-0 transition-colors" />
                            <span className="truncate max-w-xs sm:max-w-md text-xs group-hover:underline">
                              {supplier.name}
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

                      {/* Recorrência Interactive Button */}
                      <td className="py-3 px-3 text-center border-r border-slate-100 whitespace-nowrap">
                        <div className="flex items-center justify-center w-full whitespace-nowrap">
                          <button
                            type="button"
                            onClick={() => toggleSupplierRecurrence(supplier)}
                            className={`inline-flex items-center justify-center gap-1 px-2.5 py-1 rounded-lg text-[11px] font-extrabold transition cursor-pointer active:scale-95 whitespace-nowrap ${
                              isRec
                                ? 'bg-emerald-100 text-emerald-900 border border-emerald-300 hover:bg-rose-100 hover:text-rose-900 hover:border-rose-300'
                                : 'bg-slate-100 text-slate-500 border border-slate-200 hover:bg-emerald-50 hover:text-emerald-800 hover:border-emerald-300'
                            }`}
                            title={isRec ? 'Clique para desmarcar como recorrente' : 'Clique para marcar como fornecedor recorrente'}
                          >
                            <Repeat className="w-3 h-3 shrink-0" />
                            <span className="whitespace-nowrap">{isRec ? 'Recorrente' : '+ Marcar'}</span>
                          </button>
                        </div>
                      </td>

                      {/* Atendimento no Mês Atual */}
                      <td className="py-3 px-3 text-center border-r border-slate-100 whitespace-nowrap">
                        <div className="flex items-center justify-center w-full whitespace-nowrap">
                          <button
                            type="button"
                            onClick={() => toggleSupplierAttendance(supplier)}
                            className={`inline-flex items-center justify-center gap-1 px-2.5 py-1 rounded-lg text-[11px] font-extrabold transition cursor-pointer active:scale-95 shadow-2xs whitespace-nowrap ${
                              isAttended
                                ? 'bg-emerald-600 text-white hover:bg-emerald-700'
                                : 'bg-slate-100 text-slate-600 border border-slate-200 hover:bg-emerald-50 hover:text-emerald-800 hover:border-emerald-300'
                            }`}
                            title={
                              isAttended
                                ? `Atendido em ${monthFullLabel}. Clique para desmarcar.`
                                : `Marcar como atendido em ${monthFullLabel}.`
                            }
                          >
                            {isAttended ? (
                              <>
                                <CheckCircle2 className="w-3 h-3 text-white shrink-0" />
                                <span className="whitespace-nowrap">Atendido</span>
                              </>
                            ) : (
                              <>
                                <Clock className="w-3 h-3 text-slate-400 shrink-0" />
                                <span className="whitespace-nowrap">+ Atender</span>
                              </>
                            )}
                          </button>
                        </div>
                      </td>

                      {/* Extrato / Auditoria Rápida Button */}
                      <td className="py-3 px-3 text-center border-r border-slate-100 whitespace-nowrap">
                        <div className="flex items-center justify-center w-full whitespace-nowrap">
                          <button
                            type="button"
                            onClick={() => handleAuditClick({ name: supplier.name, cnpj: supplier.cnpj })}
                            className="inline-flex items-center justify-center gap-1 px-2.5 py-1 bg-slate-100 hover:bg-emerald-100 text-slate-700 hover:text-emerald-900 rounded-lg text-[11px] font-bold border border-slate-200 hover:border-emerald-300 transition cursor-pointer active:scale-95 whitespace-nowrap"
                            title="Abrir Extrato Geral de Auditoria da Empresa"
                          >
                            <span className="whitespace-nowrap">Extrato</span>
                            <ExternalLink className="w-3 h-3 text-slate-400 shrink-0" />
                          </button>
                        </div>
                      </td>

                      {/* Base de PPs */}
                      <td className="py-3 px-3 text-right font-mono border-r border-slate-100 bg-emerald-50/10 whitespace-nowrap">
                        <span className={`text-xs whitespace-nowrap ${(supplier.basePpsValue ?? supplier.ppProntasValue ?? 0) > 0 ? 'font-bold text-emerald-900' : 'text-slate-300'}`}>
                          {formatBRL(supplier.basePpsValue ?? supplier.ppProntasValue ?? 0)}
                        </span>
                      </td>

                      {/* Liquidação Obedece */}
                      <td className="py-3 px-3 text-right font-mono border-r border-slate-100 bg-blue-50/10 whitespace-nowrap">
                        <span className={`text-xs whitespace-nowrap ${supplier.liqObedeceValue > 0 ? 'font-bold text-blue-900' : 'text-slate-300'}`}>
                          {formatBRL(supplier.liqObedeceValue || 0)}
                        </span>
                      </td>

                      {/* Liquidação Não Obedece */}
                      <td className="py-3 px-3 text-right font-mono border-r border-slate-100 bg-amber-50/10 whitespace-nowrap">
                        <span className={`text-xs whitespace-nowrap ${supplier.liqNaoObedeceValue > 0 ? 'font-bold text-amber-900' : 'text-slate-300'}`}>
                          {formatBRL(supplier.liqNaoObedeceValue || 0)}
                        </span>
                      </td>

                      {/* Valor Total */}
                      <td className="py-3 px-4 text-right font-mono bg-slate-100/50 whitespace-nowrap">
                        <span className="text-sm font-black text-slate-950 tracking-tight whitespace-nowrap">
                          {formatBRL(supplier.totalValue)}
                        </span>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Table Summary and Pagination Footer */}
      <div className="flex flex-col sm:flex-row items-center justify-between gap-3 text-xs text-slate-500 pt-1 font-medium">
        <span>
          Exibindo <strong>{filteredSuppliers.length}</strong> credores
        </span>

        {/* Pagination controls */}
        {totalPages > 1 && pageSize > 0 && (
          <div className="flex items-center gap-1.5">
            <button
              onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
              disabled={currentPage === 1}
              className="p-1.5 rounded-lg border border-slate-200 bg-slate-50 hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed transition cursor-pointer active:scale-95"
              title="Página Anterior"
            >
              <ChevronLeft className="w-4 h-4" />
            </button>
            <span className="px-3 py-1 font-bold text-slate-700 bg-slate-100 rounded-lg text-xs">
              Página {currentPage} de {totalPages}
            </span>
            <button
              onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
              disabled={currentPage === totalPages}
              className="p-1.5 rounded-lg border border-slate-200 bg-slate-50 hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed transition cursor-pointer active:scale-95"
              title="Próxima Página"
            >
              <ChevronRight className="w-4 h-4" />
            </button>
          </div>
        )}
      </div>

      {/* Company Audit Modal */}
      {internalAuditedSupplier && (
        <ExtratoAuditoriaEmpresaModal
          supplier={internalAuditedSupplier}
          onClose={() => setInternalAuditedSupplier(null)}
          files={files}
          quebraOrdemItems={quebraOrdemItems}
          bankRecords={bankRecords}
          isRecurring={
            recurringKeysSet.has((internalAuditedSupplier.cnpj || '').replace(/\D/g, '')) ||
            recurringKeysSet.has((internalAuditedSupplier.name || '').trim().toUpperCase())
          }
        />
      )}
    </section>
  );
};

