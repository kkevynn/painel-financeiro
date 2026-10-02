import React, { useState, useMemo, useEffect, useDeferredValue } from 'react';
import {
  Repeat,
  Plus,
  Trash2,
  Search,
  Download,
  ArrowUpDown,
  X,
  Pencil,
  RotateCcw,
  AlertTriangle,
  CheckCircle2,
  CheckSquare,
  Square,
  Check,
  Building2,
  UserPlus,
  ChevronLeft,
  ChevronRight,
  Loader2,
} from 'lucide-react';
import * as XLSX from 'xlsx';
import {
  FileData,
  DashboardMetrics,
  RecurringSupplier,
  ManualRecurringSupplierEntry,
} from '../../types';
import {
  getManualRecurringSuppliers,
  addManualRecurringSupplier,
  addMultipleManualRecurringSuppliers,
  updateManualRecurringSupplier,
  removeManualRecurringSupplier,
  clearAllManualRecurringSuppliers,
  calculateManualRecurringMetrics,
} from '../../utils/recorrenciaService';
import { InfoHelpButton } from '../InfoHelpButton';
import { formatBRL } from '../../utils/excelParser';

interface RecorrenciaViewProps {
  metrics: DashboardMetrics;
  files: FileData[];
}

type SortField =
  | 'name'
  | 'totalValue'
  | 'ppProntasValue'
  | 'liqObedeceValue'
  | 'liqNaoObedeceValue';

export const RecorrenciaView: React.FC<RecorrenciaViewProps> = ({ metrics, files }) => {
  const [manualList, setManualList] = useState<ManualRecurringSupplierEntry[]>([]);
  const [searchTerm, setSearchTerm] = useState('');
  const deferredSearchTerm = useDeferredValue(searchTerm);
  const [filterActiveOnly, setFilterActiveOnly] = useState<boolean>(false);
  const [sortField, setSortField] = useState<SortField>('totalValue');
  const [sortOrder, setSortOrder] = useState<'desc' | 'asc'>('desc');

  // Pagination state for fluid table rendering
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState<number>(50);

  // Modal state for adding supplier
  const [isAddModalOpen, setIsAddModalOpen] = useState(false);
  const [addMode, setAddMode] = useState<'multi' | 'manual'>('multi');
  const [selectedCreditorKeys, setSelectedCreditorKeys] = useState<Set<string>>(new Set());
  const [supplierInputName, setSupplierInputName] = useState('');
  const [supplierInputCnpj, setSupplierInputCnpj] = useState('');
  const [supplierInputNotes, setSupplierInputNotes] = useState('');
  const [modalSearchFilter, setModalSearchFilter] = useState('');
  const deferredModalSearchFilter = useDeferredValue(modalSearchFilter);
  const [formError, setFormError] = useState<string | null>(null);

  // Modal state for editing supplier
  const [editingSupplier, setEditingSupplier] = useState<ManualRecurringSupplierEntry | null>(null);
  const [editName, setEditName] = useState('');
  const [editCnpj, setEditCnpj] = useState('');
  const [editNotes, setEditNotes] = useState('');
  const [editError, setEditError] = useState<string | null>(null);

  // Modal state for deleting confirmation
  const [supplierToDelete, setSupplierToDelete] = useState<{ id: string; name: string; cnpj: string } | null>(null);

  // Modal state for reset confirmation
  const [isResetConfirmOpen, setIsResetConfirmOpen] = useState(false);

  // Toast notification
  const [toastMessage, setToastMessage] = useState<string | null>(null);

  const showToast = (msg: string) => {
    setToastMessage(msg);
    setTimeout(() => {
      setToastMessage((current) => (current === msg ? null : current));
    }, 3000);
  };

  // Load manual list on mount
  useEffect(() => {
    const list = getManualRecurringSuppliers();
    setManualList(list);
  }, []);

  // Compute calculated metrics
  const summary = useMemo(() => {
    return calculateManualRecurringMetrics(files, manualList);
  }, [files, manualList]);

  // List of all suppliers found in files for easy selection in Add modal
  const availableCreditorsFromFiles = useMemo(() => {
    if (!metrics.consolidatedSuppliers) return [];
    return metrics.consolidatedSuppliers;
  }, [metrics.consolidatedSuppliers]);

  // Filtered available suppliers for modal selection
  const filteredModalCreditors = useMemo(() => {
    if (!deferredModalSearchFilter.trim()) {
      return availableCreditorsFromFiles;
    }
    const q = deferredModalSearchFilter.toLowerCase().trim();
    return availableCreditorsFromFiles.filter(
      (s) => s.name.toLowerCase().includes(q) || s.cnpj.toLowerCase().includes(q)
    );
  }, [availableCreditorsFromFiles, deferredModalSearchFilter]);

  // Filter and sort display list
  const displaySuppliers = useMemo(() => {
    return summary.suppliers
      .filter((s) => {
        if (filterActiveOnly && s.totalCount === 0) return false;
        if (!deferredSearchTerm.trim()) return true;
        const q = deferredSearchTerm.toLowerCase().trim();
        return (
          s.name.toLowerCase().includes(q) ||
          s.cnpj.toLowerCase().includes(q) ||
          (s.notes && s.notes.toLowerCase().includes(q))
        );
      })
      .sort((a, b) => {
        let valA: any = a[sortField] || 0;
        let valB: any = b[sortField] || 0;

        if (typeof valA === 'string') {
          valA = (valA as string).toLowerCase();
          valB = (valB as string).toLowerCase();
        }

        if (valA < valB) return sortOrder === 'asc' ? -1 : 1;
        if (valA > valB) return sortOrder === 'asc' ? 1 : -1;
        return 0;
      });
  }, [summary.suppliers, filterActiveOnly, deferredSearchTerm, sortField, sortOrder]);

  // Reset page when filter or sorting changes
  useEffect(() => {
    setCurrentPage(1);
  }, [deferredSearchTerm, filterActiveOnly, sortField, sortOrder, pageSize]);

  // Pagination calculation
  const totalPages = pageSize === 0 ? 1 : Math.ceil(displaySuppliers.length / pageSize) || 1;
  const paginatedSuppliers = useMemo(() => {
    if (pageSize === 0) return displaySuppliers;
    const start = (currentPage - 1) * pageSize;
    return displaySuppliers.slice(start, start + pageSize);
  }, [displaySuppliers, currentPage, pageSize]);

  const toggleSort = (field: SortField) => {
    if (sortField === field) {
      setSortOrder(sortOrder === 'asc' ? 'desc' : 'asc');
    } else {
      setSortField(field);
      setSortOrder('desc');
    }
  };

  // Toggle multi-selection for a creditor
  const toggleSelectCreditor = (creditor: { name: string; cnpj: string }) => {
    const key = `${creditor.cnpj || 'NI'}__${creditor.name}`;
    setSelectedCreditorKeys((prev) => {
      const next = new Set(prev);
      if (next.has(key)) {
        next.delete(key);
      } else {
        next.add(key);
      }
      return next;
    });
  };

  // Select/Deselect all filtered
  const handleToggleSelectAllFiltered = () => {
    const next = new Set(selectedCreditorKeys);
    const allFilteredSelected =
      filteredModalCreditors.length > 0 &&
      filteredModalCreditors.every((c) => next.has(`${c.cnpj || 'NI'}__${c.name}`));

    if (allFilteredSelected) {
      filteredModalCreditors.forEach((c) => next.delete(`${c.cnpj || 'NI'}__${c.name}`));
    } else {
      filteredModalCreditors.forEach((c) => next.add(`${c.cnpj || 'NI'}__${c.name}`));
    }
    setSelectedCreditorKeys(next);
  };

  // Add multiple selected creditors
  const handleAddMultipleSelected = () => {
    if (selectedCreditorKeys.size === 0) return;

    const creditorsToAdd: { name: string; cnpj: string; notes?: string }[] = [];
    availableCreditorsFromFiles.forEach((c) => {
      const key = `${c.cnpj || 'NI'}__${c.name}`;
      if (selectedCreditorKeys.has(key)) {
        creditorsToAdd.push({
          name: c.name,
          cnpj: c.cnpj || 'N/I',
          notes: '',
        });
      }
    });

    if (creditorsToAdd.length > 0) {
      const { updatedList, addedCount } = addMultipleManualRecurringSuppliers(creditorsToAdd);
      setManualList(updatedList);
      setSelectedCreditorKeys(new Set());
      setIsAddModalOpen(false);
      showToast(`${addedCount} fornecedor(es) adicionado(s) aos recorrentes com sucesso!`);
    }
  };

  // Handle Add Single Supplier Manually
  const handleSaveNewSupplier = (e: React.FormEvent) => {
    e.preventDefault();
    if (!supplierInputName.trim() && !supplierInputCnpj.trim()) {
      setFormError('Informe o Nome ou CNPJ do fornecedor.');
      return;
    }

    const updated = addManualRecurringSupplier({
      name: supplierInputName.trim() || 'FORNECEDOR SEM NOME',
      cnpj: supplierInputCnpj.trim() || 'N/I',
      notes: supplierInputNotes.trim(),
    });

    setManualList(updated);
    setSupplierInputName('');
    setSupplierInputCnpj('');
    setSupplierInputNotes('');
    setModalSearchFilter('');
    setFormError(null);
    setIsAddModalOpen(false);
    showToast('Fornecedor adicionado aos recorrentes com sucesso!');
  };

  // Open Edit Modal
  const handleOpenEditModal = (supplier: RecurringSupplier) => {
    const entry = manualList.find((m) => m.id === supplier.id) || {
      id: supplier.id || '',
      name: supplier.name,
      cnpj: supplier.cnpj,
      notes: supplier.notes || '',
      createdAt: new Date().toISOString(),
    };
    setEditingSupplier(entry);
    setEditName(entry.name);
    setEditCnpj(entry.cnpj !== 'N/I' ? entry.cnpj : '');
    setEditNotes(entry.notes || '');
    setEditError(null);
  };

  // Save Edit
  const handleSaveEdit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingSupplier) return;
    if (!editName.trim() && !editCnpj.trim()) {
      setEditError('Informe o Nome ou CNPJ do fornecedor.');
      return;
    }

    const updated = updateManualRecurringSupplier(editingSupplier.id, {
      name: editName.trim() || 'FORNECEDOR SEM NOME',
      cnpj: editCnpj.trim() || 'N/I',
      notes: editNotes.trim(),
    });

    setManualList(updated);
    setEditingSupplier(null);
    setEditError(null);
    showToast('Fornecedor atualizado com sucesso!');
  };

  // Quick select from modal list for manual form
  const handleSelectFromExisting = (item: { name: string; cnpj: string }) => {
    setSupplierInputName(item.name);
    setSupplierInputCnpj(item.cnpj !== 'N/I' ? item.cnpj : '');
    setAddMode('manual');
  };

  // Confirm and execute delete
  const handleConfirmDelete = () => {
    if (!supplierToDelete) return;
    const targetKey = supplierToDelete.id || supplierToDelete.cnpj || supplierToDelete.name;
    const updated = removeManualRecurringSupplier(targetKey);
    setManualList(updated);
    setSupplierToDelete(null);
    showToast(`"${supplierToDelete.name}" foi removido da lista.`);
  };

  // Reset list
  const handleConfirmReset = () => {
    const reset = clearAllManualRecurringSuppliers(true);
    setManualList(reset);
    setIsResetConfirmOpen(false);
    showToast('Lista restaurada para o padrão inicial.');
  };

  const [isExporting, setIsExporting] = useState(false);

  // Export to Excel
  const handleExportExcel = () => {
    if (!displaySuppliers.length || isExporting) return;

    setIsExporting(true);
    setTimeout(() => {
      try {
        const dataToExport = displaySuppliers.map((s) => ({
          FORNECEDOR: s.name,
          CNPJ: s.cnpj,
          OBSERVAÇÕES: s.notes || '-',
          'BASE DE PPS (R$)': (s.basePpsValue || s.ppProntasValue || 0),
          'LIQUIDAÇÃO OBEDECE (R$)': s.liqObedeceValue || 0,
          'LIQUIDAÇÃO NÃO OBEDECE (R$)': s.liqNaoObedeceValue || 0,
          'VALOR TOTAL (R$)': s.totalValue,
        }));

        const worksheet = XLSX.utils.json_to_sheet(dataToExport);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Fornecedores_Recorrentes');
        XLSX.writeFile(
          workbook,
          `Fornecedores_Recorrentes_${new Date().toISOString().slice(0, 10)}.xlsx`
        );
      } finally {
        setIsExporting(false);
      }
    }, 200);
  };

  const isAllFilteredSelected =
    filteredModalCreditors.length > 0 &&
    filteredModalCreditors.every((c) => selectedCreditorKeys.has(`${c.cnpj || 'NI'}__${c.name}`));

  return (
    <div className="space-y-6">

      {/* TOAST NOTIFICATION */}
      {toastMessage && (
        <div className="fixed bottom-6 right-6 z-50 bg-slate-900 text-white px-4 py-3 rounded-xl shadow-xl flex items-center gap-2.5 text-xs font-semibold animate-in fade-in slide-in-from-bottom-2">
          <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0" />
          <span>{toastMessage}</span>
        </div>
      )}
      
      {/* Title & Info */}
      <div className="flex items-center gap-2.5">
        <h2 className="text-xl font-black text-slate-900 tracking-tight">
          Fornecedores Recorrentes
        </h2>
        <InfoHelpButton
          title="Fornecedores Recorrentes"
          variant="light"
          content={
            <p>
              Gestão e acompanhamento contínuo dos prestadores de serviços, concessionárias e contratos permanentes da SESAP/FES, permitindo monitorar movimentações e saldos nas etapas de liquidação.
            </p>
          }
        />
      </div>

      {/* 1. TOP SUMMARY METRICS */}
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
        
        {/* Total Cadastrados */}
        <div className="bg-white rounded-2xl p-5 border border-slate-200/90 shadow-2xs">
          <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider block">
            Fornecedores Cadastrados
          </span>
          <div className="text-2xl font-black text-slate-900 font-mono mt-1">
            {summary.totalRegistered}
          </div>
        </div>

        {/* Com Movimentação no Período */}
        <div className="bg-white rounded-2xl p-5 border border-emerald-200/90 shadow-2xs">
          <span className="text-[11px] font-bold text-emerald-800 uppercase tracking-wider block">
            Com Movimentação nos Arquivos
          </span>
          <div className="text-2xl font-black text-emerald-950 font-mono mt-1">
            {summary.totalWithTransactions}{' '}
            <span className="text-xs font-semibold text-emerald-700">de {summary.totalRegistered}</span>
          </div>
        </div>

        {/* Valor Total Recorrente */}
        <div className="bg-white rounded-2xl p-5 border border-slate-200/90 shadow-2xs">
          <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider block">
            Valor Total
          </span>
          <div
            className="text-xl font-black text-slate-900 font-mono mt-1 truncate"
            title={formatBRL(summary.totalRecurringValue)}
          >
            {formatBRL(summary.totalRecurringValue)}
          </div>
        </div>

      </div>

      {/* 2. MAIN TABLE CARD */}
      <div className="bg-white rounded-2xl p-6 shadow-xs border border-slate-200/90 space-y-5">
        
        {/* Controls Bar */}
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          
          {/* Search input */}
          <div className="relative flex-1 max-w-md">
            <Search className="w-4 h-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2 pointer-events-none" />
            <input
              type="text"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              placeholder="Pesquisar fornecedor cadastrado ou CNPJ..."
              className="w-full pl-9 pr-4 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold text-slate-800 placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-emerald-500 focus:bg-white transition"
            />
          </div>

          {/* Action buttons & Filters */}
          <div className="flex flex-wrap items-center gap-2">
            
            {/* Filter Toggle */}
            <button
              type="button"
              onClick={() => setFilterActiveOnly(!filterActiveOnly)}
              className={`px-3 py-2 text-xs font-bold rounded-xl border transition cursor-pointer active:scale-98 ${
                filterActiveOnly
                  ? 'bg-emerald-600 text-white border-emerald-600 shadow-2xs'
                  : 'bg-slate-100 text-slate-700 border-slate-200 hover:bg-slate-200'
              }`}
            >
              {filterActiveOnly ? 'Apenas com Movimentação' : 'Todos os Cadastrados'}
            </button>

            {/* Page size selector */}
            <div className="flex items-center gap-1.5 text-xs text-slate-600 font-bold bg-slate-50 px-2.5 py-2 rounded-xl border border-slate-200">
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

            {/* Reset / Restore Defaults */}
            <button
              type="button"
              onClick={() => setIsResetConfirmOpen(true)}
              className="px-3 py-2 bg-white hover:bg-slate-50 text-slate-600 border border-slate-300 text-xs font-bold rounded-xl shadow-2xs transition flex items-center gap-1.5 cursor-pointer active:scale-98"
              title="Restaurar lista padrão"
            >
              <RotateCcw className="w-3.5 h-3.5 text-slate-500" />
              <span className="hidden sm:inline">Restaurar</span>
            </button>

            {/* Export button */}
            <button
              type="button"
              onClick={handleExportExcel}
              disabled={isExporting}
              className="px-3.5 py-2 bg-white hover:bg-slate-50 text-slate-700 border border-slate-300 text-xs font-bold rounded-xl shadow-2xs transition flex items-center gap-1.5 cursor-pointer active:scale-98 disabled:opacity-75 disabled:cursor-wait"
              title="Exportar para Excel"
            >
              {isExporting ? (
                <>
                  <Loader2 className="w-3.5 h-3.5 animate-spin text-emerald-700" />
                  <span>Exportando...</span>
                </>
              ) : (
                <>
                  <Download className="w-3.5 h-3.5 text-emerald-700" />
                  <span>Exportar</span>
                </>
              )}
            </button>

            {/* Add Manual Supplier Button */}
            <button
              type="button"
              onClick={() => {
                setFormError(null);
                setSelectedCreditorKeys(new Set());
                setAddMode(availableCreditorsFromFiles.length > 0 ? 'multi' : 'manual');
                setIsAddModalOpen(true);
              }}
              className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold rounded-xl shadow-md shadow-emerald-600/20 transition flex items-center gap-1.5 cursor-pointer active:scale-98"
            >
              <Plus className="w-4 h-4" />
              <span>Adicionar Fornecedor</span>
            </button>

          </div>
        </div>

        {/* Recurrence Table */}
        <div className="border border-slate-200/90 rounded-xl overflow-hidden shadow-2xs">
          <div className="overflow-x-auto overflow-y-auto max-h-[600px]">
            <table className="w-full text-left border-collapse text-xs">
              <thead className="sticky top-0 z-10 bg-slate-100/95 backdrop-blur-xs shadow-2xs">
                <tr className="border-b border-slate-200 text-slate-700 font-bold uppercase tracking-wider text-[11px]">
                  <th className="py-3 px-3 w-12 text-center text-slate-400">#</th>

                  {/* Fornecedor */}
                  <th className="py-3 px-4 min-w-[280px]">
                    <button
                      type="button"
                      onClick={() => toggleSort('name')}
                      className="flex items-center gap-1 hover:text-emerald-700 cursor-pointer font-extrabold active:scale-98"
                    >
                      <span>Fornecedor / CNPJ</span>
                      <ArrowUpDown className="w-3 h-3 text-slate-400" />
                    </button>
                  </th>

                  {/* Observação / Notas */}
                  <th className="py-3 px-3 min-w-[200px]">
                    <span>Observação / Contrato</span>
                  </th>

                  {/* Base de PPs */}
                  <th className="py-3 px-3.5 text-right min-w-[140px] whitespace-nowrap">
                    <button
                      type="button"
                      onClick={() => toggleSort('ppProntasValue')}
                      className="flex items-center justify-end gap-1 w-full hover:text-emerald-700 cursor-pointer font-bold text-slate-700 active:scale-98 whitespace-nowrap"
                    >
                      <span>Base de PPs</span>
                      <ArrowUpDown className="w-3 h-3 text-slate-400" />
                    </button>
                  </th>

                  {/* Liquidado - Obedece */}
                  <th className="py-3 px-3.5 text-right min-w-[155px] whitespace-nowrap">
                    <button
                      type="button"
                      onClick={() => toggleSort('liqObedeceValue')}
                      className="flex items-center justify-end gap-1 w-full hover:text-emerald-700 cursor-pointer font-bold text-slate-700 active:scale-98 whitespace-nowrap"
                    >
                      <span>Liquidado - Obedece</span>
                      <ArrowUpDown className="w-3 h-3 text-slate-400" />
                    </button>
                  </th>

                  {/* Liquidado - Não Obedece */}
                  <th className="py-3 px-3.5 text-right min-w-[165px] whitespace-nowrap">
                    <button
                      type="button"
                      onClick={() => toggleSort('liqNaoObedeceValue')}
                      className="flex items-center justify-end gap-1 w-full hover:text-emerald-700 cursor-pointer font-bold text-slate-700 active:scale-98 whitespace-nowrap"
                    >
                      <span>Liquidado - Não Obedece</span>
                      <ArrowUpDown className="w-3 h-3 text-slate-400" />
                    </button>
                  </th>

                  {/* Valor Total */}
                  <th className="py-3 px-4 text-right min-w-[150px] whitespace-nowrap">
                    <button
                      type="button"
                      onClick={() => toggleSort('totalValue')}
                      className="flex items-center justify-end gap-1 w-full hover:text-emerald-700 cursor-pointer font-extrabold text-slate-900 active:scale-98 whitespace-nowrap"
                    >
                      <span>Valor Total</span>
                      <ArrowUpDown className="w-3 h-3 text-slate-400" />
                    </button>
                  </th>

                  {/* Ações */}
                  <th className="py-3 px-3 text-center min-w-[130px] whitespace-nowrap">
                    <span>Ações</span>
                  </th>
                </tr>
              </thead>

              <tbody className="divide-y divide-slate-100 font-medium">
                {paginatedSuppliers.length === 0 ? (
                  <tr>
                    <td colSpan={8} className="py-12 text-center text-slate-400 italic">
                      Nenhum fornecedor recorrente cadastrado ou localizado. Clique em "Adicionar Fornecedor" para cadastrar.
                    </td>
                  </tr>
                ) : (
                  paginatedSuppliers.map((supplier, idx) => {
                    const realIndex = pageSize === 0 ? idx + 1 : (currentPage - 1) * pageSize + idx + 1;
                    return (
                      <tr
                        key={`rec_row_${supplier.id || `${supplier.cnpj || 'NI'}_${supplier.name}_${idx}`}`}
                        className="hover:bg-slate-50/80 transition-colors"
                      >
                        {/* Index */}
                        <td className="py-3 px-3 text-center font-mono text-[11px] text-slate-400 border-r border-slate-100 whitespace-nowrap">
                          {realIndex}
                        </td>

                        {/* Fornecedor e CNPJ */}
                        <td className="py-3 px-4 border-r border-slate-100">
                          <div className="space-y-0.5">
                            <span
                              className="font-extrabold text-slate-900 block truncate max-w-xs sm:max-w-md text-xs"
                              title={supplier.name}
                            >
                              {supplier.name}
                            </span>
                            <span className="text-[11px] font-mono text-slate-500 whitespace-nowrap">
                              {supplier.cnpj !== 'N/I' ? supplier.cnpj : 'CNPJ Não Informado'}
                            </span>
                          </div>
                        </td>

                        {/* Observações / Notas */}
                        <td className="py-3 px-3 border-r border-slate-100 text-slate-600 text-[11px]">
                          {supplier.notes ? (
                            <span className="bg-slate-100 text-slate-700 px-2 py-0.5 rounded-md border border-slate-200 line-clamp-2" title={supplier.notes}>
                              {supplier.notes}
                            </span>
                          ) : (
                            <span className="text-slate-300">-</span>
                          )}
                        </td>

                        {/* Base de PPs */}
                        <td className="py-3 px-3.5 text-right font-mono text-xs border-r border-slate-100 text-slate-700 whitespace-nowrap">
                          {(supplier.basePpsValue || supplier.ppProntasValue || 0) > 0 ? formatBRL(supplier.basePpsValue || supplier.ppProntasValue || 0) : '-'}
                        </td>

                        {/* Liq. Obedece */}
                        <td className="py-3 px-3.5 text-right font-mono text-xs border-r border-slate-100 text-slate-700 whitespace-nowrap">
                          {supplier.liqObedeceValue > 0 ? formatBRL(supplier.liqObedeceValue) : '-'}
                        </td>

                        {/* Liq. Não Obedece */}
                        <td className="py-3 px-3.5 text-right font-mono text-xs border-r border-slate-100 text-slate-700 whitespace-nowrap">
                          {supplier.liqNaoObedeceValue > 0 ? formatBRL(supplier.liqNaoObedeceValue) : '-'}
                        </td>

                        {/* Valor Total */}
                        <td className="py-3 px-4 text-right font-mono border-r border-slate-100 whitespace-nowrap">
                          <span className="text-xs font-black text-slate-950 whitespace-nowrap">
                            {formatBRL(supplier.totalValue)}
                          </span>
                        </td>

                        {/* Ações (Editar / Excluir) */}
                        <td className="py-3 px-3 text-center whitespace-nowrap">
                          <div className="flex items-center justify-center w-full gap-1.5 whitespace-nowrap">
                            {/* Edit Button */}
                            <button
                              type="button"
                              onClick={() => handleOpenEditModal(supplier)}
                              className="p-1.5 text-slate-400 hover:text-emerald-700 hover:bg-emerald-50 rounded-lg transition cursor-pointer active:scale-95 whitespace-nowrap"
                              title="Editar dados e observações do fornecedor"
                            >
                              <Pencil className="w-4 h-4 shrink-0" />
                            </button>

                            {/* Delete Button */}
                            <button
                              type="button"
                              onClick={() => {
                                setSupplierToDelete({
                                  id: supplier.id || '',
                                  name: supplier.name,
                                  cnpj: supplier.cnpj,
                                });
                              }}
                              className="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition cursor-pointer active:scale-95 whitespace-nowrap"
                              title="Excluir da lista de fornecedores recorrentes"
                            >
                              <Trash2 className="w-4 h-4 shrink-0" />
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
        </div>

        {/* Pagination Footer */}
        {totalPages > 1 && pageSize > 0 && (
          <div className="flex flex-col sm:flex-row items-center justify-between gap-3 text-xs text-slate-500 pt-2 font-medium">
            <span>
              Exibindo <strong>{paginatedSuppliers.length}</strong> de <strong>{displaySuppliers.length}</strong> fornecedores (Cadastrados: {manualList.length})
            </span>
            <div className="flex items-center gap-1.5">
              <button
                type="button"
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
                type="button"
                onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
                disabled={currentPage === totalPages}
                className="p-1.5 rounded-lg border border-slate-200 bg-slate-50 hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed transition cursor-pointer active:scale-95"
                title="Próxima Página"
              >
                <ChevronRight className="w-4 h-4" />
              </button>
            </div>
          </div>
        )}

      </div>

      {/* 3. MODAL: ADICIONAR FORNECEDOR COM SELEÇÃO MÚLTIPLA OU INDIVIDUAL */}
      {isAddModalOpen && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl max-w-3xl w-full p-6 shadow-2xl border border-slate-200 space-y-5 max-h-[90vh] flex flex-col">
            
            {/* Modal Header */}
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2.5">
                <div className="w-8 h-8 rounded-xl bg-emerald-100 text-emerald-700 flex items-center justify-center">
                  <Plus className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-base font-extrabold text-slate-900">
                    Adicionar Fornecedores Recorrentes
                  </h3>
                  <p className="text-xs text-slate-500">
                    Selecione múltiplos favorecidos das planilhas ou cadastre manualmente.
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setIsAddModalOpen(false)}
                className="p-1.5 text-slate-400 hover:text-slate-700 hover:bg-slate-100 rounded-xl transition cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Mode Switcher Tabs */}
            <div className="flex items-center gap-2 border-b border-slate-200 pb-3">
              <button
                type="button"
                onClick={() => setAddMode('multi')}
                className={`px-4 py-2 text-xs font-bold rounded-xl transition flex items-center gap-2 cursor-pointer ${
                  addMode === 'multi'
                    ? 'bg-emerald-600 text-white shadow-xs'
                    : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                }`}
              >
                <CheckSquare className="w-4 h-4" />
                <span>Seleção Múltipla das Planilhas ({availableCreditorsFromFiles.length})</span>
              </button>

              <button
                type="button"
                onClick={() => setAddMode('manual')}
                className={`px-4 py-2 text-xs font-bold rounded-xl transition flex items-center gap-2 cursor-pointer ${
                  addMode === 'manual'
                    ? 'bg-emerald-600 text-white shadow-xs'
                    : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                }`}
              >
                <UserPlus className="w-4 h-4" />
                <span>Cadastro Manual Individual</span>
              </button>
            </div>

            {/* MULTI-SELECT MODE */}
            {addMode === 'multi' && (
              <div className="space-y-4 flex-1 flex flex-col min-h-0">
                {availableCreditorsFromFiles.length === 0 ? (
                  <div className="p-8 text-center bg-slate-50 rounded-xl border border-dashed border-slate-200">
                    <Building2 className="w-8 h-8 text-slate-300 mx-auto mb-2" />
                    <p className="text-xs font-bold text-slate-600">
                      Nenhuma planilha importada no momento.
                    </p>
                    <p className="text-xs text-slate-400 mt-1">
                      Você pode usar a aba "Cadastro Manual Individual" para adicionar fornecedores.
                    </p>
                  </div>
                ) : (
                  <>
                    {/* Filter and Select All Bar */}
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2.5">
                      <div className="relative flex-1">
                        <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                        <input
                          type="text"
                          value={modalSearchFilter}
                          onChange={(e) => setModalSearchFilter(e.target.value)}
                          placeholder="Buscar por Favorecido ou CNPJ..."
                          className="w-full pl-8 pr-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold focus:outline-none focus:ring-2 focus:ring-emerald-500 focus:bg-white"
                        />
                      </div>

                      <div className="flex items-center gap-2 shrink-0">
                        <button
                          type="button"
                          onClick={handleToggleSelectAllFiltered}
                          className="px-3 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold rounded-xl transition cursor-pointer flex items-center gap-1.5"
                        >
                          {isAllFilteredSelected ? (
                            <>
                              <Square className="w-3.5 h-3.5 text-slate-600" />
                              <span>Desmarcar Filtrados</span>
                            </>
                          ) : (
                            <>
                              <CheckSquare className="w-3.5 h-3.5 text-emerald-700" />
                              <span>Selecionar Todos ({filteredModalCreditors.length})</span>
                            </>
                          )}
                        </button>

                        {selectedCreditorKeys.size > 0 && (
                          <button
                            type="button"
                            onClick={() => setSelectedCreditorKeys(new Set())}
                            className="px-2.5 py-2 text-rose-600 hover:bg-rose-50 text-xs font-bold rounded-xl transition cursor-pointer"
                          >
                            Limpar ({selectedCreditorKeys.size})
                          </button>
                        )}
                      </div>
                    </div>

                    {/* Suppliers List with Checkboxes */}
                    <div className="flex-1 overflow-y-auto max-h-72 border border-slate-200 rounded-xl divide-y divide-slate-100 bg-white">
                      {filteredModalCreditors.length === 0 ? (
                        <div className="p-6 text-center text-xs text-slate-400 italic">
                          Nenhum favorecido encontrado com o filtro "{modalSearchFilter}".
                        </div>
                      ) : (
                        filteredModalCreditors.map((c, i) => {
                          const key = `${c.cnpj || 'NI'}__${c.name}`;
                          const isSelected = selectedCreditorKeys.has(key);
                          const isAlreadyRecurring = manualList.some(
                            (m) =>
                              (m.cnpj !== 'N/I' && m.cnpj === c.cnpj) ||
                              m.name.trim().toLowerCase() === c.name.trim().toLowerCase()
                          );

                          return (
                            <div
                              key={`modal_cred_${i}`}
                              onClick={() => toggleSelectCreditor(c)}
                              className={`px-3.5 py-2.5 flex items-center justify-between text-xs transition cursor-pointer select-none ${
                                isSelected
                                  ? 'bg-emerald-50/90 border-l-4 border-emerald-600'
                                  : 'hover:bg-slate-50'
                              }`}
                            >
                              <div className="flex items-center gap-3 min-w-0 flex-1 pr-3">
                                <input
                                  type="checkbox"
                                  checked={isSelected}
                                  onChange={() => {}} // Controlled by container onClick
                                  className="w-4 h-4 rounded text-emerald-600 focus:ring-emerald-500 cursor-pointer shrink-0 accent-emerald-600"
                                />

                                <div className="min-w-0 flex-1">
                                  <div className="flex items-center gap-2">
                                    <span className="font-extrabold text-slate-900 truncate block">
                                      {c.name}
                                    </span>
                                    {isAlreadyRecurring && (
                                      <span className="text-[10px] font-bold bg-amber-100 text-amber-800 px-1.5 py-0.2 rounded border border-amber-200 shrink-0">
                                        Já cadastrado
                                      </span>
                                    )}
                                  </div>
                                  <div className="flex items-center gap-3 text-[11px] text-slate-500 font-mono mt-0.5">
                                    <span>{c.cnpj || 'CNPJ Não Informado'}</span>
                                    {c.totalValue > 0 && (
                                      <span className="text-slate-700 font-bold">
                                        {formatBRL(c.totalValue)}
                                      </span>
                                    )}
                                  </div>
                                </div>
                              </div>

                              <button
                                type="button"
                                onClick={(e) => {
                                  e.stopPropagation();
                                  handleSelectFromExisting(c);
                                }}
                                className="text-[11px] font-bold text-slate-400 hover:text-emerald-700 px-2 py-1 rounded hover:bg-slate-100 shrink-0 transition"
                                title="Preencher no formulário manual"
                              >
                                Editar manual &rarr;
                              </button>
                            </div>
                          );
                        })
                      )}
                    </div>

                    {/* Bottom Status & Multi-Add Action */}
                    <div className="flex items-center justify-between pt-2 border-t border-slate-100">
                      <span className="text-xs font-semibold text-slate-600">
                        <strong>{selectedCreditorKeys.size}</strong> fornecedor(es) selecionado(s)
                      </span>

                      <div className="flex items-center gap-2">
                        <button
                          type="button"
                          onClick={() => setIsAddModalOpen(false)}
                          className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold rounded-xl transition cursor-pointer"
                        >
                          Cancelar
                        </button>
                        <button
                          type="button"
                          onClick={handleAddMultipleSelected}
                          disabled={selectedCreditorKeys.size === 0}
                          className={`px-5 py-2 text-xs font-bold rounded-xl shadow-md transition flex items-center gap-1.5 cursor-pointer ${
                            selectedCreditorKeys.size > 0
                              ? 'bg-emerald-600 hover:bg-emerald-500 text-white shadow-emerald-600/20'
                              : 'bg-slate-200 text-slate-400 cursor-not-allowed'
                          }`}
                        >
                          <Check className="w-4 h-4" />
                          <span>
                            Adicionar {selectedCreditorKeys.size > 0 ? `(${selectedCreditorKeys.size}) ` : ''}aos Recorrentes
                          </span>
                        </button>
                      </div>
                    </div>
                  </>
                )}
              </div>
            )}

            {/* MANUAL SINGLE ENTRY MODE */}
            {addMode === 'manual' && (
              <form onSubmit={handleSaveNewSupplier} className="space-y-4">
                {formError && (
                  <div className="p-3 bg-rose-50 border border-rose-200 text-rose-800 text-xs font-semibold rounded-xl">
                    {formError}
                  </div>
                )}

                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-700">
                    Nome / Razão Social do Fornecedor *
                  </label>
                  <input
                    type="text"
                    value={supplierInputName}
                    onChange={(e) => setSupplierInputName(e.target.value)}
                    placeholder="Ex: WHITE MARTINS GASES INDUSTRIAIS"
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-200 rounded-xl text-xs font-semibold text-slate-900 focus:outline-none focus:ring-2 focus:ring-emerald-500"
                  />
                </div>

                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-700">
                    CNPJ (Opcional, recomendado para cruzamento exato)
                  </label>
                  <input
                    type="text"
                    value={supplierInputCnpj}
                    onChange={(e) => setSupplierInputCnpj(e.target.value)}
                    placeholder="Ex: 35.820.448/0001-30"
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-200 rounded-xl text-xs font-semibold text-slate-900 focus:outline-none focus:ring-2 focus:ring-emerald-500 font-mono"
                  />
                </div>

                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-700">
                    Observações / Tipo de Contrato (Opcional)
                  </label>
                  <input
                    type="text"
                    value={supplierInputNotes}
                    onChange={(e) => setSupplierInputNotes(e.target.value)}
                    placeholder="Ex: Contrato continuado de oxigênio / locação"
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-200 rounded-xl text-xs font-semibold text-slate-900 focus:outline-none focus:ring-2 focus:ring-emerald-500"
                  />
                </div>

                {/* Modal Buttons */}
                <div className="flex items-center justify-end gap-2 pt-3 border-t border-slate-100">
                  <button
                    type="button"
                    onClick={() => setIsAddModalOpen(false)}
                    className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold rounded-xl transition cursor-pointer"
                  >
                    Cancelar
                  </button>
                  <button
                    type="submit"
                    className="px-5 py-2 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold rounded-xl shadow-md shadow-emerald-600/20 transition cursor-pointer"
                  >
                    Salvar Fornecedor
                  </button>
                </div>
              </form>
            )}

          </div>
        </div>
      )}

      {/* 4. MODAL: EDITAR FORNECEDOR RECORRENTE */}
      {editingSupplier && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl max-w-lg w-full p-6 shadow-2xl border border-slate-200 space-y-5 flex flex-col">
            {/* Modal Header */}
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2">
                <div className="w-8 h-8 rounded-xl bg-blue-100 text-blue-700 flex items-center justify-center">
                  <Pencil className="w-4 h-4" />
                </div>
                <h3 className="text-base font-extrabold text-slate-900">
                  Editar Fornecedor Recorrente
                </h3>
              </div>
              <button
                type="button"
                onClick={() => setEditingSupplier(null)}
                className="p-1.5 text-slate-400 hover:text-slate-700 hover:bg-slate-100 rounded-xl transition cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Edit Form */}
            <form onSubmit={handleSaveEdit} className="space-y-4">
              {editError && (
                <div className="p-3 bg-rose-50 border border-rose-200 text-rose-800 text-xs font-semibold rounded-xl">
                  {editError}
                </div>
              )}

              <div>
                <label className="text-[11px] font-bold text-slate-700 uppercase tracking-wider block mb-1.5">
                  Nome / Razão Social <span className="text-rose-500">*</span>
                </label>
                <input
                  type="text"
                  value={editName}
                  onChange={(e) => setEditName(e.target.value)}
                  className="w-full px-3.5 py-2.5 bg-white border border-slate-200 rounded-xl text-xs font-semibold text-slate-900 focus:outline-none focus:ring-2 focus:ring-emerald-500"
                  required
                />
              </div>

              <div>
                <label className="text-[11px] font-bold text-slate-700 uppercase tracking-wider block mb-1.5">
                  CNPJ
                </label>
                <input
                  type="text"
                  value={editCnpj}
                  onChange={(e) => setEditCnpj(e.target.value)}
                  placeholder="00.000.000/0000-00"
                  className="w-full px-3.5 py-2.5 bg-white border border-slate-200 rounded-xl text-xs font-mono font-bold text-slate-900 focus:outline-none focus:ring-2 focus:ring-emerald-500"
                />
              </div>

              <div>
                <label className="text-[11px] font-bold text-slate-700 uppercase tracking-wider block mb-1.5">
                  Observações / Tipo de Contrato
                </label>
                <input
                  type="text"
                  value={editNotes}
                  onChange={(e) => setEditNotes(e.target.value)}
                  placeholder="Ex: Fornecimento contínuo de gases"
                  className="w-full px-3.5 py-2.5 bg-white border border-slate-200 rounded-xl text-xs font-semibold text-slate-900 focus:outline-none focus:ring-2 focus:ring-emerald-500"
                />
              </div>

              {/* Modal Buttons */}
              <div className="flex items-center justify-end gap-2 pt-3 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setEditingSupplier(null)}
                  className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold rounded-xl transition cursor-pointer"
                >
                  Cancelar
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold rounded-xl shadow-md shadow-emerald-600/20 transition cursor-pointer"
                >
                  Salvar Alterações
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* 5. MODAL: CONFIRMAÇÃO DE EXCLUSÃO */}
      {supplierToDelete && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl max-w-md w-full p-6 shadow-2xl border border-slate-200 space-y-4">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-rose-100 text-rose-600 flex items-center justify-center shrink-0">
                <AlertTriangle className="w-5 h-5" />
              </div>
              <div>
                <h3 className="text-base font-extrabold text-slate-900">
                  Excluir Fornecedor
                </h3>
                <p className="text-xs text-slate-500 mt-0.5">
                  Deseja remover este fornecedor da lista de recorrentes?
                </p>
              </div>
            </div>

            <div className="bg-slate-50 p-3.5 rounded-xl border border-slate-200 text-xs space-y-1">
              <div className="font-extrabold text-slate-900">{supplierToDelete.name}</div>
              <div className="text-slate-500 font-mono">{supplierToDelete.cnpj}</div>
            </div>

            <div className="flex items-center justify-end gap-2 pt-2">
              <button
                type="button"
                onClick={() => setSupplierToDelete(null)}
                className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold rounded-xl transition cursor-pointer"
              >
                Cancelar
              </button>
              <button
                type="button"
                onClick={handleConfirmDelete}
                className="px-5 py-2 bg-rose-600 hover:bg-rose-500 text-white text-xs font-bold rounded-xl shadow-md shadow-rose-600/20 transition cursor-pointer"
              >
                Sim, Excluir
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 6. MODAL: CONFIRMAÇÃO DE RESTAURAÇÃO */}
      {isResetConfirmOpen && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl max-w-md w-full p-6 shadow-2xl border border-slate-200 space-y-4">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-amber-100 text-amber-600 flex items-center justify-center shrink-0">
                <RotateCcw className="w-5 h-5" />
              </div>
              <div>
                <h3 className="text-base font-extrabold text-slate-900">
                  Restaurar Lista Padrão
                </h3>
                <p className="text-xs text-slate-500 mt-0.5">
                  Isso irá resetar os fornecedores recorrentes para a configuração inicial.
                </p>
              </div>
            </div>

            <div className="flex items-center justify-end gap-2 pt-2">
              <button
                type="button"
                onClick={() => setIsResetConfirmOpen(false)}
                className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold rounded-xl transition cursor-pointer"
              >
                Cancelar
              </button>
              <button
                type="button"
                onClick={handleConfirmReset}
                className="px-5 py-2 bg-amber-600 hover:bg-amber-500 text-white text-xs font-bold rounded-xl shadow-md shadow-amber-600/20 transition cursor-pointer"
              >
                Confirmar Restauração
              </button>
            </div>
          </div>
        </div>
      )}

    </div>
  );
};
