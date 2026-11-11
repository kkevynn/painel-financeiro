import React, { useState, useMemo, useEffect } from 'react';
import {
  X,
  Building,
  FileSpreadsheet,
  Download,
  FileText,
  Printer,
  Search,
  Filter,
  ArrowUpDown,
  CheckCircle2,
  Clock,
  AlertCircle,
  Repeat,
  Sparkles,
  ChevronLeft,
  ChevronRight,
  ShieldCheck,
  Check,
} from 'lucide-react';
import { FileData, QuebraOrdemItem } from '../types';
import { formatBRL } from '../utils/excelParser';
import {
  buildCompanyAuditSummary,
  exportCompanyAuditToExcel,
  exportCompanyAuditToPdf,
  formatCleanNL,
  BankAuditRecord,
} from '../utils/auditCompanyService';
import { OrdemBancariaItem } from '../utils/ordemBancariaParser';
import { paraPublicarService } from '../utils/paraPublicarService';
import { resolveCnpj, resolveSupplierName } from '../utils/appKnowledgeDirectory';

interface ExtratoAuditoriaEmpresaModalProps {
  supplier: { name: string; cnpj: string } | null;
  onClose: () => void;
  files: FileData[];
  quebraOrdemItems?: QuebraOrdemItem[];
  bankRecords?: BankAuditRecord[];
  ordensBancariasList?: OrdemBancariaItem[] | null;
  isRecurring?: boolean;
}

export const ExtratoAuditoriaEmpresaModal: React.FC<ExtratoAuditoriaEmpresaModalProps> = ({
  supplier,
  onClose,
  files,
  quebraOrdemItems,
  bankRecords,
  ordensBancariasList,
  isRecurring,
}) => {
  const [activeFilterStage, setActiveFilterStage] = useState<string>('all');
  const [searchTerm, setSearchTerm] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(50);
  const [, setPublicarTick] = useState(0);

  // Prevent background scrolling when modal is open
  useEffect(() => {
    const originalOverflow = document.body.style.overflow;
    document.body.style.overflow = 'hidden';
    return () => {
      document.body.style.overflow = originalOverflow;
    };
  }, []);

  // Subscribe to changes in paraPublicarService
  useEffect(() => {
    const unsub = paraPublicarService.subscribe(() => {
      setPublicarTick((t) => t + 1);
    });
    return () => unsub();
  }, []);

  // Close on Escape key
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [onClose]);

  // Compute the audit summary
  const summary = useMemo(() => {
    if (!supplier) return null;
    return buildCompanyAuditSummary(supplier, files, quebraOrdemItems, bankRecords, ordensBancariasList);
  }, [supplier, files, quebraOrdemItems, bankRecords, ordensBancariasList]);

  // Filter rows
  const filteredRows = useMemo(() => {
    if (!summary) return [];

    return summary.allRows.filter((r) => {
      if (activeFilterStage === 'ordens_bancarias') {
        const isOB =
          r.sourceType === 'status_banco' ||
          r.sourceType === 'pagamentos_emitidos' ||
          r.sourceType === 'ordem_bancaria' ||
          (r.status && (r.status.includes('PPCB') || r.status.includes('PAGA') || r.status.includes('PAGO') || r.status.includes('CB')));
        if (!isOB) return false;
      } else if (activeFilterStage !== 'all' && r.sourceType !== activeFilterStage) {
        return false;
      }

      const q = searchTerm.toLowerCase().trim();
      if (!q) return true;

      const matchesDoc = r.documentNumber.toLowerCase().includes(q);
      const matchesFile = r.sourceFileName.toLowerCase().includes(q);
      const matchesStatus = (r.status || '').toLowerCase().includes(q);
      const matchesFonte = (r.fonteRecurso || '').toLowerCase().includes(q);
      const matchesDetails = (r.details || '').toLowerCase().includes(q);

      return matchesDoc || matchesFile || matchesStatus || matchesFonte || matchesDetails;
    });
  }, [summary, activeFilterStage, searchTerm]);

  // Reset page when filter changes
  useEffect(() => {
    setCurrentPage(1);
  }, [activeFilterStage, searchTerm, pageSize]);

  // Paginated rows
  const totalPages = pageSize === 0 ? 1 : Math.ceil(filteredRows.length / pageSize) || 1;
  const paginatedRows = useMemo(() => {
    if (pageSize === 0) return filteredRows;
    const start = (currentPage - 1) * pageSize;
    return filteredRows.slice(start, start + pageSize);
  }, [filteredRows, currentPage, pageSize]);

  if (!supplier || !summary) return null;

  const handleExport = () => {
    exportCompanyAuditToExcel(summary);
  };

  const handlePrint = () => {
    window.print();
  };

  const getSourceBadgeStyle = (sourceType: string, status?: string) => {
    if (status && (status.includes('PPCB') || status.includes('PPNC') || status.includes('Enviada a Banco'))) {
      return 'bg-blue-100 text-blue-950 border-blue-400 font-extrabold';
    }
    if (status && (status.includes('AO') || status.includes('PP Pronta'))) {
      return 'bg-emerald-100 text-emerald-950 border-emerald-400 font-extrabold';
    }
    switch (sourceType) {
      case 'base_pps':
      case 'pp_prontas':
        return 'bg-emerald-100 text-emerald-900 border-emerald-300';
      case 'obedece':
        return 'bg-blue-100 text-blue-900 border-blue-300';
      case 'nao_obedece':
        return 'bg-amber-100 text-amber-900 border-amber-300';
      case 'pagamentos_emitidos':
        return 'bg-purple-100 text-purple-900 border-purple-300';
      case 'status_banco':
        return 'bg-teal-100 text-teal-900 border-teal-300';
      default:
        return 'bg-slate-100 text-slate-800 border-slate-200';
    }
  };

  return (
    <div className="fixed inset-0 z-50 overflow-y-auto overscroll-contain bg-black/60 backdrop-blur-xs flex items-center justify-center p-3 sm:p-6 animate-in fade-in-50 duration-200">
      <div
        className="bg-white rounded-3xl shadow-2xl border border-slate-200 w-full max-w-6xl max-h-[92vh] flex flex-col overflow-hidden overscroll-contain animate-in zoom-in-95 duration-200"
        role="dialog"
        aria-modal="true"
      >
        {/* Modal Header */}
        <div className="bg-[#061d15] text-white p-6 border-b border-emerald-900 flex items-start justify-between gap-4 shrink-0">
          <div className="flex items-start space-x-3.5">
            <div className="w-11 h-11 rounded-2xl bg-emerald-600/30 text-emerald-400 border border-emerald-500/30 flex items-center justify-center shrink-0 mt-0.5">
              <Building className="w-6 h-6" />
            </div>
            <div>
              <div className="flex items-center gap-2 flex-wrap">
                <span className="text-[10px] font-black uppercase tracking-wider px-2 py-0.5 rounded bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">
                  Extrato Geral de Auditoria
                </span>
                {isRecurring && (
                  <span className="text-[10px] font-black uppercase tracking-wider px-2 py-0.5 rounded bg-amber-500/20 text-amber-300 border border-amber-500/30 flex items-center gap-1">
                    <Repeat className="w-2.5 h-2.5" />
                    <span>Fornecedor Recorrente</span>
                  </span>
                )}
              </div>
              <h2 className="text-xl sm:text-2xl font-black tracking-tight text-white mt-1">
                {summary.supplierName}
              </h2>
              <p className="text-xs text-emerald-200/80 font-mono mt-0.5">
                CNPJ/CPF: <strong className="text-white">{summary.cnpj && summary.cnpj !== 'N/I' ? summary.cnpj : (resolveCnpj(summary.supplierName) || 'Não informado')}</strong>
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2 shrink-0">
            <button
              type="button"
              onClick={() => exportCompanyAuditToPdf(summary)}
              className="px-3.5 py-2 bg-red-600 hover:bg-red-500 text-white text-xs font-bold rounded-xl transition flex items-center gap-1.5 cursor-pointer shadow-md shadow-red-600/20 active:scale-95"
              title="Gerar Relatório em PDF com divisão mensal de pagamentos"
            >
              <FileText className="w-3.5 h-3.5" />
              <span className="hidden sm:inline">Exportar PDF</span>
            </button>
            <button
              type="button"
              onClick={handleExport}
              className="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold rounded-xl transition flex items-center gap-1.5 cursor-pointer shadow-md shadow-emerald-600/20 active:scale-95"
              title="Exportar Extrato Geral para Excel"
            >
              <Download className="w-3.5 h-3.5" />
              <span className="hidden sm:inline">Exportar Excel</span>
            </button>
            <button
              type="button"
              onClick={handlePrint}
              className="p-2 bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white rounded-xl transition cursor-pointer active:scale-95 hidden sm:flex"
              title="Imprimir extrato"
            >
              <Printer className="w-4 h-4" />
            </button>
            <button
              type="button"
              onClick={onClose}
              className="p-2 text-slate-400 hover:text-white hover:bg-slate-800 rounded-xl transition cursor-pointer active:scale-95"
              title="Fechar (Esc)"
            >
              <X className="w-5 h-5" />
            </button>
          </div>
        </div>

        {/* Modal Scrollable Body */}
        <div className="overflow-y-auto flex-1 p-6 space-y-6">
          {/* Financial Summary Cards */}
          <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3">
            {/* Soma Bruta Geral */}
            <div className="bg-slate-900 text-white rounded-2xl p-3.5 space-y-1 shadow-xs border border-slate-800">
              <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider block">
                Soma Bruta Geral
              </span>
              <span className="text-sm sm:text-base font-black font-mono block text-emerald-400 break-words leading-tight" title={formatBRL(summary.totalConsolidatedValue)}>
                {formatBRL(summary.totalConsolidatedValue)}
              </span>
              <span className="text-[10px] text-slate-400 block">
                {summary.allRows.length} lançamentos
              </span>
            </div>

            {/* Base de PPs */}
            <div className="bg-emerald-50/60 rounded-2xl p-3.5 space-y-1 border border-emerald-200">
              <span className="text-[10px] font-bold text-emerald-800 uppercase tracking-wider block">
                Base de PPs
              </span>
              <span className="text-sm sm:text-base font-black font-mono block text-emerald-950 break-words leading-tight" title={formatBRL(summary.ppProntasAOValue || summary.ppProntasValue)}>
                {formatBRL(summary.ppProntasAOValue || summary.ppProntasValue)}
              </span>
              <span className="text-[10px] text-emerald-700 block">
                {summary.ppProntasAOCount || summary.ppProntasCount} registros
              </span>
            </div>

            {/* Enviadas a Banco (PPCB / PPNC) - se houver transição */}
            {(summary.enviadasBancoCount || 0) > 0 && (
              <div className="bg-blue-50/60 rounded-2xl p-3.5 space-y-1 border border-blue-300">
                <span className="text-[10px] font-bold text-blue-900 uppercase tracking-wider block">
                  Enviadas a Banco
                </span>
                <span className="text-sm sm:text-base font-black font-mono block text-blue-950 break-words leading-tight" title={formatBRL(summary.enviadasBancoValue || 0)}>
                  {formatBRL(summary.enviadasBancoValue || 0)}
                </span>
                <span className="text-[10px] text-blue-800 font-bold block">
                  {summary.enviadasBancoCount} (PPCB / PPNC)
                </span>
              </div>
            )}

            {/* Liquidado - Obedece */}
            <div className="bg-blue-50/60 rounded-2xl p-3.5 space-y-1 border border-blue-200">
              <span className="text-[10px] font-bold text-blue-800 uppercase tracking-wider block">
                Liquidado - Obedece
              </span>
              <span className="text-sm sm:text-base font-black font-mono block text-blue-950 break-words leading-tight" title={formatBRL(summary.liqObedeceValue)}>
                {formatBRL(summary.liqObedeceValue)}
              </span>
              <span className="text-[10px] text-blue-700 block">
                {summary.liqObedeceCount} registros
              </span>
            </div>

            {/* Liquidado - Não Obedece */}
            <div className="bg-amber-50/60 rounded-2xl p-3.5 space-y-1 border border-amber-200">
              <span className="text-[10px] font-bold text-amber-800 uppercase tracking-wider block">
                Liquidado - Não Obedece
              </span>
              <span className="text-sm sm:text-base font-black font-mono block text-amber-950 break-words leading-tight" title={formatBRL(summary.liqNaoObedeceValue)}>
                {formatBRL(summary.liqNaoObedeceValue)}
              </span>
              <span className="text-[10px] text-amber-700 block">
                {summary.liqNaoObedeceCount} registros
              </span>
            </div>

            {/* Quebra de Ordem (AO) - se houver */}
            {(summary.quebraOrdemCount || 0) > 0 && (
              <div className="bg-purple-50/60 rounded-2xl p-3.5 space-y-1 border border-purple-200">
                <span className="text-[10px] font-bold text-purple-800 uppercase tracking-wider block">
                  Quebra Ordem (AO)
                </span>
                <span className="text-sm sm:text-base font-black font-mono block text-purple-950 break-words leading-tight" title={formatBRL(summary.quebraOrdemValue)}>
                  {formatBRL(summary.quebraOrdemValue)}
                </span>
                <span className="text-[10px] text-purple-700 block">
                  {summary.quebraOrdemCount} OBs emitidas
                </span>
              </div>
            )}

            {/* Status Banco */}
            <div className="bg-teal-50/60 rounded-2xl p-3.5 space-y-1 border border-teal-200">
              <span className="text-[10px] font-bold text-teal-800 uppercase tracking-wider block">
                Status Banco
              </span>
              <span className="text-sm sm:text-base font-black font-mono block text-teal-950 break-words leading-tight">
                {summary.bancoRecordsCount} reg.
              </span>
              <span className="text-[10px] text-teal-700 block">
                {summary.bancoConfirmedCount} confirmados
              </span>
            </div>
          </div>

          {/* Filter Bar */}
          <div className="bg-slate-50 p-4 rounded-2xl border border-slate-200 flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-3">
            {/* Stage filter pills */}
            <div className="flex items-center gap-1.5 overflow-x-auto pb-1 sm:pb-0">
              <button
                type="button"
                onClick={() => setActiveFilterStage('all')}
                className={`px-3 py-1.5 rounded-xl text-xs font-bold transition cursor-pointer shrink-0 ${
                  activeFilterStage === 'all'
                    ? 'bg-slate-900 text-white shadow-xs'
                    : 'bg-white text-slate-600 hover:text-slate-900 border border-slate-200'
                }`}
              >
                Todas ({summary.allRows.length})
              </button>

              <button
                type="button"
                onClick={() => setActiveFilterStage('obedece')}
                className={`px-3 py-1.5 rounded-xl text-xs font-bold transition cursor-pointer shrink-0 ${
                  activeFilterStage === 'obedece'
                    ? 'bg-blue-700 text-white shadow-xs'
                    : 'bg-white text-blue-800 hover:bg-blue-50 border border-blue-200'
                }`}
              >
                Liquidações - Obedece ({summary.liqObedeceCount})
              </button>

              <button
                type="button"
                onClick={() => setActiveFilterStage('nao_obedece')}
                className={`px-3 py-1.5 rounded-xl text-xs font-bold transition cursor-pointer shrink-0 ${
                  activeFilterStage === 'nao_obedece'
                    ? 'bg-amber-700 text-white shadow-xs'
                    : 'bg-white text-amber-800 hover:bg-amber-50 border border-amber-200'
                }`}
              >
                Liquidações - Não Obedece ({summary.liqNaoObedeceCount})
              </button>

              <button
                type="button"
                onClick={() => setActiveFilterStage('ordens_bancarias')}
                className={`px-3 py-1.5 rounded-xl text-xs font-bold transition cursor-pointer shrink-0 ${
                  activeFilterStage === 'ordens_bancarias'
                    ? 'bg-teal-700 text-white shadow-xs'
                    : 'bg-white text-teal-800 hover:bg-teal-50 border border-teal-200'
                }`}
              >
                Ordens bancárias ({summary.bancoRecordsCount || 0})
              </button>

              {summary.ppProntasCount > 0 && (
                <button
                  type="button"
                  onClick={() => setActiveFilterStage('base_pps')}
                  className={`px-3 py-1.5 rounded-xl text-xs font-bold transition cursor-pointer shrink-0 ${
                    activeFilterStage === 'base_pps'
                      ? 'bg-emerald-700 text-white shadow-xs'
                      : 'bg-white text-emerald-800 hover:bg-emerald-50 border border-emerald-200'
                  }`}
                >
                  Base de PPs ({summary.ppProntasCount})
                </button>
              )}
            </div>

            {/* Quick search input */}
            <div className="relative w-full sm:w-64 shrink-0">
              <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
              <input
                type="text"
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                placeholder="Filtrar por documento, processo, etc..."
                className="w-full pl-8 pr-3 py-1.5 bg-white border border-slate-200 rounded-xl text-xs text-slate-800 focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500"
              />
            </div>
          </div>

          {/* Occurrences Table */}
          <div className="border border-slate-200 rounded-2xl overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="bg-slate-100/80 text-slate-700 font-extrabold uppercase tracking-wider border-b border-slate-200 text-[11px]">
                    <th className="py-3 px-3 text-center w-10">#</th>
                    <th className="py-3 px-3 w-40">LIQUIDAÇÃO</th>
                    <th className="py-3 px-3 w-44 font-mono">NÚMERO DA LIQUIDAÇÃO</th>
                    <th className="py-3 px-3 text-right w-32 font-mono">VALOR (R$)</th>
                    <th className="py-3 px-3 text-center w-36">PUBLICAR</th>
                    <th className="py-3 px-3 min-w-[120px]">FONTE DE RECURSO</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 font-medium">
                  {paginatedRows.length === 0 ? (
                    <tr>
                      <td colSpan={6} className="py-10 text-center text-slate-400 italic">
                        Nenhuma ocorrência encontrada para este filtro.
                      </td>
                    </tr>
                  ) : (
                    paginatedRows.map((row, idx) => {
                      const realIndex = pageSize === 0 ? idx + 1 : (currentPage - 1) * pageSize + idx + 1;
                      const isSelected = paraPublicarService.isSelected(row.documentNumber, summary.cnpj);
                      return (
                        <tr key={row.id} className="hover:bg-slate-50 transition-colors">
                          <td className="py-2.5 px-3 text-center font-mono text-[11px] text-slate-400 border-r border-slate-100">
                            {realIndex}
                          </td>
                          <td className="py-2.5 px-3 border-r border-slate-100">
                            <span className={`inline-block px-2 py-0.5 rounded-md text-[10px] font-extrabold border ${getSourceBadgeStyle(row.sourceType, row.status)}`}>
                              {row.sourceLabel}
                            </span>
                          </td>
                          <td className="py-2.5 px-3 font-mono font-bold text-slate-900 border-r border-slate-100">
                            <span className="bg-slate-100 px-2 py-0.5 rounded border border-slate-200 text-[11px] font-mono font-bold text-slate-900 inline-block">
                              {formatCleanNL(row.documentNumber)}
                            </span>
                          </td>
                          <td className="py-2.5 px-3 text-right font-mono font-bold text-slate-900 border-r border-slate-100">
                            {formatBRL(row.value)}
                          </td>
                          <td className="py-2.5 px-3 text-center border-r border-slate-100">
                            <label className="inline-flex items-center gap-1.5 cursor-pointer select-none">
                              <input
                                type="checkbox"
                                checked={isSelected}
                                onChange={() =>
                                  paraPublicarService.toggleItem({
                                    documentNumber: row.documentNumber,
                                    supplierName: summary.supplierName,
                                    cnpj: resolveCnpj(summary.supplierName, row.documentNumber, summary.cnpj),
                                    value: row.value,
                                    fonteRecurso: row.fonteRecurso,
                                    sourceType: row.sourceType,
                                    sourceLabel: row.sourceLabel,
                                    status: row.status,
                                    date: row.date,
                                  })
                                }
                                className="w-4 h-4 rounded border-slate-300 text-emerald-600 focus:ring-emerald-500 cursor-pointer"
                              />
                              <span
                                className={`text-[10px] font-bold px-1.5 py-0.5 rounded transition ${
                                  isSelected
                                    ? 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                                    : 'text-slate-500 hover:text-slate-800'
                                }`}
                              >
                                {isSelected ? 'Selecionado' : 'Publicar'}
                              </span>
                            </label>
                          </td>
                          <td className="py-2.5 px-3 text-slate-700 text-[11px] font-mono font-bold">
                            {row.fonteRecurso || '0.6.00'}
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>
          </div>

          {/* Pagination */}
          {pageSize > 0 && totalPages > 1 && (
            <div className="flex items-center justify-between text-xs text-slate-500 pt-2">
              <div>
                Página <strong>{currentPage}</strong> de <strong>{totalPages}</strong> ({filteredRows.length} registros)
              </div>
              <div className="flex items-center gap-1.5">
                <button
                  onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                  disabled={currentPage === 1}
                  className="p-1.5 rounded-lg border border-slate-200 hover:bg-slate-100 disabled:opacity-40 disabled:pointer-events-none transition cursor-pointer"
                >
                  <ChevronLeft className="w-4 h-4" />
                </button>
                <span className="font-bold text-slate-800 px-2">
                  {currentPage} / {totalPages}
                </span>
                <button
                  onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
                  disabled={currentPage === totalPages}
                  className="p-1.5 rounded-lg border border-slate-200 hover:bg-slate-100 disabled:opacity-40 disabled:pointer-events-none transition cursor-pointer"
                >
                  <ChevronRight className="w-4 h-4" />
                </button>
              </div>
            </div>
          )}
        </div>

        {/* Modal Footer */}
        <div className="bg-slate-50 px-6 py-4 border-t border-slate-200 flex items-center justify-end shrink-0">
          <button
            type="button"
            onClick={onClose}
            className="w-full sm:w-auto px-6 py-2.5 bg-slate-800 hover:bg-slate-700 text-white font-bold text-xs rounded-xl transition cursor-pointer active:scale-95 shadow-xs"
          >
            Fechar Extrato
          </button>
        </div>
      </div>
    </div>
  );
};
