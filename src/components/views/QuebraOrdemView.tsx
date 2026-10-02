import React, { useState, useMemo, useRef } from 'react';
import {
  FileSpreadsheet,
  Upload,
  AlertTriangle,
  CheckCircle2,
  Download,
  Search,
  ArrowUpDown,
  Filter,
  Building,
  Sparkles,
  RefreshCw,
  ExternalLink,
  ChevronLeft,
  ChevronRight,
  HelpCircle,
  FileCheck2,
  TrendingDown,
  Clock,
  Trash2,
  ArrowRight,
  Loader2,
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { FileData, QuebraOrdemItem, QuebraOrdemResult } from '../../types';
import { formatBRL } from '../../utils/excelParser';
import {
  readSpreadsheetAsMatrix,
  parsePagamentosEmitidosMatrix,
  extractLiquidationCandidates,
  crossReferenceQuebraOrdem,
  extractAORowsFromWorkspaceFiles,
  RawPagamentoEmitido,
} from '../../utils/quebraOrdemParser';
import { InfoHelpButton } from '../InfoHelpButton';

interface QuebraOrdemViewProps {
  files: FileData[];
  onOpenCompanyAudit?: (supplier: { name: string; cnpj: string }) => void;
  rawPagamentosList?: RawPagamentoEmitido[] | null;
  onUpdatePagamentosList?: (list: RawPagamentoEmitido[] | null, fileName?: string) => void;
  onNavigateToImportacoes?: () => void;
}

export const QuebraOrdemView: React.FC<QuebraOrdemViewProps> = ({
  files,
  onOpenCompanyAudit,
  rawPagamentosList,
  onUpdatePagamentosList,
  onNavigateToImportacoes,
}) => {
  const [isLoading, setIsLoading] = useState(false);
  const [loadedFileName, setLoadedFileName] = useState<string>('');
  const [internalRawList, setInternalRawList] = useState<RawPagamentoEmitido[] | null>(null);

  // Local or shared raw list
  const activeRawList = rawPagamentosList !== undefined ? rawPagamentosList : internalRawList;

  // Search & Filter State
  const [searchTerm, setSearchTerm] = useState('');
  const [filterMatch, setFilterMatch] = useState<'all' | 'possivel_quebra' | 'fluxo_rapido' | 'matched_only' | 'unmatched_only'>('all');
  const [sortField, setSortField] = useState<'favorecido' | 'valorOB' | 'valorLiquidacao' | 'diferencaImpostos'>('diferencaImpostos');
  const [sortOrder, setSortOrder] = useState<'desc' | 'asc'>('desc');

  // Pagination
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState<number>(50);

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Candidate liquidations from uploaded spreadsheets
  const liquidationCandidates = useMemo(() => {
    return extractLiquidationCandidates(files);
  }, [files]);

  // Extract AO records from workspace files (including large-scale SIGEF report)
  const workspaceAoData = useMemo(() => {
    return extractAORowsFromWorkspaceFiles(files);
  }, [files]);

  // Cross-reference result
  const crossResult: QuebraOrdemResult | null = useMemo(() => {
    if (!activeRawList || activeRawList.length === 0) return null;
    return crossReferenceQuebraOrdem(activeRawList, liquidationCandidates, loadedFileName);
  }, [activeRawList, liquidationCandidates, loadedFileName]);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement> | React.DragEvent) => {
    let file: File | undefined;
    if ('dataTransfer' in e) {
      e.preventDefault();
      file = e.dataTransfer.files[0];
    } else {
      file = e.target.files?.[0];
    }

    if (!file) return;

    try {
      setIsLoading(true);
      const matrix = await readSpreadsheetAsMatrix(file);
      const parsed = parsePagamentosEmitidosMatrix(matrix);

      if (parsed.length === 0) {
        alert('Não foi possível identificar registros válidos no arquivo selecionado. Verifique o formato.');
        return;
      }

      setLoadedFileName(file.name);
      if (onUpdatePagamentosList) {
        onUpdatePagamentosList(parsed, file.name);
      } else {
        setInternalRawList(parsed);
      }
    } catch (err: any) {
      alert(`Erro ao processar arquivo: ${err?.message || err}`);
    } finally {
      setIsLoading(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  const handleClear = () => {
    setLoadedFileName('');
    if (onUpdatePagamentosList) {
      onUpdatePagamentosList(null, '');
    } else {
      setInternalRawList(null);
    }
  };

  // Filter and Sort Items
  const filteredItems = useMemo(() => {
    if (!crossResult) return [];

    return crossResult.items
      .filter((item) => {
        if (filterMatch === 'possivel_quebra' && !item.isPossivelQuebraOrdem) return false;
        if (filterMatch === 'fluxo_rapido' && !item.isFluxoRapido) return false;
        if (filterMatch === 'matched_only' && !item.liquidacaoCorrespondente?.matched) return false;
        if (filterMatch === 'unmatched_only' && item.liquidacaoCorrespondente?.matched) return false;

        const q = searchTerm.toLowerCase().trim();
        if (!q) return true;

        const matchesName = item.favorecido.toLowerCase().includes(q);
        const matchesCnpj = item.cnpj.toLowerCase().includes(q);
        const matchesPP = item.pp.toLowerCase().includes(q);
        const matchesOB = item.ob.toLowerCase().includes(q);
        const matchesProc = (item.processo || '').toLowerCase().includes(q);
        const matchesFonte = item.fonteRecurso.toLowerCase().includes(q) || (item.fonteMonitorada || '').toLowerCase().includes(q);

        return matchesName || matchesCnpj || matchesPP || matchesOB || matchesProc || matchesFonte;
      })
      .sort((a, b) => {
        let valA: any = 0;
        let valB: any = 0;

        if (sortField === 'favorecido') {
          valA = a.favorecido.toLowerCase();
          valB = b.favorecido.toLowerCase();
        } else if (sortField === 'valorOB') {
          valA = a.valorOB;
          valB = b.valorOB;
        } else if (sortField === 'valorLiquidacao') {
          valA = a.liquidacaoCorrespondente?.valorLiquidacao || a.valorOB;
          valB = b.liquidacaoCorrespondente?.valorLiquidacao || b.valorOB;
        } else if (sortField === 'diferencaImpostos') {
          valA = a.liquidacaoCorrespondente?.diferencaImpostos || 0;
          valB = b.liquidacaoCorrespondente?.diferencaImpostos || 0;
        }

        if (valA < valB) return sortOrder === 'asc' ? -1 : 1;
        if (valA > valB) return sortOrder === 'asc' ? 1 : -1;
        return 0;
      });
  }, [crossResult, filterMatch, searchTerm, sortField, sortOrder]);

  // Pagination slice
  const totalPages = pageSize === 0 ? 1 : Math.ceil(filteredItems.length / pageSize) || 1;
  const paginatedItems = useMemo(() => {
    if (pageSize === 0) return filteredItems;
    const start = (currentPage - 1) * pageSize;
    return filteredItems.slice(start, start + pageSize);
  }, [filteredItems, currentPage, pageSize]);

  const toggleSort = (field: typeof sortField) => {
    if (sortField === field) {
      setSortOrder(sortOrder === 'asc' ? 'desc' : 'asc');
    } else {
      setSortField(field);
      setSortOrder('desc');
    }
  };

  const [isExporting, setIsExporting] = useState(false);

  const handleExportExcel = () => {
    if (!crossResult || crossResult.items.length === 0 || isExporting) return;

    setIsExporting(true);
    setTimeout(() => {
      try {
        const dataToExport = filteredItems.map((item, idx) => ({
          'ITEM': idx + 1,
          'STATUS QUEBRA ORDEM': item.isPossivelQuebraOrdem ? 'POSSÍVEL QUEBRA' : 'REGULAR',
          'FONTE MONITORADA (OBEDECE)': item.fonteMonitorada || '-',
          'FAVORECIDO': item.favorecido,
          'CNPJ': item.cnpj,
          'VALOR LIQUIDAÇÃO (R$)': item.liquidacaoCorrespondente ? item.liquidacaoCorrespondente.valorLiquidacao : item.valorOB,
          'VALOR OB EMITIDA AO (R$)': item.valorOB,
          'DIFERENÇA IMPOSTOS (R$)': item.liquidacaoCorrespondente ? item.liquidacaoCorrespondente.diferencaImpostos : 0,
          '% RETENÇÃO': item.liquidacaoCorrespondente ? `${item.liquidacaoCorrespondente.percentualDiferenca.toFixed(2)}%` : '0%',
          'PP / OB': [item.pp ? `PP: ${item.pp}` : '', item.ob ? `OB: ${item.ob}` : ''].filter(Boolean).join(' / ') || item.ob || item.pp,
          'SITUAÇÃO': item.situacao,
          'FONTE RECURSO': item.fonteRecurso,
          'ETAPA ORIGEM LIQUIDAÇÃO': item.liquidacaoCorrespondente ? item.liquidacaoCorrespondente.etapaOrigem : 'Não localizada',
          'PLANILHA DE LIQUIDAÇÃO': item.liquidacaoCorrespondente ? item.liquidacaoCorrespondente.nomePlanilha : '-',
          'PROCESSO': item.processo || '-',
          'NOTA DE EMPENHO': item.notaEmpenho || '-',
        }));

        const worksheet = XLSX.utils.json_to_sheet(dataToExport);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Quebra_de_Ordem');
        XLSX.writeFile(workbook, `Quebra_de_Ordem_AO_${new Date().toISOString().slice(0, 10)}.xlsx`);
      } finally {
        setIsExporting(false);
      }
    }, 200);
  };

  return (
    <div className="space-y-8 pb-12">
      {/* Banner / Header */}
      <div className="bg-gradient-to-r from-slate-900 via-emerald-950 to-slate-900 rounded-2xl p-6 md:p-8 text-white shadow-xl relative overflow-hidden border border-slate-800">
        <div className="relative z-10 flex flex-col md:flex-row md:items-center justify-between gap-6">
          <div className="flex items-center gap-3">
            <h2 className="text-2xl font-black text-white tracking-tight">
              Monitoramento de Publicações (Quebra de Ordem)
            </h2>
            <InfoHelpButton
              title="Monitoramento de Publicações"
              variant="dark"
              content={
                <div className="space-y-2.5">
                  <p>
                    Monitora se uma liquidação que estava em aberto na fila de <strong>Obedece</strong> das fontes monitoradas teve despesa emitida com valor igual ou equivalente no relatório geral de pagamento, apontando possível quebra de ordem cronológica e dedução tributária.
                  </p>
                  <div className="p-2.5 bg-slate-800/90 rounded-xl border border-slate-700/80 text-[11px] space-y-1">
                    <span className="font-bold text-emerald-400 block">Auditoria FES/SESAP</span>
                    <span className="text-slate-300">
                      <strong>Fontes:</strong> 0. 5.00, 0. 7.00.035, 0. 7.04.121, 0. 7.07, 0. 7.20.720
                    </span>
                  </div>
                </div>
              }
            />
          </div>

          <div className="flex flex-wrap items-center gap-2.5 shrink-0">
            {crossResult && (
              <button
                type="button"
                onClick={handleExportExcel}
                disabled={isExporting}
                className="px-4 py-2.5 bg-emerald-600 hover:bg-emerald-500 text-white font-bold text-xs rounded-xl transition flex items-center gap-2 cursor-pointer shadow-md shadow-emerald-600/20 active:scale-98 disabled:opacity-75 disabled:cursor-wait"
              >
                {isExporting ? (
                  <>
                    <Loader2 className="w-4 h-4 animate-spin text-white" />
                    <span>Exportando Excel...</span>
                  </>
                ) : (
                  <>
                    <Download className="w-4 h-4" />
                    <span>Exportar Excel</span>
                  </>
                )}
              </button>
            )}

            {crossResult && (
              <button
                type="button"
                onClick={handleClear}
                className="px-3 py-2.5 bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white font-bold text-xs rounded-xl transition flex items-center gap-1.5 cursor-pointer active:scale-98"
                title="Limpar relatório atual"
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>Limpar</span>
              </button>
            )}
          </div>
        </div>
      </div>

      {/* Active File Bar or Empty State Callout */}
      {crossResult ? (
        <div className="bg-emerald-50/70 border border-emerald-200 rounded-2xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 shadow-2xs">
          <div className="flex items-center space-x-3">
            <div className="p-2 bg-emerald-600 text-white rounded-xl">
              <FileCheck2 className="w-5 h-5" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="text-xs font-bold text-slate-900">
                  Arquivo ativo: {loadedFileName || 'Relatório de Pagamentos Emitidos'}
                </span>
                <span className="text-[10px] bg-emerald-200 text-emerald-900 font-extrabold px-2 py-0.5 rounded-full">
                  {crossResult.totalValidos} registros analisados
                </span>
              </div>
              <p className="text-[11px] text-slate-500">
                Todas as importações e substituições de planilhas são gerenciadas na aba <strong>Importações</strong>.
              </p>
            </div>
          </div>

          {onNavigateToImportacoes && (
            <button
              type="button"
              onClick={onNavigateToImportacoes}
              className="px-3.5 py-1.5 bg-white hover:bg-slate-100 text-emerald-800 border border-emerald-300 text-xs font-bold rounded-xl transition flex items-center gap-1.5 cursor-pointer shrink-0 self-start sm:self-auto"
            >
              <span>Gerenciar na aba Importações</span>
              <ExternalLink className="w-3.5 h-3.5" />
            </button>
          )}
        </div>
      ) : (
        <div className="space-y-4 max-w-2xl mx-auto my-6">
          {workspaceAoData.records.length > 0 && (
            <div className="bg-emerald-50 border border-emerald-300 rounded-2xl p-5 flex flex-col sm:flex-row items-center justify-between gap-4 shadow-xs">
              <div className="flex items-center gap-3.5">
                <div className="p-2.5 bg-emerald-600 text-white rounded-xl shadow-xs">
                  <Sparkles className="w-5 h-5" />
                </div>
                <div className="text-left">
                  <h4 className="text-sm font-bold text-slate-900">
                    Detectamos {workspaceAoData.records.length} PPs Prontas (AO) no Relatório do SIGEF
                  </h4>
                  <p className="text-xs text-slate-600 mt-0.5">
                    Origem: <strong>{workspaceAoData.sourceFileName}</strong>. Deseja carregar essas PPs Prontas para análise de Quebra de Ordem?
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => {
                  setLoadedFileName(workspaceAoData.sourceFileName);
                  if (onUpdatePagamentosList) {
                    onUpdatePagamentosList(workspaceAoData.records, workspaceAoData.sourceFileName);
                  } else {
                    setInternalRawList(workspaceAoData.records);
                  }
                }}
                className="px-4 py-2.5 bg-emerald-700 hover:bg-emerald-600 text-white text-xs font-bold rounded-xl transition flex items-center gap-2 cursor-pointer shrink-0 shadow-xs"
              >
                <CheckCircle2 className="w-4 h-4" />
                <span>Usar PPs do SIGEF</span>
              </button>
            </div>
          )}

          <div className="bg-white rounded-2xl p-8 shadow-xs border border-slate-200/90 text-center space-y-4">
            <div className="w-14 h-14 mx-auto rounded-2xl bg-emerald-100 text-emerald-800 flex items-center justify-center shadow-xs">
              <FileSpreadsheet className="w-7 h-7" />
            </div>
            <div className="space-y-1.5">
              <h3 className="text-lg font-bold text-slate-900">
                Nenhum Relatório de Pagamentos Emitidos carregado
              </h3>
              <p className="text-xs text-slate-500 leading-relaxed max-w-lg mx-auto">
                Todas as importações do painel são realizadas em uma única aba. Acesse a <strong>Central de Importações</strong> para carregar o arquivo de Pagamentos Emitidos.
              </p>
            </div>

            <div className="flex flex-wrap items-center justify-center gap-3 pt-2">
              {onNavigateToImportacoes && (
                <button
                  type="button"
                  onClick={onNavigateToImportacoes}
                  className="px-5 py-2.5 bg-emerald-700 hover:bg-emerald-600 text-white font-bold text-xs rounded-xl transition flex items-center gap-2 cursor-pointer shadow-md shadow-emerald-700/20"
                >
                  <span>Ir para a Aba de Importações</span>
                  <ArrowRight className="w-4 h-4" />
                </button>
              )}
            </div>
          </div>
        </div>
      )}

      {/* Warning if no liquidation sheets were imported yet */}
      {crossResult && files.length === 0 && (
        <div className="bg-amber-50 border border-amber-200/90 rounded-xl p-4 flex items-start gap-3 text-amber-900 text-xs">
          <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
          <div>
            <strong>Dica para o cruzamento completo:</strong> Para calcular a variação de valor com as liquidações originais (dedução de impostos), importe também as planilhas de liquidações (Obedece e Não Obedece) na Central de Importações.
          </div>
        </div>
      )}

      {/* KPI Summary Cards */}
      {crossResult && (
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-6 gap-4">
          {/* Card 1: Possível Quebra de Ordem (Alerta Principal) */}
          <div className="bg-rose-50/70 border-2 border-rose-200 rounded-2xl p-5 shadow-xs space-y-1">
            <div className="flex items-center justify-between">
              <span className="text-[11px] font-extrabold text-rose-800 uppercase tracking-wider block">
                Possível Quebra de Ordem
              </span>
              <span className="text-xs font-black text-rose-700 bg-rose-100 px-2 py-0.5 rounded-md border border-rose-300">
                Alerta
              </span>
            </div>
            <div className="flex items-baseline justify-between">
              <span className="text-2xl font-black text-rose-950">
                {crossResult.totalPossivelQuebraOrdem}
              </span>
              <span className="text-xs font-mono font-bold text-rose-900">
                {formatBRL(crossResult.valorPossivelQuebraOrdem)}
              </span>
            </div>
          </div>

          {/* Card 2: Fluxo Rápido (22 Fontes Isentas) */}
          <div className="bg-blue-50/70 border border-blue-200 rounded-2xl p-5 shadow-xs space-y-1">
            <div className="flex items-center justify-between">
              <span className="text-[11px] font-extrabold text-blue-900 uppercase tracking-wider block">
                ⚡ Fluxo Rápido
              </span>
              <span className="text-xs font-black text-blue-700 bg-blue-100 px-2 py-0.5 rounded-md border border-blue-300">
                22 Fontes
              </span>
            </div>
            <div className="flex items-baseline justify-between">
              <span className="text-2xl font-black text-blue-950">
                {crossResult.totalFluxoRapido || 0}
              </span>
              <span className="text-xs font-mono font-bold text-blue-900">
                {formatBRL(crossResult.valorFluxoRapido || 0)}
              </span>
            </div>
          </div>

          {/* Card 3: Total AO Encontrados */}
          <div className="bg-white rounded-2xl p-5 shadow-xs border border-slate-200/80 space-y-1">
            <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider block">
              Publicações AO
            </span>
            <div className="flex items-baseline justify-between">
              <span className="text-2xl font-black text-slate-900">
                {crossResult.totalValidos}
              </span>
              <span className="text-xs font-semibold text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded-md border border-emerald-200">
                Situação AO
              </span>
            </div>
          </div>

          {/* Card 4: Valor Total das OBs Emitidas */}
          <div className="bg-white rounded-2xl p-5 shadow-xs border border-slate-200/80 space-y-1">
            <span className="text-[11px] font-bold text-emerald-800 uppercase tracking-wider block">
              Total OBs Emitidas
            </span>
            <span className="text-xl font-black text-emerald-950 font-mono block">
              {formatBRL(crossResult.valorTotalOB)}
            </span>
          </div>

          {/* Card 5: Dedução de Impostos (Diferença) */}
          <div className="bg-white rounded-2xl p-5 shadow-xs border border-slate-200/80 space-y-1">
            <span className="text-[11px] font-bold text-amber-800 uppercase tracking-wider block">
              Diferença (Impostos)
            </span>
            <span className="text-xl font-black text-amber-900 font-mono block">
              {formatBRL(crossResult.valorTotalImpostos)}
            </span>
          </div>

          {/* Card 6: Ignorados Fonte .600 / 0.6.00 */}
          <div className="bg-white rounded-2xl p-5 shadow-xs border border-slate-200/80 space-y-1">
            <span className="text-[11px] font-bold text-slate-600 uppercase tracking-wider block">
              Ignorados Fonte .600
            </span>
            <div className="flex items-baseline justify-between">
              <span className="text-2xl font-black text-slate-800">
                {crossResult.totalIgnoradosFonte600}
              </span>
              <span className="text-xs font-semibold text-slate-600 bg-slate-100 px-2 py-0.5 rounded-md border border-slate-200">
                Excluídos
              </span>
            </div>
          </div>
        </div>
      )}

      {/* Main Table Section */}
      {crossResult && (
        <div className="bg-white rounded-2xl shadow-xs border border-slate-200/90 overflow-hidden space-y-4 p-5">
          {/* Controls Bar: Search, Filters, Page size */}
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-slate-100 pb-4">
            {/* Search Input */}
            <div className="relative flex-1 max-w-md">
              <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
              <input
                type="text"
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                placeholder="Pesquisar por favorecido, CNPJ, PP, OB ou fonte..."
                className="w-full pl-9 pr-4 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs text-slate-800 focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition"
              />
            </div>

            {/* Filter Toggle Buttons */}
            <div className="flex flex-wrap items-center gap-2">
              <div className="flex items-center bg-slate-100 p-1 rounded-xl text-xs font-semibold">
                <button
                  type="button"
                  onClick={() => setFilterMatch('all')}
                  className={`px-3 py-1.5 rounded-lg transition cursor-pointer ${
                    filterMatch === 'all'
                      ? 'bg-white text-slate-900 shadow-xs'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  Todos ({crossResult.totalValidos})
                </button>
                <button
                  type="button"
                  onClick={() => setFilterMatch('possivel_quebra')}
                  className={`px-3 py-1.5 rounded-lg transition cursor-pointer flex items-center gap-1.5 ${
                    filterMatch === 'possivel_quebra'
                      ? 'bg-rose-600 text-white shadow-xs font-bold'
                      : 'text-rose-700 hover:text-rose-900 font-bold'
                  }`}
                >
                  <span>⚠️ Possível Quebra ({crossResult.totalPossivelQuebraOrdem})</span>
                </button>
                <button
                  type="button"
                  onClick={() => setFilterMatch('fluxo_rapido')}
                  className={`px-3 py-1.5 rounded-lg transition cursor-pointer flex items-center gap-1.5 ${
                    filterMatch === 'fluxo_rapido'
                      ? 'bg-blue-600 text-white shadow-xs font-bold'
                      : 'text-blue-700 hover:text-blue-900 font-bold'
                  }`}
                >
                  <span>⚡ Fluxo Rápido ({crossResult.totalFluxoRapido || 0})</span>
                </button>
                <button
                  type="button"
                  onClick={() => setFilterMatch('matched_only')}
                  className={`px-3 py-1.5 rounded-lg transition cursor-pointer ${
                    filterMatch === 'matched_only'
                      ? 'bg-white text-emerald-800 shadow-xs'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  Com Liquidação ({crossResult.totalComLiquidacaoCasada})
                </button>
                <button
                  type="button"
                  onClick={() => setFilterMatch('unmatched_only')}
                  className={`px-3 py-1.5 rounded-lg transition cursor-pointer ${
                    filterMatch === 'unmatched_only'
                      ? 'bg-white text-slate-900 shadow-xs'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  Sem Liq. Localizada ({crossResult.totalValidos - crossResult.totalComLiquidacaoCasada})
                </button>
              </div>

              {/* Page size select */}
              <div className="flex items-center gap-1.5 text-xs text-slate-500">
                <span>Exibir:</span>
                <select
                  value={pageSize}
                  onChange={(e) => {
                    setPageSize(Number(e.target.value));
                    setCurrentPage(1);
                  }}
                  className="bg-slate-50 border border-slate-200 rounded-lg px-2 py-1 text-xs font-bold text-slate-700 cursor-pointer focus:outline-none"
                >
                  <option value={25}>25</option>
                  <option value={50}>50</option>
                  <option value={100}>100</option>
                  <option value={0}>Todos</option>
                </select>
              </div>
            </div>
          </div>

          {/* Clean Data Table */}
          <div className="overflow-x-auto rounded-xl border border-slate-100">
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="bg-slate-50 text-slate-600 font-extrabold uppercase tracking-wider border-b border-slate-200 text-[11px]">
                  <th className="py-3.5 px-3 w-12 text-center">#</th>
                  <th className="py-3.5 px-3 w-32 text-center">STATUS ORDEM</th>
                  <th className="py-3.5 px-4 min-w-[200px]">
                    <button
                      onClick={() => toggleSort('favorecido')}
                      className="flex items-center gap-1.5 hover:text-emerald-700 cursor-pointer font-extrabold"
                    >
                      <span>FAVORECIDO</span>
                      <ArrowUpDown className="w-3 h-3 text-slate-400" />
                    </button>
                  </th>
                  <th className="py-3.5 px-3 w-36">CNPJ</th>
                  <th className="py-3.5 px-3 text-right w-36 bg-blue-50/40">
                    <button
                      onClick={() => toggleSort('valorLiquidacao')}
                      className="flex items-center justify-end gap-1 w-full hover:text-blue-700 cursor-pointer font-bold text-blue-950"
                    >
                      <span>VALOR LIQUIDAÇÃO</span>
                      <ArrowUpDown className="w-3 h-3 text-blue-600" />
                    </button>
                  </th>
                  <th className="py-3.5 px-3 text-right w-36 bg-emerald-50/40">
                    <button
                      onClick={() => toggleSort('valorOB')}
                      className="flex items-center justify-end gap-1 w-full hover:text-emerald-700 cursor-pointer font-bold text-emerald-950"
                    >
                      <span>VALOR OB (AO)</span>
                      <ArrowUpDown className="w-3 h-3 text-emerald-600" />
                    </button>
                  </th>
                  <th className="py-3.5 px-3 w-32 text-center">PP / OB</th>
                  <th className="py-3.5 px-3 text-right w-36 bg-amber-50/40">
                    <button
                      onClick={() => toggleSort('diferencaImpostos')}
                      className="flex items-center justify-end gap-1 w-full hover:text-amber-700 cursor-pointer font-bold text-amber-950"
                    >
                      <span>DIFERENÇA (IMPOSTOS)</span>
                      <ArrowUpDown className="w-3 h-3 text-amber-600" />
                    </button>
                  </th>
                  <th className="py-3.5 px-3 w-28 text-center">FONTE RECURSO</th>
                  <th className="py-3.5 px-3 w-24 text-center">SITUAÇÃO</th>
                  <th className="py-3.5 px-3 w-24 text-center">AÇÃO</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 font-medium">
                {paginatedItems.length === 0 ? (
                  <tr>
                    <td colSpan={11} className="py-12 text-center text-slate-400 italic">
                      Nenhum registro de Quebra de Ordem (AO) encontrado com os filtros aplicados.
                    </td>
                  </tr>
                ) : (
                  paginatedItems.map((item, idx) => {
                    const realIndex = pageSize === 0 ? idx + 1 : (currentPage - 1) * pageSize + idx + 1;
                    const liqVal = item.liquidacaoCorrespondente ? item.liquidacaoCorrespondente.valorLiquidacao : item.valorOB;
                    const difVal = item.liquidacaoCorrespondente ? item.liquidacaoCorrespondente.diferencaImpostos : 0;
                    const pctVal = item.liquidacaoCorrespondente ? item.liquidacaoCorrespondente.percentualDiferenca : 0;

                    return (
                      <tr
                        key={item.id}
                        className={`transition-colors ${
                          item.isPossivelQuebraOrdem
                            ? 'bg-rose-50/20 hover:bg-rose-50/40'
                            : 'hover:bg-slate-50'
                        }`}
                      >
                        {/* Index */}
                        <td className="py-3 px-3 text-center font-mono text-[11px] text-slate-400 border-r border-slate-100">
                          {realIndex}
                        </td>

                        {/* Status Ordem Cronológica */}
                        <td className="py-3 px-2 text-center border-r border-slate-100">
                          {item.isPossivelQuebraOrdem ? (
                            <div className="inline-flex flex-col items-center">
                              <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md text-[10px] font-black bg-rose-100 text-rose-800 border border-rose-300">
                                ⚠️ Possível Quebra
                              </span>
                              {item.fonteMonitorada && (
                                <span className="text-[9px] font-mono text-rose-700 mt-0.5 font-bold" title="Fonte Obedece Monitorada">
                                  {item.fonteMonitorada}
                                </span>
                              )}
                            </div>
                          ) : item.isFluxoRapido ? (
                            <div className="inline-flex flex-col items-center">
                              <span
                                className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md text-[10px] font-black bg-blue-100 text-blue-900 border border-blue-300"
                                title="Fonte de Fluxo Rápido: Paga mais rapidamente e não exige publicação de quebra de ordem."
                              >
                                ⚡ Fluxo Rápido
                              </span>
                              {item.fonteFluxoRapido && (
                                <span className="text-[9px] font-mono text-blue-700 mt-0.5 font-bold" title="Fonte Isenta de Fluxo Rápido">
                                  {item.fonteFluxoRapido}
                                </span>
                              )}
                            </div>
                          ) : item.liquidacaoCorrespondente?.matched ? (
                            <span className="inline-flex items-center px-2 py-0.5 rounded-md text-[10px] font-semibold bg-emerald-50 text-emerald-800 border border-emerald-200">
                              Regular
                            </span>
                          ) : (
                            <span className="inline-flex items-center px-2 py-0.5 rounded-md text-[10px] font-medium text-slate-400 bg-slate-50 border border-slate-200">
                              Não Vinculada
                            </span>
                          )}
                        </td>

                        {/* Favorecido Name */}
                        <td className="py-3 px-4 font-bold text-slate-900 border-r border-slate-100">
                          <div className="flex items-center space-x-2">
                            <Building className={`w-3.5 h-3.5 shrink-0 ${item.isPossivelQuebraOrdem ? 'text-rose-600' : 'text-slate-400'}`} />
                            <span className="truncate max-w-xs md:max-w-sm" title={item.favorecido}>
                              {item.favorecido}
                            </span>
                          </div>
                        </td>

                        {/* CNPJ */}
                        <td className="py-3 px-3 font-mono text-slate-700 border-r border-slate-100">
                          <span className="bg-slate-100 text-slate-800 px-2 py-0.5 rounded-md border border-slate-200/80 text-[11px] font-bold inline-block">
                            {item.cnpj || 'N/I'}
                          </span>
                        </td>

                        {/* Valor Liquidação Original */}
                        <td className="py-3 px-3 text-right font-mono border-r border-slate-100 bg-blue-50/10">
                          <span className="font-bold text-blue-900 text-xs">
                            {formatBRL(liqVal)}
                          </span>
                          {item.liquidacaoCorrespondente && (
                            <div className="text-[10px] text-blue-600 font-sans truncate max-w-[130px] ml-auto" title={item.liquidacaoCorrespondente.nomePlanilha}>
                              {item.liquidacaoCorrespondente.etapaOrigem === 'obedece' ? 'Fila Obedece' : item.liquidacaoCorrespondente.etapaOrigem === 'nao_obedece' ? 'Não Obedece' : 'PP Prontas'}
                            </div>
                          )}
                        </td>

                        {/* Valor OB Emitida (AO) */}
                        <td className="py-3 px-3 text-right font-mono border-r border-slate-100 bg-emerald-50/10">
                          <span className="font-bold text-emerald-900 text-xs">
                            {formatBRL(item.valorOB)}
                          </span>
                        </td>

                        {/* PP / OB */}
                        <td className="py-3 px-3 text-center font-mono text-slate-800 border-r border-slate-100">
                          <span className="bg-slate-100 px-2 py-0.5 rounded border border-slate-200 text-[11px] font-bold inline-block">
                            {[item.pp ? `PP:${item.pp}` : '', item.ob ? `OB:${item.ob}` : ''].filter(Boolean).join(' ') || item.ob || item.pp || '-'}
                          </span>
                        </td>

                        {/* Diferença (Impostos Retidos) */}
                        <td className="py-3 px-3 text-right font-mono border-r border-slate-100 bg-amber-50/10">
                          {item.liquidacaoCorrespondente?.matched ? (
                            <div>
                              <span className={`font-bold text-xs ${difVal > 0 ? 'text-amber-900' : 'text-slate-500'}`}>
                                {formatBRL(difVal)}
                              </span>
                              {difVal > 0 && (
                                <span className="block text-[10px] text-amber-700 font-sans">
                                  ({pctVal.toFixed(1)}% retenção)
                                </span>
                              )}
                            </div>
                          ) : (
                            <span className="text-slate-400 text-[11px] italic">
                              Sem Liq. vinculada
                            </span>
                          )}
                        </td>

                        {/* Fonte de Recurso */}
                        <td className="py-3 px-3 text-center font-mono text-slate-700 border-r border-slate-100">
                          <span className="text-[11px] font-semibold text-slate-700 bg-slate-100 px-2 py-0.5 rounded">
                            {item.fonteRecurso || 'N/I'}
                          </span>
                        </td>

                        {/* Situação AO */}
                        <td className="py-3 px-3 text-center border-r border-slate-100">
                          <span className="inline-flex items-center px-2 py-0.5 rounded-md text-[11px] font-extrabold bg-emerald-100 text-emerald-900 border border-emerald-300">
                            {item.situacao || 'AO'}
                          </span>
                        </td>

                        {/* Ação: Auditar Empresa */}
                        <td className="py-3 px-3 text-center">
                          {onOpenCompanyAudit && (
                            <button
                              type="button"
                              onClick={() => onOpenCompanyAudit({ name: item.favorecido, cnpj: item.cnpj })}
                              className="inline-flex items-center gap-1 px-2.5 py-1 bg-slate-100 hover:bg-emerald-100 text-slate-700 hover:text-emerald-900 rounded-lg text-[11px] font-bold border border-slate-200 hover:border-emerald-300 transition cursor-pointer active:scale-95"
                              title="Abrir Extrato Geral de Auditoria desta empresa"
                            >
                              <span>Extrato</span>
                              <ExternalLink className="w-3 h-3 text-slate-400 group-hover:text-emerald-700" />
                            </button>
                          )}
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>

          {/* Pagination Controls */}
          {pageSize > 0 && totalPages > 1 && (
            <div className="flex flex-col sm:flex-row items-center justify-between gap-3 pt-2 text-xs text-slate-500">
              <div>
                Exibindo página <strong>{currentPage}</strong> de <strong>{totalPages}</strong> ({filteredItems.length} registros filtrados)
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
      )}
    </div>
  );
};
