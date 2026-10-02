import React, { useState, useMemo } from 'react';
import {
  Upload,
  Download,
  Search,
  ArrowUpDown,
  Building,
  Calendar,
  Layers,
  CheckCircle2,
  Trash2,
  ArrowRight,
  Receipt,
  ChevronDown,
  ChevronRight,
  ExternalLink,
  ChevronLeft,
  ChevronRight as ChevronRightIcon,
  ChevronsUpDown,
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { OrdemBancariaItem, parseListarOrdemBancariaMatrixAsync } from '../../utils/ordemBancariaParser';
import { readSpreadsheetAsMatrix, formatBRL } from '../../utils/excelParser';
import { FileData } from '../../types';

interface PagamentosPorFonteViewProps {
  files?: FileData[];
  ordensBancariasList?: OrdemBancariaItem[] | null;
  ordensBancariasFileName?: string;
  onUpdateOrdensBancariasList?: (list: OrdemBancariaItem[] | null, fileName?: string) => void;
  onOpenCompanyAudit?: (supplier: { name: string; cnpj: string }) => void;
  onNavigateToImportacoes?: () => void;
}

interface SupplierFonteRow {
  key: string;
  name: string;
  cnpj: string;
  totalGeral: number;
  pagoNoMes: number;
  fontesValues: Record<string, number>;
  despesas: OrdemBancariaItem[];
}

interface FonteColorStyle {
  badgeBg: string;
  badgeBorder: string;
  badgeText: string;
  dot: string;
  headerBg: string;
  headerText: string;
}

const getFonteColorStyle = (fonte: string): FonteColorStyle => {
  const clean = (fonte || '').trim();
  let hash = 0;
  for (let i = 0; i < clean.length; i++) {
    hash = (hash * 31 + clean.charCodeAt(i)) & 0xffffffff;
  }
  const palettes: FonteColorStyle[] = [
    {
      badgeBg: 'bg-emerald-50',
      badgeBorder: 'border-emerald-200/90',
      badgeText: 'text-emerald-900',
      dot: 'bg-emerald-500',
      headerBg: 'bg-emerald-50/40',
      headerText: 'text-emerald-950',
    },
    {
      badgeBg: 'bg-blue-50',
      badgeBorder: 'border-blue-200/90',
      badgeText: 'text-blue-900',
      dot: 'bg-blue-500',
      headerBg: 'bg-blue-50/40',
      headerText: 'text-blue-950',
    },
    {
      badgeBg: 'bg-indigo-50',
      badgeBorder: 'border-indigo-200/90',
      badgeText: 'text-indigo-900',
      dot: 'bg-indigo-500',
      headerBg: 'bg-indigo-50/40',
      headerText: 'text-indigo-950',
    },
    {
      badgeBg: 'bg-purple-50',
      badgeBorder: 'border-purple-200/90',
      badgeText: 'text-purple-900',
      dot: 'bg-purple-500',
      headerBg: 'bg-purple-50/40',
      headerText: 'text-purple-950',
    },
    {
      badgeBg: 'bg-teal-50',
      badgeBorder: 'border-teal-200/90',
      badgeText: 'text-teal-900',
      dot: 'bg-teal-500',
      headerBg: 'bg-teal-50/40',
      headerText: 'text-teal-950',
    },
    {
      badgeBg: 'bg-amber-50',
      badgeBorder: 'border-amber-200/90',
      badgeText: 'text-amber-900',
      dot: 'bg-amber-500',
      headerBg: 'bg-amber-50/40',
      headerText: 'text-amber-950',
    },
    {
      badgeBg: 'bg-sky-50',
      badgeBorder: 'border-sky-200/90',
      badgeText: 'text-sky-900',
      dot: 'bg-sky-500',
      headerBg: 'bg-sky-50/40',
      headerText: 'text-sky-950',
    },
    {
      badgeBg: 'bg-rose-50',
      badgeBorder: 'border-rose-200/90',
      badgeText: 'text-rose-900',
      dot: 'bg-rose-500',
      headerBg: 'bg-rose-50/40',
      headerText: 'text-rose-950',
    },
  ];
  return palettes[Math.abs(hash) % palettes.length];
};

export const PagamentosPorFonteView: React.FC<PagamentosPorFonteViewProps> = ({
  files = [],
  ordensBancariasList = [],
  ordensBancariasFileName = '',
  onUpdateOrdensBancariasList,
  onOpenCompanyAudit,
  onNavigateToImportacoes,
}) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedMes, setSelectedMes] = useState<string>('all');
  const [sortField, setSortField] = useState<string>('pagoNoMes');
  const [sortOrder, setSortOrder] = useState<'asc' | 'desc'>('desc');
  const [isUploading, setIsUploading] = useState(false);
  const [expandedSuppliers, setExpandedSuppliers] = useState<Set<string>>(new Set());
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(50);
  const [confirmClear, setConfirmClear] = useState(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);

  // Consolidated items from ordensBancariasList or detected Fonte in files
  const items = useMemo<OrdemBancariaItem[]>(() => {
    if (ordensBancariasList && ordensBancariasList.length > 0) {
      return ordensBancariasList;
    }

    if (files && files.length > 0) {
      const extracted: OrdemBancariaItem[] = [];
      files.forEach((file) => {
        const fonteCol = file.detectedFonteCol;
        if (!fonteCol) return;
        const favCol = file.detectedFavorecidoCol;
        const valCol = file.detectedValorCol;
        const obCol = file.detectedObCol;
        const ppCol = file.detectedNumeroCol || file.detectedPpCol;
        const cnpjCol = file.detectedCnpjCol;
        const sitCol = file.detectedSituacaoCol;

        file.rows.forEach((row, idx) => {
          const rawFonte = String(row[fonteCol] || '').trim();
          if (!rawFonte) return;
          const rawVal = valCol ? row[valCol] : 0;
          const valor = typeof rawVal === 'number' ? rawVal : 0;
          if (valor <= 0) return;

          const rawFav = favCol ? String(row[favCol] || '').trim() : 'Favorecido';
          const ob = obCol ? String(row[obCol] || '').trim() : `ITEM-${idx + 1}`;
          const pp = ppCol ? String(row[ppCol] || '').trim() : '';
          const cnpj = cnpjCol ? String(row[cnpjCol] || '').trim() : '';
          const sit = sitCol ? String(row[sitCol] || '').trim() : 'PAGO';

          extracted.push({
            id: `${file.id}_${idx}`,
            ob: ob || `ITEM-${idx + 1}`,
            dataReferencia: '',
            mesAnoKey: '',
            mesAnoLabel: 'Exercício Atual',
            valor,
            situacao: sit,
            pp,
            fonteRecurso: rawFonte,
            favorecidoRaw: rawFav,
            favorecidoCnpj: cnpj,
            favorecidoName: rawFav,
          });
        });
      });
      if (extracted.length > 0) return extracted;
    }

    return [];
  }, [ordensBancariasList, files]);

  // Handle direct file upload in this view
  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      setIsUploading(true);
      setErrorMessage(null);
      const matrix = await readSpreadsheetAsMatrix(file);
      const parsed = await parseListarOrdemBancariaMatrixAsync(matrix);

      if (parsed.length === 0) {
        setErrorMessage(
          'Nenhum registro de Ordem Bancária conclusivo foi identificado no arquivo. Certifique-se de carregar a planilha "Listar Ordem Bancária" exportada do SIGEF.'
        );
        return;
      }

      if (onUpdateOrdensBancariasList) {
        onUpdateOrdensBancariasList(parsed, file.name);
      }
    } catch (err: any) {
      setErrorMessage('Erro ao processar planilha: ' + (err?.message || 'Arquivo corrompido ou formato inválido.'));
    } finally {
      setIsUploading(false);
      e.target.value = '';
    }
  };

  // Distinct list of Available Fontes de Recurso
  const availableFontes = useMemo(() => {
    const fontesSet = new Set<string>();
    items.forEach((it) => {
      const f = (it.fonteRecurso || '').trim();
      if (f) fontesSet.add(f);
    });
    return Array.from(fontesSet).sort();
  }, [items]);

  // Distinct list of Months
  const availableMonths = useMemo(() => {
    const map = new Map<string, string>();
    items.forEach((it) => {
      if (it.mesAnoKey && it.mesAnoLabel) {
        map.set(it.mesAnoKey, it.mesAnoLabel);
      }
    });
    return Array.from(map.entries())
      .map(([key, label]) => ({ key, label }))
      .sort((a, b) => b.key.localeCompare(a.key));
  }, [items]);

  // Set default month to the most recent month if available and currently 'all'
  React.useEffect(() => {
    if (selectedMes === 'all' && availableMonths.length > 0) {
      setSelectedMes(availableMonths[0].key);
    }
  }, [availableMonths]);

  // Active Month Label for headers
  const activeMonthLabel = useMemo(() => {
    if (selectedMes === 'all') return 'Todos os Meses';
    const found = availableMonths.find((m) => m.key === selectedMes);
    return found ? found.label : selectedMes;
  }, [selectedMes, availableMonths]);

  // Build supplier matrix: Fornecedores x Fontes + Pago no Mês
  const supplierRows = useMemo<SupplierFonteRow[]>(() => {
    const map = new Map<string, SupplierFonteRow>();

    items.forEach((it) => {
      const cnpjDigits = (it.favorecidoCnpj || '').replace(/\D/g, '');
      const key = cnpjDigits.length >= 11 ? cnpjDigits : it.favorecidoName.trim().toUpperCase();

      let row = map.get(key);
      if (!row) {
        row = {
          key,
          name: it.favorecidoName,
          cnpj: it.favorecidoCnpj || 'N/I',
          totalGeral: 0,
          pagoNoMes: 0,
          fontesValues: {},
          despesas: [],
        };
        map.set(key, row);
      }

      const f = (it.fonteRecurso || 'OUTRAS').trim();
      const belongsToSelectedMonth = selectedMes === 'all' || it.mesAnoKey === selectedMes;

      row.totalGeral += it.valor;

      if (belongsToSelectedMonth) {
        row.pagoNoMes += it.valor;
        row.fontesValues[f] = (row.fontesValues[f] || 0) + it.valor;
        row.despesas.push(it);
      }
    });

    return Array.from(map.values());
  }, [items, selectedMes]);

  // Filter suppliers by search query (name or CNPJ) and only show suppliers active in period
  const filteredSuppliers = useMemo(() => {
    const query = searchTerm.toLowerCase().trim();
    const cleanQuery = query.replace(/\D/g, '');

    const list = supplierRows.filter((s) => {
      // Must have value in the current filter context
      if (s.pagoNoMes <= 0 && selectedMes !== 'all') return false;

      if (!query) return true;

      const cleanCnpj = s.cnpj.replace(/\D/g, '');
      return (
        s.name.toLowerCase().includes(query) ||
        s.cnpj.toLowerCase().includes(query) ||
        (cleanQuery.length >= 3 && cleanCnpj.includes(cleanQuery))
      );
    });

    // Sorting
    list.sort((a, b) => {
      let valA = 0;
      let valB = 0;

      if (sortField === 'name') {
        const comp = a.name.localeCompare(b.name);
        return sortOrder === 'asc' ? comp : -comp;
      } else if (sortField === 'pagoNoMes') {
        valA = a.pagoNoMes;
        valB = b.pagoNoMes;
      } else if (sortField.startsWith('fonte_')) {
        const fonteKey = sortField.replace('fonte_', '');
        valA = a.fontesValues[fonteKey] || 0;
        valB = b.fontesValues[fonteKey] || 0;
      } else {
        valA = a.pagoNoMes;
        valB = b.pagoNoMes;
      }

      if (valA < valB) return sortOrder === 'asc' ? -1 : 1;
      if (valA > valB) return sortOrder === 'asc' ? 1 : -1;
      return 0;
    });

    return list;
  }, [supplierRows, searchTerm, selectedMes, sortField, sortOrder]);

  // Totals for table footer
  const columnTotals = useMemo(() => {
    const totals: Record<string, number> = {};
    let totalPagoNoMes = 0;
    let totalGeral = 0;
    let totalOrdens = 0;

    filteredSuppliers.forEach((s) => {
      totalPagoNoMes += s.pagoNoMes;
      totalGeral += s.totalGeral;
      totalOrdens += s.despesas.length;
      Object.entries(s.fontesValues).forEach(([f, val]) => {
        totals[f] = (totals[f] || 0) + val;
      });
    });

    return {
      fontes: totals,
      totalPagoNoMes,
      totalGeral,
      totalOrdens,
      fornecedoresCount: filteredSuppliers.length,
    };
  }, [filteredSuppliers]);

  // Pagination
  const totalPages = pageSize === 0 ? 1 : Math.ceil(filteredSuppliers.length / pageSize) || 1;
  const paginatedSuppliers = useMemo(() => {
    if (pageSize === 0) return filteredSuppliers;
    const start = (currentPage - 1) * pageSize;
    return filteredSuppliers.slice(start, start + pageSize);
  }, [filteredSuppliers, currentPage, pageSize]);

  // Toggle supplier expansion
  const toggleSupplierExpand = (key: string) => {
    setExpandedSuppliers((prev) => {
      const next = new Set(prev);
      if (next.has(key)) {
        next.delete(key);
      } else {
        next.add(key);
      }
      return next;
    });
  };

  // Expand or collapse all
  const handleToggleExpandAll = () => {
    if (expandedSuppliers.size >= filteredSuppliers.length && filteredSuppliers.length > 0) {
      setExpandedSuppliers(new Set());
    } else {
      setExpandedSuppliers(new Set(filteredSuppliers.map((s) => s.key)));
    }
  };

  // Toggle sort order
  const handleSort = (field: string) => {
    if (sortField === field) {
      setSortOrder(sortOrder === 'asc' ? 'desc' : 'asc');
    } else {
      setSortField(field);
      setSortOrder('desc');
    }
    setCurrentPage(1);
  };

  // Export to Excel: Matrix with each Fonte in a column + Pago no Mês
  const handleExportExcel = () => {
    if (filteredSuppliers.length === 0) return;

    // Sheet 1: Matrix view
    const matrixRows = filteredSuppliers.map((s, idx) => {
      const row: any = {
        '#': idx + 1,
        'Fornecedor / Favorecido': s.name,
        'CNPJ / CPF': s.cnpj,
      };

      availableFontes.forEach((fonte) => {
        row[`Fonte ${fonte} (R$)`] = s.fontesValues[fonte] || 0;
      });

      row[`Pago no Mês (${activeMonthLabel})`] = s.pagoNoMes;
      row['Total de Despesas (OBs)'] = s.despesas.length;
      return row;
    });

    const wb = XLSX.utils.book_new();
    const wsMatrix = XLSX.utils.json_to_sheet(matrixRows);
    XLSX.utils.book_append_sheet(wb, wsMatrix, 'Fornecedores por Fonte');

    // Sheet 2: Detailed expenses
    const detailRows: any[] = [];
    filteredSuppliers.forEach((s) => {
      s.despesas.forEach((d) => {
        detailRows.push({
          'Fornecedor': s.name,
          'CNPJ': s.cnpj,
          'Ordem Bancária (OB)': d.ob,
          'PP': d.pp || '',
          'Data': d.dataReferencia || '',
          'Mês': d.mesAnoLabel || '',
          'Fonte de Recursos': d.fonteRecurso || '',
          'Situação': d.situacao || '',
          'Valor Pago (R$)': d.valor,
        });
      });
    });

    if (detailRows.length > 0) {
      const wsDetails = XLSX.utils.json_to_sheet(detailRows);
      XLSX.utils.book_append_sheet(wb, wsDetails, 'Detalhamento de Despesas');
    }

    XLSX.writeFile(wb, `Pagamentos_Fornecedores_Por_Fonte_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  // Empty State: No data loaded
  if (items.length === 0) {
    return (
      <div className="space-y-8 animate-in fade-in-50 duration-200">
        <div className="bg-[#061d15] border border-emerald-900/80 text-white rounded-3xl p-8 md:p-12 shadow-xl text-center space-y-6 max-w-3xl mx-auto my-8">
          <div className="w-16 h-16 bg-emerald-600/20 text-emerald-400 rounded-2xl flex items-center justify-center mx-auto border border-emerald-500/30">
            <Receipt className="w-8 h-8" />
          </div>
          <div className="space-y-2">
            <h2 className="text-2xl font-black tracking-tight text-white">
              Nenhuma Ordem Bancária Carregada
            </h2>
          </div>

          <div className="flex flex-col sm:flex-row items-center justify-center gap-3 pt-2">
            <label className="w-full sm:w-auto px-6 py-3 bg-emerald-600 hover:bg-emerald-500 text-white font-bold text-sm rounded-xl transition flex items-center justify-center gap-2 shadow-lg shadow-emerald-600/20 cursor-pointer active:scale-98">
              <Upload className="w-4 h-4" />
              <span>{isUploading ? 'Processando Planilha...' : 'Carregar Planilha do SIGEF'}</span>
              <input
                type="file"
                accept=".xlsx, .xls, .csv"
                onChange={handleFileUpload}
                disabled={isUploading}
                className="hidden"
              />
            </label>

            {onNavigateToImportacoes && (
              <button
                type="button"
                onClick={onNavigateToImportacoes}
                className="w-full sm:w-auto px-6 py-3 bg-white/10 hover:bg-white/20 text-white font-bold text-sm rounded-xl transition flex items-center justify-center gap-2 cursor-pointer active:scale-98"
              >
                <span>Ir para Central de Importações</span>
                <ArrowRight className="w-4 h-4" />
              </button>
            )}
          </div>
        </div>
      </div>
    );
  }

  const isAllExpanded = expandedSuppliers.size >= filteredSuppliers.length && filteredSuppliers.length > 0;

  return (
    <div className="space-y-6 animate-in fade-in-50 duration-200">
      {/* Top Banner */}
      <div className="bg-[#061d15] text-white rounded-3xl p-6 sm:p-8 border border-emerald-900/80 shadow-xl flex flex-col md:flex-row items-start md:items-center justify-between gap-6">
        <div>
          <h2 className="text-2xl sm:text-3xl font-black tracking-tight text-white">
            Pagamentos por Fonte de Recursos
          </h2>
        </div>

        {/* Action Buttons */}
        <div className="flex flex-wrap items-center gap-3">
          <label className="inline-flex items-center gap-2 px-4 py-2.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-xs font-bold transition shadow-lg shadow-emerald-950/20 cursor-pointer active:scale-98">
            <Upload className="w-4 h-4" />
            <span>{isUploading ? 'Processando...' : 'Atualizar Planilha'}</span>
            <input
              type="file"
              accept=".xlsx, .xls, .csv"
              onChange={handleFileUpload}
              disabled={isUploading}
              className="hidden"
            />
          </label>

          <button
            type="button"
            onClick={handleExportExcel}
            className="inline-flex items-center gap-2 px-4 py-2.5 bg-white/10 hover:bg-white/20 text-white rounded-xl text-xs font-bold transition cursor-pointer active:scale-98"
            title="Exportar matriz de fornecedores e fontes para Excel"
          >
            <Download className="w-4 h-4 text-emerald-300" />
            <span>Exportar Excel</span>
          </button>

          {onUpdateOrdensBancariasList && (
            confirmClear ? (
              <div className="flex items-center gap-1.5 bg-rose-950/70 border border-rose-700/60 rounded-xl px-2.5 py-1.5 animate-in fade-in duration-100">
                <span className="text-[11px] text-rose-300 font-semibold">Remover OBs?</span>
                <button
                  type="button"
                  onClick={() => {
                    onUpdateOrdensBancariasList(null, '');
                    setConfirmClear(false);
                  }}
                  className="px-2 py-0.5 bg-rose-600 hover:bg-rose-700 text-white rounded text-[10px] font-bold cursor-pointer"
                >
                  Sim
                </button>
                <button
                  type="button"
                  onClick={() => setConfirmClear(false)}
                  className="px-2 py-0.5 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded text-[10px] font-bold cursor-pointer"
                >
                  Não
                </button>
              </div>
            ) : (
              <button
                type="button"
                onClick={() => setConfirmClear(true)}
                className="inline-flex items-center gap-1.5 px-3 py-2.5 bg-rose-950/40 hover:bg-rose-900/60 text-rose-300 border border-rose-800/40 rounded-xl text-xs font-bold transition cursor-pointer active:scale-98"
                title="Limpar relatório de Ordens Bancárias"
              >
                <Trash2 className="w-4 h-4" />
                <span className="hidden sm:inline">Limpar</span>
              </button>
            )
          )}
        </div>
      </div>

      {/* Error Message Banner */}
      {errorMessage && (
        <div className="bg-rose-50 border border-rose-200 text-rose-800 px-4 py-2.5 rounded-xl text-xs flex items-center justify-between animate-in fade-in duration-150">
          <span>{errorMessage}</span>
          <button
            type="button"
            onClick={() => setErrorMessage(null)}
            className="text-rose-500 hover:text-rose-700 font-bold ml-2 cursor-pointer"
          >
            ✕
          </button>
        </div>
      )}

      {/* KPI Cards Strip */}
      <div className="grid grid-cols-1 sm:grid-cols-4 gap-4">
        {/* Total Pago no Mês */}
        <div className="bg-white rounded-2xl p-5 border border-slate-200/80 shadow-sm flex items-center justify-between">
          <div>
            <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider block">
              Total Pago no Mês
            </span>
            <span className="text-xl sm:text-2xl font-black text-slate-900 tracking-tight block mt-1">
              {formatBRL(columnTotals.totalPagoNoMes)}
            </span>
          </div>
          <div className="w-12 h-12 rounded-2xl bg-emerald-50 border border-emerald-200 text-emerald-700 flex items-center justify-center shrink-0">
            <Receipt className="w-6 h-6" />
          </div>
        </div>

        {/* Fornecedores Atendidos */}
        <div className="bg-white rounded-2xl p-5 border border-slate-200/80 shadow-sm flex items-center justify-between">
          <div>
            <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider block">
              Fornecedores Atendidos
            </span>
            <span className="text-xl sm:text-2xl font-black text-slate-900 tracking-tight block mt-1">
              {columnTotals.fornecedoresCount.toLocaleString('pt-BR')}
            </span>
          </div>
          <div className="w-12 h-12 rounded-2xl bg-blue-50 border border-blue-200 text-blue-700 flex items-center justify-center shrink-0">
            <Building className="w-6 h-6" />
          </div>
        </div>

        {/* Fontes de Recurso */}
        <div className="bg-white rounded-2xl p-5 border border-slate-200/80 shadow-sm flex items-center justify-between">
          <div>
            <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider block">
              Fontes em Operação
            </span>
            <span className="text-xl sm:text-2xl font-black text-slate-900 tracking-tight block mt-1">
              {availableFontes.length}
            </span>
          </div>
          <div className="w-12 h-12 rounded-2xl bg-teal-50 border border-teal-200 text-teal-700 flex items-center justify-center shrink-0">
            <Layers className="w-6 h-6" />
          </div>
        </div>

        {/* Total Acumulado */}
        <div className="bg-white rounded-2xl p-5 border border-slate-200/80 shadow-sm flex items-center justify-between">
          <div>
            <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider block">
              Total Geral Acumulado
            </span>
            <span className="text-xl sm:text-2xl font-black text-slate-900 tracking-tight block mt-1">
              {formatBRL(columnTotals.totalGeral)}
            </span>
          </div>
          <div className="w-12 h-12 rounded-2xl bg-slate-100 border border-slate-200 text-slate-700 flex items-center justify-center shrink-0">
            <CheckCircle2 className="w-6 h-6" />
          </div>
        </div>
      </div>

      {/* Control Bar: Search Input, Month Selector & View Controls */}
      <div className="bg-white rounded-2xl p-4 border border-slate-200/80 shadow-sm space-y-4">
        <div className="flex flex-col md:flex-row items-stretch md:items-center justify-between gap-3">
          {/* Search Box - Matches Supplier Name & CNPJ */}
          <div className="relative flex-1 max-w-lg">
            <Search className="w-4 h-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              value={searchTerm}
              onChange={(e) => {
                setSearchTerm(e.target.value);
                setCurrentPage(1);
              }}
              placeholder="Pesquisar fornecedor por nome ou CNPJ..."
              className="w-full pl-9 pr-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs text-slate-900 focus:bg-white focus:outline-hidden focus:border-emerald-500 transition font-medium"
            />
            {searchTerm && (
              <button
                type="button"
                onClick={() => setSearchTerm('')}
                className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 text-xs font-bold"
              >
                ✕
              </button>
            )}
          </div>

          {/* Month Selector Filter */}
          <div className="flex flex-wrap items-center gap-2">
            <div className="flex items-center gap-2 bg-teal-50/80 border border-teal-200 px-3.5 py-2 rounded-xl text-xs font-bold text-teal-950">
              <Calendar className="w-4 h-4 text-teal-700 shrink-0" />
              <span className="text-[11px] text-teal-800">Mês de Referência:</span>
              <select
                value={selectedMes}
                onChange={(e) => {
                  setSelectedMes(e.target.value);
                  setCurrentPage(1);
                }}
                className="bg-transparent border-none font-black text-teal-950 text-xs cursor-pointer focus:outline-none"
              >
                <option value="all">Todos os Meses (Acumulado)</option>
                {availableMonths.map((m) => (
                  <option key={m.key} value={m.key}>
                    {m.label}
                  </option>
                ))}
              </select>
            </div>

            {/* Expand / Collapse All Despesas Button */}
            <button
              type="button"
              onClick={handleToggleExpandAll}
              className="px-3 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold transition flex items-center gap-1.5 cursor-pointer active:scale-98"
              title={isAllExpanded ? 'Recolher despesas de todos os fornecedores' : 'Expandir despesas de todos os fornecedores'}
            >
              <ChevronsUpDown className="w-3.5 h-3.5 text-slate-500" />
              <span>{isAllExpanded ? 'Recolher Despesas' : 'Expandir Todas as Despesas'}</span>
            </button>
          </div>
        </div>

        {/* Secondary Info and Active Filter Reset */}
        <div className="flex flex-wrap items-center justify-between gap-2 pt-2 border-t border-slate-100 text-xs text-slate-500">
          <div className="flex items-center gap-2">
            <span>
              Exibindo <strong>{filteredSuppliers.length}</strong> fornecedores
            </span>
            {searchTerm && (
              <span className="bg-emerald-50 text-emerald-800 px-2 py-0.5 rounded text-[11px] font-bold border border-emerald-200">
                Filtro: "{searchTerm}"
              </span>
            )}
          </div>

          {/* Rows per page */}
          <div className="flex items-center gap-2">
            <span className="text-[11px]">Linhas:</span>
            <select
              value={pageSize}
              onChange={(e) => {
                setPageSize(Number(e.target.value));
                setCurrentPage(1);
              }}
              className="bg-slate-50 border border-slate-200 rounded-lg px-2 py-1 text-xs font-bold text-slate-800 cursor-pointer focus:outline-none"
            >
              <option value={25}>25</option>
              <option value={50}>50</option>
              <option value={100}>100</option>
              <option value={0}>Todas</option>
            </select>
          </div>
        </div>
      </div>

      {/* ---------------- MAIN MATRIX TABLE: FORNECEDORES x FONTES + PAGO NO MÊS ---------------- */}
      <div className="bg-white rounded-3xl border border-slate-200/80 shadow-sm overflow-hidden">
        <div className="overflow-x-auto max-h-[700px] overflow-y-auto">
          <table className="w-full text-left border-collapse text-xs">
            <thead className="sticky top-0 z-20 bg-slate-100/95 backdrop-blur-xs shadow-xs">
              <tr className="border-b border-slate-200 text-slate-700 font-bold uppercase tracking-wider text-[11px]">
                {/* Index Column */}
                <th className="py-3 px-3 w-12 text-center text-slate-400 sticky left-0 z-20 bg-slate-100/95 border-r border-slate-200/80">
                  #
                </th>

                {/* Fornecedor Column (Sticky Left) */}
                <th className="py-3 px-4 min-w-[260px] sticky left-12 z-20 bg-slate-100/95 border-r border-slate-200/80 shadow-sm">
                  <button
                    type="button"
                    onClick={() => handleSort('name')}
                    className="flex items-center gap-1.5 hover:text-emerald-700 cursor-pointer font-extrabold active:scale-98"
                  >
                    <span>FORNECEDOR / FAVORECIDO</span>
                    <ArrowUpDown className="w-3.5 h-3.5 text-slate-400" />
                  </button>
                </th>

                {/* Dynamic Columns for each Fonte de Recurso */}
                {availableFontes.map((fonte) => {
                  const style = getFonteColorStyle(fonte);
                  const isSorted = sortField === `fonte_${fonte}`;
                  return (
                    <th
                      key={fonte}
                      className={`py-3 px-3.5 text-right min-w-[155px] border-r border-slate-200/80 whitespace-nowrap transition-colors ${
                        isSorted ? 'bg-slate-100/90' : 'bg-slate-50/70'
                      }`}
                    >
                      <button
                        type="button"
                        onClick={() => handleSort(`fonte_${fonte}`)}
                        className="group flex items-center justify-end gap-1.5 w-full cursor-pointer active:scale-98"
                        title={`Ordenar por pagamentos na Fonte ${fonte}`}
                      >
                        <span
                          className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg border text-[11px] font-semibold tracking-tight transition shadow-2xs ${
                            isSorted
                              ? `${style.badgeBg} ${style.badgeBorder} ${style.badgeText} ring-1 ring-slate-400/30`
                              : `bg-white border-slate-200/90 text-slate-700 group-hover:border-slate-300 group-hover:bg-slate-50`
                          }`}
                        >
                          <span className={`w-1.5 h-1.5 rounded-full ${style.dot} shrink-0`} />
                          <span className="font-mono font-bold tracking-tight text-[11px]">Fonte {fonte}</span>
                        </span>
                        <ArrowUpDown
                          className={`w-3 h-3 transition-colors shrink-0 ${
                            isSorted ? 'text-slate-800' : 'text-slate-400 group-hover:text-slate-600'
                          }`}
                        />
                      </button>
                    </th>
                  );
                })}

                {/* Final Column: Pago no Mês */}
                <th className="py-3 px-4 text-right min-w-[175px] bg-emerald-50/90 border-l-2 border-emerald-400 whitespace-nowrap sticky right-0 z-10 shadow-xs">
                  <button
                    type="button"
                    onClick={() => handleSort('pagoNoMes')}
                    className="group flex items-center justify-end gap-2 w-full cursor-pointer font-bold active:scale-98 text-emerald-950"
                    title="Ordenar pelo valor total pago no mês"
                  >
                    <div className="flex items-center gap-1.5">
                      <span className="px-1.5 py-0.5 rounded text-[10px] font-black uppercase tracking-wider bg-emerald-600 text-white shadow-2xs">
                        Total
                      </span>
                      <span className="text-xs font-bold text-emerald-950 tracking-tight group-hover:text-emerald-800">
                        PAGO NO MÊS
                      </span>
                    </div>
                    <ArrowUpDown className="w-3.5 h-3.5 text-emerald-700 shrink-0 group-hover:text-emerald-900" />
                  </button>
                </th>
              </tr>
            </thead>

            <tbody className="divide-y divide-slate-100 font-medium">
              {paginatedSuppliers.length === 0 ? (
                <tr>
                  <td
                    colSpan={availableFontes.length + 3}
                    className="py-12 text-center text-slate-400 italic"
                  >
                    Nenhum fornecedor localizado com o termo digitado ou sem pagamentos no período selecionado.
                  </td>
                </tr>
              ) : (
                paginatedSuppliers.map((supplier, idx) => {
                  const realIndex = pageSize === 0 ? idx + 1 : (currentPage - 1) * pageSize + idx + 1;
                  const isExpanded = expandedSuppliers.has(supplier.key);

                  return (
                    <React.Fragment key={supplier.key}>
                      <tr className={`hover:bg-slate-50 transition-colors ${isExpanded ? 'bg-slate-50/70' : ''}`}>
                        {/* Index */}
                        <td className="py-3 px-3 text-center font-mono text-[11px] text-slate-400 border-r border-slate-100 sticky left-0 bg-white z-10">
                          {realIndex}
                        </td>

                        {/* Fornecedor Name & Despesas Toggle */}
                        <td className="py-3 px-4 font-bold text-slate-900 border-r border-slate-100 sticky left-12 bg-white z-10 shadow-xs">
                          <div className="flex items-center justify-between gap-2">
                            <div className="flex items-center gap-2 min-w-0">
                              {/* Toggle Accordion Button */}
                              <button
                                type="button"
                                onClick={() => toggleSupplierExpand(supplier.key)}
                                className="p-1 hover:bg-slate-100 rounded text-slate-500 hover:text-emerald-700 transition cursor-pointer shrink-0"
                                title={isExpanded ? 'Ocultar despesas' : 'Ver despesas listadas'}
                              >
                                {isExpanded ? (
                                  <ChevronDown className="w-4 h-4 text-emerald-600" />
                                ) : (
                                  <ChevronRight className="w-4 h-4 text-slate-400" />
                                )}
                              </button>

                              <button
                                type="button"
                                onClick={() =>
                                  onOpenCompanyAudit &&
                                  onOpenCompanyAudit({ name: supplier.name, cnpj: supplier.cnpj })
                                }
                                className="flex items-center gap-1.5 text-left hover:text-emerald-700 transition cursor-pointer group truncate"
                                title="Abrir Extrato de Auditoria do Fornecedor"
                              >
                                <Building className="w-3.5 h-3.5 text-slate-400 group-hover:text-emerald-600 shrink-0" />
                                <span className="truncate max-w-[200px] sm:max-w-xs font-bold text-xs group-hover:underline">
                                  {supplier.name}
                                </span>
                              </button>
                            </div>

                            {/* Badge with despesas count */}
                            <button
                              type="button"
                              onClick={() => toggleSupplierExpand(supplier.key)}
                              className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-200/80 shrink-0 transition cursor-pointer"
                              title="Clique para listar todas as despesas deste fornecedor"
                            >
                              {supplier.despesas.length}{' '}
                              {supplier.despesas.length === 1 ? 'despesa' : 'despesas'}
                            </button>
                          </div>
                        </td>

                        {/* Fonte Columns Values */}
                        {availableFontes.map((fonte) => {
                          const val = supplier.fontesValues[fonte] || 0;
                          return (
                            <td
                              key={fonte}
                              className="py-3 px-3.5 text-right tabular-nums border-r border-slate-100 whitespace-nowrap"
                            >
                              {val > 0 ? (
                                <span className="font-semibold text-slate-800 text-xs tracking-tight">
                                  {formatBRL(val)}
                                </span>
                              ) : (
                                <span className="text-slate-300 font-light select-none text-xs">—</span>
                              )}
                            </td>
                          );
                        })}

                        {/* Final Column: Pago no Mês */}
                        <td className="py-3 px-4 text-right tabular-nums font-bold text-emerald-950 bg-emerald-50/40 border-l-2 border-emerald-300/80 whitespace-nowrap sticky right-0 z-10 shadow-xs text-xs">
                          <span className="inline-block px-2.5 py-1 rounded-lg bg-emerald-100/70 border border-emerald-200/80 font-mono font-bold text-emerald-950 shadow-2xs">
                            {formatBRL(supplier.pagoNoMes)}
                          </span>
                        </td>
                      </tr>

                      {/* ---------------- EXPANDABLE ROW: DEVIDAS DESPESAS LISTADAS ---------------- */}
                      {isExpanded && (
                        <tr className="bg-emerald-50/20">
                          <td
                            colSpan={availableFontes.length + 3}
                            className="p-3 pl-8 sm:pl-14 border-b border-emerald-200/60"
                          >
                            <div className="bg-white rounded-2xl border border-slate-200 p-4 shadow-sm space-y-3">
                              <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-2 border-b border-slate-100 pb-2">
                                <div className="flex items-center gap-2">
                                  <Receipt className="w-4 h-4 text-emerald-600" />
                                  <h4 className="text-xs font-black text-slate-900">
                                    Despesas Efetuadas para {supplier.name}
                                  </h4>
                                  <span className="text-[10px] font-mono text-slate-500">
                                    ({supplier.cnpj})
                                  </span>
                                </div>
                                <div className="text-xs tabular-nums font-bold text-slate-700">
                                  <span>Total das Despesas: </span>
                                  <span className="text-emerald-900 font-black font-mono">
                                    {formatBRL(supplier.pagoNoMes)}
                                  </span>
                                </div>
                              </div>

                              {supplier.despesas.length === 0 ? (
                                <p className="text-xs text-slate-400 italic py-2">
                                  Nenhuma despesa individual registrada no mês selecionado.
                                </p>
                              ) : (
                                <div className="overflow-x-auto rounded-xl border border-slate-200">
                                  <table className="w-full text-left border-collapse text-xs">
                                    <thead>
                                      <tr className="bg-slate-50 text-[10px] font-bold text-slate-600 uppercase border-b border-slate-200">
                                        <th className="py-2.5 px-3">#</th>
                                        <th className="py-2.5 px-3">Ordem Bancária (OB)</th>
                                        <th className="py-2.5 px-3">PP (Preparação)</th>
                                        <th className="py-2.5 px-3">Data</th>
                                        <th className="py-2.5 px-3">Mês</th>
                                        <th className="py-2.5 px-3">Fonte de Recursos</th>
                                        <th className="py-2.5 px-3 text-center">Situação</th>
                                        <th className="py-2.5 px-3 text-right">Valor Pago (R$)</th>
                                      </tr>
                                    </thead>
                                    <tbody className="divide-y divide-slate-100 text-[11px]">
                                      {supplier.despesas.map((despesa, dIdx) => {
                                        const fStyle = getFonteColorStyle(despesa.fonteRecurso || '');
                                        return (
                                          <tr key={despesa.id || dIdx} className="hover:bg-slate-50">
                                            <td className="py-2.5 px-3 text-slate-400 font-mono text-[10px]">
                                              {dIdx + 1}
                                            </td>
                                            <td className="py-2.5 px-3 font-mono font-bold text-slate-900">
                                              {despesa.ob}
                                            </td>
                                            <td className="py-2.5 px-3 font-mono text-slate-600">
                                              {despesa.pp || '—'}
                                            </td>
                                            <td className="py-2.5 px-3 text-slate-700">
                                              {despesa.dataReferencia || '—'}
                                            </td>
                                            <td className="py-2.5 px-3 text-slate-600 font-medium">
                                              {despesa.mesAnoLabel || '—'}
                                            </td>
                                            <td className="py-2.5 px-3">
                                              <span className={`inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-md ${fStyle.badgeBg} border ${fStyle.badgeBorder} ${fStyle.badgeText} font-semibold font-mono text-[10px]`}>
                                                <span className={`w-1.5 h-1.5 rounded-full ${fStyle.dot}`} />
                                                <span>Fonte {despesa.fonteRecurso || 'OUTRAS'}</span>
                                              </span>
                                            </td>
                                            <td className="py-2.5 px-3 text-center">
                                              <span className="px-2 py-0.5 rounded-full text-[10px] font-extrabold bg-blue-50 text-blue-800 border border-blue-200">
                                                {despesa.situacao || 'CB'}
                                              </span>
                                            </td>
                                            <td className="py-2.5 px-3 text-right font-mono font-bold text-emerald-950 tabular-nums">
                                              {formatBRL(despesa.valor)}
                                            </td>
                                          </tr>
                                        );
                                      })}
                                    </tbody>
                                  </table>
                                </div>
                              )}
                            </div>
                          </td>
                        </tr>
                      )}
                    </React.Fragment>
                  );
                })
              )}
            </tbody>

            {/* Table Footer: Column Totals */}
            {filteredSuppliers.length > 0 && (
              <tfoot className="sticky bottom-0 z-20 bg-slate-100/95 backdrop-blur-xs border-t-2 border-slate-300 font-black text-slate-900 text-xs">
                <tr>
                  <td className="py-3.5 px-3 text-center sticky left-0 bg-slate-100 border-r border-slate-200">
                    Σ
                  </td>
                  <td className="py-3.5 px-4 uppercase text-[11px] tracking-wider sticky left-12 bg-slate-100 border-r border-slate-200 shadow-sm">
                    TOTAL CONSOLIDADO ({filteredSuppliers.length} fornecedores)
                  </td>

                  {/* Totals per Fonte */}
                  {availableFontes.map((fonte) => (
                    <td
                      key={fonte}
                      className="py-3.5 px-3.5 text-right tabular-nums font-bold text-slate-900 border-r border-slate-200 whitespace-nowrap text-xs"
                    >
                      {formatBRL(columnTotals.fontes[fonte] || 0)}
                    </td>
                  ))}

                  {/* Total Pago no Mês */}
                  <td className="py-3.5 px-4 text-right tabular-nums font-black text-emerald-950 bg-emerald-100/90 border-l-2 border-emerald-400 whitespace-nowrap sticky right-0 z-10 shadow-xs text-xs">
                    <span className="inline-block px-2.5 py-1 rounded-lg bg-emerald-200/80 border border-emerald-300 font-mono font-black text-emerald-950 shadow-2xs">
                      {formatBRL(columnTotals.totalPagoNoMes)}
                    </span>
                  </td>
                </tr>
              </tfoot>
            )}
          </table>
        </div>

        {/* Pagination Bar */}
        {pageSize > 0 && totalPages > 1 && (
          <div className="p-4 border-t border-slate-100 flex flex-col sm:flex-row items-center justify-between gap-3 text-xs bg-slate-50/50">
            <span className="text-slate-500 font-medium">
              Exibindo {filteredSuppliers.length.toLocaleString('pt-BR')} fornecedores
            </span>

            <div className="flex items-center gap-1 font-bold">
              <button
                type="button"
                onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                disabled={currentPage === 1}
                className="p-1.5 rounded-lg border border-slate-200 text-slate-600 hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
                title="Página Anterior"
              >
                <ChevronLeft className="w-4 h-4" />
              </button>

              <span className="px-3 py-1 bg-white border border-slate-200 text-slate-800 rounded-lg">
                Página {currentPage} de {totalPages}
              </span>

              <button
                type="button"
                onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
                disabled={currentPage === totalPages}
                className="p-1.5 rounded-lg border border-slate-200 text-slate-600 hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
                title="Próxima Página"
              >
                <ChevronRightIcon className="w-4 h-4" />
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};
