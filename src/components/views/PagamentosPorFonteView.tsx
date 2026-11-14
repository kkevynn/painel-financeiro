import React, { useState, useMemo, useEffect, useRef } from 'react';
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
  ChevronLeft,
  ChevronRight as ChevronRightIcon,
  ChevronsUpDown,
  Settings2,
} from 'lucide-react';
import * as XLSX from 'xlsx';
import {
  OrdemBancariaItem,
  parseListarOrdemBancariaMatrixAsync,
  isObFileName,
  extractOrdensBancariasFromFileData,
  isValidFonteRecursoCode,
} from '../../utils/ordemBancariaParser';
import { readSpreadsheetAsMatrix, formatBRL, parseCurrencyValue } from '../../utils/excelParser';
import { FileData } from '../../types';
import { isAdmin } from '../../utils/authService';

interface PagamentosPorFonteViewProps {
  files?: FileData[];
  ordensBancariasList?: OrdemBancariaItem[] | null;
  ordensBancariasFileName?: string;
  onUpdateOrdensBancariasList?: (list: OrdemBancariaItem[] | null, fileName?: string) => void;
  onOpenCompanyAudit?: (supplier: { name: string; cnpj: string }) => void;
  onNavigateToImportacoes?: () => void;
  onOpenColumnMapModal?: (fileId: string) => void;
  onFileUpload?: (file: File) => void;
}

interface SupplierFonteRow {
  key: string;
  name: string;
  cnpj: string;
  qtdOrdens: number;
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

/**
 * Safely converts any value (number, string, currency string) into a finite positive number.
 */
function sanitizeNumericValue(val: unknown): number {
  if (val === null || val === undefined) return 0;
  if (typeof val === 'number') {
    return Number.isFinite(val) && !Number.isNaN(val) && val > 0 ? val : 0;
  }
  if (typeof val === 'string') {
    const trimmed = val.trim();
    if (!trimmed || trimmed === '-' || trimmed === '0' || trimmed === '0,00' || trimmed === 'R$ 0,00') {
      return 0;
    }
    const parsed = parseCurrencyValue(trimmed);
    if (typeof parsed === 'number' && Number.isFinite(parsed) && !Number.isNaN(parsed) && parsed > 0) {
      return parsed;
    }
    const cleaned = trimmed.replace(/[R$\s]/g, '').replace(/\./g, '').replace(',', '.');
    const direct = parseFloat(cleaned);
    if (typeof direct === 'number' && Number.isFinite(direct) && !Number.isNaN(direct) && direct > 0) {
      return direct;
    }
  }
  return 0;
}

export const PagamentosPorFonteView: React.FC<PagamentosPorFonteViewProps> = ({
  files = [],
  ordensBancariasList = [],
  ordensBancariasFileName = '',
  onUpdateOrdensBancariasList,
  onOpenCompanyAudit,
  onNavigateToImportacoes,
  onOpenColumnMapModal,
  onFileUpload,
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

  const userClearedRef = useRef(false);
  const autoHealedObRef = useRef<string | null>(null);

  const explicitObFile = useMemo(() => {
    if (!Array.isArray(files) || files.length === 0) return null;
    return (
      files.find(
        (f) =>
          f &&
          f.stage !== 'base_pps' &&
          (f.stage === 'ordem_bancaria' || (typeof f.fileName === 'string' && isObFileName(f.fileName)))
      ) || null
    );
  }, [files]);

  const hasValidFonteInCurrentList = useMemo(() => {
    if (!Array.isArray(ordensBancariasList) || ordensBancariasList.length === 0) return false;
    return ordensBancariasList.some((item) => {
      if (!item) return false;
      const f = String(item.fonteRecurso ?? '').trim();
      const upper = f.toUpperCase();
      if (
        !f ||
        upper === 'OUTRAS' ||
        upper === 'FONTE NÃO IDENTIFICADA' ||
        upper === 'FONTE NAO IDENTIFICADA'
      ) {
        return false;
      }
      return isValidFonteRecursoCode(f);
    });
  }, [ordensBancariasList]);

  useEffect(() => {
    userClearedRef.current = false;
  }, [explicitObFile]);

  useEffect(() => {
    if (userClearedRef.current || !explicitObFile || !onUpdateOrdensBancariasList) return;
    if (Array.isArray(ordensBancariasList) && ordensBancariasList.length > 0 && hasValidFonteInCurrentList) {
      return;
    }
    const healKey = `${explicitObFile.id}_${explicitObFile.rows?.length || 0}`;
    if (autoHealedObRef.current === healKey) return;

    const extracted = extractOrdensBancariasFromFileData(explicitObFile);
    if (extracted.length > 0 && extracted.some((it) => isValidFonteRecursoCode(it.fonteRecurso))) {
      autoHealedObRef.current = healKey;
      onUpdateOrdensBancariasList(extracted, explicitObFile.fileName);
    }
  }, [explicitObFile, ordensBancariasList, hasValidFonteInCurrentList, onUpdateOrdensBancariasList]);

  // Consolidated items: direct, fast, linear reading of ordensBancariasList (prioritizing explicit OB file if current list lacks valid fontes)
  const items = useMemo<OrdemBancariaItem[]>(() => {
    let rawList: unknown[] = [];

    if (
      explicitObFile &&
      (!Array.isArray(ordensBancariasList) ||
        ordensBancariasList.length === 0 ||
        !hasValidFonteInCurrentList)
    ) {
      rawList = extractOrdensBancariasFromFileData(explicitObFile);
    } else if (Array.isArray(ordensBancariasList) && ordensBancariasList.length > 0) {
      rawList = ordensBancariasList;
    } else if (Array.isArray(files) && files.length > 0) {
      const obFile = files.find(
        (f) =>
          f &&
          f.stage !== 'base_pps' &&
          (f.stage === 'ordem_bancaria' ||
            (typeof f.fileName === 'string' && isObFileName(f.fileName)) ||
            Boolean(f.detectedObCol && f.detectedFonteCol))
      );
      if (obFile) {
        rawList = extractOrdensBancariasFromFileData(obFile);
      }
    }

    if (!Array.isArray(rawList) || rawList.length === 0) {
      return [];
    }

    const validated: OrdemBancariaItem[] = [];
    for (let i = 0; i < rawList.length; i++) {
      const entry = rawList[i];
      if (!entry || typeof entry !== 'object') continue;

      const raw = entry as Record<string, unknown>;
      const rawVal = raw.valor ?? raw.Valor ?? raw.VALOR ?? raw.vlr ?? raw.valorPago;
      const valor = sanitizeNumericValue(rawVal);
      if (valor <= 0) continue;

      const rawOb = raw.ob ?? '';
      const cleanOb = String(rawOb).trim().toUpperCase().replace(/[\s\.\-\/]/g, '');
      const rawName = String(raw.favorecidoName ?? raw.favorecidoRaw ?? '').trim();
      const upperName = rawName.toUpperCase();

      // FILTRO RIGOROSO PARA BARRAR LINHAS TOTALIZADORAS
      if (!cleanOb || cleanOb === '-' || cleanOb === '0') continue;
      if (cleanOb.includes('TOTAL') || cleanOb.includes('RESUMO') || cleanOb.includes('SUBTOTAL')) continue;
      if (upperName.includes('TOTAL') || upperName.includes('RESUMO') || upperName.includes('SUBTOTAL')) continue;

      const favorecidoName = rawName || 'NÃO INFORMADO';
      const favorecidoCnpj = String(raw.favorecidoCnpj ?? '').trim() || 'N/I';

      const rawFonteStr = String(raw.fonteRecurso ?? '').trim();
      const fonteRecurso = isValidFonteRecursoCode(rawFonteStr)
        ? rawFonteStr
        : 'Fonte Não Identificada';
      const mesAnoKey = String(raw.mesAnoKey ?? '').trim();
      const mesAnoLabel = String(raw.mesAnoLabel ?? '').trim() || 'Sem Data';

      validated.push({
        id: String(raw.id || `ob_${cleanOb}_${i}`),
        ob: cleanOb,
        dataReferencia: String(raw.dataReferencia ?? ''),
        dataIso: typeof raw.dataIso === 'string' ? raw.dataIso : undefined,
        mesAnoKey,
        mesAnoLabel,
        valor,
        situacao: String(raw.situacao ?? 'CB'),
        pp: String(raw.pp ?? '').trim(),
        fonteRecurso,
        favorecidoRaw: String(raw.favorecidoRaw ?? ''),
        favorecidoCnpj,
        favorecidoName,
        linhaOrigem: typeof raw.linhaOrigem === 'number' ? raw.linhaOrigem : i + 1,
      });
    }

    return validated;
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

      if (onFileUpload) {
        onFileUpload(file);
      }
    } catch (err: any) {
      setErrorMessage('Erro ao processar planilha: ' + (err?.message || 'Arquivo corrompido ou formato inválido.'));
    } finally {
      setIsUploading(false);
      e.target.value = '';
    }
  };

  // Find active FileData for OB report if available in files
  const activeObFile = useMemo(() => {
    if (!Array.isArray(files) || files.length === 0) return null;
    if (ordensBancariasFileName) {
      const match = files.find((f) => f.fileName === ordensBancariasFileName);
      if (match) return match;
    }
    return (
      files.find(
        (f) =>
          f.stage === 'ordem_bancaria' ||
          isObFileName(f.fileName) ||
          Boolean(f.detectedObCol && f.detectedFonteCol)
      ) || null
    );
  }, [files, ordensBancariasFileName]);

  // Distinct list of Available Fontes de Recurso
  const availableFontes = useMemo(() => {
    const fontesSet = new Set<string>();
    items.forEach((it) => {
      const f = (it.fonteRecurso || '').trim();
      if (f && f !== '-' && f !== '0') fontesSet.add(f);
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

  // Active Month Label for headers
  const activeMonthLabel = useMemo(() => {
    if (selectedMes === 'all') return 'Todos os Meses';
    const found = availableMonths.find((m) => m.key === selectedMes);
    return found ? found.label : selectedMes;
  }, [selectedMes, availableMonths]);

  // Aggregated Supplier rows: single fast O(N) pass
  const supplierRows = useMemo<SupplierFonteRow[]>(() => {
    if (items.length === 0) return [];

    const map = new Map<string, SupplierFonteRow>();

    for (let i = 0; i < items.length; i++) {
      const it = items[i];
      const cnpjDigits = it.favorecidoCnpj.replace(/\D/g, '');
      const rawName = it.favorecidoName.trim().toUpperCase();

      const key = cnpjDigits.length >= 11 ? cnpjDigits : rawName || it.ob || String(i);

      let row = map.get(key);
      if (!row) {
        row = {
          key,
          name: it.favorecidoName,
          cnpj: it.favorecidoCnpj !== 'N/I' ? it.favorecidoCnpj : (cnpjDigits || 'N/I'),
          qtdOrdens: 0,
          totalGeral: 0,
          pagoNoMes: 0,
          fontesValues: {},
          despesas: [],
        };
        map.set(key, row);
      } else {
        if ((!row.cnpj || row.cnpj === 'N/I') && it.favorecidoCnpj && it.favorecidoCnpj !== 'N/I') {
          row.cnpj = it.favorecidoCnpj;
        }
        if ((row.name === 'NÃO INFORMADO' || row.name.length < 5) && it.favorecidoName && it.favorecidoName.length >= 5) {
          row.name = it.favorecidoName;
        }
      }

      const f = it.fonteRecurso || 'Fonte Não Identificada';
      const belongsToSelectedMonth = selectedMes === 'all' || it.mesAnoKey === selectedMes;

      row.totalGeral += it.valor;

      if (belongsToSelectedMonth) {
        row.pagoNoMes += it.valor;
        row.qtdOrdens += 1;
        row.fontesValues[f] = (row.fontesValues[f] || 0) + it.valor;
        row.despesas.push(it);
      }
    }

    return Array.from(map.values());
  }, [items, selectedMes]);

  // Filter suppliers by search query (name or CNPJ) and only show suppliers active in period
  const filteredSuppliers = useMemo(() => {
    const query = searchTerm.toLowerCase().trim();
    const cleanQuery = query.replace(/\D/g, '');

    const list = supplierRows.filter((s) => {
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
      } else if (sortField === 'cnpj') {
        const comp = a.cnpj.localeCompare(b.cnpj);
        return sortOrder === 'asc' ? comp : -comp;
      } else if (sortField === 'qtdOrdens') {
        valA = a.qtdOrdens;
        valB = b.qtdOrdens;
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

  // Totals for table footer and KPI cards
  const columnTotals = useMemo(() => {
    let totalPagoNoMes = 0;
    let totalGeral = 0;
    let totalOrdens = 0;

    for (let i = 0; i < filteredSuppliers.length; i++) {
      const s = filteredSuppliers[i];
      totalPagoNoMes += s.pagoNoMes;
      totalGeral += s.totalGeral;
      totalOrdens += s.qtdOrdens;
    }

    return {
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

  // Toggle supplier expansion (Master-Detail)
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

    // Sheet 1: Main Supplier summary with Fonte breakdown
    const matrixRows = filteredSuppliers.map((s, idx) => {
      const row: any = {
        '#': idx + 1,
        'Fornecedor / Favorecido': s.name,
        'CNPJ / CPF': s.cnpj,
        'Qtd de Ordens': s.qtdOrdens,
      };

      availableFontes.forEach((fonte) => {
        row[`Fonte ${fonte} (R$)`] = s.fontesValues[fonte] || 0;
      });

      row[`Pago no Mês (${activeMonthLabel})`] = s.pagoNoMes;
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
          {isAdmin() && (
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
          )}

          <button
            type="button"
            onClick={handleExportExcel}
            className="inline-flex items-center gap-2 px-4 py-2.5 bg-white/10 hover:bg-white/20 text-white rounded-xl text-xs font-bold transition cursor-pointer active:scale-98"
            title="Exportar fornecedores e detalhamento para Excel"
          >
            <Download className="w-4 h-4 text-emerald-300" />
            <span>Exportar Excel</span>
          </button>

          {isAdmin() && activeObFile && onOpenColumnMapModal && (
            <button
              type="button"
              onClick={() => onOpenColumnMapModal(activeObFile.id)}
              className="inline-flex items-center gap-2 px-4 py-2.5 bg-emerald-700 hover:bg-emerald-600 text-white rounded-xl text-xs font-bold transition shadow-md cursor-pointer active:scale-98"
              title="Ajustar manualmente as colunas dos dados a serem analisados"
            >
              <Settings2 className="w-4 h-4 text-emerald-200" />
              <span>Ajustar Colunas</span>
            </button>
          )}

          {isAdmin() && onUpdateOrdensBancariasList && (
            confirmClear ? (
              <div className="flex items-center gap-1.5 bg-rose-950/70 border border-rose-700/60 rounded-xl px-2.5 py-1.5 animate-in fade-in duration-100">
                <span className="text-[11px] text-rose-300 font-semibold">Remover OBs?</span>
                <button
                  type="button"
                  onClick={() => {
                    userClearedRef.current = true;
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
              {selectedMes === 'all' ? 'Total Pago (Acumulado)' : `Total Pago no Mês (${activeMonthLabel})`}
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

        {/* Fontes em Operação */}
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

            {/* Expand / Collapse All Accordions Button */}
            <button
              type="button"
              onClick={handleToggleExpandAll}
              className="px-3 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold transition flex items-center gap-1.5 cursor-pointer active:scale-98"
              title={isAllExpanded ? 'Recolher fontes de todos os fornecedores' : 'Expandir fontes de todos os fornecedores'}
            >
              <ChevronsUpDown className="w-3.5 h-3.5 text-slate-500" />
              <span>{isAllExpanded ? 'Recolher Fontes' : 'Expandir Todas as Fontes'}</span>
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

      {/* ---------------- MAIN TABLE (MASTER-DETAIL APPROACH: APENAS 4 COLUNAS PRINCIPAIS + AÇÃO) ---------------- */}
      <div className="bg-white rounded-3xl border border-slate-200/80 shadow-sm overflow-hidden">
        <div className="overflow-x-auto max-h-[700px] overflow-y-auto">
          <table className="w-full text-left border-collapse text-xs">
            <thead className="sticky top-0 z-20 bg-slate-100/95 backdrop-blur-xs shadow-xs">
              <tr className="border-b border-slate-200 text-slate-700 font-bold uppercase tracking-wider text-[11px]">
                {/* Index */}
                <th className="py-3 px-3 w-12 text-center text-slate-400">
                  #
                </th>

                {/* Fornecedor */}
                <th className="py-3 px-4 min-w-[240px]">
                  <button
                    type="button"
                    onClick={() => handleSort('name')}
                    className="flex items-center gap-1.5 hover:text-emerald-700 cursor-pointer font-extrabold active:scale-98"
                  >
                    <span>FORNECEDOR</span>
                    <ArrowUpDown className="w-3.5 h-3.5 text-slate-400" />
                  </button>
                </th>

                {/* CNPJ */}
                <th className="py-3 px-4 min-w-[160px]">
                  <button
                    type="button"
                    onClick={() => handleSort('cnpj')}
                    className="flex items-center gap-1.5 hover:text-emerald-700 cursor-pointer font-extrabold active:scale-98"
                  >
                    <span>CNPJ</span>
                    <ArrowUpDown className="w-3.5 h-3.5 text-slate-400" />
                  </button>
                </th>

                {/* Qtd de Ordens */}
                <th className="py-3 px-4 text-center min-w-[140px]">
                  <button
                    type="button"
                    onClick={() => handleSort('qtdOrdens')}
                    className="inline-flex items-center gap-1.5 hover:text-emerald-700 cursor-pointer font-extrabold active:scale-98"
                  >
                    <span>QTD DE ORDENS</span>
                    <ArrowUpDown className="w-3.5 h-3.5 text-slate-400" />
                  </button>
                </th>

                {/* Total Pago no Mês */}
                <th className="py-3 px-4 text-right min-w-[180px]">
                  <button
                    type="button"
                    onClick={() => handleSort('pagoNoMes')}
                    className="inline-flex items-center justify-end gap-1.5 w-full hover:text-emerald-700 cursor-pointer font-extrabold active:scale-98"
                  >
                    <span>{selectedMes === 'all' ? 'TOTAL PAGO (ACUMULADO)' : `TOTAL PAGO NO MÊS (${activeMonthLabel.toUpperCase()})`}</span>
                    <ArrowUpDown className="w-3.5 h-3.5 text-emerald-700" />
                  </button>
                </th>

                {/* Ação: Detalhar Fontes */}
                <th className="py-3 px-4 text-center min-w-[150px]">
                  <span>FONTES DE RECURSO</span>
                </th>
              </tr>
            </thead>

            <tbody className="divide-y divide-slate-100 font-medium">
              {paginatedSuppliers.length === 0 ? (
                <tr>
                  <td
                    colSpan={6}
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
                        <td className="py-3 px-3 text-center font-mono text-[11px] text-slate-400">
                          {realIndex}
                        </td>

                        {/* Fornecedor */}
                        <td className="py-3 px-4 font-bold text-slate-900">
                          <button
                            type="button"
                            onClick={() =>
                              onOpenCompanyAudit &&
                              onOpenCompanyAudit({ name: supplier.name, cnpj: supplier.cnpj })
                            }
                            className="flex items-center gap-2 text-left hover:text-emerald-700 transition cursor-pointer group"
                            title="Abrir Extrato de Auditoria do Fornecedor"
                          >
                            <Building className="w-3.5 h-3.5 text-slate-400 group-hover:text-emerald-600 shrink-0" />
                            <span className="font-bold text-xs group-hover:underline">
                              {supplier.name}
                            </span>
                          </button>
                        </td>

                        {/* CNPJ */}
                        <td className="py-3 px-4 font-mono text-slate-600 font-medium text-xs whitespace-nowrap">
                          {supplier.cnpj}
                        </td>

                        {/* Qtd de Ordens */}
                        <td className="py-3 px-4 text-center">
                          <span className="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-bold bg-slate-100 text-slate-700 border border-slate-200/80">
                            {supplier.qtdOrdens} {supplier.qtdOrdens === 1 ? 'ordem' : 'ordens'}
                          </span>
                        </td>

                        {/* Total Pago no Mês */}
                        <td className="py-3 px-4 text-right tabular-nums font-bold text-emerald-950 whitespace-nowrap text-xs">
                          <span className="inline-block px-2.5 py-1 rounded-lg bg-emerald-50 border border-emerald-200/80 font-mono font-bold text-emerald-950 shadow-2xs">
                            {formatBRL(supplier.pagoNoMes)}
                          </span>
                        </td>

                        {/* Botão Detalhar Fontes */}
                        <td className="py-3 px-4 text-center whitespace-nowrap">
                          <button
                            type="button"
                            onClick={() => toggleSupplierExpand(supplier.key)}
                            className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-bold transition cursor-pointer active:scale-98 ${
                              isExpanded
                                ? 'bg-emerald-600 text-white shadow-sm'
                                : 'bg-slate-100 hover:bg-slate-200 text-slate-700'
                            }`}
                            title={isExpanded ? 'Ocultar divisão de fontes' : 'Ver divisão de valores por fonte'}
                          >
                            <Layers className="w-3.5 h-3.5" />
                            <span>Detalhar Fontes</span>
                            {isExpanded ? (
                              <ChevronDown className="w-3.5 h-3.5" />
                            ) : (
                              <ChevronRight className="w-3.5 h-3.5" />
                            )}
                          </button>
                        </td>
                      </tr>

                      {/* ---------------- EXPANDED DETAIL: DIVISÃO DE VALORES POR FONTE APENAS PARA ESTE FORNECEDOR ---------------- */}
                      {isExpanded && (
                        <tr className="bg-emerald-50/20">
                          <td
                            colSpan={6}
                            className="p-4 sm:p-6 border-b border-emerald-200/60"
                          >
                            <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm space-y-4">
                              {/* Header do Detalhamento */}
                              <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 border-b border-slate-100 pb-3">
                                <div className="flex items-center gap-2">
                                  <Layers className="w-4 h-4 text-emerald-600" />
                                  <h4 className="text-xs font-black text-slate-900">
                                    Divisão de Pagamentos por Fonte de Recursos
                                  </h4>
                                  <span className="text-[11px] font-mono text-slate-500">
                                    ({supplier.name} — {supplier.cnpj})
                                  </span>
                                </div>
                                <div className="text-xs tabular-nums font-bold text-slate-700">
                                  <span>Total Pago: </span>
                                  <span className="text-emerald-900 font-black font-mono text-sm">
                                    {formatBRL(supplier.pagoNoMes)}
                                  </span>
                                </div>
                              </div>

                              {/* Badges/Cards de cada Fonte usada por este fornecedor */}
                              <div className="space-y-1.5">
                                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider block">
                                  Fontes Utilizadas
                                </span>
                                <div className="flex flex-wrap items-center gap-2.5">
                                  {Object.keys(supplier.fontesValues).length === 0 ? (
                                    <span className="text-xs text-slate-400 italic">Nenhuma fonte vinculada.</span>
                                  ) : (
                                    Object.entries(supplier.fontesValues).map(([fonte, val]) => {
                                      const style = getFonteColorStyle(fonte);
                                      const pct = supplier.pagoNoMes > 0 ? ((val / supplier.pagoNoMes) * 100).toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) : '0,0';
                                      return (
                                        <div
                                          key={fonte}
                                          className={`flex items-center gap-2 px-3.5 py-2 rounded-xl border ${style.badgeBg} ${style.badgeBorder} ${style.badgeText} text-xs font-semibold shadow-2xs`}
                                        >
                                          <span className={`w-2 h-2 rounded-full ${style.dot} shrink-0`} />
                                          <span className="font-mono font-bold">Fonte {fonte}:</span>
                                          <span className="font-mono font-black">{formatBRL(val)}</span>
                                          <span className="text-[10px] opacity-75">({pct}%)</span>
                                        </div>
                                      );
                                    })
                                  )}
                                </div>
                              </div>

                              {/* Lista de Despesas Individuais do Fornecedor */}
                              {supplier.despesas.length > 0 && (
                                <div className="pt-2 space-y-1.5">
                                  <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider block">
                                    Ordens Bancárias Individuais ({supplier.despesas.length})
                                  </span>
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
                                                <span
                                                  className={`inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-md ${fStyle.badgeBg} border ${fStyle.badgeBorder} ${fStyle.badgeText} font-semibold font-mono text-[10px]`}
                                                >
                                                  <span className={`w-1.5 h-1.5 rounded-full ${fStyle.dot}`} />
                                                  <span>Fonte {despesa.fonteRecurso || 'Fonte Não Identificada'}</span>
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

            {/* Table Footer: Totals */}
            {filteredSuppliers.length > 0 && (
              <tfoot className="sticky bottom-0 z-20 bg-slate-100/95 backdrop-blur-xs border-t-2 border-slate-300 font-black text-slate-900 text-xs">
                <tr>
                  <td className="py-3.5 px-3 text-center text-slate-400">
                    Σ
                  </td>
                  <td className="py-3.5 px-4 uppercase text-[11px] tracking-wider">
                    TOTAL CONSOLIDADO ({filteredSuppliers.length} fornecedores)
                  </td>
                  <td className="py-3.5 px-4 text-slate-400">
                    —
                  </td>
                  <td className="py-3.5 px-4 text-center font-mono font-bold text-slate-800">
                    {columnTotals.totalOrdens} {columnTotals.totalOrdens === 1 ? 'ordem' : 'ordens'}
                  </td>
                  <td className="py-3.5 px-4 text-right tabular-nums font-black text-emerald-950">
                    <span className="inline-block px-3 py-1 rounded-lg bg-emerald-200/80 border border-emerald-300 font-mono font-black text-emerald-950 shadow-2xs">
                      {formatBRL(columnTotals.totalPagoNoMes)}
                    </span>
                  </td>
                  <td className="py-3.5 px-4 text-center text-slate-400">
                    —
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
