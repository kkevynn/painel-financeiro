import React, { useState, useMemo, useRef } from 'react';
import {
  FileText,
  Upload,
  Download,
  CheckCircle2,
  Clock,
  Search,
  Check,
  ArrowRight,
  TrendingUp,
  AlertCircle,
  FileSpreadsheet,
  Layers,
  Calendar,
  AlertTriangle,
  Zap,
  Link2,
  X,
  Sparkles,
} from 'lucide-react';
import { FileData, PublicacaoItem } from '../../types';
import { OrdemBancariaItem } from '../../utils/ordemBancariaParser';
import { analyzePublicacaoRecords, saveManualPublicacaoResolution } from '../../utils/publicacaoService';
import { exportPublicacaoToExcel, exportPublicacaoToPdf } from '../../utils/publicacaoExporter';
import { formatBRL } from '../../utils/excelParser';
import { SuggestionInput } from '../common/SuggestionInput';
import { resolveCnpj, resolveSupplierName } from '../../utils/appKnowledgeDirectory';
import {
  extractAvailableNls,
  extractAvailablePPs,
  extractAvailableOBs,
  searchNls,
  searchPPs,
  searchOBs,
  recordCustomPP,
  recordCustomOB,
  cleanDoc,
  NlSuggestionItem,
  PpSuggestionItem,
  ObSuggestionItem,
} from '../../utils/suggestionService';

interface PublicacaoViewProps {
  files: FileData[];
  ordensBancariasList?: OrdemBancariaItem[] | null;
  onFileUpload: (file: File, stage?: string) => void;
  onNavigateToImportacoes: () => void;
}

export const PublicacaoView: React.FC<PublicacaoViewProps> = ({
  files,
  ordensBancariasList,
  onFileUpload,
  onNavigateToImportacoes,
}) => {
  const [searchTerm, setSearchTerm] = useState('');
  // Filter options:
  // 'todos': All items
  // 'pendentes_obedece': Specifically NLs that are STILL listed in Obedece and went to publication (actionable items)
  // 'concluido': Completed/resolved items (PP/OB emitted in AO)
  // 'todas_pendentes': All pending items
  const [statusFilter, setStatusFilter] = useState<'todos' | 'pendentes_obedece' | 'concluido' | 'todas_pendentes'>('todos');
  const [selectedDia, setSelectedDia] = useState<string>('todos');
  const [isExporting, setIsExporting] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Manual PP/OB binding modal state
  const [manualModalItem, setManualModalItem] = useState<PublicacaoItem | null>(null);
  const [manualNL, setManualNL] = useState('');
  const [manualCredor, setManualCredor] = useState('');
  const [manualCnpj, setManualCnpj] = useState('');
  const [manualPP, setManualPP] = useState('');
  const [manualOB, setManualOB] = useState('');
  const [manualSituacao, setManualSituacao] = useState<'AO' | 'PPCB' | 'CB'>('AO');
  const [manualValorNL, setManualValorNL] = useState<string>('');
  const [manualValor, setManualValor] = useState<string>('');
  const [manualFonte, setManualFonte] = useState<string>('0.600');
  const [fonteAutoDetected, setFonteAutoDetected] = useState<boolean>(false);
  const [refreshTrigger, setRefreshTrigger] = useState(0);

  // Analyze records from uploaded files (with persistent resolution caching)
  const publicacaoItems = useMemo(() => {
    return analyzePublicacaoRecords(files, ordensBancariasList);
  }, [files, ordensBancariasList, refreshTrigger]);

  // Extract all available suggestions across files and OB lists
  const allAvailableNls = useMemo(() => extractAvailableNls(files), [files]);
  const allAvailablePPs = useMemo(() => extractAvailablePPs(files, ordensBancariasList), [files, ordensBancariasList]);
  const allAvailableOBs = useMemo(() => extractAvailableOBs(files, ordensBancariasList), [files, ordensBancariasList]);

  // Search filtered suggestions in real-time as user types
  const nlSuggestions = useMemo(() => {
    return searchNls(manualNL, allAvailableNls, 10).map((it) => {
      const docCnpj = it.cnpj || resolveCnpj(it.fornecedor, it.nlNumero);
      return {
        id: `nl_${it.cleanDoc}`,
        title: it.nlNumero,
        badge: it.fonteLabel || (it.fonteTipo === 'federal' ? 'Fonte 0.600 (Fed)' : 'Fonte 0.500 (Est)'),
        badgeColor:
          it.fonteTipo === 'federal'
            ? 'bg-blue-50 text-blue-900 border-blue-200'
            : 'bg-emerald-50 text-emerald-900 border-emerald-200',
        subtitle: `${it.fornecedor}${docCnpj ? ` • CNPJ: ${docCnpj}` : ''} • ${
          it.stage === 'obedece' ? 'Fila Obedece' : 'Não Obedece'
        }`,
        extraInfo: formatBRL(it.valor),
        data: it,
      };
    });
  }, [manualNL, allAvailableNls]);

  const ppSuggestions = useMemo(() => {
    return searchPPs(manualPP, allAvailablePPs, 10).map((it) => ({
      id: `pp_${it.cleanNum}`,
      title: it.ppNumero,
      badge: it.situacao || 'AO',
      badgeColor: it.situacao?.includes('CB')
        ? 'bg-emerald-50 text-emerald-800 border-emerald-200'
        : 'bg-blue-50 text-blue-800 border-blue-200',
      subtitle: `${it.fornecedor || 'Credor'} • ${it.obNumero ? `OB: ${it.obNumero}` : 'Sem OB associada'}`,
      extraInfo: it.valor ? formatBRL(it.valor) : undefined,
      data: it,
    }));
  }, [manualPP, allAvailablePPs]);

  const obSuggestions = useMemo(() => {
    return searchOBs(manualOB, allAvailableOBs, 10).map((it) => ({
      id: `ob_${it.cleanNum}`,
      title: it.obNumero,
      badge: it.situacao || 'CB',
      badgeColor: 'bg-teal-50 text-teal-800 border-teal-200',
      subtitle: `${it.fornecedor || 'Credor'} • ${it.ppNumero ? `PP: ${it.ppNumero}` : 'Sem PP vinculada'}`,
      extraInfo: it.valor ? formatBRL(it.valor) : undefined,
      data: it,
    }));
  }, [manualOB, allAvailableOBs]);

  const handleOpenManualModal = (item: PublicacaoItem) => {
    setManualModalItem(item);
    setManualNL(item.nlNumero || '');
    setManualCredor(item.fornecedor || '');
    const resolvedCnpjVal = item.cnpj || resolveCnpj(item.fornecedor, item.nlNumero);
    setManualCnpj(resolvedCnpjVal || '');
    setManualPP(item.equivalentePP || '');
    setManualOB(item.equivalenteOB || '');
    const sitNorm = (item.situacaoEquivalente || '').toUpperCase();
    setManualSituacao(sitNorm.includes('CB') ? 'CB' : sitNorm.includes('PPCB') ? 'PPCB' : 'AO');
    setManualValorNL(item.valorNL ? String(item.valorNL) : '');
    setManualValor(item.valorEquivalente ? String(item.valorEquivalente) : String(item.valorNL || ''));

    // Check if imported Obedece/Não Obedece files contain this NL and pull its Fonte automatically
    let detectedFonte = item.fonteRecurso || '';
    let autoDet = false;
    const cleanItemNl = cleanDoc(item.nlNumero);
    const foundNl = allAvailableNls.find((n) => n.cleanDoc === cleanItemNl);
    if (foundNl && foundNl.fonteRecurso) {
      detectedFonte = foundNl.fonteRecurso;
      autoDet = true;
    }
    setManualFonte(detectedFonte || '0.600');
    setFonteAutoDetected(autoDet || !!item.fonteRecurso);
  };

  const handleSelectNlSuggestion = (selected: NlSuggestionItem) => {
    setManualNL(selected.nlNumero);
    setManualCredor(selected.fornecedor);
    if (selected.cnpj) setManualCnpj(selected.cnpj);
    if (selected.valor > 0) {
      setManualValorNL(String(selected.valor));
      if (!manualValor || manualValor === '0') {
        setManualValor(String(selected.valor));
      }
    }
    // Auto-pull Fonte de Recursos without requiring manual selection
    if (selected.fonteRecurso) {
      setManualFonte(selected.fonteRecurso);
      setFonteAutoDetected(true);
    }
    // Auto-select OB or PP if detected in records
    if (selected.suggestedOB && !manualOB) {
      setManualOB(selected.suggestedOB);
    }
    if (selected.suggestedPP && !manualPP) {
      setManualPP(selected.suggestedPP);
    }
  };

  const handleSelectPpSuggestion = (selected: PpSuggestionItem) => {
    setManualPP(selected.ppNumero);
    // Auto-select linked OB if available
    if (selected.obNumero) {
      setManualOB(selected.obNumero);
    }
    if (selected.valor && selected.valor > 0) {
      setManualValor(String(selected.valor));
    }
    if (selected.situacao) {
      const sitNorm = selected.situacao.toUpperCase();
      setManualSituacao(sitNorm.includes('CB') ? 'CB' : sitNorm.includes('PPCB') ? 'PPCB' : 'AO');
    }
  };

  const handleSelectObSuggestion = (selected: ObSuggestionItem) => {
    setManualOB(selected.obNumero);
    // Auto-select linked PP if available
    if (selected.ppNumero && (!manualPP || manualPP === '-')) {
      setManualPP(selected.ppNumero);
    }
    if (selected.valor && selected.valor > 0 && (!manualValor || manualValor === '0')) {
      setManualValor(String(selected.valor));
    }
    if (selected.situacao) {
      const sitNorm = selected.situacao.toUpperCase();
      setManualSituacao(sitNorm.includes('AO') ? 'AO' : 'CB');
    }
  };

  const handleSaveManualBinding = (e: React.FormEvent) => {
    e.preventDefault();
    if (!manualModalItem) return;
    const cleanValNL = manualValorNL.replace(/\./g, '').replace(',', '.');
    const parsedValorNL = parseFloat(cleanValNL) || manualModalItem.valorNL;

    const cleanVal = manualValor.replace(/\./g, '').replace(',', '.');
    const valorNum = parseFloat(cleanVal) || parsedValorNL;

    const finalNL = manualNL.trim() || manualModalItem.nlNumero;
    const finalCredor = manualCredor.trim() || manualModalItem.fornecedor;
    const finalCnpj = manualCnpj.trim() || manualModalItem.cnpj;

    saveManualPublicacaoResolution({
      nlNumero: finalNL,
      cnpj: finalCnpj,
      fornecedor: finalCredor,
      valorNL: parsedValorNL,
      ppNumero: manualPP,
      obNumero: manualOB,
      situacao: manualSituacao,
      valorPago: valorNum,
      fonteRecurso: manualFonte.trim(),
    });

    // Record user-entered PP and OB so the app remembers and auto-suggests them everywhere
    if (manualPP.trim()) {
      recordCustomPP({
        ppNumero: manualPP.trim(),
        obNumero: manualOB.trim() || undefined,
        valor: valorNum,
        fornecedor: finalCredor,
        cnpj: finalCnpj,
        situacao: manualSituacao,
      });
    }
    if (manualOB.trim()) {
      recordCustomOB({
        obNumero: manualOB.trim(),
        ppNumero: manualPP.trim() || undefined,
        valor: valorNum,
        fornecedor: finalCredor,
        cnpj: finalCnpj,
        situacao: manualSituacao,
      });
    }

    setRefreshTrigger((prev) => prev + 1);
    setManualModalItem(null);
  };

  // Publicacao files list
  const publicacaoFiles = useMemo(() => {
    return files.filter((f) => f.stage === 'publicacao');
  }, [files]);

  // Compute available days with statistics
  const availableDias = useMemo(() => {
    const map = new Map<
      string,
      {
        dia: string;
        total: number;
        pendentesObedece: number;
        concluidos: number;
        todasPendentes: number;
        valorTotal: number;
        valorPendenteObedece: number;
      }
    >();

    publicacaoItems.forEach((it) => {
      const d = it.diaPublicacao || 'Geral';
      const cur = map.get(d) || {
        dia: d,
        total: 0,
        pendentesObedece: 0,
        concluidos: 0,
        todasPendentes: 0,
        valorTotal: 0,
        valorPendenteObedece: 0,
      };

      cur.total += 1;
      cur.valorTotal += it.valorNL || 0;

      if (it.status === 'concluido') {
        cur.concluidos += 1;
      } else {
        cur.todasPendentes += 1;
        // Strictly check if it is STILL in Obedece (actionable pending)
        if (!it.sumiuObedece) {
          cur.pendentesObedece += 1;
          cur.valorPendenteObedece += it.valorNL || 0;
        }
      }

      map.set(d, cur);
    });

    return Array.from(map.values()).sort((a, b) => {
      if (a.dia === 'Geral') return 1;
      if (b.dia === 'Geral') return -1;
      return b.dia.localeCompare(a.dia);
    });
  }, [publicacaoItems]);

  // Items for the currently selected day (for day-level KPIs)
  const itemsForActiveDia = useMemo(() => {
    if (selectedDia === 'todos') return publicacaoItems;
    return publicacaoItems.filter((i) => (i.diaPublicacao || 'Geral') === selectedDia);
  }, [publicacaoItems, selectedDia]);

  // Filtered items based on day, status and search
  const filteredItems = useMemo(() => {
    return publicacaoItems.filter((item) => {
      // Day filter
      if (selectedDia !== 'todos') {
        const itemDia = item.diaPublicacao || 'Geral';
        if (itemDia !== selectedDia) return false;
      }

      // Status filter
      if (statusFilter === 'pendentes_obedece') {
        // User rule: "ajuste a busca se ha pendencias apenas das nls que ainda estao listadas em obedece e que tambem foram para publicação para que possamos ir nelas para resolver"
        if (item.status !== 'pendente' || item.sumiuObedece) {
          return false;
        }
      } else if (statusFilter === 'concluido') {
        if (item.status !== 'concluido') return false;
      } else if (statusFilter === 'todas_pendentes') {
        if (item.status !== 'pendente') return false;
      }

      // Search term
      if (searchTerm.trim()) {
        const term = searchTerm.toLowerCase().trim();
        const matchesNL = item.nlNumero.toLowerCase().includes(term);
        const matchesName = item.fornecedor.toLowerCase().includes(term);
        const matchesCnpj = item.cnpj.replace(/\D/g, '').includes(term.replace(/\D/g, ''));
        const matchesPP = item.equivalentePP ? item.equivalentePP.toLowerCase().includes(term) : false;
        const matchesOB = item.equivalenteOB ? item.equivalenteOB.toLowerCase().includes(term) : false;
        const matchesDia = item.diaPublicacao ? item.diaPublicacao.toLowerCase().includes(term) : false;

        if (!matchesNL && !matchesName && !matchesCnpj && !matchesPP && !matchesOB && !matchesDia) {
          return false;
        }
      }

      return true;
    });
  }, [publicacaoItems, selectedDia, statusFilter, searchTerm]);

  // KPI calculations
  const totalCount = itemsForActiveDia.length;
  const concluidos = itemsForActiveDia.filter((i) => i.status === 'concluido');
  const pendentesObedece = itemsForActiveDia.filter((i) => i.status === 'pendente' && !i.sumiuObedece);
  const todasPendentes = itemsForActiveDia.filter((i) => i.status === 'pendente');

  const totalValorNL = itemsForActiveDia.reduce((acc, i) => acc + (i.valorNL || 0), 0);
  const valorConcluido = concluidos.reduce((acc, i) => acc + (i.valorNL || 0), 0);
  const valorPendenteObedece = pendentesObedece.reduce((acc, i) => acc + (i.valorNL || 0), 0);

  const taxaEfetivacao = totalCount > 0 ? (concluidos.length / totalCount) * 100 : 0;

  // Global pending counts across all days
  const totalGlobalPendentesObedece = useMemo(() => {
    return publicacaoItems.filter((i) => i.status === 'pendente' && !i.sumiuObedece).length;
  }, [publicacaoItems]);

  // Handle direct file upload for publicacao
  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      Array.from(e.target.files).forEach((f) => onFileUpload(f, 'publicacao'));
      e.target.value = '';
    }
  };

  const handleExportPdf = () => {
    setIsExporting(true);
    try {
      const suffix = selectedDia !== 'todos' ? `_Dia_${selectedDia.replace(/\//g, '-')}` : '_Todos_Dias';
      exportPublicacaoToPdf(filteredItems, `Acompanhamento_NLs_Publicacao${suffix}`);
    } finally {
      setIsExporting(false);
    }
  };

  const handleExportExcel = () => {
    setIsExporting(true);
    try {
      const suffix = selectedDia !== 'todos' ? `_Dia_${selectedDia.replace(/\//g, '-')}` : '_Todos_Dias';
      exportPublicacaoToExcel(filteredItems, `Acompanhamento_NLs_Publicacao${suffix}`);
    } finally {
      setIsExporting(false);
    }
  };

  return (
    <div className="space-y-6 animate-in fade-in-50 duration-200">
      {/* Top Institutional Header Banner */}
      <div className="bg-[#061d15] text-white rounded-3xl p-6 sm:p-8 border border-emerald-900/80 shadow-xl flex flex-col md:flex-row items-start md:items-center justify-between gap-6">
        <div className="space-y-1.5">
          <div className="flex items-center gap-2">
            <span className="px-2.5 py-0.5 rounded-lg bg-emerald-500/20 text-emerald-300 text-[10px] font-black uppercase tracking-wider border border-emerald-500/30">
              Controle DOE & Publicação
            </span>
            <span className="text-xs text-emerald-400 font-semibold">• FES / SESAP</span>
          </div>
          <h2 className="text-2xl sm:text-3xl font-black tracking-tight text-white">
            Publicação
          </h2>
          <p className="text-xs text-emerald-200/80 max-w-2xl font-medium">
            Separação diária das Liquidações enviadas para publicação. O sistema monitora quais NLs ainda constam na
            fila de <strong>Obedece</strong> para resolução imediata e localiza PPs e OBs emitidas em situação <strong>AO</strong> no SIGEF.
          </p>
        </div>

        {/* Action Buttons */}
        <div className="flex flex-wrap items-center gap-2.5">
          <input
            type="file"
            ref={fileInputRef}
            accept=".xlsx, .xls, .csv"
            multiple
            className="hidden"
            onChange={handleFileChange}
          />

          <button
            type="button"
            onClick={() => fileInputRef.current?.click()}
            className="inline-flex items-center gap-2 px-4 py-2.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-xs font-bold transition shadow-lg shadow-emerald-950/20 cursor-pointer active:scale-98"
            title="Importar planilha de NLs para Publicação (ex: Publicação - 05/11.xlsx)"
          >
            <Upload className="w-4 h-4" />
            <span>Adicionar Planilha de Publicação</span>
          </button>

          <button
            type="button"
            onClick={handleExportPdf}
            disabled={filteredItems.length === 0 || isExporting}
            className={`inline-flex items-center gap-2 px-4 py-2.5 rounded-xl text-xs font-bold transition shadow-md active:scale-98 cursor-pointer ${
              filteredItems.length > 0 && !isExporting
                ? 'bg-rose-700 hover:bg-rose-600 text-white shadow-rose-950/20'
                : 'bg-slate-800 text-slate-500 cursor-not-allowed opacity-60'
            }`}
            title="Exportar Relatório em PDF no padrão institucional SIGEF"
          >
            <FileText className="w-4 h-4 text-rose-200" />
            <span>Exportar PDF (SIGEF)</span>
          </button>

          <button
            type="button"
            onClick={handleExportExcel}
            disabled={filteredItems.length === 0 || isExporting}
            className={`inline-flex items-center gap-2 px-4 py-2.5 rounded-xl text-xs font-bold transition shadow-md active:scale-98 cursor-pointer ${
              filteredItems.length > 0 && !isExporting
                ? 'bg-emerald-800/80 hover:bg-emerald-700 text-white border border-emerald-600/40'
                : 'bg-slate-800 text-slate-500 cursor-not-allowed opacity-60'
            }`}
            title="Exportar dados completos para planilha Excel"
          >
            <Download className="w-4 h-4 text-emerald-300" />
            <span>Exportar Excel</span>
          </button>
        </div>
      </div>

      {/* Day Selector Navigation Bar */}
      {availableDias.length > 0 && (
        <div className="bg-white rounded-2xl p-4 border border-slate-200/90 shadow-xs space-y-3">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
            <div className="flex items-center gap-2 text-xs font-bold text-slate-800">
              <Calendar className="w-4 h-4 text-emerald-700" />
              <span>Separação por Dia de Publicação:</span>
              <span className="text-[11px] font-normal text-slate-500">
                (Clique para filtrar as NLs publicadas em cada data)
              </span>
            </div>
            {totalGlobalPendentesObedece > 0 && (
              <button
                type="button"
                onClick={() => setStatusFilter('pendentes_obedece')}
                className="inline-flex items-center gap-1.5 px-3 py-1 rounded-xl bg-amber-500 hover:bg-amber-600 text-slate-950 text-xs font-black shadow-xs transition cursor-pointer active:scale-98"
                title="Ver apenas as NLs que ainda estão listadas em Obedece para resolver"
              >
                <AlertTriangle className="w-3.5 h-3.5" />
                <span>{totalGlobalPendentesObedece} pendentes em Obedece a resolver</span>
              </button>
            )}
          </div>

          {/* Day Pills */}
          <div className="flex flex-wrap items-center gap-2 pt-1">
            {/* "Todos os Dias" Pill */}
            <button
              type="button"
              onClick={() => setSelectedDia('todos')}
              className={`px-3.5 py-2 rounded-xl text-xs font-bold transition cursor-pointer flex items-center gap-2 ${
                selectedDia === 'todos'
                  ? 'bg-slate-900 text-white shadow-sm ring-2 ring-slate-900/20'
                  : 'bg-slate-100 hover:bg-slate-200 text-slate-700'
              }`}
            >
              <span>Todos os Dias</span>
              <span
                className={`px-1.5 py-0.2 rounded-md text-[10px] font-black ${
                  selectedDia === 'todos' ? 'bg-slate-800 text-slate-200' : 'bg-slate-200 text-slate-700'
                }`}
              >
                {publicacaoItems.length}
              </span>
            </button>

            {/* Individual Day Pills */}
            {availableDias.map((d) => {
              const isSelected = selectedDia === d.dia;
              const hasPendentesObedece = d.pendentesObedece > 0;

              return (
                <button
                  key={d.dia}
                  type="button"
                  onClick={() => setSelectedDia(d.dia)}
                  className={`px-3.5 py-2 rounded-xl text-xs font-bold transition cursor-pointer flex items-center gap-2 ${
                    isSelected
                      ? 'bg-emerald-800 text-white shadow-sm ring-2 ring-emerald-800/30'
                      : 'bg-white hover:bg-slate-50 border border-slate-200 text-slate-800'
                  }`}
                >
                  <span className="font-mono">{d.dia}</span>
                  <span
                    className={`px-1.5 py-0.2 rounded-md text-[10px] font-black ${
                      isSelected ? 'bg-emerald-950 text-emerald-200' : 'bg-slate-100 text-slate-700'
                    }`}
                  >
                    {d.total}
                  </span>
                  {hasPendentesObedece ? (
                    <span
                      className={`inline-flex items-center gap-0.5 px-1.5 py-0.2 rounded text-[10px] font-black ${
                        isSelected ? 'bg-amber-400 text-amber-950' : 'bg-amber-100 text-amber-900'
                      }`}
                      title={`${d.pendentesObedece} NL(s) ainda listadas em Obedece para resolver`}
                    >
                      <Clock className="w-2.5 h-2.5" />
                      <span>{d.pendentesObedece} em Obedece</span>
                    </span>
                  ) : (
                    <span
                      className={`inline-flex items-center gap-0.5 px-1.5 py-0.2 rounded text-[10px] font-bold ${
                        isSelected ? 'bg-emerald-400/20 text-emerald-200' : 'bg-emerald-50 text-emerald-700'
                      }`}
                      title="Nenhuma pendência na fila Obedece para esta data"
                    >
                      <Check className="w-2.5 h-2.5" />
                    </span>
                  )}
                </button>
              );
            })}
          </div>
        </div>
      )}

      {/* Selected Day Status Alert Banner */}
      {selectedDia !== 'todos' && (
        <div
          className={`rounded-2xl p-4 border flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 shadow-xs ${
            pendentesObedece.length > 0
              ? 'bg-amber-50/90 border-amber-300 text-amber-950'
              : 'bg-emerald-50/80 border-emerald-200 text-emerald-950'
          }`}
        >
          <div className="flex items-start gap-3">
            <div
              className={`w-8 h-8 rounded-xl flex items-center justify-center shrink-0 mt-0.5 ${
                pendentesObedece.length > 0 ? 'bg-amber-200 text-amber-900' : 'bg-emerald-100 text-emerald-800'
              }`}
            >
              {pendentesObedece.length > 0 ? (
                <AlertTriangle className="w-4 h-4" />
              ) : (
                <CheckCircle2 className="w-4 h-4" />
              )}
            </div>
            <div>
              <div className="font-bold text-xs sm:text-sm">
                {pendentesObedece.length > 0
                  ? `Dia ${selectedDia}: Constam ${pendentesObedece.length} NL(s) ainda listadas em Obedece a resolver`
                  : `Dia ${selectedDia}: Nenhuma NL pendente na fila Obedece`}
              </div>
              <div className="text-[11px] opacity-85 mt-0.5">
                {pendentesObedece.length > 0
                  ? `Valor a resolver neste dia: ${formatBRL(valorPendenteObedece)}. As outras ${concluidos.length} NL(s) (${formatBRL(
                      valorConcluido
                    )}) já tiveram PP e OB emitidas em situação 'AO' no SIGEF.`
                  : `Todas as NLs enviadas nesta publicação já saíram de Obedece e foram processadas com PP/OB emitida em situação AO/PPCB.`}
              </div>
            </div>
          </div>
          <div className="flex items-center gap-2 shrink-0 self-end sm:self-center">
            {pendentesObedece.length > 0 && statusFilter !== 'pendentes_obedece' && (
              <button
                type="button"
                onClick={() => setStatusFilter('pendentes_obedece')}
                className="px-3 py-1.5 bg-amber-500 hover:bg-amber-600 text-slate-950 rounded-xl text-xs font-black shadow-xs transition cursor-pointer"
              >
                Filtrar apenas as {pendentesObedece.length} a resolver
              </button>
            )}
            <button
              type="button"
              onClick={() => setSelectedDia('todos')}
              className="text-xs font-bold text-slate-600 hover:text-slate-900 underline cursor-pointer"
            >
              Ver todos os dias
            </button>
          </div>
        </div>
      )}

      {/* KPI Cards */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        {/* Card 1: Total Publicado */}
        <div className="bg-white rounded-2xl p-5 border border-slate-200/90 shadow-xs flex flex-col justify-between space-y-3">
          <div className="flex items-center justify-between">
            <span className="text-xs font-bold text-slate-500 uppercase tracking-wider">
              {selectedDia === 'todos' ? 'NLs na Publicação' : `NLs (${selectedDia})`}
            </span>
            <div className="w-8 h-8 rounded-xl bg-slate-100 text-slate-700 flex items-center justify-center font-bold text-xs">
              <Layers className="w-4 h-4" />
            </div>
          </div>
          <div>
            <div className="text-2xl font-black text-slate-900 tracking-tight">
              {totalCount}
            </div>
            <div className="text-xs font-bold text-slate-600 mt-0.5">
              {formatBRL(totalValorNL)}
            </div>
          </div>
          <div className="text-[11px] text-slate-400 font-medium pt-2 border-t border-slate-100 flex items-center gap-1">
            <span>
              {selectedDia === 'todos'
                ? `${publicacaoFiles.length} planilha(s) • ${availableDias.length} data(s)`
                : `Data selecionada: ${selectedDia}`}
            </span>
          </div>
        </div>

        {/* Card 2: Pendentes em Obedece (A Resolver) - PRIMARY FOCUS */}
        <div className="bg-white rounded-2xl p-5 border border-amber-300 shadow-xs flex flex-col justify-between space-y-3 bg-gradient-to-br from-white to-amber-50/50">
          <div className="flex items-center justify-between">
            <span className="text-xs font-black text-amber-900 uppercase tracking-wider flex items-center gap-1.5">
              <Zap className="w-3.5 h-3.5 text-amber-600" />
              <span>A Resolver em Obedece</span>
            </span>
            <div className="w-8 h-8 rounded-xl bg-amber-200 text-amber-950 flex items-center justify-center font-black text-xs">
              <Clock className="w-4 h-4" />
            </div>
          </div>
          <div>
            <div className="text-2xl font-black text-amber-950 tracking-tight">
              {pendentesObedece.length}
            </div>
            <div className="text-xs font-black text-amber-800 mt-0.5">
              {formatBRL(valorPendenteObedece)}
            </div>
          </div>
          <div className="text-[11px] text-amber-900 font-semibold pt-2 border-t border-amber-200 flex items-center gap-1">
            <AlertTriangle className="w-3.5 h-3.5 text-amber-700 shrink-0" />
            <span>Listadas em Obedece aguardando ação</span>
          </div>
        </div>

        {/* Card 3: Concluídos (Feito) */}
        <div className="bg-white rounded-2xl p-5 border border-emerald-200/90 shadow-xs flex flex-col justify-between space-y-3 bg-gradient-to-br from-white to-emerald-50/40">
          <div className="flex items-center justify-between">
            <span className="text-xs font-bold text-emerald-800 uppercase tracking-wider">
              Concluídas (Feito)
            </span>
            <div className="w-8 h-8 rounded-xl bg-emerald-100 text-emerald-700 flex items-center justify-center font-bold text-xs">
              <CheckCircle2 className="w-4 h-4" />
            </div>
          </div>
          <div>
            <div className="text-2xl font-black text-emerald-950 tracking-tight">
              {concluidos.length}
            </div>
            <div className="text-xs font-bold text-emerald-700 mt-0.5">
              {formatBRL(valorConcluido)}
            </div>
          </div>
          <div className="text-[11px] text-emerald-800/80 font-medium pt-2 border-t border-emerald-100 flex items-center gap-1">
            <Check className="w-3.5 h-3.5 text-emerald-600" />
            <span>PP/OB emitida em situação AO</span>
          </div>
        </div>

        {/* Card 4: Taxa de Resolução */}
        <div className="bg-white rounded-2xl p-5 border border-indigo-200/90 shadow-xs flex flex-col justify-between space-y-3 bg-gradient-to-br from-white to-indigo-50/40">
          <div className="flex items-center justify-between">
            <span className="text-xs font-bold text-indigo-800 uppercase tracking-wider">
              Taxa de Efetivação
            </span>
            <div className="w-8 h-8 rounded-xl bg-indigo-100 text-indigo-700 flex items-center justify-center font-bold text-xs">
              <TrendingUp className="w-4 h-4" />
            </div>
          </div>
          <div>
            <div className="text-2xl font-black text-indigo-950 tracking-tight">
              {taxaEfetivacao.toFixed(1).replace('.', ',')}%
            </div>
            <div className="text-xs font-bold text-indigo-700 mt-0.5">
              {concluidos.length} de {totalCount} resolvidas
            </div>
          </div>
          <div className="text-[11px] text-indigo-800/80 font-medium pt-2 border-t border-indigo-100">
            <div className="w-full bg-indigo-100 rounded-full h-1.5 mt-1 overflow-hidden">
              <div
                className="bg-indigo-600 h-1.5 rounded-full transition-all duration-500"
                style={{ width: `${Math.min(100, Math.max(0, taxaEfetivacao))}%` }}
              />
            </div>
          </div>
        </div>
      </div>

      {/* Main Table Card */}
      <div className="bg-white rounded-2xl shadow-xs border border-slate-200/90 overflow-hidden">
        {/* Table Filters Header */}
        <div className="p-4 sm:p-5 border-b border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-4 bg-slate-50/40">
          {/* Search Input */}
          <div className="relative flex-1 max-w-md">
            <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              placeholder="Buscar por NL, Credor, CNPJ, PP, OB ou Dia..."
              className="w-full pl-9 pr-4 py-2 bg-white border border-slate-200 rounded-xl text-xs font-medium text-slate-800 focus:outline-hidden focus:border-emerald-500 focus:ring-2 focus:ring-emerald-500/20"
            />
          </div>

          {/* Status Filter Buttons */}
          <div className="flex flex-wrap items-center gap-1.5 bg-slate-100 p-1 rounded-xl">
            {/* Filter: Pendentes em Obedece (A Resolver) - User Priority Request */}
            <button
              type="button"
              onClick={() => setStatusFilter('pendentes_obedece')}
              className={`px-3 py-1.5 rounded-lg text-xs font-black transition cursor-pointer flex items-center gap-1.5 ${
                statusFilter === 'pendentes_obedece'
                  ? 'bg-amber-500 text-slate-950 shadow-xs'
                  : 'text-amber-950 hover:bg-amber-200/60'
              }`}
              title="Filtrar apenas as NLs que ainda estão listadas em Obedece e foram para publicação"
            >
              <AlertTriangle className="w-3.5 h-3.5 text-amber-800" />
              <span>A Resolver em Obedece ({pendentesObedece.length})</span>
            </button>

            {/* Filter: Concluídos */}
            <button
              type="button"
              onClick={() => setStatusFilter('concluido')}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold transition cursor-pointer flex items-center gap-1.5 ${
                statusFilter === 'concluido'
                  ? 'bg-emerald-600 text-white shadow-xs'
                  : 'text-emerald-700 hover:bg-emerald-100/60'
              }`}
            >
              <Check className="w-3.5 h-3.5" />
              <span>Concluídos ({concluidos.length})</span>
            </button>

            {/* Filter: Todos */}
            <button
              type="button"
              onClick={() => setStatusFilter('todos')}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold transition cursor-pointer ${
                statusFilter === 'todos'
                  ? 'bg-white text-slate-900 shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              Todos ({totalCount})
            </button>

            {/* Filter: Todas as Pendências */}
            <button
              type="button"
              onClick={() => setStatusFilter('todas_pendentes')}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold transition cursor-pointer flex items-center gap-1.5 ${
                statusFilter === 'todas_pendentes'
                  ? 'bg-slate-800 text-white shadow-xs'
                  : 'text-slate-600 hover:bg-slate-200'
              }`}
            >
              <Clock className="w-3.5 h-3.5" />
              <span>Todas Pendências ({todasPendentes.length})</span>
            </button>
          </div>
        </div>

        {/* Records Table */}
        {publicacaoItems.length === 0 ? (
          <div className="p-12 text-center space-y-4">
            <div className="w-14 h-14 rounded-2xl bg-indigo-50 text-indigo-600 flex items-center justify-center mx-auto border border-indigo-200 shadow-xs">
              <FileSpreadsheet className="w-7 h-7" />
            </div>
            <div className="max-w-md mx-auto space-y-1.5">
              <h3 className="text-base font-bold text-slate-900">
                Nenhuma planilha de Publicação carregada
              </h3>
              <p className="text-xs text-slate-500 font-medium">
                Importe planilhas no padrão <strong>Publicação - DD/MM</strong> (ex: <em>Publicação - 05/11.xlsx</em>). O sistema
                irá organizar as NLs por dia, verificar a permanência ou saída de Obedece e localizar PPs e OBs emitidas em situação <strong>AO</strong> no SIGEF.
              </p>
            </div>
            <div className="pt-2 flex flex-wrap items-center justify-center gap-3">
              <button
                type="button"
                onClick={() => fileInputRef.current?.click()}
                className="px-5 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold transition shadow-md flex items-center gap-2 cursor-pointer active:scale-98"
              >
                <Upload className="w-4 h-4" />
                <span>Carregar Planilha de Publicação</span>
              </button>
              <button
                type="button"
                onClick={onNavigateToImportacoes}
                className="px-5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold transition flex items-center gap-2 cursor-pointer active:scale-98"
              >
                <span>Ir para Central de Importações</span>
                <ArrowRight className="w-4 h-4" />
              </button>
            </div>
          </div>
        ) : filteredItems.length === 0 ? (
          <div className="p-12 text-center text-slate-400 space-y-2">
            <AlertCircle className="w-8 h-8 mx-auto text-slate-300" />
            <p className="text-xs font-bold text-slate-600">Nenhum registro encontrado com os filtros aplicados.</p>
            <button
              onClick={() => {
                setSearchTerm('');
                setStatusFilter('todos');
                setSelectedDia('todos');
              }}
              className="text-xs text-emerald-700 font-bold hover:underline cursor-pointer"
            >
              Limpar filtros de busca
            </button>
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs border-collapse">
              <thead className="bg-slate-100/90 text-slate-700 font-bold uppercase tracking-wider text-[11px] border-b border-slate-200 sticky top-0">
                <tr>
                  <th className="py-3 px-3.5 text-center w-12">#</th>
                  <th className="py-3 px-3.5 text-center">Dia</th>
                  <th className="py-3 px-3.5">Status</th>
                  <th className="py-3 px-3.5">NL (Publicação)</th>
                  <th className="py-3 px-3.5">Fornecedor / Favorecido</th>
                  <th className="py-3 px-3.5 text-right">Valor NL</th>
                  <th className="py-3 px-3.5 text-center">Fila Obedece</th>
                  <th className="py-3 px-3.5">PP Emitida</th>
                  <th className="py-3 px-3.5">OB Emitida</th>
                  <th className="py-3 px-3.5 text-right">Valor PP/OB</th>
                  <th className="py-3 px-3.5 text-center">Situação</th>
                  <th className="py-3 px-3.5">Diagnóstico SIGEF</th>
                  <th className="py-3 px-3.5 text-center w-28">Ações</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 font-sans">
                {filteredItems.map((item, idx) => {
                  const isConcluido = item.status === 'concluido';
                  const isPendenteObedece = !isConcluido && !item.sumiuObedece;
                  const isAO = item.situacaoEquivalente?.includes('AO');

                  return (
                    <tr
                      key={item.id}
                      className={`hover:bg-slate-50/80 transition-colors ${
                        isConcluido
                          ? 'bg-emerald-50/20'
                          : isPendenteObedece
                          ? 'bg-amber-50/30'
                          : 'bg-white'
                      }`}
                    >
                      {/* # */}
                      <td className="py-3 px-3.5 text-slate-400 font-mono text-[10px] text-center">
                        {idx + 1}
                      </td>

                      {/* Dia Publicação */}
                      <td className="py-3 px-3.5 text-center whitespace-nowrap">
                        <span className="px-2 py-0.5 rounded-md bg-slate-100 text-slate-700 font-mono font-bold text-[10px] border border-slate-200">
                          {item.diaPublicacao || 'Geral'}
                        </span>
                      </td>

                      {/* Status Badge */}
                      <td className="py-3 px-3.5 whitespace-nowrap">
                        {isConcluido ? (
                          <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg bg-emerald-100 text-emerald-800 border border-emerald-300 font-bold text-[10px] uppercase tracking-wide">
                            <Check className="w-3 h-3 text-emerald-700" />
                            Feito
                          </span>
                        ) : isPendenteObedece ? (
                          <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg bg-amber-200 text-amber-950 border border-amber-400 font-black text-[10px] uppercase tracking-wide">
                            <AlertTriangle className="w-3 h-3 text-amber-800" />
                            A Resolver
                          </span>
                        ) : (
                          <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg bg-slate-100 text-slate-700 border border-slate-300 font-bold text-[10px] uppercase tracking-wide">
                            <Clock className="w-3 h-3 text-slate-500" />
                            Pendente
                          </span>
                        )}
                      </td>

                      {/* NL Número */}
                      <td className="py-3 px-3.5 font-mono font-bold text-slate-900 whitespace-nowrap">
                        {item.nlNumero}
                      </td>

                      {/* Fornecedor */}
                      <td className="py-3 px-3.5 max-w-xs">
                        <div className="font-bold text-slate-900 truncate" title={item.fornecedor}>
                          {item.fornecedor}
                        </div>
                        {(() => {
                          const c = item.cnpj || resolveCnpj(item.fornecedor, item.nlNumero);
                          return c ? (
                            <div className="text-[10px] text-slate-500 font-mono">
                              CNPJ: {c}
                            </div>
                          ) : null;
                        })()}
                      </td>

                      {/* Valor Bruto NL */}
                      <td className="py-3 px-3.5 text-right whitespace-nowrap">
                        <div className="font-bold text-slate-900">
                          {formatBRL(item.valorNL)}
                        </div>
                        {item.valorPublicacao && Math.abs(item.valorPublicacao - item.valorNL) > 0.01 ? (
                          <div className="text-[10px] text-slate-400 font-medium" title="Valor original do arquivo de publicação">
                            Pub: {formatBRL(item.valorPublicacao)}
                          </div>
                        ) : null}
                      </td>

                      {/* Fila Obedece Status */}
                      <td className="py-3 px-3.5 text-center whitespace-nowrap">
                        {!item.sumiuObedece ? (
                          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md bg-amber-100 text-amber-900 border border-amber-300 font-bold text-[10px]" title="A NL ainda consta na fila de Obedece e precisa ser liberada">
                            <AlertTriangle className="w-3 h-3 text-amber-700" />
                            <span>Em Fila Obedece</span>
                          </span>
                        ) : (
                          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md bg-emerald-50 text-emerald-800 font-bold text-[10px]">
                            <Check className="w-3 h-3 text-emerald-600" />
                            <span>Saiu da Fila</span>
                          </span>
                        )}
                      </td>

                      {/* PP Equivalente */}
                      <td className="py-3 px-3.5 font-mono text-slate-700 whitespace-nowrap">
                        {item.equivalentePP ? (
                          <span className="px-2 py-0.5 rounded bg-blue-50 text-blue-900 border border-blue-200 font-bold text-[11px]">
                            {item.equivalentePP}
                          </span>
                        ) : (
                          <span className="text-slate-400 italic text-[11px]">-</span>
                        )}
                      </td>

                      {/* OB Equivalente */}
                      <td className="py-3 px-3.5 font-mono text-slate-700 whitespace-nowrap">
                        {item.equivalenteOB ? (
                          <span className="px-2 py-0.5 rounded bg-teal-50 text-teal-900 border border-teal-200 font-bold text-[11px]">
                            {item.equivalenteOB}
                          </span>
                        ) : (
                          <span className="text-slate-400 italic text-[11px]">-</span>
                        )}
                      </td>

                      {/* Valor PP/OB */}
                      <td className="py-3 px-3.5 text-right whitespace-nowrap">
                        {item.valorEquivalente ? (
                          <div>
                            <span className="font-bold text-emerald-800">
                              {formatBRL(item.valorEquivalente)}
                            </span>
                            {item.percentualRetencao && item.percentualRetencao > 0 ? (
                              <div className="text-[10px] text-slate-500 font-medium">
                                -{(item.percentualRetencao * 100).toFixed(1).replace('.', ',')}% trib.
                              </div>
                            ) : null}
                          </div>
                        ) : (
                          <span className="text-slate-400 italic text-[11px]">-</span>
                        )}
                      </td>

                      {/* Situação */}
                      <td className="py-3 px-3.5 text-center whitespace-nowrap">
                        {item.situacaoEquivalente ? (
                          <span
                            className={`px-2 py-0.5 rounded font-mono font-bold text-[10px] border ${
                              isAO
                                ? 'bg-emerald-50 text-emerald-800 border-emerald-300'
                                : 'bg-slate-100 text-slate-800 border-slate-200'
                            }`}
                            title={isAO ? 'AO: Associada aos Ordenadores (Emitida no SIGEF)' : item.situacaoEquivalente}
                          >
                            {item.situacaoEquivalente}
                          </span>
                        ) : (
                          <span className="text-slate-400 italic text-[11px]">-</span>
                        )}
                      </td>

                      {/* Diagnóstico */}
                      <td className="py-3 px-3.5 max-w-sm text-slate-600 text-[11px] leading-relaxed">
                        {item.statusDetalhe}
                      </td>

                      {/* Ações */}
                      <td className="py-3 px-3.5 text-center whitespace-nowrap">
                        <button
                          type="button"
                          onClick={() => handleOpenManualModal(item)}
                          className={`inline-flex items-center gap-1 px-2.5 py-1 rounded-lg text-[11px] font-bold transition cursor-pointer active:scale-98 ${
                            isConcluido
                              ? 'bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-200'
                              : 'bg-emerald-600 hover:bg-emerald-500 text-white shadow-2xs'
                          }`}
                          title="Vincular PP e OB manualmente para calcular e registrar padrão"
                        >
                          <Link2 className="w-3.5 h-3.5" />
                          <span>{isConcluido ? 'Editar' : 'Vincular'}</span>
                        </button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}

        {/* Footer info bar */}
        {publicacaoItems.length > 0 && (
          <div className="p-4 bg-slate-50 border-t border-slate-100 flex flex-col sm:flex-row items-center justify-between gap-3 text-xs text-slate-600">
            <div>
              Exibindo <strong>{filteredItems.length}</strong> de <strong>{totalCount}</strong> liquidações
              {selectedDia !== 'todos' ? ` do dia ${selectedDia}` : ' de todas as datas'}.
            </div>
            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={handleExportPdf}
                className="text-rose-700 hover:text-rose-900 font-bold flex items-center gap-1 cursor-pointer"
              >
                <FileText className="w-3.5 h-3.5" />
                <span>Baixar PDF</span>
              </button>
              <span className="text-slate-300">•</span>
              <button
                type="button"
                onClick={handleExportExcel}
                className="text-emerald-700 hover:text-emerald-900 font-bold flex items-center gap-1 cursor-pointer"
              >
                <Download className="w-3.5 h-3.5" />
                <span>Baixar Excel</span>
              </button>
            </div>
          </div>
        )}
      </div>

      {/* Modal de Vinculação Manual de PP e OB */}
      {manualModalItem && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/70 backdrop-blur-xs animate-in fade-in duration-200">
          <div className="bg-white rounded-3xl shadow-2xl border border-slate-200 w-full max-w-lg overflow-hidden flex flex-col">
            {/* Header */}
            <div className="p-5 bg-slate-900 text-white flex items-center justify-between border-b border-slate-800">
              <div className="flex items-center gap-2.5">
                <div className="w-9 h-9 rounded-xl bg-emerald-500/20 border border-emerald-500/30 flex items-center justify-center text-emerald-400">
                  <Link2 className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-base font-black text-white">Vincular PP e OB Manualmente</h3>
                  <p className="text-[11px] text-slate-400">Regularização e aprendizado de padrão de retenção</p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setManualModalItem(null)}
                className="p-1.5 text-slate-400 hover:text-white rounded-xl hover:bg-slate-800 transition cursor-pointer"
                title="Fechar"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Form */}
            <form onSubmit={handleSaveManualBinding} className="p-6 space-y-4 text-xs">
              {/* NL Selector with Real-time Suggestions from Obedece and Não Obedece */}
              <div className="bg-slate-50 p-3.5 rounded-2xl border border-slate-200/80 space-y-2">
                <div className="flex items-center justify-between">
                  <span className="text-[11px] font-bold text-slate-700">Liquidação Selecionada (NL)</span>
                  <span className="font-mono font-bold text-slate-900 text-sm">
                    {formatBRL(parseFloat(manualValorNL) || manualModalItem.valorNL)}
                  </span>
                </div>
                <SuggestionInput
                  value={manualNL}
                  onChange={(val) => setManualNL(val)}
                  onSelect={handleSelectNlSuggestion}
                  suggestions={nlSuggestions}
                  placeholder="Digite o número da NL para buscar nas planilhas..."
                  helperText="Conforme você digita, o sistema sugere NLs das planilhas Obedece e Não Obedece e puxa a Fonte automaticamente."
                />
                <div className="flex items-center justify-between text-[11px] text-slate-600 font-semibold pt-1">
                  <span className="truncate max-w-[280px]" title={manualCredor || manualModalItem.fornecedor}>
                    {manualCredor || manualModalItem.fornecedor}
                  </span>
                  <span className="font-mono text-slate-500">
                    CNPJ: {manualCnpj || manualModalItem.cnpj || resolveCnpj(manualCredor || manualModalItem.fornecedor, manualNL || manualModalItem.nlNumero) || 'Não informado'}
                  </span>
                </div>
              </div>

              {/* Input Fields Grid */}
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                {/* PP Suggestion Input */}
                <div className="space-y-1">
                  <SuggestionInput
                    label="Número da PP"
                    required
                    value={manualPP}
                    onChange={(val) => setManualPP(val)}
                    onSelect={handleSelectPpSuggestion}
                    suggestions={ppSuggestions}
                    placeholder="Ex: 2026PP000123"
                    helperText="Selecione das PPs registradas ou digite uma nova para adicionar."
                  />
                </div>

                {/* OB Suggestion Input */}
                <div className="space-y-1">
                  <SuggestionInput
                    label="Número da OB"
                    value={manualOB}
                    onChange={(val) => setManualOB(val)}
                    onSelect={handleSelectObSuggestion}
                    suggestions={obSuggestions}
                    placeholder="Ex: 2026OB000456"
                    helperText="Selecione das OBs do app ou digite para cadastrar."
                  />
                </div>

                <div className="space-y-1">
                  <label className="text-[11px] font-bold text-slate-700">Situação no SIGEF</label>
                  <select
                    value={manualSituacao}
                    onChange={(e) => setManualSituacao(e.target.value as any)}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl font-bold text-xs focus:bg-white focus:outline-hidden focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500"
                  >
                    <option value="AO">AO (Associada aos Ordenadores / A Pagar)</option>
                    <option value="PPCB">PPCB (Paga Confirmada Banco)</option>
                    <option value="CB">CB (Confirmada Banco)</option>
                  </select>
                </div>

                {/* Fonte de Recursos (0.600 Federal / 0.500 Estadual) with Auto-Detection badge */}
                <div className="space-y-1">
                  <div className="flex items-center justify-between">
                    <div className="flex items-center gap-1.5">
                      <label className="text-[11px] font-bold text-slate-700">Fonte de Recursos</label>
                      {fonteAutoDetected && (
                        <span className="text-[9px] font-bold text-emerald-800 bg-emerald-100 px-1 py-0.5 rounded border border-emerald-300">
                          Auto da Planilha
                        </span>
                      )}
                    </div>
                    <div className="flex items-center gap-1">
                      <button
                        type="button"
                        onClick={() => {
                          setManualFonte('0.600');
                          setFonteAutoDetected(false);
                        }}
                        className={`px-1.5 py-0.5 rounded text-[10px] font-mono font-bold transition cursor-pointer ${
                          manualFonte.startsWith('0.600')
                            ? 'bg-blue-600 text-white'
                            : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                        }`}
                        title="Fonte Federal 0.600 (SUS / FNS)"
                      >
                        0.600 (Fed)
                      </button>
                      <button
                        type="button"
                        onClick={() => {
                          setManualFonte('0.500');
                          setFonteAutoDetected(false);
                        }}
                        className={`px-1.5 py-0.5 rounded text-[10px] font-mono font-bold transition cursor-pointer ${
                          manualFonte.startsWith('0.500')
                            ? 'bg-emerald-600 text-white'
                            : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                        }`}
                        title="Fonte Estadual 0.500 (FES / Tesouro)"
                      >
                        0.500 (Est)
                      </button>
                    </div>
                  </div>
                  <input
                    type="text"
                    value={manualFonte}
                    onChange={(e) => {
                      setManualFonte(e.target.value);
                      setFonteAutoDetected(false);
                    }}
                    placeholder="Ex: 0.600 ou 0.500"
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl font-mono text-xs focus:bg-white focus:outline-hidden focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500"
                  />
                </div>

                {/* Editable Valor da Parcela / NL */}
                <div className="space-y-1">
                  <div className="flex items-center justify-between">
                    <label className="text-[11px] font-bold text-slate-700">Valor da Parcela / NL (R$)</label>
                    <span className="text-[10px] text-slate-400">Rateio de NF</span>
                  </div>
                  <input
                    type="text"
                    value={manualValorNL}
                    onChange={(e) => setManualValorNL(e.target.value)}
                    placeholder={String(manualModalItem.valorNL)}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl font-mono font-bold text-xs focus:bg-white focus:outline-hidden focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500"
                  />
                </div>

                {/* Editable Valor Pago Líquido PP/OB */}
                <div className="space-y-1">
                  <div className="flex items-center justify-between">
                    <label className="text-[11px] font-bold text-slate-700">Valor Pago Efetivo PP/OB (R$)</label>
                    <span className="text-[10px] text-slate-400">Pós retenções</span>
                  </div>
                  <input
                    type="text"
                    value={manualValor}
                    onChange={(e) => setManualValor(e.target.value)}
                    placeholder={manualValorNL || String(manualModalItem.valorNL)}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl font-mono font-bold text-xs focus:bg-white focus:outline-hidden focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500"
                  />
                </div>
              </div>

              {/* Notice about Split and Retention */}
              <div className="p-2.5 bg-blue-50/60 rounded-xl border border-blue-200/80 text-[11px] text-blue-900 flex items-start gap-2">
                <span className="font-bold shrink-0 text-blue-700">Divisão de Recursos:</span>
                <span>
                  Quando uma NF é paga com recursos federais (<strong>0.600</strong>) e estaduais (<strong>0.500</strong>), você pode ajustar o valor da parcela desta NL e a fonte correspondente para o cálculo correto.
                </span>
              </div>

              {/* Dynamic Retention Calculation */}
              {(() => {
                const cleanValNL = manualValorNL.replace(/\./g, '').replace(',', '.');
                const parsedValNL = parseFloat(cleanValNL) || manualModalItem.valorNL;
                const cleanVal = manualValor.replace(/\./g, '').replace(',', '.');
                const parsedVal = parseFloat(cleanVal) || parsedValNL;
                const dif = parsedValNL - parsedVal;
                const retPct = dif > 0 && parsedValNL > 0 ? (dif / parsedValNL) * 100 : 0;

                if (dif > 0.01) {
                  return (
                    <div className="p-3 bg-amber-50 rounded-xl border border-amber-200 text-amber-900 space-y-1">
                      <div className="flex items-center gap-1.5 font-bold">
                        <Sparkles className="w-3.5 h-3.5 text-amber-700" />
                        <span>Retenção tributária calculada: {formatBRL(dif)} ({retPct.toFixed(2).replace('.', ',')}%)</span>
                      </div>
                      <p className="text-[11px] text-amber-800 leading-relaxed">
                        O sistema registrará esse percentual para o credor <strong>{manualModalItem.fornecedor}</strong> na Fonte <strong>{manualFonte}</strong>, aprendendo o padrão para associar automaticamente futuras liquidações semelhantes.
                      </p>
                    </div>
                  );
                }

                return (
                  <div className="p-2.5 bg-slate-50 rounded-xl border border-slate-200 text-slate-600 text-[11px]">
                    Pagamento com valor integral ({formatBRL(parsedValNL)}) na Fonte <strong>{manualFonte}</strong>. Sem retenção de impostos.
                  </div>
                );
              })()}

              {/* Action Buttons */}
              <div className="pt-2 flex items-center justify-end gap-2">
                <button
                  type="button"
                  onClick={() => setManualModalItem(null)}
                  className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl font-bold cursor-pointer"
                >
                  Cancelar
                </button>
                <button
                  type="submit"
                  className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl font-bold transition shadow-sm cursor-pointer active:scale-98 flex items-center gap-1.5"
                >
                  <Check className="w-4 h-4" />
                  <span>Salvar e Memorizar</span>
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};
