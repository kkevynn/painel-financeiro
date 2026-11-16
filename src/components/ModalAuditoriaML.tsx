import React, { useState, useEffect } from 'react';
import {
  BrainCircuit,
  CheckCircle2,
  XCircle,
  Sparkles,
  ArrowRight,
  ShieldCheck,
  Building,
  Trash2,
  CheckCheck,
  X,
  Database,
  TrendingDown,
  RefreshCw,
  Plus,
  Edit3,
  Check,
  AlertCircle,
} from 'lucide-react';
import { InferredPaymentMatch, LearnedCreditorPattern, FileData } from '../types';
import { mlInferenceService } from '../utils/machineLearningInferenceService';
import { formatBRL } from '../utils/excelParser';
import { OrdemBancariaItem } from '../utils/ordemBancariaParser';
import { SuggestionInput } from './common/SuggestionInput';
import {
  extractAvailableNls,
  extractAvailablePPs,
  extractAvailableOBs,
  searchNls,
  searchPPs,
  searchOBs,
  recordCustomPP,
  recordCustomOB,
  NlSuggestionItem,
  PpSuggestionItem,
  ObSuggestionItem,
} from '../utils/suggestionService';

interface ModalAuditoriaMLProps {
  isOpen: boolean;
  onClose: () => void;
  files?: FileData[];
  ordensBancariasList?: OrdemBancariaItem[] | null;
}

export const ModalAuditoriaML: React.FC<ModalAuditoriaMLProps> = ({
  isOpen,
  onClose,
  files = [],
  ordensBancariasList = null,
}) => {
  const [activeTab, setActiveTab] = useState<'validacao' | 'padroes'>('validacao');
  const [filterMode, setFilterMode] = useState<'pendentes' | 'f600' | 'confirmados' | 'all'>('pendentes');
  const [matches, setMatches] = useState<InferredPaymentMatch[]>(() => mlInferenceService.getInferredMatches());
  const [patterns, setPatterns] = useState<LearnedCreditorPattern[]>(() => mlInferenceService.getLearnedPatterns());
  const [feedbackMsg, setFeedbackMsg] = useState<string | null>(null);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const [isScanning, setIsScanning] = useState(false);

  // Available suggestions from app files and OB list
  const availableNls = React.useMemo(() => extractAvailableNls(files), [files]);
  const availablePPs = React.useMemo(() => extractAvailablePPs(files, ordensBancariasList), [files, ordensBancariasList]);
  const availableOBs = React.useMemo(() => extractAvailableOBs(files, ordensBancariasList), [files, ordensBancariasList]);

  // Manual addition state
  const [isAddingManual, setIsAddingManual] = useState(false);
  const [manualNL, setManualNL] = useState('');
  const [manualCredor, setManualCredor] = useState('');
  const [manualCnpj, setManualCnpj] = useState('');
  const [manualValorNL, setManualValorNL] = useState('');
  const [manualPP, setManualPP] = useState('');
  const [manualOB, setManualOB] = useState('');
  const [manualSituacao, setManualSituacao] = useState('AO');
  const [manualValorPP, setManualValorPP] = useState('');

  // Inline edit state
  const [editingMatchId, setEditingMatchId] = useState<string | null>(null);
  const [editPP, setEditPP] = useState('');
  const [editOB, setEditOB] = useState('');
  const [editSituacao, setEditSituacao] = useState('AO');
  const [editValorPP, setEditValorPP] = useState('');

  // Suggestions for manual addition form
  const manualNlSuggestions = React.useMemo(() => {
    return searchNls(manualNL, availableNls, 10).map((it) => ({
      id: `man_nl_${it.cleanDoc}`,
      title: it.nlNumero,
      badge: it.fonteLabel || (it.fonteTipo === 'federal' ? 'Fonte 0.600' : 'Fonte 0.500'),
      badgeColor: it.fonteTipo === 'federal' ? 'bg-blue-50 text-blue-900 border-blue-200' : 'bg-emerald-50 text-emerald-900 border-emerald-200',
      subtitle: `${it.fornecedor} • ${it.cnpj ? `CNPJ: ${it.cnpj}` : ''} • ${it.stage === 'obedece' ? 'Fila Obedece' : 'Não Obedece'}`,
      extraInfo: formatBRL(it.valor),
      data: it,
    }));
  }, [manualNL, availableNls]);

  const manualPpSuggestions = React.useMemo(() => {
    return searchPPs(manualPP, availablePPs, 10).map((it) => ({
      id: `man_pp_${it.cleanNum}`,
      title: it.ppNumero,
      badge: it.situacao || 'AO',
      badgeColor: it.situacao?.includes('CB') ? 'bg-emerald-50 text-emerald-800 border-emerald-200' : 'bg-blue-50 text-blue-800 border-blue-200',
      subtitle: `${it.fornecedor || 'Credor'} • ${it.obNumero ? `OB: ${it.obNumero}` : 'Sem OB'}`,
      extraInfo: it.valor ? formatBRL(it.valor) : undefined,
      data: it,
    }));
  }, [manualPP, availablePPs]);

  const manualObSuggestions = React.useMemo(() => {
    return searchOBs(manualOB, availableOBs, 10).map((it) => ({
      id: `man_ob_${it.cleanNum}`,
      title: it.obNumero,
      badge: it.situacao || 'CB',
      badgeColor: 'bg-teal-50 text-teal-800 border-teal-200',
      subtitle: `${it.fornecedor || 'Credor'} • ${it.ppNumero ? `PP: ${it.ppNumero}` : ''}`,
      extraInfo: it.valor ? formatBRL(it.valor) : undefined,
      data: it,
    }));
  }, [manualOB, availableOBs]);

  // Suggestions for inline edit form
  const editPpSuggestions = React.useMemo(() => {
    return searchPPs(editPP, availablePPs, 10).map((it) => ({
      id: `edit_pp_${it.cleanNum}`,
      title: it.ppNumero,
      badge: it.situacao || 'AO',
      badgeColor: 'bg-blue-50 text-blue-800 border-blue-200',
      subtitle: `${it.fornecedor || ''} • ${it.obNumero ? `OB: ${it.obNumero}` : ''}`,
      extraInfo: it.valor ? formatBRL(it.valor) : undefined,
      data: it,
    }));
  }, [editPP, availablePPs]);

  const editObSuggestions = React.useMemo(() => {
    return searchOBs(editOB, availableOBs, 10).map((it) => ({
      id: `edit_ob_${it.cleanNum}`,
      title: it.obNumero,
      badge: it.situacao || 'CB',
      badgeColor: 'bg-teal-50 text-teal-800 border-teal-200',
      subtitle: `${it.fornecedor || ''} • ${it.ppNumero ? `PP: ${it.ppNumero}` : ''}`,
      extraInfo: it.valor ? formatBRL(it.valor) : undefined,
      data: it,
    }));
  }, [editOB, availableOBs]);

  // Lock background scroll when open
  useEffect(() => {
    if (!isOpen) return;
    const origOverflow = document.body.style.overflow;
    document.body.style.overflow = 'hidden';
    return () => {
      document.body.style.overflow = origOverflow;
    };
  }, [isOpen]);

  const handleRunAudit = () => {
    setIsScanning(true);
    // Defer execution so the UI thread updates the spinner without lag
    setTimeout(() => {
      mlInferenceService.processFilesForInference(files, ordensBancariasList, true);
      setMatches(mlInferenceService.getInferredMatches());
      setPatterns(mlInferenceService.getLearnedPatterns());
      const count = mlInferenceService.getInferredMatches().length;
      setFeedbackMsg(`Auditoria finalizada! ${count} amostras detectadas.`);
      setIsScanning(false);
      setTimeout(() => setFeedbackMsg(null), 3500);
    }, 40);
  };

  useEffect(() => {
    if (!isOpen) return;

    // Load available cached matches instantly so the modal opens with 0 latency
    setMatches(mlInferenceService.getInferredMatches());
    setPatterns(mlInferenceService.getLearnedPatterns());

    // If no matches are cached yet and files are present, process asynchronously
    if (mlInferenceService.getInferredMatches().length === 0 && (files.length > 0 || (ordensBancariasList && ordensBancariasList.length > 0))) {
      setIsScanning(true);
      const timer = setTimeout(() => {
        mlInferenceService.processFilesForInference(files, ordensBancariasList);
        setMatches(mlInferenceService.getInferredMatches());
        setPatterns(mlInferenceService.getLearnedPatterns());
        setIsScanning(false);
      }, 50);
      return () => clearTimeout(timer);
    }

    const refreshData = () => {
      setMatches(mlInferenceService.getInferredMatches());
      setPatterns(mlInferenceService.getLearnedPatterns());
    };

    const unsubscribe = mlInferenceService.subscribe(refreshData);
    return () => unsubscribe();
  }, [isOpen, files, ordensBancariasList]);

  if (!isOpen) return null;

  // RULE: "Na análise de auditoria, está me mostrando liquidações de cpf, quero que me mostre apenas de cnpj"
  const cnpjMatches = matches.filter((m) => {
    const digits = (m.cnpj || '').replace(/\D/g, '');
    return digits.length !== 11;
  });

  const cnpjPatterns = patterns.filter((p) => {
    const digits = (p.cnpj || '').replace(/\D/g, '');
    return digits.length !== 11;
  });

  const pendingMatches = cnpjMatches.filter((m) => m.statusValidacao === 'pendente');
  const f600Matches = cnpjMatches.filter((m) => m.isFonte600);
  const confirmedMatches = cnpjMatches.filter((m) => m.statusValidacao === 'confirmado');

  const displayedMatches = cnpjMatches.filter((m) => {
    if (filterMode === 'pendentes') return m.statusValidacao === 'pendente';
    if (filterMode === 'f600') return m.isFonte600;
    if (filterMode === 'confirmados') return m.statusValidacao === 'confirmado';
    return true;
  });

  const handleConfirm = (id: string, credor: string) => {
    mlInferenceService.confirmMatch(id);
    setFeedbackMsg(`Conversão confirmada e memorizada para "${credor}".`);
    setTimeout(() => setFeedbackMsg(null), 4000);
  };

  const handleReject = (id: string) => {
    mlInferenceService.rejectMatch(id);
    setFeedbackMsg('Conversão descartada.');
    setTimeout(() => setFeedbackMsg(null), 3000);
  };

  const handleReopen = (id: string) => {
    mlInferenceService.reopenMatch(id);
  };

  const handleConfirmAll = () => {
    mlInferenceService.confirmAllHighConfidence();
    setFeedbackMsg('Todas as amostras consistentes foram confirmadas e memorizadas!');
    setTimeout(() => setFeedbackMsg(null), 4000);
  };

  const handleSaveManualMatch = () => {
    if (!manualNL.trim() || !manualCredor.trim() || !manualValorNL || !manualPP.trim()) {
      setErrorMsg('Por favor, preencha pelo menos a NL, o Credor, o Valor da NL e o número da PP.');
      setTimeout(() => setErrorMsg(null), 4500);
      return;
    }
    setErrorMsg(null);

    const valNL = parseFloat(manualValorNL) || 0;
    const valPP = manualValorPP ? parseFloat(manualValorPP) : valNL;

    mlInferenceService.addManualInferredMatch({
      nlNumero: manualNL.trim(),
      credor: manualCredor.trim(),
      cnpj: manualCnpj.trim() || undefined,
      valorNL: valNL,
      ppNumero: manualPP.trim(),
      obNumero: manualOB.trim() || undefined,
      situacao: manualSituacao.trim() || 'AO',
      valorPP: valPP,
    });

    // Record custom PP and OB so the app remembers and auto-suggests them everywhere
    if (manualPP.trim()) {
      recordCustomPP({
        ppNumero: manualPP.trim(),
        obNumero: manualOB.trim() || undefined,
        valor: valPP,
        fornecedor: manualCredor.trim(),
        cnpj: manualCnpj.trim() || undefined,
        situacao: manualSituacao.trim() || 'AO',
      });
    }
    if (manualOB.trim()) {
      recordCustomOB({
        obNumero: manualOB.trim(),
        ppNumero: manualPP.trim() || undefined,
        valor: valPP,
        fornecedor: manualCredor.trim(),
        cnpj: manualCnpj.trim() || undefined,
        situacao: manualSituacao.trim() || 'AO',
      });
    }

    setMatches(mlInferenceService.getInferredMatches());
    setPatterns(mlInferenceService.getLearnedPatterns());

    setFeedbackMsg(`Ligação manual registrada e padrão aprendido para "${manualCredor.trim()}"!`);
    setIsAddingManual(false);

    setManualNL('');
    setManualCredor('');
    setManualCnpj('');
    setManualValorNL('');
    setManualPP('');
    setManualOB('');
    setManualSituacao('AO');
    setManualValorPP('');

    setTimeout(() => setFeedbackMsg(null), 4000);
  };

  const handleStartEdit = (match: InferredPaymentMatch) => {
    setEditingMatchId(match.id);
    setEditPP(match.ppNumero || '');
    setEditOB(match.obNumero || '');
    setEditSituacao(match.statusSituacaoPP || 'AO');
    setEditValorPP(String(match.valorPP || match.valorNL));
  };

  const handleSaveEdit = (matchId: string) => {
    if (!editPP.trim()) {
      setErrorMsg('O número da PP não pode ficar vazio.');
      setTimeout(() => setErrorMsg(null), 4000);
      return;
    }
    setErrorMsg(null);

    const valPP = parseFloat(editValorPP);

    mlInferenceService.updateInferredMatch(matchId, {
      ppNumero: editPP.trim(),
      obNumero: editOB.trim(),
      situacao: editSituacao.trim(),
      valorPP: !isNaN(valPP) && valPP > 0 ? valPP : undefined,
    });

    const match = matches.find((m) => m.id === matchId);
    if (match) {
      if (editPP.trim()) {
        recordCustomPP({
          ppNumero: editPP.trim(),
          obNumero: editOB.trim() || undefined,
          valor: !isNaN(valPP) && valPP > 0 ? valPP : match.valorPP,
          fornecedor: match.credor,
          cnpj: match.cnpj,
          situacao: editSituacao.trim(),
        });
      }
      if (editOB.trim()) {
        recordCustomOB({
          obNumero: editOB.trim(),
          ppNumero: editPP.trim() || undefined,
          valor: !isNaN(valPP) && valPP > 0 ? valPP : match.valorPP,
          fornecedor: match.credor,
          cnpj: match.cnpj,
          situacao: editSituacao.trim(),
        });
      }
    }

    setMatches(mlInferenceService.getInferredMatches());
    setPatterns(mlInferenceService.getLearnedPatterns());
    setEditingMatchId(null);
    setFeedbackMsg('Ligação atualizada com sucesso e novo padrão memorizado!');
    setTimeout(() => setFeedbackMsg(null), 3500);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/70 backdrop-blur-xs overscroll-contain animate-in fade-in duration-200">
      <div className="bg-white rounded-3xl shadow-2xl border border-slate-200 w-full max-w-5xl max-h-[90vh] flex flex-col overflow-hidden overscroll-contain">
        
        {/* Header */}
        <div className="p-6 bg-slate-900 text-white flex items-center justify-between border-b border-slate-800">
          <div className="flex items-center space-x-3">
            <div className="w-10 h-10 rounded-xl bg-emerald-500/20 border border-emerald-500/30 flex items-center justify-center text-emerald-400">
              <BrainCircuit className="w-6 h-6" />
            </div>
            <div>
              <h3 className="text-lg font-black tracking-tight text-white">
                Auditoria IA
              </h3>
            </div>
          </div>

          <button
            onClick={onClose}
            className="p-2 text-slate-400 hover:text-white rounded-xl hover:bg-slate-800 transition cursor-pointer"
            title="Fechar modal"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Tab Selector & Metrics */}
        <div className="bg-slate-50 border-b border-slate-200 px-6 py-3 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={() => setActiveTab('validacao')}
              className={`px-4 py-2 rounded-xl text-xs font-bold transition flex items-center gap-2 cursor-pointer ${
                activeTab === 'validacao'
                  ? 'bg-slate-900 text-white shadow-xs'
                  : 'text-slate-600 hover:bg-slate-200/60'
              }`}
            >
              <span>Amostras de Validação</span>
              {pendingMatches.length > 0 && (
                <span className="bg-amber-500 text-slate-950 px-1.5 py-0.2 rounded-full text-[10px] font-black">
                  {pendingMatches.length}
                </span>
              )}
            </button>

            <button
              onClick={() => setActiveTab('padroes')}
              className={`px-4 py-2 rounded-xl text-xs font-bold transition flex items-center gap-2 cursor-pointer ${
                activeTab === 'padroes'
                  ? 'bg-slate-900 text-white shadow-xs'
                  : 'text-slate-600 hover:bg-slate-200/60'
              }`}
            >
              <Database className="w-3.5 h-3.5" />
              <span>Padrões Memorizados</span>
              <span className="bg-slate-200 text-slate-700 px-1.5 py-0.2 rounded-full text-[10px] font-black">
                {patterns.length}
              </span>
            </button>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <button
              type="button"
              onClick={() => setIsAddingManual((v) => !v)}
              className="px-3.5 py-1.5 bg-emerald-700 hover:bg-emerald-600 text-white text-xs font-bold rounded-xl transition flex items-center gap-1.5 shadow-xs cursor-pointer active:scale-98"
              title="Adicionar manualmente se uma NL virou PP e OB para o modelo aprender o padrão"
            >
              <Plus className="w-3.5 h-3.5" />
              <span>+ Adicionar Ligação Manual</span>
            </button>

            <button
              onClick={handleRunAudit}
              disabled={isScanning}
              className="px-3.5 py-1.5 bg-indigo-50 hover:bg-indigo-100 text-indigo-900 border border-indigo-200 text-xs font-bold rounded-xl transition flex items-center gap-1.5 shadow-2xs cursor-pointer active:scale-98 disabled:opacity-50"
              title="Reexecutar Auditoria"
            >
              <RefreshCw className={`w-3.5 h-3.5 text-indigo-600 ${isScanning ? 'animate-spin' : ''}`} />
              <span>{isScanning ? 'Analisando...' : 'Reexecutar Auditoria'}</span>
            </button>

            {activeTab === 'validacao' && pendingMatches.length > 0 && (
              <button
                onClick={handleConfirmAll}
                className="px-3.5 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold rounded-xl transition flex items-center gap-1.5 shadow-sm cursor-pointer active:scale-98"
              >
                <CheckCheck className="w-3.5 h-3.5" />
                <span>Aprovar Amostras de Alta Confiança</span>
              </button>
            )}
          </div>
        </div>

        {/* Filter Sub-Bar when on Validation Tab */}
        {activeTab === 'validacao' && matches.length > 0 && (
          <div className="bg-white border-b border-slate-200 px-6 py-2.5 flex flex-wrap items-center gap-2 text-xs">
            <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider mr-1">Filtrar:</span>
            <button
              type="button"
              onClick={() => setFilterMode('pendentes')}
              className={`px-2.5 py-1 rounded-lg text-xs font-bold transition cursor-pointer ${
                filterMode === 'pendentes'
                  ? 'bg-amber-100 text-amber-900 border border-amber-300'
                  : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
              }`}
            >
              Pendentes ({pendingMatches.length})
            </button>
            <button
              type="button"
              onClick={() => setFilterMode('f600')}
              className={`px-2.5 py-1 rounded-lg text-xs font-bold transition flex items-center gap-1.5 cursor-pointer ${
                filterMode === 'f600'
                  ? 'bg-emerald-100 text-emerald-950 border border-emerald-300'
                  : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
              }`}
            >
              <span className="w-1.5 h-1.5 rounded-full bg-emerald-500" />
              <span>Fonte 600 ({f600Matches.length})</span>
            </button>
            <button
              type="button"
              onClick={() => setFilterMode('confirmados')}
              className={`px-2.5 py-1 rounded-lg text-xs font-bold transition cursor-pointer ${
                filterMode === 'confirmados'
                  ? 'bg-blue-100 text-blue-900 border border-blue-300'
                  : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
              }`}
            >
              Confirmados ({confirmedMatches.length})
            </button>
            <button
              type="button"
              onClick={() => setFilterMode('all')}
              className={`px-2.5 py-1 rounded-lg text-xs font-bold transition cursor-pointer ${
                filterMode === 'all'
                  ? 'bg-slate-800 text-white'
                  : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
              }`}
            >
              Todos ({matches.length})
            </button>
          </div>
        )}

        {/* Error Alert */}
        {errorMsg && (
          <div className="bg-rose-50 border-b border-rose-200 px-6 py-2.5 text-xs text-rose-900 font-semibold flex items-center justify-between gap-2 animate-in fade-in">
            <div className="flex items-center gap-2">
              <AlertCircle className="w-4 h-4 text-rose-600 shrink-0" />
              <span>{errorMsg}</span>
            </div>
            <button
              type="button"
              onClick={() => setErrorMsg(null)}
              className="text-rose-500 hover:text-rose-800 p-0.5 cursor-pointer"
            >
              <X className="w-3.5 h-3.5" />
            </button>
          </div>
        )}

        {/* Feedback Alert */}
        {feedbackMsg && (
          <div className="bg-emerald-50 border-b border-emerald-200 px-6 py-2.5 text-xs text-emerald-900 font-semibold flex items-center gap-2 animate-in fade-in">
            <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
            <span>{feedbackMsg}</span>
          </div>
        )}

        {/* Content Area */}
        <div className="flex-1 overflow-y-auto p-6 space-y-4">
          {activeTab === 'validacao' && (
            <>
              {/* Form to add manual match (NL virou PP/OB) */}
              {isAddingManual && (
                <div className="bg-emerald-50/70 border border-emerald-300 rounded-2xl p-5 space-y-4 shadow-sm animate-in fade-in duration-200">
                  <div className="flex items-center justify-between border-b border-emerald-200 pb-2">
                    <div className="flex items-center gap-2">
                      <Sparkles className="w-4 h-4 text-emerald-700" />
                      <h4 className="text-sm font-bold text-emerald-950">
                        Adicionar Ligação Manual (NL → PP / OB)
                      </h4>
                    </div>
                    <button
                      type="button"
                      onClick={() => setIsAddingManual(false)}
                      className="text-slate-400 hover:text-slate-700 p-1 rounded-lg cursor-pointer"
                    >
                      <X className="w-4 h-4" />
                    </button>
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-3 text-xs">
                    <div>
                      <SuggestionInput
                        label="Nº da Liquidação (NL)"
                        required
                        value={manualNL}
                        onChange={(val) => setManualNL(val)}
                        onSelect={(item: NlSuggestionItem) => {
                          setManualNL(item.nlNumero);
                          setManualCredor(item.fornecedor);
                          if (item.cnpj) setManualCnpj(item.cnpj);
                          if (item.valor > 0) setManualValorNL(String(item.valor));
                          if (item.suggestedOB && !manualOB) setManualOB(item.suggestedOB);
                          if (item.suggestedPP && !manualPP) setManualPP(item.suggestedPP);
                        }}
                        suggestions={manualNlSuggestions}
                        placeholder="ex: 2026NL00123"
                        helperText="Sugere NLs das planilhas Obedece e Não Obedece"
                      />
                    </div>

                    <div>
                      <label className="block text-[11px] font-bold text-slate-700 mb-1">
                        Credor / Razão Social *
                      </label>
                      <input
                        type="text"
                        value={manualCredor}
                        onChange={(e) => setManualCredor(e.target.value)}
                        placeholder="Nome da Empresa"
                        className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-xl text-slate-900 focus:outline-none focus:ring-2 focus:ring-emerald-500"
                      />
                    </div>

                    <div>
                      <label className="block text-[11px] font-bold text-slate-700 mb-1">
                        CNPJ (Pessoa Jurídica)
                      </label>
                      <input
                        type="text"
                        value={manualCnpj}
                        onChange={(e) => setManualCnpj(e.target.value)}
                        placeholder="00.000.000/0000-00"
                        className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-xl text-slate-900 focus:outline-none focus:ring-2 focus:ring-emerald-500 font-mono"
                      />
                    </div>

                    <div>
                      <label className="block text-[11px] font-bold text-slate-700 mb-1">
                        Valor da NL (R$) *
                      </label>
                      <input
                        type="number"
                        step="0.01"
                        value={manualValorNL}
                        onChange={(e) => setManualValorNL(e.target.value)}
                        placeholder="0.00"
                        className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-xl text-slate-900 focus:outline-none focus:ring-2 focus:ring-emerald-500 font-mono"
                      />
                    </div>

                    <div>
                      <SuggestionInput
                        label="Nº da PP"
                        required
                        value={manualPP}
                        onChange={(val) => setManualPP(val)}
                        onSelect={(item: PpSuggestionItem) => {
                          setManualPP(item.ppNumero);
                          if (item.obNumero && !manualOB) setManualOB(item.obNumero);
                          if (item.valor && !manualValorPP) setManualValorPP(String(item.valor));
                          if (item.situacao) setManualSituacao(item.situacao.includes('CB') ? 'PPCB' : 'AO');
                        }}
                        suggestions={manualPpSuggestions}
                        placeholder="ex: 2026PP00456"
                        helperText="Selecione das PPs existentes ou digite nova"
                      />
                    </div>

                    <div>
                      <SuggestionInput
                        label="Nº da OB (Opcional)"
                        value={manualOB}
                        onChange={(val) => setManualOB(val)}
                        onSelect={(item: ObSuggestionItem) => {
                          setManualOB(item.obNumero);
                          if (item.ppNumero && !manualPP) setManualPP(item.ppNumero);
                          if (item.valor && !manualValorPP) setManualValorPP(String(item.valor));
                        }}
                        suggestions={manualObSuggestions}
                        placeholder="ex: 2026OB00789"
                        helperText="Selecione das OBs do app ou digite nova"
                      />
                    </div>

                    <div>
                      <label className="block text-[11px] font-bold text-slate-700 mb-1">
                        Situação da PP
                      </label>
                      <select
                        value={manualSituacao}
                        onChange={(e) => setManualSituacao(e.target.value)}
                        className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-xl text-slate-900 focus:outline-none focus:ring-2 focus:ring-emerald-500"
                      >
                        <option value="AO">AO (Aguardando Ordem / Pronta)</option>
                        <option value="PPCB">PPCB (Paga no Banco)</option>
                        <option value="PPNC">PPNC (Não Confirmada)</option>
                        <option value="PP Pronta">PP Pronta</option>
                      </select>
                    </div>

                    <div>
                      <label className="block text-[11px] font-bold text-slate-700 mb-1">
                        Valor da PP / Líquido (R$)
                      </label>
                      <input
                        type="number"
                        step="0.01"
                        value={manualValorPP}
                        onChange={(e) => setManualValorPP(e.target.value)}
                        placeholder="Deixe vazio se for igual à NL"
                        className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-xl text-slate-900 focus:outline-none focus:ring-2 focus:ring-emerald-500 font-mono"
                      />
                    </div>
                  </div>

                  <div className="flex items-center justify-end gap-2 pt-2 border-t border-emerald-200">
                    <button
                      type="button"
                      onClick={() => setIsAddingManual(false)}
                      className="px-3 py-1.5 bg-white text-slate-700 border border-slate-300 rounded-xl text-xs font-bold hover:bg-slate-50 cursor-pointer"
                    >
                      Cancelar
                    </button>
                    <button
                      type="button"
                      onClick={handleSaveManualMatch}
                      className="px-4 py-1.5 bg-emerald-700 hover:bg-emerald-600 text-white rounded-xl text-xs font-bold transition flex items-center gap-1.5 cursor-pointer shadow-sm active:scale-98"
                    >
                      <Check className="w-3.5 h-3.5" />
                      <span>Salvar e Memorizar Padrão</span>
                    </button>
                  </div>
                </div>
              )}

              {displayedMatches.length === 0 ? (
                <div className="py-14 text-center text-slate-500 space-y-3 max-w-lg mx-auto">
                  <div className="w-12 h-12 bg-indigo-50 border border-indigo-200 text-indigo-700 rounded-2xl flex items-center justify-center mx-auto">
                    <BrainCircuit className="w-6 h-6" />
                  </div>
                  <h4 className="text-sm font-bold text-slate-800">
                    {matches.length === 0
                      ? 'Nenhuma amostra pendente de validação'
                      : 'Nenhuma amostra encontrada para o filtro selecionado'}
                  </h4>
                  <div className="pt-2">
                    <button
                      type="button"
                      onClick={handleRunAudit}
                      disabled={isScanning}
                      className="px-4 py-2 bg-indigo-600 hover:bg-indigo-500 text-white rounded-xl text-xs font-bold transition flex items-center gap-2 mx-auto cursor-pointer shadow-md disabled:opacity-50"
                    >
                      <RefreshCw className={`w-3.5 h-3.5 ${isScanning ? 'animate-spin' : ''}`} />
                      <span>{isScanning ? 'Executando Auditoria...' : 'Executar Auditoria Agora'}</span>
                    </button>
                  </div>
                </div>
              ) : (
                <div className="space-y-3">
                  {displayedMatches.map((item) => {
                    const isPending = item.statusValidacao === 'pendente';
                    const isConfirmed = item.statusValidacao === 'confirmado';
                    const isRejected = item.statusValidacao === 'rejeitado';
                    const isEditing = editingMatchId === item.id;

                    return (
                      <div
                        key={item.id}
                        className={`rounded-2xl p-4 border transition-all ${
                          isConfirmed
                            ? 'bg-emerald-50/40 border-emerald-200'
                            : isRejected
                            ? 'bg-slate-50 border-slate-200 opacity-60'
                            : item.isFonte600
                            ? 'bg-amber-50/20 border-amber-300 shadow-xs'
                            : 'bg-white border-slate-200/90 shadow-sm hover:border-slate-300'
                        }`}
                      >
                        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
                          
                          {/* Creditor Info & Transition */}
                          <div className="space-y-2 flex-1 min-w-0">
                            <div className="flex items-center gap-2 flex-wrap">
                              <span className="font-bold text-slate-900 text-sm flex items-center gap-1.5">
                                <Building className="w-3.5 h-3.5 text-slate-400" />
                                <span className="truncate max-w-sm">{item.credor}</span>
                              </span>
                              <span className="text-[11px] font-mono bg-slate-100 text-slate-700 px-2 py-0.5 rounded border border-slate-200">
                                {item.cnpj || 'N/I'}
                              </span>
                              {item.isFonte600 && (
                                <span className="inline-flex items-center gap-1 text-[10px] font-black bg-emerald-100 text-emerald-900 border border-emerald-300 px-2 py-0.5 rounded-full">
                                  <span className="w-1.5 h-1.5 rounded-full bg-emerald-500" />
                                  Fonte 600 ({item.fonteRecurso || '0.6.00'})
                                </span>
                              )}
                              {item.automaticamenteConfirmado && (
                                <span className="inline-flex items-center gap-1 text-[10px] font-black bg-blue-100 text-blue-800 border border-blue-300 px-2 py-0.5 rounded-full">
                                  <Sparkles className="w-3 h-3 text-blue-600" />
                                  Auto-Confirmado via Padrão Memorizado
                                </span>
                              )}
                              {isConfirmed && !item.automaticamenteConfirmado && (
                                <span className="text-[10px] font-black bg-emerald-100 text-emerald-800 border border-emerald-300 px-2 py-0.5 rounded-full">
                                  ✓ Confirmado pelo Usuário
                                </span>
                              )}
                              {isRejected && (
                                <span className="text-[10px] font-black bg-rose-100 text-rose-800 border border-rose-300 px-2 py-0.5 rounded-full">
                                  ✕ Descartado
                                </span>
                              )}
                            </div>

                            {/* Transition Visual Block */}
                            {!isEditing ? (
                              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 bg-slate-50 p-3 rounded-xl border border-slate-200/80 text-xs">
                                {/* Left: NL Original (Liquidado) */}
                                <div className="space-y-1">
                                  <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider block">
                                    Liquidado ({item.baseOrigem === 'obedece' ? 'Obedece' : 'Não Obedece'})
                                  </span>
                                  <div className="font-mono font-bold text-slate-900">
                                    NL {item.nlNumero}
                                  </div>
                                  <div className="text-slate-600 font-mono font-bold text-xs">
                                    {formatBRL(item.valorNL)}
                                  </div>
                                </div>

                                {/* Center: Imposto Retido */}
                                <div className="space-y-1 border-y sm:border-y-0 sm:border-x border-slate-200/80 py-1 sm:py-0 sm:px-3">
                                  <span className="text-[10px] font-bold text-amber-800 uppercase tracking-wider block flex items-center gap-1">
                                    <TrendingDown className="w-3 h-3" />
                                    <span>Retenção Tributária</span>
                                  </span>
                                  <div className="text-amber-900 font-mono font-bold text-xs">
                                    {formatBRL(item.diferencaImpostos)}
                                  </div>
                                  <div className="text-[10px] font-semibold text-amber-700">
                                    {(item.percentualRetencao * 100).toFixed(2).replace('.', ',')}% retido
                                  </div>
                                </div>

                                {/* Right: A Pagar (PP / OB) */}
                                <div className="space-y-1">
                                  <span className="text-[10px] font-bold text-emerald-700 uppercase tracking-wider block flex items-center gap-1">
                                    <ArrowRight className="w-3 h-3" />
                                    <span>A Pagar ({item.statusSituacaoPP})</span>
                                  </span>
                                  <div className="font-mono font-bold text-emerald-950">
                                    PP {item.ppNumero} {item.obNumero ? `| OB ${item.obNumero}` : ''}
                                  </div>
                                  <div className="text-emerald-800 font-mono font-bold text-xs">
                                    {formatBRL(item.valorPP)}
                                  </div>
                                </div>
                              </div>
                            ) : (
                              /* Inline Editing Mode */
                              <div className="bg-amber-50/60 p-3.5 rounded-xl border border-amber-300 space-y-3 text-xs animate-in fade-in duration-150">
                                <div className="flex items-center justify-between">
                                  <span className="font-bold text-amber-950 text-xs flex items-center gap-1.5">
                                    <Edit3 className="w-3.5 h-3.5 text-amber-700" />
                                    <span>Editando Dados da PP / OB para NL {item.nlNumero}</span>
                                  </span>
                                  <span className="text-[11px] text-slate-500">
                                    Valor da NL: <strong className="font-mono text-slate-800">{formatBRL(item.valorNL)}</strong>
                                  </span>
                                </div>

                                <div className="grid grid-cols-1 sm:grid-cols-4 gap-2.5">
                                  <div>
                                    <SuggestionInput
                                      label="Nº da PP"
                                      required
                                      value={editPP}
                                      onChange={(val) => setEditPP(val)}
                                      onSelect={(item: PpSuggestionItem) => {
                                        setEditPP(item.ppNumero);
                                        if (item.obNumero && !editOB) setEditOB(item.obNumero);
                                        if (item.valor) setEditValorPP(String(item.valor));
                                        if (item.situacao) setEditSituacao(item.situacao.includes('CB') ? 'PPCB' : 'AO');
                                      }}
                                      suggestions={editPpSuggestions}
                                      placeholder="ex: 2026PP00123"
                                    />
                                  </div>

                                  <div>
                                    <SuggestionInput
                                      label="Nº da OB"
                                      value={editOB}
                                      onChange={(val) => setEditOB(val)}
                                      onSelect={(item: ObSuggestionItem) => {
                                        setEditOB(item.obNumero);
                                        if (item.ppNumero && !editPP) setEditPP(item.ppNumero);
                                        if (item.valor && !editValorPP) setEditValorPP(String(item.valor));
                                      }}
                                      suggestions={editObSuggestions}
                                      placeholder="Opcional"
                                    />
                                  </div>

                                  <div>
                                    <label className="block text-[10px] font-bold text-slate-700 mb-0.5">
                                      Situação
                                    </label>
                                    <select
                                      value={editSituacao}
                                      onChange={(e) => setEditSituacao(e.target.value)}
                                      className="w-full px-2.5 py-1 bg-white border border-slate-300 rounded-lg text-slate-900 text-xs focus:ring-2 focus:ring-emerald-500"
                                    >
                                      <option value="AO">AO (Aguardando Ordem)</option>
                                      <option value="PPCB">PPCB (Paga no Banco)</option>
                                      <option value="PPNC">PPNC (Não Confirmada)</option>
                                      <option value="PP Pronta">PP Pronta</option>
                                    </select>
                                  </div>

                                  <div>
                                    <label className="block text-[10px] font-bold text-slate-700 mb-0.5">
                                      Valor Líquido PP (R$)
                                    </label>
                                    <input
                                      type="number"
                                      step="0.01"
                                      value={editValorPP}
                                      onChange={(e) => setEditValorPP(e.target.value)}
                                      className="w-full px-2.5 py-1 bg-white border border-slate-300 rounded-lg text-slate-900 font-mono text-xs focus:ring-2 focus:ring-emerald-500"
                                    />
                                  </div>
                                </div>

                                <div className="flex items-center justify-end gap-2 pt-1 border-t border-amber-200">
                                  <button
                                    type="button"
                                    onClick={() => setEditingMatchId(null)}
                                    className="px-2.5 py-1 bg-white text-slate-700 border border-slate-300 rounded-lg text-[11px] font-bold hover:bg-slate-50 cursor-pointer"
                                  >
                                    Cancelar
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => handleSaveEdit(item.id)}
                                    className="px-3 py-1 bg-emerald-700 hover:bg-emerald-600 text-white rounded-lg text-[11px] font-bold transition flex items-center gap-1 cursor-pointer shadow-xs active:scale-98"
                                  >
                                    <Check className="w-3 h-3" />
                                    <span>Salvar Alterações</span>
                                  </button>
                                </div>
                              </div>
                            )}

                            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-1 text-[11px] text-slate-500">
                              <span>Motivo: {item.motivo} (Confiança estimada: {item.confianca}%)</span>
                              {isPending && (
                                <span className="font-bold text-amber-700">
                                  Aguardando confirmação do usuário
                                </span>
                              )}
                            </div>
                          </div>

                          {/* Action Buttons */}
                          <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2 shrink-0 self-end lg:self-center">
                            {/* Edit Button */}
                            {!isEditing && (
                              <button
                                type="button"
                                onClick={() => handleStartEdit(item)}
                                className="px-2.5 py-2 bg-slate-100 hover:bg-indigo-50 text-slate-700 hover:text-indigo-700 border border-slate-200 text-xs font-bold rounded-xl transition flex items-center justify-center gap-1 cursor-pointer active:scale-98"
                                title="Editar a PP e OB associadas a esta liquidação"
                              >
                                <Edit3 className="w-3.5 h-3.5" />
                                <span>Editar PP/OB</span>
                              </button>
                            )}

                            {isPending ? (
                              <>
                                <button
                                  type="button"
                                  onClick={() => handleConfirm(item.id, item.credor)}
                                  className="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold rounded-xl transition flex items-center justify-center gap-1.5 shadow-sm cursor-pointer active:scale-98"
                                  title="Confirmar esta ligação de despesa equivalente e memorizar para o fornecedor"
                                >
                                  <CheckCircle2 className="w-4 h-4" />
                                  <span>Sim, Confirmar Ligação</span>
                                </button>

                                <button
                                  type="button"
                                  onClick={() => handleReject(item.id)}
                                  className="px-3.5 py-2 bg-slate-100 hover:bg-rose-50 text-slate-700 hover:text-rose-700 border border-slate-200 text-xs font-bold rounded-xl transition flex items-center justify-center gap-1.5 cursor-pointer active:scale-98"
                                  title="Descartar esta correspondência como incorreta"
                                >
                                  <XCircle className="w-4 h-4 text-slate-500" />
                                  <span>Descartar</span>
                                </button>
                              </>
                            ) : (
                              <button
                                type="button"
                                onClick={() => handleReopen(item.id)}
                                className="px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-600 text-xs font-semibold rounded-xl transition cursor-pointer"
                                title="Reabrir validação desta amostra"
                              >
                                Reabrir Validação
                              </button>
                            )}
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </>
          )}

          {activeTab === 'padroes' && (
            <>
              {patterns.length === 0 ? (
                <div className="py-16 text-center text-slate-400 space-y-2">
                  <Database className="w-10 h-10 mx-auto text-slate-300" />
                  <p className="text-sm font-semibold text-slate-600">Nenhum padrão memorizado ainda</p>
                  <p className="text-xs text-slate-400 max-w-md mx-auto">
                    Ao confirmar que uma NL virou uma PP na aba de validação, a taxa de retenção tributária daquele credor será aprendida e guardada no navegador para auto-validação futura.
                  </p>
                </div>
              ) : (
                <div className="border border-slate-200 rounded-2xl overflow-hidden shadow-xs">
                  <table className="w-full text-left border-collapse text-xs">
                    <thead>
                      <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 font-extrabold text-[11px] uppercase">
                        <th className="py-3 px-4">CREDOR / FORNECEDOR</th>
                        <th className="py-3 px-3">CNPJ</th>
                        <th className="py-3 px-3 text-right">ALÍQUOTA DE RETENÇÃO APRENDIDA</th>
                        <th className="py-3 px-3 text-center">CONFIRMAÇÕES</th>
                        <th className="py-3 px-3">ÚLTIMA AMOSTRA</th>
                        <th className="py-3 px-3 text-center w-16">AÇÃO</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 font-medium">
                      {patterns.map((p) => (
                        <tr key={p.cnpj || p.credor} className="hover:bg-slate-50 transition-colors">
                          <td className="py-3 px-4 font-bold text-slate-900">
                            {p.credor}
                          </td>
                          <td className="py-3 px-3 font-mono text-slate-600">
                            {p.cnpj || 'N/I'}
                          </td>
                          <td className="py-3 px-3 text-right font-mono font-bold text-emerald-800">
                            {(p.aliquotaImpostoRetidoMedia * 100).toFixed(2).replace('.', ',')}% (±{(p.toleranciaPercentual * 100).toFixed(1).replace('.', ',')}%)
                          </td>
                          <td className="py-3 px-3 text-center">
                            <span className="bg-emerald-100 text-emerald-800 px-2 py-0.5 rounded-full font-bold text-[10px]">
                              {p.totalConfirmacoes}x confirmada
                            </span>
                          </td>
                          <td className="py-3 px-3 text-slate-500 font-mono text-[11px]">
                            {p.sampleNL} → {p.samplePP}
                          </td>
                          <td className="py-3 px-3 text-center">
                            <button
                              onClick={() => mlInferenceService.deleteLearnedPattern(p.cnpj || p.credor)}
                              className="p-1.5 text-slate-400 hover:text-rose-600 rounded-lg hover:bg-rose-50 transition cursor-pointer"
                              title="Remover padrão memorizado"
                            >
                              <Trash2 className="w-4 h-4" />
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </>
          )}
        </div>

        {/* Footer */}
        <div className="p-4 bg-slate-50 border-t border-slate-200 flex items-center justify-end">
          <button
            onClick={onClose}
            className="px-5 py-2 bg-slate-800 hover:bg-slate-700 text-white font-bold rounded-xl transition cursor-pointer"
          >
            Concluir
          </button>
        </div>

      </div>
    </div>
  );
};
