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
} from 'lucide-react';
import { InferredPaymentMatch, LearnedCreditorPattern, FileData } from '../types';
import { mlInferenceService } from '../utils/machineLearningInferenceService';
import { formatBRL } from '../utils/excelParser';
import { OrdemBancariaItem } from '../utils/ordemBancariaParser';

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
  const [matches, setMatches] = useState<InferredPaymentMatch[]>([]);
  const [patterns, setPatterns] = useState<LearnedCreditorPattern[]>([]);
  const [feedbackMsg, setFeedbackMsg] = useState<string | null>(null);
  const [isScanning, setIsScanning] = useState(false);

  const handleRunAudit = () => {
    setIsScanning(true);
    mlInferenceService.processFilesForInference(files, ordensBancariasList);
    setMatches(mlInferenceService.getInferredMatches());
    setPatterns(mlInferenceService.getLearnedPatterns());
    const count = mlInferenceService.getInferredMatches().length;
    setFeedbackMsg(`Auditoria finalizada! ${count} amostras detectadas.`);
    setTimeout(() => setIsScanning(false), 500);
    setTimeout(() => setFeedbackMsg(null), 3500);
  };

  useEffect(() => {
    if (!isOpen) return;

    if (files.length > 0 || (ordensBancariasList && ordensBancariasList.length > 0)) {
      mlInferenceService.processFilesForInference(files, ordensBancariasList);
    }

    const refreshData = () => {
      setMatches(mlInferenceService.getInferredMatches());
      setPatterns(mlInferenceService.getLearnedPatterns());
    };

    refreshData();
    const unsubscribe = mlInferenceService.subscribe(refreshData);
    return () => unsubscribe();
  }, [isOpen, files, ordensBancariasList]);

  if (!isOpen) return null;

  const pendingMatches = matches.filter((m) => m.statusValidacao === 'pendente');
  const confirmedMatches = matches.filter((m) => m.statusValidacao === 'confirmado');

  const handleConfirm = (id: string, credor: string) => {
    mlInferenceService.confirmMatch(id);
    setFeedbackMsg(`Padrão memorizado para "${credor}". Os próximos uploads reconhecerão essa taxa automaticamente.`);
    setTimeout(() => setFeedbackMsg(null), 4000);
  };

  const handleReject = (id: string) => {
    mlInferenceService.rejectMatch(id);
  };

  const handleConfirmAll = () => {
    mlInferenceService.confirmAllHighConfidence();
    setFeedbackMsg('Todas as amostras consistentes foram confirmadas e memorizadas!');
    setTimeout(() => setFeedbackMsg(null), 4000);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/70 backdrop-blur-xs animate-in fade-in duration-200">
      <div className="bg-white rounded-3xl shadow-2xl border border-slate-200 w-full max-w-5xl max-h-[90vh] flex flex-col overflow-hidden">
        
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
          <div className="flex items-center gap-2">
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

          <div className="flex items-center gap-2">
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
              {matches.length === 0 ? (
                <div className="py-14 text-center text-slate-500 space-y-3 max-w-lg mx-auto">
                  <div className="w-12 h-12 bg-indigo-50 border border-indigo-200 text-indigo-700 rounded-2xl flex items-center justify-center mx-auto">
                    <BrainCircuit className="w-6 h-6" />
                  </div>
                  <h4 className="text-sm font-bold text-slate-800">
                    Nenhuma amostra pendente de validação
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
                  {matches.map((item) => {
                    const isPending = item.statusValidacao === 'pendente';
                    const isConfirmed = item.statusValidacao === 'confirmado';
                    const isRejected = item.statusValidacao === 'rejeitado';

                    return (
                      <div
                        key={item.id}
                        className={`rounded-2xl p-4 border transition-all ${
                          isConfirmed
                            ? 'bg-emerald-50/40 border-emerald-200'
                            : isRejected
                            ? 'bg-slate-50 border-slate-200 opacity-60'
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
                            <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 bg-slate-50 p-3 rounded-xl border border-slate-200/80 text-xs">
                              {/* Left: NL Original */}
                              <div className="space-y-1">
                                <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider block">
                                  NL Original ({item.baseOrigem.toUpperCase()})
                                </span>
                                <div className="font-mono font-bold text-slate-900">
                                  {item.nlNumero}
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
                                  {(item.percentualRetencao * 100).toFixed(2)}% retido
                                </div>
                              </div>

                              {/* Right: PP Gerada */}
                              <div className="space-y-1">
                                <span className="text-[10px] font-bold text-emerald-700 uppercase tracking-wider block flex items-center gap-1">
                                  <ArrowRight className="w-3 h-3" />
                                  <span>PP Emitida ({item.statusSituacaoPP})</span>
                                </span>
                                <div className="font-mono font-bold text-emerald-950">
                                  {item.ppNumero} {item.obNumero ? `| ${item.obNumero}` : ''}
                                </div>
                                <div className="text-emerald-800 font-mono font-bold text-xs">
                                  {formatBRL(item.valorPP)}
                                </div>
                              </div>
                            </div>

                            <p className="text-[11px] text-slate-500 italic">
                              Motivo: {item.motivo} (Confiança estimada: {item.confianca}%)
                            </p>
                          </div>

                          {/* Action Buttons */}
                          {isPending && (
                            <div className="flex items-center gap-2 shrink-0 self-end lg:self-center">
                              <button
                                onClick={() => handleConfirm(item.id, item.credor)}
                                className="px-3 py-2 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold rounded-xl transition flex items-center gap-1.5 shadow-sm cursor-pointer active:scale-98"
                                title="Confirmar e memorizar padrão no localStorage"
                              >
                                <CheckCircle2 className="w-4 h-4" />
                                <span>Correto (Aprender Padrão)</span>
                              </button>

                              <button
                                onClick={() => handleReject(item.id)}
                                className="px-3 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold rounded-xl transition flex items-center gap-1.5 cursor-pointer"
                                title="Descartar correspondência"
                              >
                                <XCircle className="w-4 h-4 text-slate-500" />
                                <span>Incorreto</span>
                              </button>
                            </div>
                          )}
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
                            {(p.aliquotaImpostoRetidoMedia * 100).toFixed(2)}% (±{(p.toleranciaPercentual * 100).toFixed(1)}%)
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
