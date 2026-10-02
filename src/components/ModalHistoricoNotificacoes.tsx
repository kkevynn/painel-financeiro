import React, { useState, useEffect } from 'react';
import { X, Search, RotateCw, CheckCircle2, AlertCircle, Clock, ExternalLink, Check } from 'lucide-react';
import { ToastAlert, FileData } from '../types';
import { toastAlertService } from '../utils/toastAlertService';
import { formatBRL } from '../utils/excelParser';
import { OrdemBancariaItem } from '../utils/ordemBancariaParser';

interface ModalHistoricoNotificacoesProps {
  isOpen: boolean;
  onClose: () => void;
  onSelectToast: (toast: ToastAlert) => void;
  files?: FileData[];
  ordensBancariasList?: OrdemBancariaItem[] | null;
  onSyncAlerts?: () => void;
}

export const ModalHistoricoNotificacoes: React.FC<ModalHistoricoNotificacoesProps> = ({
  isOpen,
  onClose,
  onSelectToast,
  files = [],
  ordensBancariasList,
  onSyncAlerts,
}) => {
  const [filter, setFilter] = useState<'all' | 'sucesso' | 'virou_ob' | 'critico'>('sucesso');
  const [searchTerm, setSearchTerm] = useState('');
  const [isSyncing, setIsSyncing] = useState(false);
  const [syncFeedback, setSyncFeedback] = useState<string | null>(null);
  const [confirmClear, setConfirmClear] = useState(false);
  const [allAlerts, setAllAlerts] = useState<ToastAlert[]>(() => toastAlertService.getAlertsHistory());

  // Subscribe to toastAlertService changes
  useEffect(() => {
    const handleUpdate = () => {
      setAllAlerts([...toastAlertService.getAlertsHistory()]);
    };
    handleUpdate();
    const unsub = toastAlertService.subscribe(handleUpdate);
    return () => unsub();
  }, []);

  // Load alerts when modal opens
  useEffect(() => {
    if (isOpen) {
      setAllAlerts([...toastAlertService.getAlertsHistory()]);
    }
  }, [isOpen]);

  if (!isOpen) return null;

  const countCritico = allAlerts.filter((a) => a.tipo === 'critico').length;
  const countConfirmadas = allAlerts.filter(
    (a) => a.tipo === 'sucesso' && a.situacaoNova !== 'Virou OB' && !a.titulo.toLowerCase().includes('virou ob')
  ).length;
  const countVirouOB = allAlerts.filter(
    (a) => a.situacaoNova === 'Virou OB' || a.titulo.toLowerCase().includes('virou ob')
  ).length;

  const handleSync = () => {
    setIsSyncing(true);
    setTimeout(() => {
      try {
        if (onSyncAlerts) {
          onSyncAlerts();
        } else {
          toastAlertService.syncConfirmations(files, ordensBancariasList, true);
        }
        setAllAlerts([...toastAlertService.getAlertsHistory()]);
        setSyncFeedback('Confirmações sincronizadas com sucesso com os arquivos atuais!');
        setTimeout(() => setSyncFeedback(null), 3500);
      } finally {
        setIsSyncing(false);
      }
    }, 200);
  };

  const handleResetBaseline = () => {
    toastAlertService.resetBaseline();
    setAllAlerts([]);
    setSyncFeedback('Base de comparação redefinida! O próximo upload definirá um novo ponto de partida.');
    setTimeout(() => setSyncFeedback(null), 3500);
  };

  const formatConfirmationTime = (d: Date | null, isCritico?: boolean, isVirouOB?: boolean) => {
    if (!d || isNaN(d.getTime())) return '';
    const now = new Date();
    const isToday =
      d.getDate() === now.getDate() &&
      d.getMonth() === now.getMonth() &&
      d.getFullYear() === now.getFullYear();
    const timeStr = d.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
    const prefix = isCritico ? 'Rejeitado' : isVirouOB ? 'Gerou OB' : 'Confirmado';

    if (isToday) return `${prefix} hoje às ${timeStr}`;

    const yesterday = new Date(now);
    yesterday.setDate(yesterday.getDate() - 1);
    const isYesterday =
      d.getDate() === yesterday.getDate() &&
      d.getMonth() === yesterday.getMonth() &&
      d.getFullYear() === yesterday.getFullYear();
    if (isYesterday) return `${prefix} ontem às ${timeStr}`;

    return `${prefix} em ${d.toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit', year: 'numeric' })} às ${timeStr}`;
  };

  const filtered = allAlerts.filter((a) => {
    // Filter type
    if (filter === 'critico' && a.tipo !== 'critico') return false;
    if (filter === 'sucesso' && (a.tipo !== 'sucesso' || a.situacaoNova === 'Virou OB' || a.titulo.toLowerCase().includes('virou ob'))) return false;
    if (filter === 'virou_ob' && a.situacaoNova !== 'Virou OB' && !a.titulo.toLowerCase().includes('virou ob')) return false;

    // Search term
    if (searchTerm.trim()) {
      const q = searchTerm.toLowerCase().trim();
      const matchesSupplier = (a.fornecedor || '').toLowerCase().includes(q);
      const matchesCnpj = (a.cnpj || '').includes(q);
      const matchesPp = (a.pp || '').toLowerCase().includes(q);
      const matchesOb = (a.ob || '').toLowerCase().includes(q);
      const matchesSit = (a.situacaoNova || '').toLowerCase().includes(q);
      if (!matchesSupplier && !matchesCnpj && !matchesPp && !matchesOb && !matchesSit) {
        return false;
      }
    }

    return true;
  });

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 backdrop-blur-xs animate-in fade-in duration-150"
      onClick={onClose}
    >
      <div
        className="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-xl max-h-[85vh] flex flex-col overflow-hidden animate-in zoom-in-95 duration-150"
        onClick={(e) => e.stopPropagation()}
        role="dialog"
        aria-modal="true"
      >
        {/* Header */}
        <div className="px-5 py-3.5 border-b border-slate-200 bg-slate-900 text-white flex items-center justify-between">
          <div className="flex items-center gap-2">
            <h3 className="text-sm font-bold text-white tracking-tight">
              Central de Notificações e Confirmações
            </h3>
            <span className="text-xs text-slate-400 font-mono">({allAlerts.length})</span>
          </div>

          <div className="flex items-center gap-2">
            <button
              onClick={handleSync}
              disabled={isSyncing}
              className="text-xs text-emerald-300 hover:text-white bg-slate-800 hover:bg-slate-700 px-3 py-1.5 rounded-lg transition flex items-center gap-1.5 cursor-pointer disabled:opacity-50 border border-slate-700 hover:border-emerald-500/50"
              title="Sincronizar e atualizar confirmações dos arquivos atuais"
            >
              <RotateCw className={`w-3.5 h-3.5 ${isSyncing ? 'animate-spin text-emerald-400' : ''}`} />
              <span>{isSyncing ? 'Sincronizando...' : 'Atualizar Confirmações'}</span>
            </button>

            <button
              onClick={onClose}
              className="p-1 rounded-lg text-slate-400 hover:text-white transition cursor-pointer"
              title="Fechar"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        </div>

        {/* Sync feedback notification if triggered */}
        {syncFeedback && (
          <div className="px-4 py-2 bg-emerald-50 border-b border-emerald-200 text-emerald-900 text-xs flex items-center gap-2 animate-in fade-in duration-150">
            <Check className="w-4 h-4 text-emerald-600 shrink-0" />
            <span className="font-semibold">{syncFeedback}</span>
          </div>
        )}

        {/* Search Input Bar */}
        <div className="px-4 py-2 bg-slate-50 border-b border-slate-200">
          <div className="relative">
            <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              placeholder="Buscar por credor, CNPJ, PP ou OB..."
              className="w-full pl-8 pr-3 py-1.5 text-xs bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-1 focus:ring-emerald-500"
            />
          </div>
        </div>

        {/* Filter bar */}
        <div className="px-4 py-2 bg-white border-b border-slate-200 flex flex-wrap items-center justify-between gap-2 text-xs">
          <div className="flex flex-wrap items-center gap-1.5">
            <button
              onClick={() => setFilter('sucesso')}
              className={`px-2.5 py-1 rounded-lg font-bold text-xs transition cursor-pointer ${
                filter === 'sucesso'
                  ? 'bg-emerald-600 text-white shadow-2xs'
                  : 'text-emerald-700 bg-emerald-50 hover:bg-emerald-100'
              }`}
            >
              Confirmadas ({countConfirmadas})
            </button>
            <button
              onClick={() => setFilter('virou_ob')}
              className={`px-2.5 py-1 rounded-lg font-bold text-xs transition cursor-pointer ${
                filter === 'virou_ob'
                  ? 'bg-blue-600 text-white shadow-2xs'
                  : 'text-blue-700 bg-blue-50 hover:bg-blue-100'
              }`}
            >
              Viraram OB ({countVirouOB})
            </button>
            <button
              onClick={() => setFilter('critico')}
              className={`px-2.5 py-1 rounded-lg font-bold text-xs transition cursor-pointer ${
                filter === 'critico'
                  ? 'bg-rose-600 text-white shadow-2xs'
                  : 'text-rose-700 bg-rose-50 hover:bg-rose-100'
              }`}
            >
              Rejeitadas ({countCritico})
            </button>
            <button
              onClick={() => setFilter('all')}
              className={`px-2.5 py-1 rounded-lg font-bold text-xs transition cursor-pointer ${
                filter === 'all'
                  ? 'bg-slate-800 text-white shadow-2xs'
                  : 'text-slate-600 hover:bg-slate-100'
              }`}
            >
              Todos ({allAlerts.length})
            </button>
          </div>

          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={() => {
                toastAlertService.resetBaseline();
                setAllAlerts([]);
                setSyncFeedback('Base de comparação redefinida.');
                setTimeout(() => setSyncFeedback(null), 2500);
              }}
              className="text-[11px] text-slate-400 hover:text-slate-700 font-medium transition cursor-pointer"
              title="Redefinir base para iniciar novo ciclo de comparação"
            >
              Redefinir Base
            </button>

            {allAlerts.length > 0 && (
              confirmClear ? (
                <div className="flex items-center gap-1.5 animate-in fade-in duration-100">
                  <span className="text-[11px] text-rose-700 font-semibold">Limpar tudo?</span>
                  <button
                    type="button"
                    onClick={() => {
                      toastAlertService.clearAlertsHistory();
                      setConfirmClear(false);
                    }}
                    className="px-2 py-0.5 bg-rose-600 hover:bg-rose-700 text-white rounded text-[10px] font-bold cursor-pointer"
                  >
                    Sim
                  </button>
                  <button
                    type="button"
                    onClick={() => setConfirmClear(false)}
                    className="px-2 py-0.5 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded text-[10px] font-bold cursor-pointer"
                  >
                    Não
                  </button>
                </div>
              ) : (
                <button
                  type="button"
                  onClick={() => setConfirmClear(true)}
                  className="text-[11px] text-slate-400 hover:text-rose-700 font-medium transition cursor-pointer"
                >
                  Limpar histórico
                </button>
              )
            )}
          </div>
        </div>

        {/* List of alerts */}
        <div className="p-3 overflow-y-auto space-y-1.5 flex-1 divide-y divide-slate-100">
          {filtered.length === 0 ? (
            <div className="text-center py-12 space-y-2 text-slate-400 text-xs">
              <AlertCircle className="w-8 h-8 text-slate-300 mx-auto" />
              <p>Nenhuma notificação localizada com os filtros selecionados.</p>
              <button
                type="button"
                onClick={handleSync}
                className="mt-2 inline-flex items-center gap-1.5 px-3 py-1.5 bg-emerald-50 text-emerald-800 border border-emerald-200 rounded-lg font-bold text-xs hover:bg-emerald-100 transition cursor-pointer"
              >
                <RotateCw className="w-3.5 h-3.5 text-emerald-600" />
                <span>Sincronizar com base nos arquivos atuais</span>
              </button>
            </div>
          ) : (
            filtered.map((item) => {
              const isCritico = item.tipo === 'critico';
              const isVirouOB = item.situacaoNova === 'Virou OB' || item.titulo.toLowerCase().includes('virou ob');
              const itemDate = item.dataHora ? new Date(item.dataHora) : null;
              const formattedDate = itemDate && !isNaN(itemDate.getTime())
                ? itemDate.toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit', year: 'numeric' })
                : '';
              const formattedTime = itemDate && !isNaN(itemDate.getTime())
                ? itemDate.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })
                : '';

              return (
                <div
                  key={item.id}
                  onClick={() => {
                    onSelectToast(item);
                    onClose();
                  }}
                  className="pt-2 first:pt-0 p-3 rounded-xl hover:bg-slate-50 transition cursor-pointer flex items-center justify-between gap-3 text-xs border border-transparent hover:border-slate-200"
                >
                  <div className="flex items-start gap-2.5 min-w-0 flex-1">
                    <span
                      className={`w-2.5 h-2.5 rounded-full shrink-0 mt-1 ${
                        isCritico ? 'bg-rose-500' : isVirouOB ? 'bg-blue-500' : 'bg-emerald-500'
                      }`}
                    />
                    <div className="min-w-0 flex-1">
                      <div className="flex flex-wrap items-center gap-2">
                        <span className="font-mono font-bold text-slate-900">
                          {item.ob ? `OB ${item.ob}` : item.pp ? `PP ${item.pp}` : item.ne ? `NE ${item.ne}` : 'S/N'}
                        </span>
                        {item.pp && item.ob && (
                          <span className="font-mono text-[11px] text-slate-500">
                            (PP {item.pp})
                          </span>
                        )}
                        <span
                          className={`text-[10px] font-extrabold uppercase px-2 py-0.5 rounded-md ${
                            isCritico
                              ? 'bg-rose-100 text-rose-800 border border-rose-200'
                              : isVirouOB
                              ? 'bg-blue-100 text-blue-800 border border-blue-200'
                              : 'bg-emerald-100 text-emerald-800 border border-emerald-200'
                          }`}
                        >
                          {isVirouOB ? 'Virou OB' : item.situacaoNova || (isCritico ? 'Rejeitada' : 'Confirmada')}
                        </span>
                        {item.situacaoAnterior && item.situacaoAnterior !== item.situacaoNova && (
                          <span className="text-[10px] text-slate-400 font-medium font-mono">
                            ({item.situacaoAnterior} → {item.situacaoNova || 'Confirmado'})
                          </span>
                        )}
                      </div>

                      <div className="text-[11px] text-slate-600 truncate mt-1">
                        <strong>{item.fornecedor || 'Credor não identificado'}</strong>
                        {item.cnpj && item.cnpj !== 'N/I' && (
                          <span className="text-slate-400 font-mono ml-1">({item.cnpj})</span>
                        )}
                      </div>

                      {item.dataReferencia && (
                        <div className="text-[10px] text-slate-600 font-medium mt-0.5 flex items-center gap-1">
                          <CheckCircle2 className="w-3 h-3 text-emerald-600 shrink-0" />
                          <span>Data no Documento: <strong className="text-slate-700">{item.dataReferencia}</strong></span>
                        </div>
                      )}
                    </div>
                  </div>

                  <div className="text-right shrink-0 flex flex-col items-end justify-center">
                    {item.valor !== undefined && item.valor > 0 && (
                      <span className="font-mono font-bold text-slate-900 block text-xs">
                        {formatBRL(item.valor)}
                      </span>
                    )}
                    <span
                      className={`text-[10px] font-semibold block mt-0.5 ${
                        isCritico ? 'text-rose-700' : isVirouOB ? 'text-blue-700' : 'text-emerald-700'
                      }`}
                      title={itemDate ? itemDate.toLocaleString('pt-BR') : ''}
                    >
                      {formatConfirmationTime(itemDate, isCritico, isVirouOB)}
                    </span>
                    <span className="text-[9px] text-slate-400 font-mono block">
                      {isCritico ? 'Rejeitado no Banco' : isVirouOB ? 'Liquidado → OB' : 'Confirmado no Banco'}
                    </span>
                  </div>
                </div>
              );
            })
          )}
        </div>

        {/* Footer */}
        <div className="px-5 py-2.5 bg-slate-50 border-t border-slate-200 flex items-center justify-between text-xs text-slate-500">
          <button
            onClick={handleResetBaseline}
            className="text-[11px] text-slate-500 hover:text-rose-600 transition underline cursor-pointer"
            title="Redefinir base comparativa para que o próximo upload sirva como nova referência"
          >
            Redefinir Base Comparativa
          </button>
          <button
            onClick={onClose}
            className="px-3.5 py-1.5 text-xs font-bold text-slate-600 hover:text-slate-900 bg-white hover:bg-slate-200 border border-slate-200 rounded-lg transition cursor-pointer"
          >
            Fechar
          </button>
        </div>
      </div>
    </div>
  );
};

