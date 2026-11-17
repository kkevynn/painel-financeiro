import React, { useEffect, useState } from 'react';
import { X, CheckCheck } from 'lucide-react';
import { ToastAlert } from '../types';
import { toastAlertService } from '../utils/toastAlertService';
import { formatBRL } from '../utils/excelParser';

interface ToastAlertContainerProps {
  onSelectToast?: (toast: ToastAlert) => void;
}

export const ToastAlertContainer: React.FC<ToastAlertContainerProps> = ({ onSelectToast }) => {
  const [toasts, setToasts] = useState<ToastAlert[]>([]);

  useEffect(() => {
    const unsubscribe = toastAlertService.subscribe((list) => {
      setToasts(list);
    });
    return () => unsubscribe();
  }, []);

  if (toasts.length === 0) return null;

  return (
    <aside
      aria-label="Notificações do Sistema"
      className="fixed bottom-5 right-5 z-50 flex flex-col gap-2 max-w-sm w-full pointer-events-none"
    >
      {toasts.length > 1 && (
        <div className="flex justify-end pointer-events-auto">
          <button
            onClick={() => toastAlertService.clearAll()}
            className="text-[11px] font-semibold text-slate-500 hover:text-slate-800 bg-white/95 px-2.5 py-1 rounded-full shadow-xs border border-slate-200 transition cursor-pointer flex items-center gap-1"
          >
            <CheckCheck className="w-3 h-3 text-emerald-600" />
            <span>Limpar ({toasts.length})</span>
          </button>
        </div>
      )}

      {toasts.map((toast) => {
        const isCritico = toast.tipo === 'critico';
        const isRequerValidacao = !!toast.requerValidacao;
        const isVirouOB = toast.situacaoNova === 'Virou OB' || toast.titulo.toLowerCase().includes('virou ob');

        return (
          <div
            key={toast.id}
            role="alert"
            onClick={() => onSelectToast?.(toast)}
            className={`pointer-events-auto transition duration-200 rounded-xl p-3 shadow-lg border backdrop-blur-md cursor-pointer hover:shadow-xl active:scale-[0.99] flex flex-col gap-2.5 ${
              isCritico
                ? 'bg-slate-900/95 text-white border-rose-500/80 shadow-rose-950/20'
                : isRequerValidacao
                ? 'bg-slate-900/98 text-white border-amber-400 shadow-amber-950/30 ring-1 ring-amber-400/30'
                : isVirouOB
                ? 'bg-slate-900/95 text-white border-blue-500/80 shadow-blue-950/20'
                : 'bg-slate-900/95 text-white border-emerald-500/80 shadow-emerald-950/20'
            }`}
            title="Clique para verificar detalhes da auditoria"
          >
            <div className="flex items-start justify-between gap-3">
              <div className="flex items-start gap-2.5 min-w-0 flex-1">
                <span
                  className={`w-2 h-2 rounded-full shrink-0 mt-1 ${
                    isCritico
                      ? 'bg-rose-500 ring-4 ring-rose-500/20'
                      : isRequerValidacao
                      ? 'bg-amber-400 ring-4 ring-amber-400/30 animate-pulse'
                      : isVirouOB
                      ? 'bg-blue-400 ring-4 ring-blue-400/20'
                      : 'bg-emerald-400 ring-4 ring-emerald-400/20'
                  }`}
                />

                <div className="min-w-0 flex-1">
                  <div className="flex items-center gap-1.5 flex-wrap">
                    <span className="font-mono font-bold text-xs text-white">
                      {toast.nl ? `NL ${toast.nl}` : toast.pp ? `PP ${toast.pp}` : toast.ob ? `OB ${toast.ob}` : 'S/N'}
                    </span>
                    {toast.pp && toast.nl && (
                      <span className="text-[10px] text-slate-300 font-mono">→ PP {toast.pp}</span>
                    )}
                    <span
                      className={`text-[9px] font-bold uppercase px-1.5 py-0.2 rounded ${
                        isCritico
                          ? 'bg-rose-500/20 text-rose-300 border border-rose-500/30'
                          : isRequerValidacao
                          ? 'bg-amber-500/20 text-amber-300 border border-amber-500/40 font-extrabold'
                          : isVirouOB
                          ? 'bg-blue-500/20 text-blue-300 border border-blue-500/30'
                          : 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/30'
                      }`}
                    >
                      {isRequerValidacao ? 'Validar Vínculo' : isVirouOB ? 'Virou OB' : toast.situacaoNova || (isCritico ? 'Rejeitada' : 'Confirmada')}
                    </span>
                    {toast.isFonte600 && (
                      <span className="text-[9px] font-black uppercase px-1.5 py-0.2 rounded bg-amber-500/20 text-amber-300 border border-amber-500/40">
                        F600
                      </span>
                    )}
                  </div>

                  <div className="text-[11px] font-medium text-slate-300 mt-0.5 line-clamp-2">
                    {toast.mensagem}
                  </div>

                  <div className="flex items-center justify-between gap-2 mt-1 text-xs text-slate-400">
                    <span className="truncate text-[11px]" title={toast.fornecedor}>
                      {toast.fornecedor || 'Credor'}
                    </span>
                    {toast.valor !== undefined && toast.valor > 0 && (
                      <span className="font-mono font-bold text-emerald-400 text-[11px] shrink-0">
                        {formatBRL(toast.valor)}
                      </span>
                    )}
                  </div>
                </div>
              </div>

              <button
                onClick={(e) => {
                  e.stopPropagation();
                  toastAlertService.dismissToast(toast.id);
                }}
                className="p-1 rounded-lg text-slate-400 hover:text-white transition shrink-0 cursor-pointer"
                title="Fechar notificação"
              >
                <X className="w-3.5 h-3.5" />
              </button>
            </div>

            {/* Interactive Confirmation Bar if Requires Validation */}
            {isRequerValidacao && (
              <div
                className="pt-2 border-t border-slate-700/80 flex items-center justify-between gap-2"
                onClick={(e) => e.stopPropagation()}
              >
                <span className="text-[10px] font-bold text-amber-300">
                  Está correto?
                </span>
                <div className="flex items-center gap-1.5">
                  <button
                    type="button"
                    onClick={() => toastAlertService.confirmTransition(toast.id)}
                    className="px-2.5 py-1 bg-emerald-600 hover:bg-emerald-500 text-white text-[10px] font-bold rounded-lg transition cursor-pointer shadow-xs active:scale-95"
                  >
                    ✓ Sim, Correto
                  </button>
                  <button
                    type="button"
                    onClick={() => toastAlertService.rejectTransition(toast.id)}
                    className="px-2 py-1 bg-slate-800 hover:bg-rose-950/80 text-rose-300 hover:text-rose-200 border border-slate-700 hover:border-rose-700/50 text-[10px] font-bold rounded-lg transition cursor-pointer active:scale-95"
                  >
                    ✕ Incorreto
                  </button>
                </div>
              </div>
            )}
          </div>
        );
      })}
    </aside>
  );
};
