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
        const isVirouOB = toast.situacaoNova === 'Virou OB' || toast.titulo.toLowerCase().includes('virou ob');

        return (
          <div
            key={toast.id}
            role="alert"
            onClick={() => onSelectToast?.(toast)}
            className={`pointer-events-auto transition duration-200 rounded-xl p-3 shadow-lg border backdrop-blur-md cursor-pointer hover:shadow-xl active:scale-[0.99] flex items-center justify-between gap-3 ${
              isCritico
                ? 'bg-slate-900/95 text-white border-rose-500/80 shadow-rose-950/20'
                : isVirouOB
                ? 'bg-slate-900/95 text-white border-blue-500/80 shadow-blue-950/20'
                : 'bg-slate-900/95 text-white border-emerald-500/80 shadow-emerald-950/20'
            }`}
            title="Clique para verificar ou alterar"
          >
            <div className="flex items-center gap-2.5 min-w-0 flex-1">
              <span
                className={`w-2 h-2 rounded-full shrink-0 ${
                  isCritico
                    ? 'bg-rose-500 ring-4 ring-rose-500/20'
                    : isVirouOB
                    ? 'bg-blue-400 ring-4 ring-blue-400/20'
                    : 'bg-emerald-400 ring-4 ring-emerald-400/20'
                }`}
              />

              <div className="min-w-0 flex-1">
                <div className="flex items-center gap-2">
                  <span className="font-mono font-bold text-xs text-white">
                    {isVirouOB && toast.ob
                      ? `OB ${toast.ob}`
                      : toast.pp
                      ? `PP ${toast.pp}`
                      : toast.ob
                      ? `OB ${toast.ob}`
                      : 'S/N'}
                  </span>
                  <span
                    className={`text-[10px] font-bold uppercase px-1.5 py-0.2 rounded ${
                      isCritico
                        ? 'bg-rose-500/20 text-rose-300 border border-rose-500/30'
                        : isVirouOB
                        ? 'bg-blue-500/20 text-blue-300 border border-blue-500/30'
                        : 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/30'
                    }`}
                  >
                    {isVirouOB ? 'Virou OB' : toast.situacaoNova || (isCritico ? 'Rejeitada' : 'Confirmada')}
                  </span>
                </div>

                <div className="flex items-center justify-between gap-2 mt-0.5 text-xs text-slate-300">
                  <span className="truncate" title={toast.fornecedor}>
                    {toast.fornecedor || 'Credor'}
                  </span>
                  {toast.valor !== undefined && toast.valor > 0 && (
                    <span className="font-mono font-semibold text-white shrink-0">
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
              title="Fechar"
            >
              <X className="w-3.5 h-3.5" />
            </button>
          </div>
        );
      })}
    </aside>
  );
};
