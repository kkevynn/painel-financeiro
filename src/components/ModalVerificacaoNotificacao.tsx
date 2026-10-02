import React, { useState, useEffect } from 'react';
import {
  X,
  ArrowRight,
  Save,
  Check,
  ExternalLink,
} from 'lucide-react';
import { ToastAlert, FileData } from '../types';
import { formatBRL } from '../utils/excelParser';

interface ModalVerificacaoNotificacaoProps {
  toast: ToastAlert | null;
  isOpen: boolean;
  onClose: () => void;
  files: FileData[];
  onSaveAlteration: (
    toast: ToastAlert,
    changes: {
      situacao: string;
      valor?: number;
      observacao?: string;
    }
  ) => void;
  onOpenCompanyAudit?: (supplier: { name: string; cnpj: string }) => void;
}

const SITUACOES_SUGERIDAS = [
  { value: 'AO', label: 'AO - Associada 2 Ordenadores (PP Pronta)' },
  { value: 'PPCB', label: 'PPCB - Confirmado no Banco' },
  { value: 'PPNC', label: 'PPNC - Não Confirmado no Banco' },
  { value: 'PPRJ', label: 'PPRJ - Rejeitada pelo Banco' },
  { value: 'RJ', label: 'RJ - Rejeitada' },
  { value: 'PAGA', label: 'PAGA - Paga' },
  { value: 'CANCELADA', label: 'CANCELADA' },
  { value: 'LIQUIDADA', label: 'LIQUIDADA' },
];

export const ModalVerificacaoNotificacao: React.FC<ModalVerificacaoNotificacaoProps> = ({
  toast,
  isOpen,
  onClose,
  files,
  onSaveAlteration,
  onOpenCompanyAudit,
}) => {
  if (!isOpen || !toast) return null;

  const isCritico = toast.tipo === 'critico';

  const [situacao, setSituacao] = useState<string>(toast.situacaoNova || 'AO');
  const [valorStr, setValorStr] = useState<string>(
    toast.valor !== undefined ? toast.valor.toString() : ''
  );
  const [observacao, setObservacao] = useState<string>(toast.observacao || '');
  const [isSaved, setIsSaved] = useState(false);

  useEffect(() => {
    if (toast) {
      setSituacao(toast.situacaoNova || 'AO');
      setValorStr(toast.valor !== undefined ? toast.valor.toString() : '');
      setObservacao(toast.observacao || '');
      setIsSaved(false);
    }
  }, [toast]);

  const handleSave = () => {
    const parsedVal = valorStr ? parseFloat(valorStr.replace(',', '.')) : toast.valor;
    onSaveAlteration(toast, {
      situacao: situacao.trim(),
      valor: isNaN(parsedVal as number) ? toast.valor : parsedVal,
      observacao: observacao.trim(),
    });
    setIsSaved(true);
    setTimeout(() => {
      setIsSaved(false);
      onClose();
    }, 800);
  };

  const handleAuditCompany = () => {
    if (onOpenCompanyAudit && toast.fornecedor) {
      onOpenCompanyAudit({
        name: toast.fornecedor,
        cnpj: toast.cnpj || '',
      });
      onClose();
    }
  };

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 backdrop-blur-xs animate-in fade-in duration-150"
      onClick={onClose}
    >
      <div
        className="bg-white rounded-2xl shadow-xl border border-slate-200 w-full max-w-lg overflow-hidden animate-in zoom-in-95 duration-150"
        onClick={(e) => e.stopPropagation()}
        role="dialog"
        aria-modal="true"
      >
        {/* Header */}
        <div className="px-5 py-3.5 border-b border-slate-200 flex items-center justify-between bg-slate-900 text-white">
          <div className="flex items-center gap-2.5">
            <span
              className={`w-2 h-2 rounded-full ${
                isCritico ? 'bg-rose-500' : 'bg-emerald-400'
              }`}
            />
            <h3 className="text-sm font-bold font-mono tracking-tight text-white">
              PP {toast.pp || toast.ob || 'S/N'}
            </h3>
            <span
              className={`text-[10px] font-bold uppercase px-1.5 py-0.2 rounded ${
                isCritico
                  ? 'bg-rose-500/20 text-rose-300 border border-rose-500/30'
                  : 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/30'
              }`}
            >
              {toast.situacaoNova || (isCritico ? 'Rejeitada' : 'Confirmada')}
            </span>
          </div>

          <button
            onClick={onClose}
            className="p-1 rounded-lg text-slate-400 hover:text-white transition cursor-pointer"
            title="Fechar"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Body */}
        <div className="p-5 space-y-4 text-slate-800 text-xs">
          
          {/* Informações Principais */}
          <div className="bg-slate-50 rounded-xl p-3.5 border border-slate-200/80 space-y-2.5">
            <div className="flex items-start justify-between gap-2">
              <div className="min-w-0 flex-1">
                <span className="font-bold text-slate-900 block truncate" title={toast.fornecedor}>
                  {toast.fornecedor || 'Credor não identificado'}
                </span>
                {toast.cnpj && (
                  <span className="text-[11px] font-mono text-slate-500">
                    CNPJ: {toast.cnpj}
                  </span>
                )}
              </div>

              {toast.fornecedor && onOpenCompanyAudit && (
                <button
                  type="button"
                  onClick={handleAuditCompany}
                  className="text-emerald-700 hover:text-emerald-800 font-semibold flex items-center gap-1 shrink-0 cursor-pointer text-[11px]"
                  title="Abrir extrato do credor"
                >
                  <span>Auditar</span>
                  <ExternalLink className="w-3 h-3" />
                </button>
              )}
            </div>

            <div className="flex items-center justify-between pt-2 border-t border-slate-200/60 font-mono text-xs">
              <div className="flex items-center gap-1.5">
                <span className="px-1.5 py-0.5 rounded bg-slate-200 text-slate-700 font-bold">
                  {toast.situacaoAnterior || 'AO'}
                </span>
                <ArrowRight className="w-3 h-3 text-slate-400" />
                <span
                  className={`px-1.5 py-0.5 rounded font-bold ${
                    isCritico
                      ? 'bg-rose-100 text-rose-800 border border-rose-200'
                      : 'bg-emerald-100 text-emerald-800 border border-emerald-200'
                  }`}
                >
                  {toast.situacaoNova || 'N/I'}
                </span>
              </div>

              {toast.valor !== undefined && toast.valor > 0 && (
                <span className="font-bold text-slate-900 text-sm">
                  {formatBRL(toast.valor)}
                </span>
              )}
            </div>
          </div>

          {/* Campos de Alteração */}
          <div className="space-y-3 pt-1">
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div>
                <label className="block text-[11px] font-bold text-slate-600 mb-1">
                  Situação:
                </label>
                <select
                  value={situacao}
                  onChange={(e) => setSituacao(e.target.value)}
                  className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs font-semibold text-slate-800 focus:ring-1 focus:ring-emerald-500 focus:border-emerald-500 transition cursor-pointer"
                >
                  {SITUACOES_SUGERIDAS.map((s) => (
                    <option key={s.value} value={s.value}>
                      {s.label}
                    </option>
                  ))}
                  {!SITUACOES_SUGERIDAS.some((s) => s.value === situacao) && (
                    <option value={situacao}>{situacao}</option>
                  )}
                </select>
              </div>

              <div>
                <label className="block text-[11px] font-bold text-slate-600 mb-1">
                  Valor (R$):
                </label>
                <input
                  type="text"
                  value={valorStr}
                  onChange={(e) => setValorStr(e.target.value)}
                  placeholder="0.00"
                  className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs font-mono font-bold text-slate-800 focus:ring-1 focus:ring-emerald-500 focus:border-emerald-500 transition"
                />
              </div>
            </div>

            <div>
              <label className="block text-[11px] font-bold text-slate-600 mb-1">
                Observação (opcional):
              </label>
              <textarea
                value={observacao}
                onChange={(e) => setObservacao(e.target.value)}
                placeholder="Motivo ou nota da alteração..."
                rows={2}
                className="w-full bg-white border border-slate-300 rounded-lg p-2 text-xs text-slate-800 focus:ring-1 focus:ring-emerald-500 focus:border-emerald-500 transition resize-none"
              />
            </div>
          </div>

        </div>

        {/* Footer */}
        <div className="px-5 py-3 bg-slate-50 border-t border-slate-200 flex items-center justify-end gap-2">
          <button
            type="button"
            onClick={onClose}
            className="px-3 py-1.5 text-xs font-semibold text-slate-600 hover:text-slate-800 hover:bg-slate-200 rounded-lg transition cursor-pointer"
          >
            Cancelar
          </button>

          <button
            type="button"
            onClick={handleSave}
            className={`px-4 py-1.5 rounded-lg text-xs font-bold transition flex items-center gap-1.5 shadow-xs cursor-pointer ${
              isSaved
                ? 'bg-emerald-600 text-white'
                : 'bg-emerald-700 hover:bg-emerald-800 text-white'
            }`}
          >
            {isSaved ? (
              <>
                <Check className="w-3.5 h-3.5" />
                <span>Salvo!</span>
              </>
            ) : (
              <>
                <Save className="w-3.5 h-3.5" />
                <span>Salvar</span>
              </>
            )}
          </button>
        </div>
      </div>
    </div>
  );
};
