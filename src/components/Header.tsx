import React, { useState } from 'react';
import {
  Menu,
  Save,
  FileDown,
  Trash2,
  Calendar,
  Check,
  BrainCircuit,
  Bell,
  Loader2,
} from 'lucide-react';

interface HeaderProps {
  activeTab: string;
  onToggleSidebar?: () => void;
  onExportStandaloneHtml: () => void;
  onClearData: () => void;
  onSaveData?: () => Promise<void> | void;
  onOpenAuditoriaML?: () => void;
  pendingInferencesCount?: number;
  onOpenNotifications?: () => void;
  unreadAlertsCount?: number;
  isProcessed: boolean;
  lastProcessedAt?: Date | string | null;
}

const TAB_TITLES: Record<string, { title: string; subtitle: string }> = {
  painel: {
    title: 'Painel Geral',
    subtitle: 'Relatórios de Liquidação e Ordens de Pagamento',
  },
  pagamentos_fonte: {
    title: 'Pagamentos por Fonte de Recursos',
    subtitle: 'Divisão, detalhamento e consolidação das Ordens Bancárias por Fonte e Credor',
  },
  fornecedores_mes: {
    title: 'Controle Mensal de Fornecedores',
    subtitle: 'Acompanhamento e registro mensal de fornecedores atendidos e pendentes',
  },
  quebra_ordem: {
    title: 'Monitoramento de Publicações',
    subtitle: 'Cruzamento de Liquidações e Ordens Bancárias (AO) por Quebra de Ordem Cronológica',
  },
  recorrencia: {
    title: 'Fornecedores Recorrentes',
    subtitle: 'Identificação e Análise de Credores Recorrentes',
  },
  categorias: {
    title: 'Grupos de Fornecedores',
    subtitle: 'Classificação por Categoria e Métricas de Liquidação e PPs',
  },
  importacoes: {
    title: 'Central de Importações',
    subtitle: 'Carregamento e Mapeamento de Planilhas',
  },
};

export const Header: React.FC<HeaderProps> = ({
  activeTab,
  onToggleSidebar,
  onExportStandaloneHtml,
  onClearData,
  onSaveData,
  onOpenAuditoriaML,
  pendingInferencesCount,
  onOpenNotifications,
  unreadAlertsCount,
  isProcessed,
  lastProcessedAt,
}) => {
  const [isSaving, setIsSaving] = useState(false);
  const [savedNotification, setSavedNotification] = useState(false);
  const [confirmClear, setConfirmClear] = useState(false);

  const activeInfo = TAB_TITLES[activeTab] || TAB_TITLES.painel;

  const formattedUpdateDate = React.useMemo(() => {
    if (!lastProcessedAt) {
      return 'Nenhuma atualização realizada';
    }
    const d = typeof lastProcessedAt === 'string' ? new Date(lastProcessedAt) : lastProcessedAt;
    if (isNaN(d.getTime())) return 'Nenhuma atualização realizada';

    return (
      'Atualizado em: ' +
      d.toLocaleDateString('pt-BR', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric',
      }) +
      ' - ' +
      d.toLocaleTimeString('pt-BR')
    );
  }, [lastProcessedAt]);

  const handleSaveClick = async () => {
    if (onSaveData && !isSaving) {
      setIsSaving(true);
      try {
        await onSaveData();
        setSavedNotification(true);
        setTimeout(() => setSavedNotification(false), 2500);
      } finally {
        setIsSaving(false);
      }
    }
  };

  return (
    <header className="bg-white/95 backdrop-blur-md border-b border-slate-200/90 shadow-xs sticky top-0 z-30 px-4 sm:px-6 py-3.5 transition-all">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        
        {/* Toggle Button & Title */}
        <div className="flex items-center gap-3.5">
          {onToggleSidebar && (
            <button
              type="button"
              onClick={onToggleSidebar}
              className="p-2 rounded-xl text-slate-700 bg-slate-100 hover:bg-emerald-50 hover:text-emerald-800 border border-slate-200 shadow-2xs transition-all flex items-center justify-center cursor-pointer active:scale-95 shrink-0"
              title="Abrir Menu de Navegação"
              aria-label="Abrir Menu de Navegação"
            >
              <Menu className="w-5 h-5" />
            </button>
          )}

          <div>
            <h1 className="text-xl font-bold text-slate-900 tracking-tight">
              {activeInfo.title}
            </h1>
          </div>
        </div>

        {/* Date/Time + Action Buttons */}
        <div className="flex flex-wrap items-center gap-3">
          
          {/* Updated Date & Time Display */}
          <div className="bg-emerald-50/80 text-emerald-900 px-3.5 py-1.5 rounded-xl border border-emerald-200 text-xs font-semibold flex items-center gap-2">
            <Calendar className="w-3.5 h-3.5 text-emerald-700" />
            <span className="font-mono">{formattedUpdateDate}</span>
          </div>

          {/* Auditoria IA Button */}
          {onOpenAuditoriaML && (
            <button
              type="button"
              onClick={onOpenAuditoriaML}
              className="px-3.5 py-2 rounded-xl text-xs font-bold transition flex items-center gap-1.5 border shadow-2xs cursor-pointer bg-indigo-50 hover:bg-indigo-100 text-indigo-900 border-indigo-200 active:scale-95"
              title="Auditoria de Inferências e Aprendizado de Máquina (Conversão de NL em PP)"
            >
              <BrainCircuit className="w-4 h-4 text-indigo-600" />
              <span>Auditoria IA</span>
              {typeof pendingInferencesCount === 'number' && pendingInferencesCount > 0 && (
                <span className="bg-indigo-600 text-white text-[10px] font-black px-1.5 py-0.5 rounded-full leading-none">
                  {pendingInferencesCount}
                </span>
              )}
            </button>
          )}

          {/* Central de Notificações e Alertas */}
          {onOpenNotifications && (
            <button
              type="button"
              onClick={onOpenNotifications}
              className="px-3.5 py-2 rounded-xl text-xs font-bold transition flex items-center gap-1.5 border shadow-2xs cursor-pointer bg-amber-50 hover:bg-amber-100 text-amber-900 border-amber-200 active:scale-95"
              title="Central de Notificações e Alertas do Sistema"
            >
              <Bell className="w-4 h-4 text-amber-600" />
              <span>Alertas</span>
              {typeof unreadAlertsCount === 'number' && unreadAlertsCount > 0 && (
                <span className="bg-amber-600 text-white text-[10px] font-black px-1.5 py-0.5 rounded-full leading-none">
                  {unreadAlertsCount}
                </span>
              )}
            </button>
          )}

          {/* Salvar Button */}
          <button
            onClick={handleSaveClick}
            disabled={isSaving}
            className={`px-3.5 py-2 rounded-xl text-xs font-bold transition flex items-center gap-1.5 border shadow-xs active:scale-95 ${
              savedNotification
                ? 'bg-emerald-600 text-white border-emerald-600'
                : isSaving
                ? 'bg-slate-100 text-emerald-800 border-emerald-300 cursor-wait'
                : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50 hover:border-slate-400 cursor-pointer'
            }`}
          >
            {isSaving ? (
              <>
                <Loader2 className="w-4 h-4 animate-spin text-emerald-600" />
                <span>Salvando...</span>
              </>
            ) : savedNotification ? (
              <>
                <Check className="w-4 h-4 text-white" />
                <span>Salvo!</span>
              </>
            ) : (
              <>
                <Save className="w-4 h-4 text-slate-600" />
                <span>Salvar</span>
              </>
            )}
          </button>

          {/* Clear Button */}
          {isProcessed && (
            confirmClear ? (
              <div className="flex items-center gap-1.5 bg-rose-50 border border-rose-300 rounded-xl px-2.5 py-1.5 animate-in fade-in duration-100">
                <span className="text-[11px] font-bold text-rose-800">Limpar tudo?</span>
                <button
                  type="button"
                  onClick={() => {
                    setConfirmClear(false);
                    onClearData();
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
                className="px-3 py-2 bg-white text-rose-700 hover:bg-rose-50 border border-rose-200 text-xs font-bold rounded-xl transition flex items-center gap-1.5 cursor-pointer"
                title="Limpar dados carregados"
              >
                <Trash2 className="w-4 h-4" />
                <span>Limpar</span>
              </button>
            )
          )}

        </div>

      </div>
    </header>
  );
};
