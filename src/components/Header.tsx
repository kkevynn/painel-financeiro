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
  RefreshCw,
  LogOut,
} from 'lucide-react';
import { GoogleDriveState } from '../utils/googleDriveService';
import { CloudSyncState } from '../utils/cloudSyncService';
import { UserRole } from '../utils/authService';

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

  // Google Drive & Sheets Integration
  driveState?: GoogleDriveState;
  onConnectDrive?: () => Promise<void> | void;
  onDisconnectDrive?: () => void;
  onSyncDriveNow?: () => Promise<void> | void;
  isDriveSyncing?: boolean;

  // Cloud Multi-Device Synchronization (Tablet & PC)
  cloudSyncState?: CloudSyncState;

  // Discreet Progress Bar
  isProcessing?: boolean;
  processingProgress?: number;
  processingMessage?: string;

  // RBAC Profile & Session
  userRole?: UserRole;
  userName?: string;
  onLogout?: () => void;
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
  publicacao: {
    title: 'NLs para Publicação por Dia',
    subtitle: 'Acompanhamento diário de publicação, emissão de PP/OB e pendências em Obedece',
  },
  para_publicar: {
    title: 'Para Publicar',
    subtitle: 'Liquidações selecionadas para publicação no Diário Oficial (DOE)',
  },
  fornecedores_mes: {
    title: 'Controle Mensal de Fornecedores',
    subtitle: 'Acompanhamento e registro mensal de fornecedores atendidos e pendentes',
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
  driveState,
  onConnectDrive,
  onDisconnectDrive,
  onSyncDriveNow,
  isDriveSyncing,
  cloudSyncState,
  isProcessing,
  processingProgress,
  processingMessage,
  userRole,
  userName,
  onLogout,
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
          
          {/* Updated Date & Time Display OR Discreet Updating Indicator */}
          {isProcessing ? (
            <div className="bg-emerald-50 text-emerald-900 px-3.5 py-1.5 rounded-xl border border-emerald-300 text-xs font-bold flex items-center gap-2 shadow-2xs">
              <RefreshCw className="w-3.5 h-3.5 animate-spin text-emerald-600" />
              <span>{processingMessage || 'Atualizando informações...'} ({processingProgress || 100}%)</span>
            </div>
          ) : (
            <div className="bg-emerald-50/80 text-emerald-900 px-3.5 py-1.5 rounded-xl border border-emerald-200 text-xs font-semibold flex items-center gap-2">
              <Calendar className="w-3.5 h-3.5 text-emerald-700" />
              <span className="font-mono">{formattedUpdateDate}</span>
            </div>
          )}

          {/* Auditoria IA Button - Admin Only */}
          {userRole === 'admin' && onOpenAuditoriaML && (
            <button
              type="button"
              onClick={onOpenAuditoriaML}
              className={`px-3.5 py-2 rounded-xl text-xs font-bold transition flex items-center gap-1.5 border shadow-2xs cursor-pointer active:scale-95 ${
                typeof pendingInferencesCount === 'number' && pendingInferencesCount > 0
                  ? 'bg-amber-50 hover:bg-amber-100 text-amber-950 border-amber-300 ring-2 ring-amber-400/30'
                  : 'bg-indigo-50 hover:bg-indigo-100 text-indigo-900 border-indigo-200'
              }`}
              title="Auditoria de Inferências e Validação de Conversões (NL em PP / OB)"
            >
              <BrainCircuit
                className={`w-4 h-4 ${
                  typeof pendingInferencesCount === 'number' && pendingInferencesCount > 0
                    ? 'text-amber-600 animate-pulse'
                    : 'text-indigo-600'
                }`}
              />
              <span>Auditoria IA</span>
              {typeof pendingInferencesCount === 'number' && pendingInferencesCount > 0 && (
                <span className="bg-amber-600 text-white text-[10px] font-black px-1.5 py-0.5 rounded-full leading-none">
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

          {/* User Role Badge & Logout Button */}
          {userRole && (
            <div className="flex items-center gap-2">
              <span
                className={`px-3 py-1.5 rounded-xl text-xs font-black border flex items-center gap-1.5 ${
                  userRole === 'admin'
                    ? 'bg-emerald-50 text-emerald-900 border-emerald-300'
                    : 'bg-sky-50 text-sky-900 border-sky-300'
                }`}
                title={`Perfil ativo: ${userName || (userRole === 'admin' ? 'Administrador' : 'Modo Consulta')}`}
              >
                <span className={`w-2 h-2 rounded-full ${userRole === 'admin' ? 'bg-emerald-500' : 'bg-sky-500'}`} />
                <span>{userRole === 'admin' ? 'Administrador' : 'Modo Consulta'}</span>
              </span>

              {onLogout && (
                <button
                  type="button"
                  onClick={onLogout}
                  className="px-3 py-1.5 bg-white hover:bg-rose-50 text-slate-600 hover:text-rose-700 border border-slate-300 hover:border-rose-300 rounded-xl text-xs font-bold transition flex items-center gap-1.5 cursor-pointer active:scale-95 shadow-2xs"
                  title="Trocar de conta / Sair do sistema"
                >
                  <LogOut className="w-3.5 h-3.5 text-slate-500 hover:text-rose-600" />
                  <span className="hidden sm:inline">Sair</span>
                </button>
              )}
            </div>
          )}

          {/* Google Integration Controls (Drive + Cloud Sync Tablet & PC) - Admin Only */}
          {driveState && userRole !== 'viewer' && (
            <div className="flex items-center gap-2">
              {!driveState.isConnected ? (
                <button
                  type="button"
                  onClick={onConnectDrive}
                  disabled={driveState.status === 'connecting'}
                  className="px-3.5 py-2 rounded-xl text-xs font-bold transition flex items-center gap-2 border shadow-2xs cursor-pointer active:scale-95 bg-white hover:bg-slate-50 text-slate-700 border-slate-300"
                  title="Conectar Conta Google para sincronização automática entre Tablet, PC e relatórios do SIGEF"
                >
                  {driveState.status === 'connecting' ? (
                    <Loader2 className="w-3.5 h-3.5 animate-spin text-emerald-600" />
                  ) : (
                    <svg className="w-4 h-4 shrink-0" viewBox="0 0 24 24" fill="none">
                      <path d="M7.71 3.5L1.15 15l3.43 6 6.55-11.5L7.71 3.5z" fill="#0066DA" />
                      <path d="M16.29 3.5H7.71l6.56 11.5h8.58L16.29 3.5z" fill="#00AC47" />
                      <path d="M14.27 15l-3.43 6h11.96l3.43-6H14.27z" fill="#EA4335" />
                      <path d="M7.71 3.5l3.43 6-3.43 6-3.43-6 3.43-6z" fill="#00832D" />
                      <path d="M14.27 15l-3.14-5.5-3.42 6h6.56z" fill="#2684FC" />
                      <path d="M22.85 15h-8.58l3.44-6h5.14z" fill="#FFBA00" />
                    </svg>
                  )}
                  <span>{driveState.status === 'connecting' ? 'Conectando...' : 'Conectar Conta Google'}</span>
                </button>
              ) : (
                <div className="flex items-center gap-1.5 bg-slate-50 border border-slate-200 p-1 rounded-xl">
                  {/* Status Indicator */}
                  <div
                    className="flex items-center gap-1.5 px-2.5 py-1 text-xs font-medium text-slate-700"
                    title={`Conta Google Conectada (${driveState.userEmail || ''}). Sincronização em nuvem e Drive ativas entre Tablet e PC.`}
                  >
                    <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse" />
                    <span className="font-semibold text-slate-800 max-w-[120px] sm:max-w-[160px] truncate">
                      {driveState.userEmail || 'Conta Google'}
                    </span>
                  </div>

                  {/* Desconectar */}
                  {onDisconnectDrive && (
                    <button
                      type="button"
                      onClick={onDisconnectDrive}
                      className="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition cursor-pointer"
                      title="Desconectar Conta Google"
                    >
                      <LogOut className="w-3.5 h-3.5" />
                    </button>
                  )}
                </div>
              )}
            </div>
          )}

          {/* Salvar Button - Admin Only */}
          {userRole !== 'viewer' && onSaveData && (
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
          )}

          {/* Clear Button - Admin Only */}
          {isProcessed && userRole !== 'viewer' && (
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

      {/* Discreet Animated Progress Bar */}
      {isProcessing && (
        <div className="absolute bottom-0 left-0 right-0 h-1 bg-emerald-100 overflow-hidden">
          <div
            className="h-full bg-emerald-600 transition-all duration-300 ease-out shadow-xs shadow-emerald-500/50"
            style={{ width: `${Math.min(100, Math.max(8, processingProgress || 100))}%` }}
          />
        </div>
      )}
    </header>
  );
};
