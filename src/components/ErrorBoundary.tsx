import React, { Component, ErrorInfo, ReactNode } from 'react';
import { AlertTriangle, RefreshCw, Trash2 } from 'lucide-react';
import { clearAppState } from '../utils/storageService';
import { clearStoredOrdensBancarias } from '../utils/ordemBancariaParser';

interface Props {
  children: ReactNode;
  fallbackTitle?: string;
}

interface State {
  hasError: boolean;
  error: Error | null;
  errorInfo: ErrorInfo | null;
}

export class ErrorBoundary extends Component<Props, State> {
  public state: State = {
    hasError: false,
    error: null,
    errorInfo: null,
  };

  public static getDerivedStateFromError(error: Error): State {
    return { hasError: true, error, errorInfo: null };
  }

  public componentDidCatch(error: Error, errorInfo: ErrorInfo) {
    console.error('Uncaught error caught by ErrorBoundary:', error, errorInfo);
    this.setState({ errorInfo });
  }

  private handleReset = () => {
    this.setState({ hasError: false, error: null, errorInfo: null });
  };

  private handleFullReload = () => {
    this.setState({ hasError: false, error: null, errorInfo: null });
    window.location.reload();
  };

  private handleClearAndReset = async () => {
    try {
      await clearAppState();
      clearStoredOrdensBancarias();
      localStorage.clear();
      sessionStorage.clear();
    } catch (e) {
      console.error('Erro ao limpar dados locais:', e);
    }
    window.location.reload();
  };

  public render() {
    if (this.state.hasError) {
      const isSubmodule = !!this.props.fallbackTitle;

      if (isSubmodule) {
        return (
          <div className="py-12 px-4 flex items-center justify-center antialiased">
            <div className="max-w-lg w-full bg-slate-900 text-white border border-rose-500/40 rounded-3xl p-6 shadow-xl space-y-4">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-xl bg-rose-500/20 border border-rose-500/30 flex items-center justify-center text-rose-400 shrink-0">
                  <AlertTriangle className="w-5 h-5" />
                </div>
                <div>
                  <h2 className="text-sm font-bold text-white">
                    {this.props.fallbackTitle}
                  </h2>
                  <p className="text-[11px] text-slate-400">
                    Ocorreu uma instabilidade neste módulo.
                  </p>
                </div>
              </div>

              {this.state.error && (
                <div className="bg-slate-950/80 rounded-xl p-3 border border-slate-700/60 font-mono text-[11px] text-rose-300 overflow-x-auto max-h-32">
                  {this.state.error.message || String(this.state.error)}
                </div>
              )}

              <div className="flex items-center gap-2 pt-1">
                <button
                  type="button"
                  onClick={this.handleReset}
                  className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-bold text-xs rounded-xl transition flex items-center justify-center gap-1.5 cursor-pointer"
                >
                  <RefreshCw className="w-3.5 h-3.5" />
                  <span>Tentar Novamente</span>
                </button>
                <button
                  type="button"
                  onClick={this.handleFullReload}
                  className="px-3 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 font-bold text-xs rounded-xl transition flex items-center justify-center gap-1.5 cursor-pointer border border-slate-700"
                >
                  <span>Recarregar Página</span>
                </button>
              </div>
            </div>
          </div>
        );
      }

      return (
        <div className="min-h-screen bg-slate-900 text-white flex items-center justify-center p-6 antialiased font-sans">
          <div className="max-w-xl w-full bg-slate-800/90 border border-rose-500/40 rounded-3xl p-8 shadow-2xl backdrop-blur-md space-y-6">
            <div className="flex items-center gap-4">
              <div className="w-14 h-14 rounded-2xl bg-rose-500/20 border border-rose-500/30 flex items-center justify-center text-rose-400 shrink-0">
                <AlertTriangle className="w-7 h-7" />
              </div>
              <div>
                <h1 className="text-xl font-black text-white">
                  Instabilidade Temporária Detectada
                </h1>
                <p className="text-xs text-slate-400 mt-0.5">
                  O painel identificou uma exceção pontual e manteve os seus dados seguros.
                </p>
              </div>
            </div>

            {this.state.error && (
              <div className="bg-slate-950/80 rounded-2xl p-4 border border-slate-700/60 font-mono text-xs text-rose-300 overflow-x-auto max-h-48 leading-relaxed">
                <strong>Detalhe:</strong> {this.state.error.message || String(this.state.error)}
              </div>
            )}

            <div className="flex flex-col sm:flex-row items-center gap-3 pt-2">
              <button
                type="button"
                onClick={this.handleReset}
                className="w-full sm:w-auto flex-1 px-5 py-3 bg-emerald-600 hover:bg-emerald-500 text-white font-bold text-xs rounded-xl transition flex items-center justify-center gap-2 cursor-pointer shadow-lg shadow-emerald-600/20 active:scale-98"
              >
                <RefreshCw className="w-4 h-4" />
                <span>Restaurar Painel</span>
              </button>

              <button
                type="button"
                onClick={this.handleFullReload}
                className="w-full sm:w-auto px-5 py-3 bg-slate-800 hover:bg-slate-700 text-slate-300 font-bold text-xs rounded-xl transition flex items-center justify-center gap-2 cursor-pointer border border-slate-700 active:scale-98"
              >
                <span>Recarregar Página</span>
              </button>

              <button
                type="button"
                onClick={this.handleClearAndReset}
                className="w-full sm:w-auto px-4 py-3 bg-slate-800/80 hover:bg-rose-950/40 text-slate-400 hover:text-rose-300 font-bold text-xs rounded-xl transition flex items-center justify-center gap-1.5 cursor-pointer border border-slate-700/80 active:scale-98"
              >
                <Trash2 className="w-3.5 h-3.5 text-rose-400" />
                <span>Limpar e Reiniciar</span>
              </button>
            </div>
          </div>
        </div>
      );
    }

    return this.props.children;
  }
}
