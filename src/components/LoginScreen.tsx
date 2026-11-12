import React, { useState } from 'react';
import { ShieldCheck, Eye, EyeOff, Lock, User, LogIn, AlertCircle } from 'lucide-react';
import { authService, AuthSession } from '../utils/authService';

interface LoginScreenProps {
  onLoginSuccess: (session: AuthSession) => void;
}

export const LoginScreen: React.FC<LoginScreenProps> = ({ onLoginSuccess }) => {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [isLoading, setIsLoading] = useState(false);

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    setErrorMessage(null);

    if (!username.trim()) {
      setErrorMessage('Por favor, informe o nome de usuário.');
      return;
    }
    if (!password) {
      setErrorMessage('Por favor, informe a senha.');
      return;
    }

    setIsLoading(true);
    setTimeout(() => {
      const result = authService.login(username, password);
      setIsLoading(false);

      if (result.success && result.session) {
        onLoginSuccess(result.session);
      } else {
        setErrorMessage(result.error || 'Credenciais inválidas.');
      }
    }, 150);
  };

  return (
    <div className="min-h-screen w-full bg-[#04140e] flex items-center justify-center p-4 sm:p-6 relative overflow-hidden font-sans">
      {/* Background Decorative Gradients */}
      <div className="absolute -top-32 -left-32 w-96 h-96 bg-emerald-600/10 rounded-full blur-3xl pointer-events-none" />
      <div className="absolute -bottom-32 -right-32 w-96 h-96 bg-teal-500/10 rounded-full blur-3xl pointer-events-none" />

      {/* Main Login Card */}
      <div className="w-full max-w-md bg-[#061d15] border border-emerald-800/60 rounded-3xl p-8 sm:p-10 shadow-2xl relative z-10 text-white">
        {/* Brand Header */}
        <div className="flex flex-col items-center text-center mb-8">
          <div className="w-16 h-16 rounded-2xl bg-gradient-to-tr from-emerald-500 to-teal-400 p-0.5 shadow-xl shadow-emerald-500/20 flex items-center justify-center mb-4">
            <div className="w-full h-full bg-[#061d15] rounded-[14px] flex items-center justify-center">
              <ShieldCheck className="w-8 h-8 text-emerald-400" />
            </div>
          </div>
          <h1 className="text-xl sm:text-2xl font-black text-white tracking-tight">
            Painel Financeiro - FES
          </h1>
          <p className="text-xs text-emerald-300/80 font-medium mt-1">
            SESAP - Secretaria de Estado da Saúde Pública
          </p>
          <span className="inline-block mt-3 px-3 py-1 rounded-full text-[11px] font-semibold bg-emerald-950/80 text-emerald-300 border border-emerald-800/60">
            Acesso Restrito ao Sistema
          </span>
        </div>

        {/* Error Feedback */}
        {errorMessage && (
          <div className="mb-6 p-3.5 bg-rose-950/60 border border-rose-800/70 rounded-xl text-rose-200 text-xs flex items-center gap-2.5 animate-in fade-in duration-150">
            <AlertCircle className="w-4 h-4 shrink-0 text-rose-400" />
            <span>{errorMessage}</span>
          </div>
        )}

        {/* Login Form */}
        <form onSubmit={handleSubmit} className="space-y-5">
          {/* Usuário Input */}
          <div>
            <label className="block text-xs font-bold text-emerald-200 mb-1.5 uppercase tracking-wider">
              Usuário
            </label>
            <div className="relative">
              <div className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none text-emerald-400/60">
                <User className="w-4 h-4" />
              </div>
              <input
                type="text"
                value={username}
                onChange={(e) => setUsername(e.target.value)}
                placeholder="Ex: kevyn ou Usuário"
                autoComplete="username"
                className="w-full pl-10 pr-4 py-2.5 bg-[#03100b] border border-emerald-900/80 rounded-xl text-sm text-white placeholder-emerald-600/50 focus:outline-hidden focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500 transition"
              />
            </div>
          </div>

          {/* Senha Input */}
          <div>
            <label className="block text-xs font-bold text-emerald-200 mb-1.5 uppercase tracking-wider">
              Senha
            </label>
            <div className="relative">
              <div className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none text-emerald-400/60">
                <Lock className="w-4 h-4" />
              </div>
              <input
                type={showPassword ? 'text' : 'password'}
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                placeholder="Digite sua senha"
                autoComplete="current-password"
                className="w-full pl-10 pr-11 py-2.5 bg-[#03100b] border border-emerald-900/80 rounded-xl text-sm text-white placeholder-emerald-600/50 focus:outline-hidden focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500 transition"
              />
              <button
                type="button"
                onClick={() => setShowPassword(!showPassword)}
                className="absolute inset-y-0 right-0 pr-3.5 flex items-center text-emerald-400/60 hover:text-emerald-300 transition cursor-pointer"
                title={showPassword ? 'Ocultar senha' : 'Exibir senha'}
              >
                {showPassword ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
              </button>
            </div>
          </div>

          {/* Submit Button */}
          <button
            type="submit"
            disabled={isLoading}
            className="w-full mt-2 py-3 px-4 bg-emerald-600 hover:bg-emerald-500 active:scale-98 text-white text-sm font-bold rounded-xl transition shadow-lg shadow-emerald-950/40 flex items-center justify-center gap-2 cursor-pointer disabled:opacity-60"
          >
            <LogIn className="w-4 h-4" />
            <span>{isLoading ? 'Autenticando...' : 'Entrar no Painel'}</span>
          </button>
        </form>

        {/* Footer info */}
        <div className="mt-8 pt-6 border-t border-emerald-900/60 text-center">
          <p className="text-[11px] text-emerald-400/60 leading-relaxed">
            Acesso autorizado para Administração e Consulta da Diretoria.
          </p>
        </div>
      </div>
    </div>
  );
};
