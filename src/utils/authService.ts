/**
 * Authentication and Role-Based Access Control (RBAC) Service
 * Manages user sessions, roles (admin vs viewer), and permissions for Painel Financeiro - FES/SESAP
 */

export type UserRole = 'admin' | 'viewer';

export interface AuthSession {
  username: string;
  displayName: string;
  role: UserRole;
  loginAt: string;
}

const AUTH_STORAGE_KEY = 'painel_fes_auth_session_v1';

export const authService = {
  getCurrentSession(): AuthSession | null {
    try {
      const stored = localStorage.getItem(AUTH_STORAGE_KEY);
      if (!stored) return null;
      const parsed = JSON.parse(stored) as AuthSession;
      if (parsed && (parsed.role === 'admin' || parsed.role === 'viewer')) {
        return parsed;
      }
      return null;
    } catch {
      return null;
    }
  },

  login(usernameInput: string, passwordInput: string): { success: boolean; session?: AuthSession; error?: string } {
    const username = (usernameInput || '').trim();
    const password = (passwordInput || '').trim();

    // Administrador: "kevyn" | "Casa134*"
    if (username.toLowerCase() === 'kevyn') {
      if (password === 'Casa134*') {
        const session: AuthSession = {
          username: 'kevyn',
          displayName: 'Administrador',
          role: 'admin',
          loginAt: new Date().toISOString(),
        };
        localStorage.setItem(AUTH_STORAGE_KEY, JSON.stringify(session));
        return { success: true, session };
      }
      return { success: false, error: 'Senha incorreta para o perfil Administrador.' };
    }

    // Consulta: "Usuário" ou "usuario" | "fes2026@"
    if (
      username.toLowerCase() === 'usuário' ||
      username.toLowerCase() === 'usuario' ||
      username.toLowerCase() === 'consulta'
    ) {
      if (password === 'fes2026@') {
        const session: AuthSession = {
          username: 'Usuário',
          displayName: 'Consulta',
          role: 'viewer',
          loginAt: new Date().toISOString(),
        };
        localStorage.setItem(AUTH_STORAGE_KEY, JSON.stringify(session));
        return { success: true, session };
      }
      return { success: false, error: 'Senha incorreta para o perfil Consulta.' };
    }

    return { success: false, error: 'Usuário não encontrado. Verifique suas credenciais.' };
  },

  logout(): void {
    try {
      localStorage.removeItem(AUTH_STORAGE_KEY);
    } catch (e) {
      console.warn('Erro ao limpar sessão:', e);
    }
  },

  isAdmin(): boolean {
    const session = this.getCurrentSession();
    return session?.role === 'admin';
  },
};

export const getCurrentSession = () => authService.getCurrentSession();
export const isAdmin = () => authService.isAdmin();
export const logout = () => authService.logout();
