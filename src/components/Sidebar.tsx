import React, { useEffect } from 'react';
import { ShieldCheck, X } from 'lucide-react';

interface SidebarProps {
  isOpen: boolean;
  onClose: () => void;
  activeTab: string;
  setActiveTab: (tab: string) => void;
}

export const Sidebar: React.FC<SidebarProps> = ({
  isOpen,
  onClose,
  activeTab,
  setActiveTab,
}) => {
  // Close on Escape key press
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape' && isOpen) {
        onClose();
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [isOpen, onClose]);

  const menuItems = [
    { id: 'painel', label: 'Painel' },
    { id: 'pagamentos_fonte', label: 'Pagamentos por Fonte' },
    { id: 'fornecedores_mes', label: 'Fornecedores Atendidos' },
    { id: 'quebra_ordem', label: 'Quebra de Ordem' },
    { id: 'recorrencia', label: 'Fornecedores Recorrentes' },
    { id: 'categorias', label: 'Categorias' },
    { id: 'importacoes', label: 'Importações' },
  ];

  return (
    <>
      {/* Backdrop overlay */}
      <div
        onClick={onClose}
        className={`fixed inset-0 bg-black/60 backdrop-blur-xs z-40 transition-opacity duration-300 ${
          isOpen ? 'opacity-100 pointer-events-auto' : 'opacity-0 pointer-events-none'
        }`}
        aria-hidden="true"
      />

      {/* Retractable Off-Canvas Sidebar Drawer */}
      <aside
        className={`fixed top-0 left-0 bottom-0 w-72 max-w-[85vw] bg-[#061d15] text-emerald-100 flex flex-col justify-between h-screen overflow-y-auto border-r border-emerald-900/80 shadow-2xl z-50 transform transition-transform duration-300 ease-in-out ${
          isOpen ? 'translate-x-0' : '-translate-x-full'
        }`}
      >
        <div>
          {/* Brand Header & Close Button */}
          <div className="p-5 border-b border-emerald-900/80 flex items-center justify-between bg-[#04140e]">
            <div className="flex items-center space-x-3">
              <div className="w-9 h-9 rounded-xl bg-gradient-to-tr from-emerald-500 to-teal-400 p-0.5 shadow-lg shadow-emerald-500/20 flex items-center justify-center">
                <div className="w-full h-full bg-[#061d15] rounded-[10px] flex items-center justify-center">
                  <ShieldCheck className="w-5 h-5 text-emerald-400" />
                </div>
              </div>
              <div>
                <h1 className="text-sm font-extrabold text-white tracking-wide font-sans">
                  Painel Financeiro
                </h1>
                <span className="text-[11px] font-semibold text-emerald-400 block -mt-0.5">
                  FES / SESAP
                </span>
              </div>
            </div>

            {/* Close Button */}
            <button
              type="button"
              onClick={onClose}
              className="p-1.5 text-emerald-400/80 hover:text-white hover:bg-emerald-900/60 rounded-xl transition cursor-pointer active:scale-95"
              title="Fechar Menu"
              aria-label="Fechar Menu"
            >
              <X className="w-5 h-5" />
            </button>
          </div>

          {/* Navigation Section (Clean Text Only - No Icons) */}
          <div className="p-4 space-y-1.5">
            <nav className="space-y-1.5">
              {menuItems.map((item) => {
                const isActive = activeTab === item.id;

                return (
                  <button
                    key={item.id}
                    type="button"
                    onClick={() => {
                      setActiveTab(item.id);
                      onClose();
                    }}
                    className={`w-full text-left px-4 py-3 rounded-xl text-sm font-semibold transition-all duration-150 cursor-pointer active:scale-98 ${
                      isActive
                        ? 'bg-emerald-600 text-white shadow-md shadow-emerald-600/30 font-bold'
                        : 'text-emerald-200/80 hover:text-white hover:bg-emerald-900/50'
                    }`}
                  >
                    <span>{item.label}</span>
                  </button>
                );
              })}
            </nav>
          </div>
        </div>

        {/* Footer Info */}
        <div className="p-4 border-t border-emerald-900/80 bg-[#04140e]">
          <div className="flex items-center justify-between text-[11px] text-emerald-400/90 font-medium">
            <span>FES / SESAP</span>
            <span>Governo do RN</span>
          </div>
        </div>
      </aside>
    </>
  );
};
