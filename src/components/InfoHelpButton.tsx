import React, { useState, useRef, useEffect } from 'react';
import { createPortal } from 'react-dom';
import { HelpCircle, X } from 'lucide-react';

interface InfoHelpButtonProps {
  title?: string;
  content: React.ReactNode;
  variant?: 'dark' | 'light';
  align?: 'left' | 'right' | 'center';
}

export const InfoHelpButton: React.FC<InfoHelpButtonProps> = ({
  title,
  content,
  variant = 'dark',
  align = 'left',
}) => {
  const [isOpen, setIsOpen] = useState(false);
  const buttonRef = useRef<HTMLButtonElement>(null);
  const popoverRef = useRef<HTMLDivElement>(null);
  const [coords, setCoords] = useState<{
    top?: number;
    bottom?: number;
    left: number;
    width: number;
  } | null>(null);

  // Position calculation
  const updatePosition = () => {
    if (!buttonRef.current) return;
    const rect = buttonRef.current.getBoundingClientRect();
    const popoverWidth = Math.min(380, window.innerWidth - 32);

    // If space below the button is less than 240px and space above is greater, place above
    const spaceBelow = window.innerHeight - rect.bottom;
    const spaceAbove = rect.top;
    const placeAbove = spaceBelow < 250 && spaceAbove > 220;

    let left = rect.left;
    if (align === 'right') {
      left = rect.right - popoverWidth;
    } else if (align === 'center') {
      left = rect.left + rect.width / 2 - popoverWidth / 2;
    }

    // Clamp left within viewport margins
    left = Math.max(16, Math.min(left, window.innerWidth - popoverWidth - 16));

    setCoords({
      top: placeAbove ? undefined : rect.bottom + 8,
      bottom: placeAbove ? window.innerHeight - rect.top + 8 : undefined,
      left,
      width: popoverWidth,
    });
  };

  useEffect(() => {
    if (!isOpen) return;
    updatePosition();

    const handleScrollOrResize = () => {
      updatePosition();
    };

    window.addEventListener('scroll', handleScrollOrResize, true);
    window.addEventListener('resize', handleScrollOrResize);

    return () => {
      window.removeEventListener('scroll', handleScrollOrResize, true);
      window.removeEventListener('resize', handleScrollOrResize);
    };
  }, [isOpen, align]);

  // Close on click outside
  useEffect(() => {
    if (!isOpen) return;

    function handleClickOutside(event: MouseEvent) {
      const target = event.target as Node;
      if (
        buttonRef.current &&
        !buttonRef.current.contains(target) &&
        popoverRef.current &&
        !popoverRef.current.contains(target)
      ) {
        setIsOpen(false);
      }
    }

    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [isOpen]);

  // Close on Escape key
  useEffect(() => {
    if (!isOpen) return;

    function handleKeyDown(event: KeyboardEvent) {
      if (event.key === 'Escape') {
        setIsOpen(false);
      }
    }

    document.addEventListener('keydown', handleKeyDown);
    return () => document.removeEventListener('keydown', handleKeyDown);
  }, [isOpen]);

  const buttonStyle =
    variant === 'dark'
      ? 'bg-white/10 hover:bg-white/20 text-slate-300 hover:text-white border-white/20'
      : 'bg-slate-100 hover:bg-slate-200 text-slate-500 hover:text-slate-800 border-slate-300';

  return (
    <>
      <button
        ref={buttonRef}
        type="button"
        onClick={() => setIsOpen((prev) => !prev)}
        className={`w-6 h-6 rounded-full border flex items-center justify-center text-xs font-bold transition cursor-pointer select-none active:scale-95 shrink-0 ${buttonStyle}`}
        title="Clique para ver informações"
        aria-label="Ajuda e informações"
      >
        <span>?</span>
      </button>

      {isOpen &&
        coords &&
        createPortal(
          <div
            ref={popoverRef}
            style={{
              position: 'fixed',
              top: coords.top !== undefined ? `${coords.top}px` : undefined,
              bottom: coords.bottom !== undefined ? `${coords.bottom}px` : undefined,
              left: `${coords.left}px`,
              width: `${coords.width}px`,
              zIndex: 99999,
            }}
            className="p-4 bg-slate-900 text-slate-100 rounded-2xl shadow-2xl border border-slate-700/90 animate-in fade-in-50 zoom-in-95 duration-150"
          >
            <div className="flex items-start justify-between gap-3 border-b border-slate-800 pb-2.5 mb-2.5">
              <div className="flex items-center gap-1.5 text-xs font-bold text-emerald-400">
                <HelpCircle className="w-4 h-4 shrink-0" />
                <span>{title || 'Informações'}</span>
              </div>
              <button
                type="button"
                onClick={() => setIsOpen(false)}
                className="text-slate-400 hover:text-white p-1 rounded-lg hover:bg-slate-800 transition cursor-pointer"
                title="Fechar"
              >
                <X className="w-3.5 h-3.5" />
              </button>
            </div>

            <div className="text-xs text-slate-300 leading-relaxed space-y-2 font-normal">
              {content}
            </div>
          </div>,
          document.body
        )}
    </>
  );
};

