import React, { useState, useRef, useEffect } from 'react';
import { Search, ChevronDown, Check, Sparkles } from 'lucide-react';

export interface SuggestionItem<T = any> {
  id: string;
  title: string;
  subtitle?: string;
  badge?: string;
  badgeColor?: string;
  extraInfo?: string;
  data: T;
}

interface SuggestionInputProps<T = any> {
  value: string;
  onChange: (val: string) => void;
  onSelect?: (item: T) => void;
  suggestions: SuggestionItem<T>[];
  placeholder?: string;
  label?: string;
  required?: boolean;
  className?: string;
  inputClassName?: string;
  disabled?: boolean;
  autoFocus?: boolean;
  helperText?: string;
  allowCustom?: boolean;
  prefixIcon?: React.ReactNode;
}

export function SuggestionInput<T = any>({
  value,
  onChange,
  onSelect,
  suggestions,
  placeholder,
  label,
  required,
  className = '',
  inputClassName = '',
  disabled = false,
  autoFocus = false,
  helperText,
  allowCustom = true,
  prefixIcon,
}: SuggestionInputProps<T>) {
  const [isOpen, setIsOpen] = useState(false);
  const [highlightedIndex, setHighlightedIndex] = useState(-1);
  const containerRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  // Close dropdown on click outside
  useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      if (containerRef.current && !containerRef.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    }
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (!isOpen || suggestions.length === 0) {
      if (e.key === 'ArrowDown' && suggestions.length > 0) {
        setIsOpen(true);
        setHighlightedIndex(0);
        e.preventDefault();
      }
      return;
    }

    if (e.key === 'ArrowDown') {
      e.preventDefault();
      setHighlightedIndex((prev) => (prev < suggestions.length - 1 ? prev + 1 : 0));
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      setHighlightedIndex((prev) => (prev > 0 ? prev - 1 : suggestions.length - 1));
    } else if (e.key === 'Enter') {
      if (highlightedIndex >= 0 && highlightedIndex < suggestions.length) {
        e.preventDefault();
        handleSelectItem(suggestions[highlightedIndex]);
      }
    } else if (e.key === 'Escape') {
      setIsOpen(false);
    }
  };

  const handleSelectItem = (item: SuggestionItem<T>) => {
    onChange(item.title);
    if (onSelect) {
      onSelect(item.data);
    }
    setIsOpen(false);
    setHighlightedIndex(-1);
  };

  return (
    <div ref={containerRef} className={`relative ${className}`}>
      {label && (
        <div className="flex items-center justify-between mb-1">
          <label className="text-[11px] font-bold text-slate-700 flex items-center gap-1">
            <span>{label}</span>
            {required && <span className="text-rose-500">*</span>}
          </label>
          {suggestions.length > 0 && (
            <span className="text-[10px] text-emerald-600 font-semibold flex items-center gap-0.5">
              <Sparkles className="w-2.5 h-2.5" />
              <span>{suggestions.length} sugestões</span>
            </span>
          )}
        </div>
      )}

      <div className="relative flex items-center">
        {prefixIcon && (
          <div className="absolute left-2.5 text-slate-400 pointer-events-none">
            {prefixIcon}
          </div>
        )}
        <input
          ref={inputRef}
          type="text"
          value={value}
          onChange={(e) => {
            onChange(e.target.value);
            setIsOpen(true);
            setHighlightedIndex(-1);
          }}
          onFocus={() => {
            if (suggestions.length > 0) {
              setIsOpen(true);
            }
          }}
          onKeyDown={handleKeyDown}
          placeholder={placeholder}
          required={required}
          disabled={disabled}
          autoFocus={autoFocus}
          className={`w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl font-mono text-xs focus:bg-white focus:outline-hidden focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500 transition ${
            prefixIcon ? 'pl-8' : ''
          } ${inputClassName}`}
        />
        {suggestions.length > 0 && (
          <button
            type="button"
            tabIndex={-1}
            onClick={() => setIsOpen((prev) => !prev)}
            className="absolute right-2 p-1 text-slate-400 hover:text-slate-600 cursor-pointer rounded-md hover:bg-slate-100 transition"
          >
            <ChevronDown className={`w-3.5 h-3.5 transition-transform duration-150 ${isOpen ? 'rotate-180' : ''}`} />
          </button>
        )}
      </div>

      {helperText && (
        <p className="mt-1 text-[10px] text-slate-400 leading-tight">{helperText}</p>
      )}

      {/* Dropdown Suggestions Menu */}
      {isOpen && suggestions.length > 0 && (
        <div className="absolute z-50 left-0 right-0 mt-1 max-h-60 overflow-y-auto bg-white rounded-xl shadow-xl border border-slate-200 divide-y divide-slate-100 animate-in fade-in-50 zoom-in-95 duration-100">
          <div className="p-1.5 text-[10px] font-bold text-slate-400 uppercase tracking-wider bg-slate-50/80 px-3">
            Sugestões encontradas ({suggestions.length})
          </div>
          {suggestions.map((item, idx) => {
            const isHighlighted = idx === highlightedIndex;
            const isSelected = item.title.trim().toUpperCase() === value.trim().toUpperCase();

            return (
              <div
                key={item.id}
                onMouseEnter={() => setHighlightedIndex(idx)}
                onClick={() => handleSelectItem(item)}
                className={`px-3 py-2 text-xs cursor-pointer transition flex items-center justify-between gap-2 ${
                  isHighlighted ? 'bg-emerald-50/80 text-emerald-950' : 'hover:bg-slate-50 text-slate-800'
                } ${isSelected ? 'font-bold bg-emerald-50/50' : ''}`}
              >
                <div className="min-w-0 flex-1">
                  <div className="flex items-center gap-1.5 flex-wrap">
                    <span className="font-mono font-bold text-slate-900">{item.title}</span>
                    {item.badge && (
                      <span
                        className={`px-1.5 py-0.5 rounded text-[10px] font-mono font-bold border ${
                          item.badgeColor || 'bg-slate-100 text-slate-700 border-slate-200'
                        }`}
                      >
                        {item.badge}
                      </span>
                    )}
                  </div>
                  {item.subtitle && (
                    <div className="text-[11px] text-slate-500 truncate mt-0.5" title={item.subtitle}>
                      {item.subtitle}
                    </div>
                  )}
                </div>

                {item.extraInfo && (
                  <div className="text-right shrink-0">
                    <span className="text-[11px] font-mono font-bold text-slate-700">
                      {item.extraInfo}
                    </span>
                  </div>
                )}

                {isSelected && (
                  <Check className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}
