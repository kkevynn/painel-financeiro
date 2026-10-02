import React from 'react';
import { Users, Wallet, CheckCircle2, AlertCircle, FileCheck2 } from 'lucide-react';
import { DashboardMetrics } from '../types';
import { formatBRL } from '../utils/excelParser';

interface KpiCardsProps {
  metrics: DashboardMetrics;
}

export const KpiCards: React.FC<KpiCardsProps> = ({ metrics }) => {
  const {
    grandTotalValue,
    consolidatedSuppliers,
    totalPpProntas = 0,
    totalLiqObedece = 0,
    totalLiqNaoObedece = 0,
  } = metrics;

  const uniqueFavorecidosCount = consolidatedSuppliers?.length || 0;
  const formattedGrandTotal = formatBRL(grandTotalValue);

  return (
    <div className="space-y-6">
      
      {/* 1. CARD PRINCIPAL (BANNER ESCURO COM RESUMO GERAL) */}
      <div className="bg-gradient-to-r from-[#064e3b] via-[#047857] to-[#059669] rounded-2xl p-6 sm:p-8 text-white shadow-xl shadow-emerald-950/20 relative overflow-hidden">
        {/* Decorative background curves */}
        <div className="absolute -right-10 -top-10 w-64 h-64 bg-white/5 rounded-full blur-2xl pointer-events-none" />
        <div className="absolute right-40 -bottom-20 w-80 h-80 bg-emerald-300/10 rounded-full blur-3xl pointer-events-none" />

        <div className="relative z-10 flex flex-col xl:flex-row xl:items-center justify-between gap-6">
          <div className="min-w-0 flex-1">
            <h2 className="text-2xl md:text-3xl font-black text-white tracking-tight">
              Visão Geral de Credores e Liquidações
            </h2>
          </div>

          {/* Valor Total Destaque Box */}
          <div className="bg-black/20 backdrop-blur-md p-5 rounded-2xl border border-white/10 w-full xl:w-auto shrink-0 min-w-0 flex flex-col justify-center">
            <div className="flex items-center space-x-2 text-emerald-200 mb-1">
              <Wallet className="w-5 h-5 text-emerald-300 shrink-0" />
              <span className="text-xs font-bold uppercase tracking-wider">Valor Total</span>
            </div>
            <div
              className="text-2xl sm:text-3xl xl:text-4xl font-black text-white font-mono tracking-tight"
              title={formattedGrandTotal}
            >
              {formattedGrandTotal}
            </div>
          </div>
        </div>
      </div>

      {/* 2. CARDS DE ESTÁGIOS E MÉTRICAS */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        
        {/* Card 1: Base de PPs */}
        <div className="bg-white rounded-2xl p-5 shadow-xs border border-emerald-200 hover:shadow-md transition-all relative overflow-hidden group min-w-0 flex items-center gap-3.5">
          <div className="w-11 h-11 rounded-xl bg-emerald-100 text-emerald-800 flex items-center justify-center shrink-0">
            <FileCheck2 className="w-5 h-5" />
          </div>
          <div className="min-w-0 flex-1">
            <span className="text-[11px] font-bold text-emerald-800 uppercase tracking-wider block">
              Base de PPs
            </span>
            <div className="text-base sm:text-lg font-black text-emerald-950 tracking-tight font-mono truncate" title={formatBRL(metrics.totalBasePps ?? totalPpProntas)}>
              {formatBRL(metrics.totalBasePps ?? totalPpProntas)}
            </div>
          </div>
        </div>

        {/* Card 2: Liquidação Obedece */}
        <div className="bg-white rounded-2xl p-5 shadow-xs border border-blue-200 hover:shadow-md transition-all relative overflow-hidden group min-w-0 flex items-center gap-3.5">
          <div className="w-11 h-11 rounded-xl bg-blue-100 text-blue-800 flex items-center justify-center shrink-0">
            <CheckCircle2 className="w-5 h-5" />
          </div>
          <div className="min-w-0 flex-1">
            <span className="text-[11px] font-bold text-blue-800 uppercase tracking-wider block">
              Liquidado - Obedece
            </span>
            <div className="text-base sm:text-lg font-black text-blue-950 tracking-tight font-mono truncate" title={formatBRL(totalLiqObedece)}>
              {formatBRL(totalLiqObedece)}
            </div>
          </div>
        </div>

        {/* Card 3: Liquidação Não Obedece */}
        <div className="bg-white rounded-2xl p-5 shadow-xs border border-amber-200 hover:shadow-md transition-all relative overflow-hidden group min-w-0 flex items-center gap-3.5">
          <div className="w-11 h-11 rounded-xl bg-amber-100 text-amber-800 flex items-center justify-center shrink-0">
            <AlertCircle className="w-5 h-5" />
          </div>
          <div className="min-w-0 flex-1">
            <span className="text-[11px] font-bold text-amber-800 uppercase tracking-wider block">
              Liquidado - Não Obedece
            </span>
            <div className="text-base sm:text-lg font-black text-amber-950 tracking-tight font-mono truncate" title={formatBRL(totalLiqNaoObedece)}>
              {formatBRL(totalLiqNaoObedece)}
            </div>
          </div>
        </div>

        {/* Card 4: Favorecidos */}
        <div className="bg-white rounded-2xl p-5 shadow-xs border border-slate-200 hover:shadow-md transition-all relative overflow-hidden group min-w-0 flex items-center gap-3.5">
          <div className="w-11 h-11 rounded-xl bg-slate-100 text-slate-700 flex items-center justify-center shrink-0">
            <Users className="w-5 h-5" />
          </div>
          <div className="min-w-0 flex-1">
            <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider block">
              Favorecidos
            </span>
            <div className="text-xl sm:text-2xl font-black text-slate-900 tracking-tight font-mono">
              {uniqueFavorecidosCount}
            </div>
          </div>
        </div>

      </div>

    </div>
  );
};
