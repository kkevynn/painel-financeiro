import React from 'react';
import { Trophy, Building2 } from 'lucide-react';
import { DashboardMetrics } from '../types';
import { formatBRL } from '../utils/excelParser';

interface TopSuppliersTablesProps {
  metrics: DashboardMetrics;
}

export const TopSuppliersTables: React.FC<TopSuppliersTablesProps> = ({ metrics }) => {
  const topList = metrics.consolidatedSuppliers.slice(0, 5);
  const totalValue = metrics.grandTotalValue || 0;

  return (
    <section className="bg-white rounded-2xl p-6 shadow-xs border border-slate-200 space-y-4">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-slate-200 pb-3">
        <div>
          <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
            <Trophy className="w-5 h-5 text-amber-500" />
            <span>Top 5 Maiores Credores</span>
          </h3>
          <p className="text-xs text-slate-500">
            Fornecedores com maior volume financeiro acumulado.
          </p>
        </div>
      </div>

      {topList.length === 0 ? (
        <div className="py-8 text-center bg-slate-50 rounded-xl border border-dashed border-slate-200">
          <Building2 className="w-8 h-8 text-slate-300 mx-auto mb-2" />
          <p className="text-xs font-semibold text-slate-400">
            Nenhum registro para exibir
          </p>
        </div>
      ) : (
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3">
          {topList.map((item, index) => {
            const rank = index + 1;
            const percentage = totalValue > 0 ? ((item.totalValue / totalValue) * 100).toFixed(1) : '0';

            return (
              <div
                key={`${item.name}-${index}`}
                className="p-3.5 rounded-xl bg-slate-50 border border-slate-200 space-y-2"
              >
                <div className="flex items-center justify-between">
                  <span className="w-6 h-6 rounded-lg text-[10px] font-black bg-emerald-100 text-emerald-900 border border-emerald-300 flex items-center justify-center">
                    #{rank}
                  </span>
                  <span className="text-[10px] font-bold text-slate-500">
                    {percentage}% do total
                  </span>
                </div>

                <div>
                  <h5 className="text-xs font-extrabold text-slate-900 truncate" title={item.name}>
                    {item.name}
                  </h5>
                  <p className="text-[10px] font-mono text-slate-500">{item.cnpj}</p>
                </div>

                <div className="pt-1 border-t border-slate-200 font-mono text-right">
                  <span className="text-xs font-black text-slate-900 block">
                    {formatBRL(item.totalValue)}
                  </span>
                </div>
              </div>
            );
          })}
        </div>
      )}
    </section>
  );
};
