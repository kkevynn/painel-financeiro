import React from 'react';
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  BarElement,
  Title,
  Tooltip,
  Legend,
} from 'chart.js';
import { Bar } from 'react-chartjs-2';
import { TrendingUp, Info, Building } from 'lucide-react';
import { DashboardMetrics } from '../types';
import { formatBRL } from '../utils/excelParser';

ChartJS.register(
  CategoryScale,
  LinearScale,
  BarElement,
  Title,
  Tooltip,
  Legend
);

interface ChartsSectionProps {
  metrics: DashboardMetrics;
}

export const ChartsSection: React.FC<ChartsSectionProps> = ({ metrics }) => {
  const { consolidatedSuppliers, isProcessed } = metrics;

  const topSuppliers = (consolidatedSuppliers || []).slice(0, 8);

  const labels = topSuppliers.map((s) => {
    if (s.name.length > 22) {
      return s.name.substring(0, 20) + '...';
    }
    return s.name;
  });

  const values = topSuppliers.map((s) => s.totalValue);

  const chartData = {
    labels,
    datasets: [
      {
        label: 'Valor Total (R$)',
        data: values,
        backgroundColor: 'rgba(5, 150, 105, 0.85)',
        borderColor: '#059669',
        borderWidth: 1.5,
        borderRadius: 8,
        barThickness: 28,
      },
    ],
  };

  const chartOptions = {
    responsive: true,
    maintainAspectRatio: false,
    indexAxis: 'y' as const, // Horizontal bars for clean reading of company names
    plugins: {
      legend: {
        display: false,
      },
      tooltip: {
        backgroundColor: '#0f172a',
        padding: 12,
        titleFont: { family: 'sans-serif', size: 12, weight: 700 },
        bodyFont: { family: 'sans-serif', size: 12 },
        callbacks: {
          title: (items: any) => {
            if (!items.length) return '';
            const idx = items[0].dataIndex;
            return topSuppliers[idx]?.name || items[0].label;
          },
          label: (context: any) => {
            return ` Valor Total: ${formatBRL(context.raw)}`;
          },
        },
      },
    },
    scales: {
      x: {
        beginAtZero: true,
        ticks: {
          font: { family: 'sans-serif', size: 10 },
          callback: (val: any) => {
            if (val >= 1000000) return `R$ ${(val / 1000000).toFixed(1)}M`;
            if (val >= 1000) return `R$ ${(val / 1000).toFixed(0)}k`;
            return `R$ ${val}`;
          },
        },
        grid: { color: '#f1f5f9' },
      },
      y: {
        ticks: { font: { family: 'sans-serif', size: 11, weight: 600 } },
        grid: { display: false },
      },
    },
  };

  return (
    <section className="bg-white rounded-2xl p-6 shadow-xs border border-slate-200/90 space-y-4">
      <div className="flex items-center justify-between">
        <div>
          <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
            <TrendingUp className="w-5 h-5 text-emerald-600" />
            <span>Maiores Credores por Valor Total</span>
          </h3>
          <p className="text-xs text-slate-500 mt-0.5">
            Distribuição dos principais fornecedores e seus respectivos montantes acumulados em Reais (R$).
          </p>
        </div>

        <span className="text-[10px] font-bold text-emerald-800 bg-emerald-50 px-2.5 py-1 rounded-full border border-emerald-200 hidden sm:inline-block">
          Top Credores
        </span>
      </div>

      <div className="h-80 relative flex items-center justify-center py-2">
        {isProcessed && topSuppliers.length > 0 ? (
          <Bar data={chartData} options={chartOptions} />
        ) : (
          <div className="text-center space-y-2">
            <Info className="w-8 h-8 text-slate-300 mx-auto" />
            <p className="text-xs text-slate-400 font-medium">
              Carregue e processe planilhas para renderizar o gráfico comparativo de credores.
            </p>
          </div>
        )}
      </div>
    </section>
  );
};
