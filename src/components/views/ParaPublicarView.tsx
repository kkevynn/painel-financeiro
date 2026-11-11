import React, { useState, useEffect, useMemo } from 'react';
import {
  Send,
  Trash2,
  FileSpreadsheet,
  FileText,
  Search,
  Filter,
  CheckCircle2,
  AlertCircle,
  Building,
  DollarSign,
  Layers,
  ArrowUpDown,
  Download,
} from 'lucide-react';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { paraPublicarService, ItemParaPublicar } from '../../utils/paraPublicarService';
import { formatBRL } from '../../utils/excelParser';
import { resolveCnpj } from '../../utils/appKnowledgeDirectory';

export const ParaPublicarView: React.FC = () => {
  const [items, setItems] = useState<ItemParaPublicar[]>(paraPublicarService.getItems());
  const [searchTerm, setSearchTerm] = useState('');
  const [filterSource, setFilterSource] = useState<string>('all');
  const [sortBy, setSortBy] = useState<'date' | 'value' | 'supplier' | 'doc'>('date');
  const [sortAsc, setSortAsc] = useState(false);

  useEffect(() => {
    const unsub = paraPublicarService.subscribe((newItems) => {
      setItems([...(newItems || paraPublicarService.getItems())]);
    });
    return () => unsub();
  }, []);

  const totalValue = useMemo(() => {
    return items.reduce((sum, item) => sum + (Number(item.value) || 0), 0);
  }, [items]);

  const uniqueSuppliersCount = useMemo(() => {
    const set = new Set(items.map((i) => i.cnpj || i.supplierName));
    return set.size;
  }, [items]);

  const sourceCounts = useMemo(() => {
    const map: Record<string, number> = {};
    for (const item of items) {
      const src = item.sourceLabel || item.sourceType || 'Outro';
      map[src] = (map[src] || 0) + 1;
    }
    return map;
  }, [items]);

  const filteredItems = useMemo(() => {
    return items
      .filter((item) => {
        if (filterSource !== 'all' && item.sourceType !== filterSource) return false;
        if (!searchTerm.trim()) return true;
        const q = searchTerm.toLowerCase();
        return (
          item.documentNumber.toLowerCase().includes(q) ||
          item.supplierName.toLowerCase().includes(q) ||
          (item.cnpj && item.cnpj.toLowerCase().includes(q)) ||
          (item.fonteRecurso && item.fonteRecurso.toLowerCase().includes(q))
        );
      })
      .sort((a, b) => {
        let cmp = 0;
        if (sortBy === 'value') {
          cmp = (a.value || 0) - (b.value || 0);
        } else if (sortBy === 'supplier') {
          cmp = a.supplierName.localeCompare(b.supplierName);
        } else if (sortBy === 'doc') {
          cmp = a.documentNumber.localeCompare(b.documentNumber);
        } else {
          // date
          cmp = (a.addedAt || 0) - (b.addedAt || 0);
        }
        return sortAsc ? cmp : -cmp;
      });
  }, [items, filterSource, searchTerm, sortBy, sortAsc]);

  const handleExportExcel = () => {
    if (items.length === 0) return;

    const data = items.map((item, idx) => ({
      '#': idx + 1,
      'Número da Liquidação': item.documentNumber.startsWith('NL') ? item.documentNumber : `NL ${item.documentNumber}`,
      'Fornecedor / Razão Social': item.supplierName,
      'CNPJ / CPF': item.cnpj || resolveCnpj(item.supplierName, item.documentNumber) || '-',
      'Origem': item.sourceLabel || item.sourceType,
      'Fonte de Recurso': item.fonteRecurso || '0.6.00',
      'Valor (R$)': item.value,
      'Status Atual': item.status || 'Pendente de Publicação',
      'Data de Seleção': new Date(item.addedAt).toLocaleString('pt-BR'),
    }));

    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Para Publicar');

    const fileName = `Para_Publicar_Liquidacoes_${new Date().toISOString().slice(0, 10)}.xlsx`;
    XLSX.writeFile(wb, fileName);
  };

  const handleExportPdf = () => {
    if (items.length === 0) return;

    const doc = new jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });
    const primaryColor: [number, number, number] = [6, 29, 21];

    // Header banner
    doc.setFillColor(...primaryColor);
    doc.rect(0, 0, 842, 60, 'F');

    doc.setTextColor(255, 255, 255);
    doc.setFontSize(16);
    doc.setFont('helvetica', 'bold');
    doc.text('RELAÇÃO DE LIQUIDAÇÕES PARA PUBLICAÇÃO (DOE)', 30, 36);

    doc.setFontSize(9);
    doc.setFont('helvetica', 'normal');
    doc.text(`Gerado em: ${new Date().toLocaleString('pt-BR')} | Total: ${items.length} itens | Valor: ${formatBRL(totalValue)}`, 500, 36);

    const rows = filteredItems.map((item, idx) => [
      String(idx + 1),
      item.documentNumber.startsWith('NL') ? item.documentNumber : `NL ${item.documentNumber}`,
      item.supplierName,
      item.cnpj || resolveCnpj(item.supplierName, item.documentNumber) || '-',
      item.sourceLabel || item.sourceType,
      item.fonteRecurso || '0.6.00',
      formatBRL(item.value),
    ]);

    autoTable(doc, {
      startY: 75,
      head: [['#', 'Nº LIQUIDAÇÃO', 'FORNECEDOR', 'CNPJ/CPF', 'ORIGEM', 'FONTE', 'VALOR (R$)']],
      body: rows,
      theme: 'grid',
      headStyles: {
        fillColor: primaryColor,
        textColor: [255, 255, 255],
        fontStyle: 'bold',
        fontSize: 9,
      },
      styles: {
        fontSize: 8,
        cellPadding: 4,
      },
      columnStyles: {
        0: { cellWidth: 25, halign: 'center' },
        1: { cellWidth: 90, fontStyle: 'bold' },
        2: { cellWidth: 280 },
        3: { cellWidth: 100 },
        4: { cellWidth: 110 },
        5: { cellWidth: 60, halign: 'center' },
        6: { cellWidth: 90, halign: 'right', fontStyle: 'bold' },
      },
      foot: [
        ['', '', 'TOTAL CONSOLIDADO', '', '', `${items.length} itens`, formatBRL(totalValue)],
      ],
      footStyles: {
        fillColor: [241, 245, 249],
        textColor: [15, 23, 42],
        fontStyle: 'bold',
        fontSize: 9,
      },
    });

    doc.save(`Para_Publicar_Liquidacoes_${new Date().toISOString().slice(0, 10)}.pdf`);
  };

  const handleClearAll = () => {
    if (window.confirm('Deseja realmente limpar toda a lista de liquidações para publicar?')) {
      paraPublicarService.clearAll();
    }
  };

  return (
    <div className="space-y-6 animate-in fade-in-50 duration-200">
      {/* Top Banner / KPIs */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        {/* Total Liquidações */}
        <div className="bg-white rounded-2xl p-5 border border-slate-200 shadow-xs flex items-center justify-between">
          <div>
            <span className="text-xs font-bold uppercase tracking-wider text-slate-500 block">
              Liquidações na Fila
            </span>
            <span className="text-2xl font-black text-slate-900 mt-1 block">
              {items.length}
            </span>
            <span className="text-[11px] text-emerald-600 font-semibold block mt-0.5">
              Prontas para publicação
            </span>
          </div>
          <div className="w-12 h-12 rounded-2xl bg-emerald-50 text-emerald-600 border border-emerald-200 flex items-center justify-center">
            <Send className="w-6 h-6" />
          </div>
        </div>

        {/* Total Bruto */}
        <div className="bg-white rounded-2xl p-5 border border-slate-200 shadow-xs flex items-center justify-between">
          <div>
            <span className="text-xs font-bold uppercase tracking-wider text-slate-500 block">
              Valor Total Bruto
            </span>
            <span className="text-xl sm:text-2xl font-black font-mono text-emerald-700 mt-1 block truncate max-w-[200px]" title={formatBRL(totalValue)}>
              {formatBRL(totalValue)}
            </span>
            <span className="text-[11px] text-slate-500 block mt-0.5">
              Soma das NLs selecionadas
            </span>
          </div>
          <div className="w-12 h-12 rounded-2xl bg-slate-900 text-emerald-400 border border-slate-800 flex items-center justify-center">
            <DollarSign className="w-6 h-6" />
          </div>
        </div>

        {/* Fornecedores Únicos */}
        <div className="bg-white rounded-2xl p-5 border border-slate-200 shadow-xs flex items-center justify-between">
          <div>
            <span className="text-xs font-bold uppercase tracking-wider text-slate-500 block">
              Fornecedores Únicos
            </span>
            <span className="text-2xl font-black text-slate-900 mt-1 block">
              {uniqueSuppliersCount}
            </span>
            <span className="text-[11px] text-slate-500 block mt-0.5">
              Empresas distintas
            </span>
          </div>
          <div className="w-12 h-12 rounded-2xl bg-blue-50 text-blue-600 border border-blue-200 flex items-center justify-center">
            <Building className="w-6 h-6" />
          </div>
        </div>

        {/* Distribuição por Origem */}
        <div className="bg-white rounded-2xl p-5 border border-slate-200 shadow-xs flex flex-col justify-between">
          <div>
            <span className="text-xs font-bold uppercase tracking-wider text-slate-500 block">
              Origem dos Dados
            </span>
            <div className="flex items-center gap-2 mt-2 flex-wrap">
              {Object.entries(sourceCounts).map(([src, count]) => (
                <span
                  key={src}
                  className="px-2 py-0.5 bg-slate-100 text-slate-700 border border-slate-200 rounded-md text-[11px] font-bold"
                >
                  {src}: {count}
                </span>
              ))}
              {Object.keys(sourceCounts).length === 0 && (
                <span className="text-xs text-slate-400 italic">Nenhum item</span>
              )}
            </div>
          </div>
        </div>
      </div>

      {/* Main Content Card */}
      <div className="bg-white rounded-3xl border border-slate-200 shadow-xs overflow-hidden">
        {/* Header toolbar */}
        <div className="p-5 border-b border-slate-200 flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-4">
          <div className="flex items-center gap-2 flex-wrap">
            <div className="relative w-full sm:w-72">
              <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
              <input
                type="text"
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                placeholder="Buscar por NL, fornecedor, CNPJ..."
                className="w-full pl-9 pr-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs text-slate-800 focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500"
              />
            </div>

            <div className="flex items-center gap-1.5">
              <button
                type="button"
                onClick={() => setFilterSource('all')}
                className={`px-3 py-1.5 rounded-xl text-xs font-bold transition cursor-pointer ${
                  filterSource === 'all'
                    ? 'bg-slate-900 text-white shadow-xs'
                    : 'bg-white text-slate-600 hover:text-slate-900 border border-slate-200'
                }`}
              >
                Todas ({items.length})
              </button>
              <button
                type="button"
                onClick={() => setFilterSource('obedece')}
                className={`px-3 py-1.5 rounded-xl text-xs font-bold transition cursor-pointer ${
                  filterSource === 'obedece'
                    ? 'bg-blue-700 text-white shadow-xs'
                    : 'bg-white text-blue-700 hover:bg-blue-50 border border-blue-200'
                }`}
              >
                Obedece
              </button>
              <button
                type="button"
                onClick={() => setFilterSource('nao_obedece')}
                className={`px-3 py-1.5 rounded-xl text-xs font-bold transition cursor-pointer ${
                  filterSource === 'nao_obedece'
                    ? 'bg-amber-700 text-white shadow-xs'
                    : 'bg-white text-amber-700 hover:bg-amber-50 border border-amber-200'
                }`}
              >
                Não Obedece
              </button>
            </div>
          </div>

          {/* Action buttons */}
          <div className="flex items-center gap-2 flex-wrap justify-end">
            <button
              type="button"
              onClick={handleExportPdf}
              disabled={items.length === 0}
              className="px-3.5 py-2 bg-red-600 hover:bg-red-500 disabled:opacity-40 text-white text-xs font-bold rounded-xl transition flex items-center gap-1.5 cursor-pointer shadow-xs active:scale-95"
              title="Exportar relação em PDF"
            >
              <FileText className="w-3.5 h-3.5" />
              <span>Exportar PDF</span>
            </button>
            <button
              type="button"
              onClick={handleExportExcel}
              disabled={items.length === 0}
              className="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-500 disabled:opacity-40 text-white text-xs font-bold rounded-xl transition flex items-center gap-1.5 cursor-pointer shadow-xs active:scale-95"
              title="Exportar relação em Excel"
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              <span>Exportar Excel</span>
            </button>
            {items.length > 0 && (
              <button
                type="button"
                onClick={handleClearAll}
                className="px-3.5 py-2 bg-slate-100 hover:bg-red-50 text-slate-600 hover:text-red-600 text-xs font-bold rounded-xl transition flex items-center gap-1.5 border border-slate-200 hover:border-red-200 cursor-pointer active:scale-95"
                title="Limpar todas as liquidações da fila"
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>Limpar Fila</span>
              </button>
            )}
          </div>
        </div>

        {/* Table or Empty State */}
        {filteredItems.length === 0 ? (
          <div className="p-12 text-center space-y-3">
            <div className="w-14 h-14 rounded-2xl bg-slate-100 text-slate-400 flex items-center justify-center mx-auto">
              <Send className="w-7 h-7" />
            </div>
            <h3 className="text-base font-bold text-slate-800">
              Nenhuma liquidação selecionada para publicação
            </h3>
            <p className="text-xs text-slate-500 max-w-md mx-auto">
              Para adicionar liquidações a esta fila, vá até a página inicial, clique no botão{' '}
              <strong className="text-slate-700">"Extrato"</strong> do fornecedor correspondente e
              marque a opção <strong className="text-slate-700">"Publicar"</strong> na coluna da tabela.
            </p>
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="bg-slate-100/80 text-slate-700 font-extrabold uppercase tracking-wider border-b border-slate-200 text-[11px]">
                  <th className="py-3 px-4 text-center w-12">#</th>
                  <th className="py-3 px-4 w-44 font-mono">Nº LIQUIDAÇÃO</th>
                  <th className="py-3 px-4 min-w-[240px]">FORNECEDOR</th>
                  <th className="py-3 px-4 w-36">ORIGEM</th>
                  <th className="py-3 px-4 w-28 text-center">FONTE</th>
                  <th className="py-3 px-4 text-right w-36 font-mono">VALOR (R$)</th>
                  <th className="py-3 px-4 text-center w-20">AÇÃO</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 font-medium">
                {filteredItems.map((item, idx) => (
                  <tr key={item.id} className="hover:bg-slate-50 transition-colors">
                    <td className="py-3 px-4 text-center font-mono text-[11px] text-slate-400 border-r border-slate-100">
                      {idx + 1}
                    </td>
                    <td className="py-3 px-4 font-mono font-bold text-slate-900 border-r border-slate-100">
                      <span className="bg-slate-100 px-2 py-0.5 rounded border border-slate-200 text-[11px] font-bold text-slate-900">
                        {item.documentNumber.startsWith('NL') ? item.documentNumber : `NL ${item.documentNumber}`}
                      </span>
                    </td>
                    <td className="py-3 px-4 border-r border-slate-100">
                      <div className="flex flex-col">
                        <span className="font-bold text-slate-900 text-xs">
                          {item.supplierName}
                        </span>
                        <span className="text-[10px] text-slate-500 font-mono">
                          CNPJ: {item.cnpj || resolveCnpj(item.supplierName, item.documentNumber) || 'Não informado'}
                        </span>
                      </div>
                    </td>
                    <td className="py-3 px-4 border-r border-slate-100">
                      <span className="inline-block px-2 py-0.5 rounded-md text-[10px] font-bold bg-blue-50 text-blue-800 border border-blue-200">
                        {item.sourceLabel || item.sourceType}
                      </span>
                    </td>
                    <td className="py-3 px-4 text-center font-mono font-bold text-slate-700 border-r border-slate-100">
                      {item.fonteRecurso || '0.6.00'}
                    </td>
                    <td className="py-3 px-4 text-right font-mono font-bold text-slate-900 border-r border-slate-100">
                      {formatBRL(item.value)}
                    </td>
                    <td className="py-3 px-4 text-center">
                      <button
                        type="button"
                        onClick={() => paraPublicarService.removeItem(item.id)}
                        className="p-1.5 text-slate-400 hover:text-red-600 hover:bg-red-50 rounded-lg transition cursor-pointer"
                        title="Remover da fila"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
              <tfoot>
                <tr className="bg-slate-100/90 font-black text-slate-900 border-t border-slate-200 text-xs">
                  <td colSpan={5} className="py-3 px-4 text-right uppercase tracking-wider">
                    Total Consolidado ({filteredItems.length} itens):
                  </td>
                  <td className="py-3 px-4 text-right font-mono text-emerald-800 text-sm">
                    {formatBRL(
                      filteredItems.reduce((sum, item) => sum + (Number(item.value) || 0), 0)
                    )}
                  </td>
                  <td></td>
                </tr>
              </tfoot>
            </table>
          </div>
        )}
      </div>
    </div>
  );
};
