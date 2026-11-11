import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import * as XLSX from 'xlsx';
import { PublicacaoItem } from '../types';
import { formatBRL } from './excelParser';

/**
 * Generates an Excel spreadsheet (.xlsx) with the complete status of Publicação NLs
 */
export function exportPublicacaoToExcel(
  items: PublicacaoItem[],
  titulo: string = 'Acompanhamento_NLs_Publicacao'
) {
  if (!items || items.length === 0) return;

  const rows = items.map((it, idx) => ({
    '#': idx + 1,
    'Dia Publicação': it.diaPublicacao || '-',
    'Status': it.status === 'concluido' ? 'CONCLUÍDO (FEITO)' : 'PENDENTE',
    'Número NL': it.nlNumero,
    'Fornecedor / Credor': it.fornecedor,
    'CNPJ / CPF': it.cnpj,
    'Valor Bruto NL (R$)': it.valorNL,
    'Saiu de Obedece?': it.sumiuObedece ? 'SIM (Saiu)' : 'NÃO (Ainda em Fila)',
    'A Resolver em Obedece?': it.pendenteEmObedece ? 'SIM (A Resolver)' : 'NÃO',
    'PP Equivalente': it.equivalentePP || '-',
    'OB Equivalente': it.equivalenteOB || '-',
    'Valor Líquido PP/OB (R$)': it.valorEquivalente ? it.valorEquivalente : '-',
    'Diferença Tributária (R$)': it.diferencaValor ? it.diferencaValor : 0,
    '% Retenção': it.percentualRetencao ? `${(it.percentualRetencao * 100).toFixed(1).replace('.', ',')}%` : '0,0%',
    'Situação': it.situacaoEquivalente || '-',
    'Diagnóstico / Detalhamento': it.statusDetalhe,
    'Arquivo Origem': it.fileName || '-',
  }));

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(rows);

  // Set column widths
  ws['!cols'] = [
    { wch: 6 },
    { wch: 18 },
    { wch: 16 },
    { wch: 38 },
    { wch: 18 },
    { wch: 18 },
    { wch: 18 },
    { wch: 16 },
    { wch: 16 },
    { wch: 20 },
    { wch: 20 },
    { wch: 12 },
    { wch: 12 },
    { wch: 50 },
    { wch: 25 },
  ];

  XLSX.utils.book_append_sheet(wb, ws, 'Publicação e Pagamentos');
  const dateStr = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `${titulo}_${dateStr}.xlsx`);
}

/**
 * Generates an official SIGEF-style layout PDF report
 * Formatted cleanly with header, institutional banners, KPIs and structured tables.
 */
export function exportPublicacaoToPdf(
  items: PublicacaoItem[],
  titulo: string = 'Relatório de Acompanhamento de Publicação'
) {
  if (!items || items.length === 0) return;

  // Landscape A4 orientation: width 297mm, height 210mm
  const doc = new jsPDF({
    orientation: 'landscape',
    unit: 'mm',
    format: 'a4',
  });

  const pageWidth = doc.internal.pageSize.getWidth();
  const totalNLs = items.length;
  const concluidos = items.filter((i) => i.status === 'concluido');
  const pendentes = items.filter((i) => i.status === 'pendente');
  const valorTotalNL = items.reduce((acc, i) => acc + (i.valorNL || 0), 0);
  const valorConcluido = concluidos.reduce((acc, i) => acc + (i.valorNL || 0), 0);
  const valorPendente = pendentes.reduce((acc, i) => acc + (i.valorNL || 0), 0);

  // --- HEADER INSTITUCIONAL (ESTILO SIGEF / GOVERNO DO ESTADO) ---
  // Top green institutional bar
  doc.setFillColor(6, 29, 21); // #061d15
  doc.rect(0, 0, pageWidth, 20, 'F');

  doc.setTextColor(255, 255, 255);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(11);
  doc.text('GOVERNO DO ESTADO • SISTEMA INTEGRADO DE GESTÃO FINANCEIRA (SIGEF)', 14, 8);

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(8.5);
  doc.setTextColor(167, 243, 208); // emerald-200
  doc.text('FUNDO ESTADUAL DE SAÚDE - FES / SESAP • COORDENADORIA FINANCEIRA', 14, 14);

  const now = new Date();
  const dateFormatted = now.toLocaleDateString('pt-BR') + ' às ' + now.toLocaleTimeString('pt-BR');
  doc.setFontSize(7.5);
  doc.setTextColor(209, 250, 229);
  doc.text(`Emissão: ${dateFormatted}`, pageWidth - 14, 14, { align: 'right' });

  // Subtitle / Document Title Bar
  doc.setFillColor(243, 244, 246);
  doc.rect(0, 20, pageWidth, 12, 'F');
  doc.setDrawColor(229, 231, 235);
  doc.line(0, 32, pageWidth, 32);

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(11);
  doc.setTextColor(15, 23, 42); // slate-900
  doc.text(
    'DEMONSTRATIVO DE LIQUIDAÇÕES PUBLICADAS X BAIXA CRONOLÓGICA E PAGAMENTO',
    14,
    27.5
  );

  doc.setFontSize(8);
  doc.setTextColor(100, 116, 139);
  doc.text('Controle de Saída da Fila Obedece e Localização de PPs/OBs Equivalentes', pageWidth - 14, 27.5, {
    align: 'right',
  });

  // --- SUMMARY KPI METRICS BAR ---
  const startY = 36;
  const boxWidth = (pageWidth - 28 - 9) / 4;
  const boxHeight = 15;

  // Box 1: Total Publicado
  doc.setFillColor(248, 250, 252);
  doc.setDrawColor(203, 213, 225);
  doc.roundedRect(14, startY, boxWidth, boxHeight, 1.5, 1.5, 'FD');
  doc.setFontSize(7);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(100, 116, 139);
  doc.text('TOTAL DE NLS PUBLICADAS', 17, startY + 5);
  doc.setFontSize(10);
  doc.setTextColor(15, 23, 42);
  doc.text(`${totalNLs} (${formatBRL(valorTotalNL)})`, 17, startY + 11);

  // Box 2: Concluídos (Feito)
  doc.setFillColor(236, 253, 245); // emerald-50
  doc.setDrawColor(167, 243, 208); // emerald-200
  doc.roundedRect(14 + boxWidth + 3, startY, boxWidth, boxHeight, 1.5, 1.5, 'FD');
  doc.setFontSize(7);
  doc.setTextColor(4, 120, 87); // emerald-700
  doc.text('CONCLUÍDAS / PAGAS (FEITO)', 14 + boxWidth + 6, startY + 5);
  doc.setFontSize(10);
  doc.setTextColor(6, 95, 70); // emerald-800
  doc.text(`${concluidos.length} (${formatBRL(valorConcluido)})`, 14 + boxWidth + 6, startY + 11);

  // Box 3: Pendentes
  doc.setFillColor(254, 243, 199); // amber-50
  doc.setDrawColor(253, 230, 138); // amber-200
  doc.roundedRect(14 + (boxWidth + 3) * 2, startY, boxWidth, boxHeight, 1.5, 1.5, 'FD');
  doc.setFontSize(7);
  doc.setTextColor(180, 83, 9); // amber-700
  doc.text('PENDENTES / EM FILA', 14 + (boxWidth + 3) * 2 + 3, startY + 5);
  doc.setFontSize(10);
  doc.setTextColor(146, 64, 14); // amber-800
  doc.text(`${pendentes.length} (${formatBRL(valorPendente)})`, 14 + (boxWidth + 3) * 2 + 3, startY + 11);

  // Box 4: Taxa de Efetivação
  const taxaEfetivacao = totalNLs > 0 ? (concluidos.length / totalNLs) * 100 : 0;
  doc.setFillColor(238, 242, 255); // indigo-50
  doc.setDrawColor(199, 210, 254); // indigo-200
  doc.roundedRect(14 + (boxWidth + 3) * 3, startY, boxWidth, boxHeight, 1.5, 1.5, 'FD');
  doc.setFontSize(7);
  doc.setTextColor(67, 56, 202); // indigo-700
  doc.text('TAXA DE EFETIVAÇÃO', 14 + (boxWidth + 3) * 3 + 3, startY + 5);
  doc.setFontSize(10);
  doc.setTextColor(49, 46, 129); // indigo-900
  doc.text(`${taxaEfetivacao.toFixed(1).replace('.', ',')}% resolvidas`, 14 + (boxWidth + 3) * 3 + 3, startY + 11);

  // --- TABELA DE ITENS (AUTOTABLE) ---
  const tableHead = [
    [
      '#',
      'DIA',
      'STATUS',
      'NL',
      'CREDOR / FAVORECIDO',
      'CNPJ / CPF',
      'VALOR NL (R$)',
      'OBEDECE?',
      'PP EQUIV.',
      'OB EQUIV.',
      'VALOR PAGO (R$)',
      'DIF. / TRIB.',
      'SITUAÇÃO',
    ],
  ];

  const tableBody = items.map((it, idx) => {
    const isConcluido = it.status === 'concluido';
    const statusText = isConcluido ? 'CONCLUÍDO' : 'PENDENTE';
    const obedeceText = it.sumiuObedece ? 'Saiu' : 'Em Fila';
    const difText = it.diferencaValor && it.diferencaValor > 0 ? formatBRL(it.diferencaValor) : '-';

    return [
      idx + 1,
      it.diaPublicacao || '-',
      statusText,
      it.nlNumero,
      it.fornecedor.substring(0, 30),
      it.cnpj,
      formatBRL(it.valorNL),
      obedeceText,
      it.equivalentePP || '-',
      it.equivalenteOB || '-',
      it.valorEquivalente ? formatBRL(it.valorEquivalente) : '-',
      difText,
      it.situacaoEquivalente || '-',
    ];
  });

  autoTable(doc, {
    startY: 55,
    head: tableHead,
    body: tableBody,
    theme: 'grid',
    styles: {
      fontSize: 7,
      cellPadding: 1.8,
      overflow: 'linebreak',
      textColor: [30, 41, 59],
      lineColor: [226, 232, 240],
      lineWidth: 0.1,
    },
    headStyles: {
      fillColor: [6, 29, 21], // #061d15 - institutional dark green
      textColor: [255, 255, 255],
      fontStyle: 'bold',
      fontSize: 7,
      halign: 'left',
    },
    columnStyles: {
      0: { cellWidth: 8, halign: 'center' },
      1: { cellWidth: 14, halign: 'center', fontStyle: 'bold' },
      2: { cellWidth: 18, fontStyle: 'bold' },
      3: { cellWidth: 20, fontStyle: 'bold' },
      4: { cellWidth: 50 },
      5: { cellWidth: 26 },
      6: { cellWidth: 24, halign: 'right', fontStyle: 'bold' },
      7: { cellWidth: 15, halign: 'center' },
      8: { cellWidth: 20 },
      9: { cellWidth: 20 },
      10: { cellWidth: 22, halign: 'right' },
      11: { cellWidth: 16, halign: 'right' },
      12: { cellWidth: 16, halign: 'center' },
    },
    didParseCell: (data) => {
      // Color-coding for status (column 2)
      if (data.section === 'body' && data.column.index === 2) {
        if (data.cell.raw === 'CONCLUÍDO') {
          data.cell.styles.textColor = [4, 120, 87]; // emerald-700
          data.cell.styles.fillColor = [236, 253, 245];
        } else {
          data.cell.styles.textColor = [180, 83, 9]; // amber-700
          data.cell.styles.fillColor = [254, 243, 199];
        }
      }
      // Color-coding for obedece status (column 7)
      if (data.section === 'body' && data.column.index === 7) {
        if (data.cell.raw === 'Saiu') {
          data.cell.styles.textColor = [4, 120, 87];
        } else {
          data.cell.styles.textColor = [59, 130, 246]; // blue
        }
      }
    },
    didDrawPage: (data) => {
      // Footer on every page
      const pageCount = doc.internal.pages.length - 1;
      doc.setFontSize(7);
      doc.setTextColor(148, 163, 184);
      doc.text(
        `Documento gerado automaticamente pelo Painel Financeiro FES/SESAP • Página ${data.pageNumber} de ${pageCount}`,
        14,
        205
      );
      doc.text(
        'Controle Interno e Auditoria Financeira',
        pageWidth - 14,
        205,
        { align: 'right' }
      );
    },
  });

  const dateStr = new Date().toISOString().slice(0, 10);
  doc.save(`${titulo}_${dateStr}.pdf`);
}
