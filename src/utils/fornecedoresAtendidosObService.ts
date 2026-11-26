import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { DashboardMetrics, FileData } from '../types';
import {
  formatBRL,
  parseCurrencyValue,
  parseFavorecidoField,
  isSituacaoAO,
  isSituacaoConclusivaPagamento,
  isSituacaoRejeitadaOuCancelada,
} from './excelParser';
import {
  OrdemBancariaItem,
  isValidFonteRecursoCode,
  findFonteRecursoInRow,
  parseDateToMonthInfo,
  isObNumber,
} from './ordemBancariaParser';
import {
  isSituacaoConfirmadaBanco,
  isSituacaoRejeitada,
  isSituacaoRejeitadaBanco,
  isSituacaoAOOuVazia,
  isSituacaoNaoPaga,
} from './toastAlertService';
import { SITUACAO_SIGLAS } from './liberacaoParser';

export type ObStatusCategoria = 'confirmada' | 'pendente' | 'rejeitada';

export interface AttendedSupplierInput {
  name: string;
  cnpj: string;
  isRec?: boolean;
  observacao?: string;
  dataAtendimento?: string | null;
}

export interface PanelOBRecord {
  ob: string;
  obClean: string;
  pp: string;
  favorecido: string;
  cnpj: string;
  valor: number;
  data: string;
  dataIso: string;
  fonteRecurso: string;
  situacaoSigla: string;
  situacaoLabel: string;
  statusCategoria: ObStatusCategoria;
  origem: string;
  isOfficialStatus: boolean;
}

export interface AttendedSupplierOBRecord extends PanelOBRecord {
  fornecedorAtendido: string;
  cnpjAtendido: string;
  classificacao: 'Recorrente' | 'Comum';
  observacaoAtendimento: string;
}

export interface AttendedSupplierOBGroup {
  fornecedor: string;
  cnpj: string;
  classificacao: 'Recorrente' | 'Comum';
  observacaoAtendimento: string;
  obs: AttendedSupplierOBRecord[];
  totalValor: number;
}

function normalizeName(name: string): string {
  return (name || '')
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeCnpjDigits(cnpj?: string): string {
  const digits = (cnpj || '').replace(/\D/g, '');
  return digits.length === 14 ? digits : '';
}

function cleanDocKey(val?: any): string {
  return String(val || '')
    .trim()
    .toUpperCase()
    .replace(/[\s\.\-\/]/g, '');
}

function stripObLeadingZeros(cleanOb: string): string {
  const m = cleanOb.match(/^(\d{4})?OB0*(\d+)$/i);
  if (!m) return cleanOb;
  return `${m[1] || ''}OB${m[2]}`.toUpperCase();
}

export function extractCanonicalObCode(raw: any): string {
  if (raw === null || raw === undefined) return '';
  const str = String(raw).trim().toUpperCase();
  if (!str || str === '-' || str === '0') return '';

  const fullMatch = str.match(/\b(\d{4})[\s\.\-\/]?OB[\s\.\-\/]?(\d{3,})\b/i);
  if (fullMatch) {
    return `${fullMatch[1]}OB${fullMatch[2]}`.toUpperCase();
  }

  const shortMatch = str.match(/\bOB[\s\.\-\/_]?(\d{3,})\b/i);
  if (shortMatch) {
    return `OB${shortMatch[1]}`.toUpperCase();
  }

  if (str.includes('OB') && isObNumber(str)) {
    return cleanDocKey(str);
  }

  return '';
}

function classifySituacao(rawSit: string): {
  categoria: ObStatusCategoria;
  sigla: string;
  label: string;
  isTerminal: boolean;
} {
  const sigla = (rawSit || '').trim().toUpperCase();
  if (!sigla) {
    return {
      categoria: 'pendente',
      sigla: '',
      label: 'Pendente (Sem situação)',
      isTerminal: false,
    };
  }

  const conf = isSituacaoConfirmadaBanco(sigla);
  if (conf.isMatch || isSituacaoConclusivaPagamento(sigla)) {
    const desc = SITUACAO_SIGLAS[sigla] || SITUACAO_SIGLAS[conf.normalized] || 'Confirmada no Banco';
    const label = sigla.length <= 6 && desc !== sigla ? `${sigla} - ${desc}` : sigla;
    return {
      categoria: 'confirmada',
      sigla,
      label,
      isTerminal: true,
    };
  }

  const rej = isSituacaoRejeitadaBanco(sigla);
  if (rej.isMatch || isSituacaoRejeitada(sigla) || isSituacaoRejeitadaOuCancelada(sigla)) {
    const desc = SITUACAO_SIGLAS[sigla] || SITUACAO_SIGLAS[rej.normalized] || 'Rejeitada / Cancelada';
    const label = sigla.length <= 6 && desc !== sigla ? `${sigla} - ${desc}` : sigla;
    return {
      categoria: 'rejeitada',
      sigla,
      label,
      isTerminal: true,
    };
  }

  const desc = SITUACAO_SIGLAS[sigla];
  const label = desc && desc !== sigla ? `${sigla} - ${desc}` : sigla;
  return {
    categoria: 'pendente',
    sigla,
    label,
    isTerminal: false,
  };
}

function matchesSupplierTarget(
  targetCnpjDigits: string,
  targetNormName: string,
  candidateCnpj?: string,
  candidateName?: string
): boolean {
  const candDigits = normalizeCnpjDigits(candidateCnpj);
  if (targetCnpjDigits && candDigits && targetCnpjDigits === candDigits) {
    return true;
  }

  const candNorm = normalizeName(candidateName || '');
  if (!targetNormName || !candNorm) return false;
  if (targetNormName === candNorm) return true;
  if (
    targetNormName.length >= 8 &&
    candNorm.length >= 8 &&
    (candNorm.includes(targetNormName) || targetNormName.includes(candNorm))
  ) {
    return true;
  }
  return false;
}

function isSanitizedMonetaryValue(val: number): boolean {
  if (!Number.isFinite(val) || val <= 0 || val >= 1e11) return false;
  if ([600, 500, 1500, 1600, 0.6, 0.5].includes(val)) return false;
  return true;
}

function formatDateBrFromRaw(rawDate: any): { data: string; dataIso: string } {
  if (rawDate === null || rawDate === undefined || rawDate === '') {
    return { data: '', dataIso: '' };
  }
  const info = parseDateToMonthInfo(rawDate);
  if (info.dateIso) {
    const [y, m, d] = info.dateIso.split('-');
    return { data: `${d}/${m}/${y}`, dataIso: info.dateIso };
  }
  const s = String(rawDate).trim();
  return { data: s, dataIso: '' };
}

function isHeaderFonteLike(k: string): boolean {
  const lk = k.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  return lk.includes('fonte') || lk.includes('fr ') || lk === 'fr' || lk.includes('destinacao');
}

function shouldReplaceStatus(
  existing: PanelOBRecord,
  incomingSit: string,
  incomingIsOfficial: boolean
): boolean {
  const existingCls = classifySituacao(existing.situacaoSigla);
  const incomingCls = classifySituacao(incomingSit);

  if (existingCls.isTerminal) {
    const incomingIsOpen =
      isSituacaoAOOuVazia(incomingSit) || (!incomingCls.isTerminal && isSituacaoNaoPaga(incomingSit));
    if (incomingIsOpen) return false;
    if (existing.isOfficialStatus && !incomingIsOfficial) return false;
    return false;
  }

  if (!existing.isOfficialStatus && incomingIsOfficial && incomingSit) {
    return true;
  }

  if (existing.isOfficialStatus && !incomingIsOfficial) {
    return !existing.situacaoSigla && Boolean(incomingSit);
  }

  if (incomingCls.isTerminal) return true;

  if (isSituacaoAOOuVazia(existing.situacaoSigla) && !isSituacaoAOOuVazia(incomingSit)) {
    return true;
  }

  return !existing.situacaoSigla && Boolean(incomingSit);
}

export function buildPanelOBStatusIndex(
  files: FileData[],
  ordensBancariasList?: OrdemBancariaItem[] | null,
  _metrics?: DashboardMetrics | null
): {
  byObClean: Map<string, PanelOBRecord>;
  byObLoose: Map<string, PanelOBRecord>;
  byPpClean: Map<string, PanelOBRecord>;
  allRecords: PanelOBRecord[];
} {
  const byObClean = new Map<string, PanelOBRecord>();
  const byObLoose = new Map<string, PanelOBRecord>();
  const byPpClean = new Map<string, PanelOBRecord>();

  const upsert = (candidate: Omit<PanelOBRecord, 'situacaoLabel' | 'statusCategoria'>) => {
    const obClean = cleanDocKey(candidate.obClean || candidate.ob);
    if (!obClean) return;

    const cls = classifySituacao(candidate.situacaoSigla);
    const fullCandidate: PanelOBRecord = {
      ...candidate,
      obClean,
      situacaoSigla: cls.sigla,
      situacaoLabel: cls.label,
      statusCategoria: cls.categoria,
    };

    const existing = byObClean.get(obClean);
    if (!existing) {
      byObClean.set(obClean, fullCandidate);
    } else {
      const useIncomingSit = shouldReplaceStatus(existing, fullCandidate.situacaoSigla, fullCandidate.isOfficialStatus);
      const mergedSit = useIncomingSit ? fullCandidate.situacaoSigla : existing.situacaoSigla;
      const mergedCls = classifySituacao(mergedSit);

      const merged: PanelOBRecord = {
        ob: existing.ob.length >= fullCandidate.ob.length ? existing.ob : fullCandidate.ob,
        obClean,
        pp: existing.pp || fullCandidate.pp,
        favorecido: existing.favorecido || fullCandidate.favorecido,
        cnpj: normalizeCnpjDigits(existing.cnpj) ? existing.cnpj : (fullCandidate.cnpj || existing.cnpj),
        valor: isSanitizedMonetaryValue(existing.valor)
          ? existing.valor
          : isSanitizedMonetaryValue(fullCandidate.valor)
          ? fullCandidate.valor
          : existing.valor || fullCandidate.valor,
        data: existing.data || fullCandidate.data,
        dataIso: existing.dataIso || fullCandidate.dataIso,
        fonteRecurso: isValidFonteRecursoCode(fullCandidate.fonteRecurso)
          ? fullCandidate.fonteRecurso
          : isValidFonteRecursoCode(existing.fonteRecurso)
          ? existing.fonteRecurso
          : existing.fonteRecurso || fullCandidate.fonteRecurso,
        situacaoSigla: mergedCls.sigla,
        situacaoLabel: mergedCls.label,
        statusCategoria: mergedCls.categoria,
        origem: useIncomingSit ? fullCandidate.origem : existing.origem,
        isOfficialStatus: existing.isOfficialStatus || fullCandidate.isOfficialStatus,
      };
      byObClean.set(obClean, merged);
    }

    const latest = byObClean.get(obClean)!;
    const looseKey = stripObLeadingZeros(obClean);
    if (looseKey) {
      byObLoose.set(looseKey, latest);
    }
    const ppKey = cleanDocKey(latest.pp);
    if (ppKey) {
      const prevPp = byPpClean.get(ppKey);
      if (!prevPp || shouldReplaceStatus(prevPp, latest.situacaoSigla, latest.isOfficialStatus)) {
        byPpClean.set(ppKey, latest);
      }
    }
  };

  (files || []).forEach((file) => {
    if (!file.rows || file.rows.length === 0) return;
    const isOfficial = file.stage !== 'ordem_bancaria';
    const headers =
      file.headers && file.headers.length > 0 ? file.headers : Object.keys(file.rows[0] || {});

    const obCol = file.detectedObCol;
    const numCol = file.detectedNumeroCol || file.detectedPpCol;
    const sitObCol = file.detectedSituacaoObCol;
    const sitCol = file.detectedSituacaoCol;
    const favCol = file.detectedFavorecidoCol;
    const cnpjCol = file.detectedCnpjCol;
    const dataCol = file.detectedDataCol;
    const fonteCol = file.detectedFonteCol || headers.find((h) => isHeaderFonteLike(h));
    let valCol = file.detectedValorCol;
    if (valCol && (isHeaderFonteLike(valCol) || valCol === fonteCol)) valCol = '';

    const obFallbackCols: string[] = [];
    const sitFallbackCols: string[] = [];
    const valFallbackCols: string[] = [];
    const dateFallbackCols: string[] = [];

    for (const h of headers) {
      const lh = h.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
      if (lh.includes('ob') || lh.includes('ordem bancaria') || lh.includes('documento') || lh.includes('numero')) {
        obFallbackCols.push(h);
      }
      if (lh.includes('situac') || lh.includes('status') || lh === 'sigla' || lh === 'sit') {
        sitFallbackCols.push(h);
      }
      if (!isHeaderFonteLike(h) && h !== fonteCol && (lh.includes('valor') || lh.includes('vlr') || lh.includes('pago') || lh.includes('liquid'))) {
        valFallbackCols.push(h);
      }
      if (lh.includes('data') || lh.includes('emissao') || lh.includes('pagto') || lh.includes('referencia')) {
        dateFallbackCols.push(h);
      }
    }

    for (const row of file.rows) {
      if (!row || typeof row !== 'object') continue;

      let obCode = obCol ? extractCanonicalObCode(row[obCol]) : '';
      if (!obCode) {
        for (const col of obFallbackCols) {
          obCode = extractCanonicalObCode(row[col]);
          if (obCode) break;
        }
      }
      if (!obCode) continue;

      let rawSit = '';
      if (sitObCol && row[sitObCol] !== undefined && row[sitObCol] !== null && String(row[sitObCol]).trim() !== '') {
        rawSit = String(row[sitObCol]).trim().toUpperCase();
      }
      if (!rawSit && sitCol && row[sitCol] !== undefined && row[sitCol] !== null) {
        rawSit = String(row[sitCol]).trim().toUpperCase();
      }
      if (!rawSit) {
        for (const col of sitFallbackCols) {
          const v = String(row[col] || '').trim().toUpperCase();
          if (v) {
            rawSit = v;
            break;
          }
        }
      }
      if ((!rawSit || isSituacaoAOOuVazia(rawSit)) && sitFallbackCols.length > 0) {
        for (const col of sitFallbackCols) {
          const v = String(row[col] || '').trim().toUpperCase();
          if (v && (isSituacaoConfirmadaBanco(v).isMatch || isSituacaoRejeitadaBanco(v).isMatch)) {
            rawSit = v;
            break;
          }
        }
      }

      const rawFav = favCol ? row[favCol] : '';
      const parsedFav = parseFavorecidoField(rawFav);
      const cnpjRaw = cnpjCol ? String(row[cnpjCol] || '').trim() : '';
      const resolvedCnpj = normalizeCnpjDigits(cnpjRaw) ? cnpjRaw : parsedFav.cnpj;

      const rawPp = numCol ? String(row[numCol] || '').trim() : '';
      const ppClean = rawPp && cleanDocKey(rawPp) !== cleanDocKey(obCode) ? rawPp : '';

      let val = valCol ? parseCurrencyValue(row[valCol]) : 0;
      if (!isSanitizedMonetaryValue(val)) {
        val = 0;
        for (const col of valFallbackCols) {
          const cand = parseCurrencyValue(row[col]);
          if (isSanitizedMonetaryValue(cand)) {
            val = cand;
            break;
          }
        }
      }

      let dateInfo = dataCol ? formatDateBrFromRaw(row[dataCol]) : { data: '', dataIso: '' };
      if (!dateInfo.data) {
        for (const col of dateFallbackCols) {
          const cand = formatDateBrFromRaw(row[col]);
          if (cand.data) {
            dateInfo = cand;
            break;
          }
        }
      }

      let rawFonte = fonteCol ? String(row[fonteCol] || '').trim() : '';
      if (!isValidFonteRecursoCode(rawFonte)) {
        const found = findFonteRecursoInRow(Object.values(row));
        if (found) rawFonte = found;
      }

      upsert({
        ob: obCode,
        obClean: cleanDocKey(obCode),
        pp: ppClean,
        favorecido: parsedFav.name || String(rawFav || '').trim(),
        cnpj: resolvedCnpj || '',
        valor: val,
        data: dateInfo.data,
        dataIso: dateInfo.dataIso,
        fonteRecurso: isValidFonteRecursoCode(rawFonte) ? rawFonte.toUpperCase() : '',
        situacaoSigla: rawSit,
        origem: file.fileName,
        isOfficialStatus: isOfficial,
      });
    }
  });

  if (ordensBancariasList && ordensBancariasList.length > 0) {
    ordensBancariasList.forEach((item) => {
      const obCode = extractCanonicalObCode(item.ob);
      if (!obCode) return;
      const dateInfo = formatDateBrFromRaw(item.dataReferencia || item.dataIso);
      upsert({
        ob: obCode,
        obClean: cleanDocKey(obCode),
        pp: item.pp || '',
        favorecido: item.favorecidoName || parseFavorecidoField(item.favorecidoRaw).name || '',
        cnpj: item.favorecidoCnpj || parseFavorecidoField(item.favorecidoRaw).cnpj || '',
        valor: isSanitizedMonetaryValue(item.valor) ? item.valor : 0,
        data: dateInfo.data,
        dataIso: item.dataIso || dateInfo.dataIso,
        fonteRecurso: isValidFonteRecursoCode(item.fonteRecurso) ? item.fonteRecurso.toUpperCase() : '',
        situacaoSigla: item.isFonteProviderOnly ? '' : (item.situacao || '').trim().toUpperCase(),
        origem: 'OBs - sigef - Principal',
        isOfficialStatus: !item.isFonteProviderOnly,
      });
    });
  }

  return {
    byObClean,
    byObLoose,
    byPpClean,
    allRecords: Array.from(byObClean.values()),
  };
}

/**
 * Coleta EXCLUSIVAMENTE as Ordens Bancárias na situação AO pertencentes aos fornecedores atendidos selecionados,
 * ordenadas alfabeticamente por fornecedor e depois por número da OB.
 */
export function collectAttendedSuppliersOBs(
  attendedSuppliers: AttendedSupplierInput[],
  files: FileData[],
  ordensBancariasList?: OrdemBancariaItem[] | null,
  metrics?: DashboardMetrics | null
): AttendedSupplierOBRecord[] {
  if (!attendedSuppliers || attendedSuppliers.length === 0) return [];

  const { allRecords } = buildPanelOBStatusIndex(files, ordensBancariasList, metrics);
  if (allRecords.length === 0) return [];

  const preparedTargets = attendedSuppliers.map((sup) => ({
    sup,
    cnpjDigits: normalizeCnpjDigits(sup.cnpj),
    normName: normalizeName(sup.name),
  }));

  const matched: AttendedSupplierOBRecord[] = [];

  for (const obRec of allRecords) {
    // Somente Ordens Bancárias cuja situação consolidada seja estritamente AO
    if (obRec.statusCategoria !== 'pendente' || !isSituacaoAO(obRec.situacaoSigla)) {
      continue;
    }

    const foundTarget = preparedTargets.find((t) =>
      matchesSupplierTarget(t.cnpjDigits, t.normName, obRec.cnpj, obRec.favorecido)
    );
    if (!foundTarget) continue;

    matched.push({
      ...obRec,
      fornecedorAtendido: foundTarget.sup.name,
      cnpjAtendido:
        normalizeCnpjDigits(foundTarget.sup.cnpj)
          ? foundTarget.sup.cnpj
          : obRec.cnpj || foundTarget.sup.cnpj || 'N/I',
      classificacao: foundTarget.sup.isRec ? 'Recorrente' : 'Comum',
      observacaoAtendimento: foundTarget.sup.observacao || '',
    });
  }

  return matched.sort((a, b) => {
    const nameCmp = a.fornecedorAtendido.localeCompare(b.fornecedorAtendido, 'pt-BR', {
      sensitivity: 'base',
    });
    if (nameCmp !== 0) return nameCmp;
    return a.ob.localeCompare(b.ob, 'pt-BR', { numeric: true });
  });
}

/**
 * Agrupa as OBs em situação AO por fornecedor em ordem alfabética (A → Z).
 */
export function groupAttendedOBsBySupplier(
  obs: AttendedSupplierOBRecord[]
): AttendedSupplierOBGroup[] {
  const map = new Map<string, AttendedSupplierOBGroup>();

  for (const item of obs) {
    const cnpjDigits = normalizeCnpjDigits(item.cnpjAtendido);
    const groupKey = cnpjDigits ? `cnpj_${cnpjDigits}` : `name_${normalizeName(item.fornecedorAtendido)}`;
    let group = map.get(groupKey);
    if (!group) {
      group = {
        fornecedor: item.fornecedorAtendido,
        cnpj: item.cnpjAtendido || 'N/I',
        classificacao: item.classificacao,
        observacaoAtendimento: item.observacaoAtendimento || '',
        obs: [],
        totalValor: 0,
      };
      map.set(groupKey, group);
    }
    group.obs.push(item);
    group.totalValor += item.valor || 0;
  }

  const groups = Array.from(map.values());
  groups.forEach((g) => {
    g.obs.sort((a, b) => a.ob.localeCompare(b.ob, 'pt-BR', { numeric: true }));
  });
  groups.sort((a, b) =>
    a.fornecedor.localeCompare(b.fornecedor, 'pt-BR', { sensitivity: 'base' })
  );
  return groups;
}

export function exportAttendedSuppliersOBsToExcel(
  obs: AttendedSupplierOBRecord[],
  monthLabel: string,
  monthKey: string
): void {
  const groups = groupAttendedOBsBySupplier(obs);
  const totalValor = obs.reduce((acc, o) => acc + (o.valor || 0), 0);

  // Aba 1: Detalhamento organizado por fornecedor em ordem alfabética (somente OBs em situação AO)
  const detailRows: Record<string, any>[] = [];
  groups.forEach((group, groupIdx) => {
    detailRows.push({
      '#': `${groupIdx + 1}`,
      'Fornecedor / Favorecido': group.fornecedor,
      'CNPJ': group.cnpj,
      'Ordem Bancária (OB)': `${group.obs.length} OB(s) em AO`,
      'PP': '',
      'Data': '',
      'Fonte': '',
      'Situação': 'AO',
      'Valor (R$)': group.totalValor,
      'Mês Ref. Atendimento': monthLabel,
    });

    group.obs.forEach((item, obIdx) => {
      detailRows.push({
        '#': `${groupIdx + 1}.${obIdx + 1}`,
        'Fornecedor / Favorecido': item.fornecedorAtendido,
        'CNPJ': item.cnpjAtendido || 'N/I',
        'Ordem Bancária (OB)': item.ob,
        'PP': item.pp || '-',
        'Data': item.data || '-',
        'Fonte': item.fonteRecurso || '-',
        'Situação': item.situacaoSigla || 'AO',
        'Valor (R$)': item.valor,
        'Mês Ref. Atendimento': monthLabel,
      });
    });

    detailRows.push({
      '#': '',
      'Fornecedor / Favorecido': `Subtotal — ${group.fornecedor}`,
      'CNPJ': group.cnpj,
      'Ordem Bancária (OB)': `${group.obs.length} OB(s) AO`,
      'PP': '',
      'Data': '',
      'Fonte': '',
      'Situação': 'AO',
      'Valor (R$)': group.totalValor,
      'Mês Ref. Atendimento': '',
    });

    // Linha em branco para separação visual entre fornecedores
    detailRows.push({
      '#': '',
      'Fornecedor / Favorecido': '',
      'CNPJ': '',
      'Ordem Bancária (OB)': '',
      'PP': '',
      'Data': '',
      'Fonte': '',
      'Situação': '',
      'Valor (R$)': '',
      'Mês Ref. Atendimento': '',
    });
  });

  detailRows.push({
    '#': '',
    'Fornecedor / Favorecido': `TOTAL GERAL (${groups.length} Fornecedor(es))`,
    'CNPJ': '',
    'Ordem Bancária (OB)': `${obs.length} OB(s) em AO`,
    'PP': '',
    'Data': '',
    'Fonte': '',
    'Situação': 'AO',
    'Valor (R$)': totalValor,
    'Mês Ref. Atendimento': monthLabel,
  });

  const detailSheet = XLSX.utils.json_to_sheet(detailRows);
  detailSheet['!cols'] = [
    { wch: 8 },
    { wch: 46 },
    { wch: 20 },
    { wch: 20 },
    { wch: 16 },
    { wch: 12 },
    { wch: 14 },
    { wch: 12 },
    { wch: 18 },
    { wch: 20 },
  ];

  // Aba 2: Resumo consolidado por fornecedor em ordem alfabética
  const summaryRows: Record<string, any>[] = groups.map((group, idx) => ({
    '#': idx + 1,
    'Fornecedor / Favorecido': group.fornecedor,
    'CNPJ': group.cnpj,
    'Classificação': group.classificacao,
    'Qtd. OBs (AO)': group.obs.length,
    'Valor Total AO (R$)': group.totalValor,
    'Anotação': group.observacaoAtendimento || '',
  }));
  summaryRows.push({
    '#': '',
    'Fornecedor / Favorecido': `TOTAL GERAL (${groups.length} Fornecedor(es))`,
    'CNPJ': '',
    'Classificação': '',
    'Qtd. OBs (AO)': obs.length,
    'Valor Total AO (R$)': totalValor,
    'Anotação': '',
  });

  const summarySheet = XLSX.utils.json_to_sheet(summaryRows);
  summarySheet['!cols'] = [
    { wch: 6 },
    { wch: 46 },
    { wch: 20 },
    { wch: 14 },
    { wch: 15 },
    { wch: 20 },
    { wch: 32 },
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, detailSheet, 'OBs_AO_Por_Fornecedor');
  XLSX.utils.book_append_sheet(workbook, summarySheet, 'Resumo_Por_Fornecedor');
  XLSX.writeFile(workbook, `OBs_AO_Fornecedores_Atendidos_${monthKey}.xlsx`);
}

export function exportAttendedSuppliersOBsToPdf(
  obs: AttendedSupplierOBRecord[],
  monthLabel: string,
  monthKey: string,
  attendedSuppliersCount: number
): void {
  const doc = new jsPDF({
    orientation: 'landscape',
    unit: 'mm',
    format: 'a4',
  });

  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const marginX = 12;

  const groups = groupAttendedOBsBySupplier(obs);
  const totalValor = obs.reduce((acc, o) => acc + (o.valor || 0), 0);

  doc.setFillColor(6, 29, 21);
  doc.rect(0, 0, pageWidth, 24, 'F');

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(10.5);
  doc.setTextColor(255, 255, 255);
  doc.text('GOVERNO DO ESTADO DO RIO GRANDE DO NORTE — SESAP / FES', marginX, 9);

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(9.5);
  doc.setTextColor(167, 243, 208);
  doc.text(`ORDENS BANCÁRIAS (SITUAÇÃO AO) — FORNECEDORES ATENDIDOS (${monthLabel.toUpperCase()})`, marginX, 17);

  const emissao = `${new Date().toLocaleDateString('pt-BR')} às ${new Date().toLocaleTimeString('pt-BR')}`;
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(7.5);
  doc.setTextColor(209, 250, 229);
  doc.text(`Emissão: ${emissao}`, pageWidth - marginX, 17, { align: 'right' });

  const kpiY = 29;
  const gap = 4;
  const boxW = (pageWidth - marginX * 2 - gap * 2) / 3;

  const drawBox = (
    idx: number,
    label: string,
    value: string,
    sub: string,
    fill: [number, number, number],
    border: [number, number, number],
    text: [number, number, number]
  ) => {
    const x = marginX + idx * (boxW + gap);
    doc.setFillColor(...fill);
    doc.setDrawColor(...border);
    doc.roundedRect(x, kpiY, boxW, 15, 2, 2, 'FD');
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(6.5);
    doc.setTextColor(...text);
    doc.text(label, x + 3, kpiY + 4.8);
    doc.setFontSize(9.5);
    doc.text(value, x + 3, kpiY + 10.2);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(6.5);
    doc.text(sub, x + 3, kpiY + 13.5);
  };

  drawBox(
    0,
    'FORNECEDORES COM OBs (AO)',
    String(groups.length),
    `de ${attendedSuppliersCount} fornecedor(es) atendido(s) no mês`,
    [248, 250, 252],
    [226, 232, 240],
    [15, 23, 42]
  );
  drawBox(
    1,
    'TOTAL DE OBs EM SITUAÇÃO AO',
    String(obs.length),
    'Ordem alfabética • Agrupado por fornecedor',
    [255, 247, 237],
    [254, 215, 170],
    [154, 52, 18]
  );
  drawBox(
    2,
    'VALOR TOTAL (SITUAÇÃO AO)',
    formatBRL(totalValor),
    `${groups.length} fornecedor(es) • ${obs.length} OB(s)`,
    [236, 253, 245],
    [167, 243, 208],
    [6, 95, 70]
  );

  const body: any[] = [];

  if (groups.length === 0) {
    body.push([
      {
        content: 'Nenhuma ordem bancária na situação AO localizada para os fornecedores atendidos selecionados.',
        colSpan: 7,
        styles: { halign: 'center', textColor: [100, 116, 139] },
      },
    ]);
  } else {
    groups.forEach((group, groupIdx) => {
      // Cabeçalho do grupo do fornecedor (A → Z) com CNPJ, contagem de OBs e subtotal
      body.push([
        {
          content: `${groupIdx + 1}. ${group.fornecedor}   •   CNPJ: ${group.cnpj}   (${group.obs.length} OB${group.obs.length === 1 ? '' : 's'} em AO)`,
          colSpan: 6,
          styles: {
            fillColor: [236, 253, 245],
            textColor: [6, 78, 59],
            fontStyle: 'bold',
            fontSize: 7.5,
            halign: 'left',
          },
        },
        {
          content: formatBRL(group.totalValor),
          styles: {
            fillColor: [236, 253, 245],
            textColor: [6, 78, 59],
            fontStyle: 'bold',
            fontSize: 7.5,
            halign: 'right',
          },
        },
      ]);

      // Linhas das OBs AO pertencentes a este fornecedor
      group.obs.forEach((item, obIdx) => {
        body.push([
          `${groupIdx + 1}.${obIdx + 1}`,
          item.ob,
          item.pp || '-',
          item.data || '-',
          item.fonteRecurso || '-',
          {
            content: item.situacaoSigla || 'AO',
            styles: { textColor: [180, 83, 9], fontStyle: 'bold', halign: 'center' },
          },
          formatBRL(item.valor),
        ]);
      });
    });

    body.push([
      {
        content: `TOTAL GERAL (${groups.length} Fornecedor(es) • ${obs.length} OB(s) em AO):`,
        colSpan: 6,
        styles: { halign: 'right', fontStyle: 'bold', fillColor: [6, 78, 59], textColor: [255, 255, 255], fontSize: 8 },
      },
      {
        content: formatBRL(totalValor),
        styles: { halign: 'right', fontStyle: 'bold', fillColor: [6, 78, 59], textColor: [255, 255, 255], fontSize: 8 },
      },
    ]);
  }

  autoTable(doc, {
    startY: kpiY + 19,
    head: [['#', 'ORDEM BANCÁRIA (OB)', 'PP', 'DATA', 'FONTE', 'SITUAÇÃO', 'VALOR (R$)']],
    body,
    theme: 'grid',
    headStyles: {
      fillColor: [6, 78, 59],
      textColor: [255, 255, 255],
      fontSize: 7.5,
      fontStyle: 'bold',
      halign: 'center',
    },
    bodyStyles: {
      fontSize: 7.2,
      textColor: [30, 41, 59],
      cellPadding: 1.9,
    },
    columnStyles: {
      0: { cellWidth: 16, halign: 'center' },
      1: { cellWidth: 52, halign: 'left', fontStyle: 'bold' },
      2: { cellWidth: 46, halign: 'center' },
      3: { cellWidth: 32, halign: 'center' },
      4: { cellWidth: 34, halign: 'center' },
      5: { cellWidth: 38, halign: 'center' },
      6: { cellWidth: 55, halign: 'right', fontStyle: 'bold' },
    },
    margin: { left: marginX, right: marginX, bottom: 14 },
  });

  const pageCount = doc.getNumberOfPages();
  for (let i = 1; i <= pageCount; i++) {
    doc.setPage(i);
    doc.setFontSize(7);
    doc.setTextColor(100, 116, 139);
    doc.text(
      `Fornecedores Atendidos (${monthLabel}) — Somente Ordens Bancárias em Situação AO (Ordem Alfabética por Fornecedor)`,
      marginX,
      pageHeight - 7
    );
    doc.text(`Página ${i} de ${pageCount}`, pageWidth - marginX, pageHeight - 7, { align: 'right' });
  }

  doc.save(`OBs_AO_Fornecedores_Atendidos_${monthKey}.pdf`);
}
