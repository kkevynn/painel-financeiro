export type StageKey = 'base_pps' | 'obedece' | 'nao_obedece' | 'outros' | 'pp_prontas';

export interface StageInfo {
  id: StageKey;
  label: string;
  shortLabel: string;
  badgeBg: string;
  badgeText: string;
  border: string;
  text: string;
  bgLight: string;
  description: string;
}

export const STAGES_CONFIG: Record<StageKey, StageInfo> = {
  base_pps: {
    id: 'base_pps',
    label: 'Arquivo Base de PPs (Relatório SIGEF)',
    shortLabel: 'Base de PPs',
    badgeBg: 'bg-emerald-600',
    badgeText: 'bg-emerald-100 text-emerald-900 border-emerald-300',
    border: 'border-emerald-300 hover:border-emerald-500',
    text: 'text-emerald-950',
    bgLight: 'bg-emerald-50/70',
    description: 'Arquivo base contendo todas as informações das PPs (ex: Relatorio_*.xls do SIGEF).',
  },
  pp_prontas: {
    id: 'pp_prontas',
    label: 'Arquivo Base de PPs (Relatório SIGEF)',
    shortLabel: 'Base de PPs',
    badgeBg: 'bg-emerald-600',
    badgeText: 'bg-emerald-100 text-emerald-900 border-emerald-300',
    border: 'border-emerald-300 hover:border-emerald-500',
    text: 'text-emerald-950',
    bgLight: 'bg-emerald-50/70',
    description: 'Arquivo base contendo todas as informações das PPs (ex: Relatorio_*.xls do SIGEF).',
  },
  obedece: {
    id: 'obedece',
    label: 'Liquidação - Obedece à Ordem Cronológica',
    shortLabel: 'Liquidado - Obedece',
    badgeBg: 'bg-blue-500',
    badgeText: 'bg-blue-100 text-blue-900 border-blue-300',
    border: 'border-blue-300 hover:border-blue-500',
    text: 'text-blue-950',
    bgLight: 'bg-blue-50/70',
    description: 'Upload de liquidações que obedecem à ordem cronológica',
  },
  nao_obedece: {
    id: 'nao_obedece',
    label: 'Liquidação - Não Obedece à Ordem Cronológica',
    shortLabel: 'Liquidado - Não Obedece',
    badgeBg: 'bg-amber-500',
    badgeText: 'bg-amber-100 text-amber-900 border-amber-300',
    border: 'border-amber-300 hover:border-amber-500',
    text: 'text-amber-950',
    bgLight: 'bg-amber-50/70',
    description: 'Upload de liquidações com exceção à ordem cronológica',
  },
  outros: {
    id: 'outros',
    label: 'Outras Planilhas / Geral',
    shortLabel: 'Geral',
    badgeBg: 'bg-slate-500',
    badgeText: 'bg-slate-100 text-slate-900 border-slate-300',
    border: 'border-slate-300 hover:border-slate-500',
    text: 'text-slate-950',
    bgLight: 'bg-slate-50/70',
    description: 'Planilhas financeiras de credores gerais',
  },
};

export interface FileData {
  id: string;
  fileName: string;
  headers: string[];
  rows: Record<string, any>[];
  detectedFavorecidoCol: string;
  detectedValorCol: string;
  detectedCnpjCol?: string;
  detectedSituacaoCol?: string;
  detectedPpCol?: string;
  detectedObCol?: string;
  detectedNumeroCol?: string;
  detectedNotaEmpenhoCol?: string;
  detectedFavorecidoNECol?: string;
  detectedProcessoCol?: string;
  detectedFonteCol?: string;
  stage?: StageKey;
  totalCount?: number;
}

export interface ColumnMappingConfig {
  favorecidoCol: string;
  valorCol: string;
  stage?: StageKey;
  cnpjCol?: string;
  situacaoCol?: string;
  numeroCol?: string;
  obCol?: string;
  notaEmpenhoCol?: string;
  favorecidoNECol?: string;
  processoCol?: string;
  fonteCol?: string;
}

export interface ConsolidatedSupplier {
  name: string;
  cnpj: string;
  basePpsValue?: number;
  ppProntasValue: number;
  liqObedeceValue: number;
  liqNaoObedeceValue: number;
  outrosValue: number;
  enviadasBancoValue?: number;
  countPpProntasAO?: number;
  countEnviadasBanco?: number;
  valorPagoAno?: number; // Total acumulado pago no ano (status conclusivos: PPCB, CB, etc.)
  countPagoAno?: number;
  pagamentosDetalhePorFonte?: Record<string, number>; // Agrupamento { [fonteRecurso]: valorTotal }
  pagamentosPorMes?: Record<string, { total: number; count: number; porFonte: Record<string, number> }>; // Por mesKey (ex: "2026-01")
  totalValue: number;
  totalCount: number;
  obs?: string[];
  pps?: string[];
  notasEmpenho?: string[];
  favorecidoNE?: string;
  hasFullSigefFields?: boolean;
}

export interface CategoryDefinition {
  id: string;
  name: string;
  color: string;
  keywords: string[];
  description?: string;
  isDefault?: boolean;
}

export interface CategoryMetrics {
  category: CategoryDefinition;
  totalValue: number;
  totalCount: number;
  suppliers: ConsolidatedSupplier[];
}

export interface DashboardMetrics {
  grandTotalValue: number;
  grandTotalCount: number;
  totalBasePps?: number;
  totalPpProntas: number;
  totalLiqObedece: number;
  totalLiqNaoObedece: number;
  totalOutros: number;
  totalEnviadasBanco?: number;
  countPpProntasAO?: number;
  countEnviadasBanco?: number;
  totalPagoAno?: number; // Total pago no ano em toda a base
  countPagoAnoTotal?: number;
  consolidatedSuppliers: ConsolidatedSupplier[];
  isProcessed: boolean;
}

export type SituacaoCategory = 'pp_pronta' | 'enviada_banco' | 'confirmada' | 'rejeitada' | 'pendente' | 'nao_localizada';

export interface LiberacaoItem {
  id: string;
  pp: string;
  ob: string;
  fornecedor: string;
  valor: number;
  sigla: string;
  situacaoTraduzida: string;
  categoria: SituacaoCategory;
  origemLinha?: number;
  notaEmpenho?: string;
  favorecidoNE?: string;
  hasFullSigefFields?: boolean;
  foundInDatabase: boolean;
  searchQueryKey?: string;
}

export interface LiberacaoCrossResult {
  totalPesquisado: number;
  totalEncontrado: number;
  totalNaoEncontrado: number;
  valorTotalEncontrado: number;
  items: LiberacaoItem[];
  databaseFileName?: string;
  pesquisaFileName?: string;
  processedAt?: Date;
}

export interface RecurringSupplierDetail {
  id: string;
  fileName: string;
  stage: StageKey;
  rawFavorecido: string;
  valor: number;
  rowNumber: number;
  rowData: Record<string, any>;
}

export interface ManualRecurringSupplierEntry {
  id: string;
  name: string;
  cnpj: string;
  categoryTag?: string;
  notes?: string;
  createdAt: string;
}

export interface RecurringSupplier {
  id?: string;
  name: string;
  cnpj: string;
  notes?: string;
  categoryTag?: string;
  createdAt?: string;
  basePpsValue?: number;
  ppProntasValue: number;
  liqObedeceValue: number;
  liqNaoObedeceValue: number;
  outrosValue: number;
  totalValue: number;
  totalCount: number;
  averageValue: number;
  stagesCount: number;
  isMultiStage: boolean;
  recurrenceLevel: 'alta' | 'media' | 'moderada';
  details?: RecurringSupplierDetail[];
}

export interface QuebraOrdemItem {
  id: string;
  favorecido: string;
  cnpj: string;
  valorOB: number;
  pp: string;
  ob: string;
  processo?: string;
  notaEmpenho?: string;
  fonteRecurso: string;
  situacao: string; // e.g. "AO"
  isPossivelQuebraOrdem?: boolean;
  fonteMonitorada?: string;
  isFluxoRapido?: boolean;
  fonteFluxoRapido?: string;
  liquidacaoCorrespondente?: {
    valorLiquidacao: number;
    etapaOrigem: StageKey;
    nomePlanilha: string;
    documento?: string;
    diferencaImpostos: number;
    percentualDiferenca: number;
    matched: boolean;
    fonteRecurso?: string;
    isFonteMonitorada?: boolean;
    isFluxoRapido?: boolean;
    fonteFluxoRapido?: string;
  };
  rawRow?: Record<string, any>;
}

export interface QuebraOrdemResult {
  totalProcessados: number;
  totalAOEncontrados: number;
  totalIgnoradosFonte600: number;
  totalValidos: number;
  totalComLiquidacaoCasada: number;
  totalPossivelQuebraOrdem: number;
  valorPossivelQuebraOrdem: number;
  totalFluxoRapido?: number;
  valorFluxoRapido?: number;
  valorTotalOB: number;
  valorTotalLiquidacao: number;
  valorTotalImpostos: number;
  items: QuebraOrdemItem[];
  fileName?: string;
  processedAt?: Date;
}

// ----------------------------------------------------
// Toast Alert Notification System
// ----------------------------------------------------
export type ToastAlertType = 'sucesso' | 'critico' | 'info';

export interface ToastAlert {
  id: string;
  tipo: ToastAlertType;
  titulo: string;
  mensagem: string;
  fornecedor?: string;
  cnpj?: string;
  pp?: string;
  ob?: string;
  ne?: string;
  fileName?: string;
  valor?: number;
  situacaoAnterior?: string;
  situacaoNova?: string;
  observacao?: string;
  dataHora: Date;
  dataReferencia?: string;
  read?: boolean;
  resolvido?: boolean;
}

// ----------------------------------------------------
// Inferência de Pagamento e Machine Learning Local
// ----------------------------------------------------
export interface InferredPaymentMatch {
  id: string;
  nlNumero: string; // Nota de Lançamento que sumiu do Obedece / Não Obedece
  processo?: string;
  credor: string;
  cnpj: string;
  valorNL: number;
  ppNumero: string;
  obNumero?: string;
  valorPP: number;
  diferencaImpostos: number;
  percentualRetencao: number;
  statusSituacaoPP: string;
  dataInferido: string;
  statusValidacao: 'pendente' | 'confirmado' | 'rejeitado';
  confianca: number; // 0 a 100
  motivo: string;
  baseOrigem: 'obedece' | 'nao_obedece' | 'pp_prontas';
  automaticamenteConfirmado?: boolean;
}

export interface LearnedCreditorPattern {
  cnpj: string;
  credor: string;
  aliquotaImpostoRetidoMedia: number; // e.g. 0.0585 (5.85%)
  toleranciaPercentual: number; // e.g. 0.02 (2%)
  totalConfirmacoes: number;
  ultimaConfirmacao: string;
  sampleNL: string;
  samplePP: string;
}

export interface AuditRowItem {
  id: string;
  sourceType: 'base_pps' | 'pp_prontas' | 'obedece' | 'nao_obedece' | 'outros' | 'status_banco' | 'pagamentos_emitidos';
  sourceLabel: string;
  sourceFileName: string;
  documentNumber: string;
  date?: string;
  value: number;
  status?: string;
  fonteRecurso?: string;
  details?: string;
  pp?: string;
  ob?: string;
  notaEmpenho?: string;
  favorecidoNE?: string;
  rawData: Record<string, any>;
}

export interface CompanyAuditSummary {
  supplierName: string;
  cnpj: string;
  totalConsolidatedValue: number;
  ppProntasValue: number;
  ppProntasCount: number;
  ppProntasAOValue?: number;
  ppProntasAOCount?: number;
  enviadasBancoValue?: number;
  enviadasBancoCount?: number;
  liqObedeceValue: number;
  liqObedeceCount: number;
  liqNaoObedeceValue: number;
  liqNaoObedeceCount: number;
  bancoRecordsCount: number;
  bancoConfirmedCount: number;
  quebraOrdemCount: number;
  quebraOrdemValue: number;
  allRows: AuditRowItem[];
}

export interface RawPagamentoEmitido {
  favorecidoRaw: string;
  cnpj: string;
  name: string;
  valorOB: number;
  pp: string;
  ob: string;
  processo: string;
  notaEmpenho: string;
  fonteRecurso: string;
  situacao: string;
  rawRow: Record<string, any>;
  isAO: boolean;
  isIgnoradaFonte600: boolean;
}

