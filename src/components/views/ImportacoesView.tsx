import React, { useState, useRef } from 'react';
import {
  FileSpreadsheet,
  Table as TableIcon,
  ArrowRight,
  HelpCircle,
  Upload,
  CheckCircle2,
  Trash2,
  Sparkles,
  ExternalLink,
  Layers,
  FileCheck2,
  AlertCircle,
  Save,
  Loader2,
  Check,
} from 'lucide-react';
import { FileData, DashboardMetrics, StageKey } from '../../types';
import { UploadSection } from '../UploadSection';
import {
  readSpreadsheetAsMatrix,
  parsePagamentosEmitidosMatrix,
  extractAORowsFromWorkspaceFiles,
  RawPagamentoEmitido,
} from '../../utils/quebraOrdemParser';
import {
  parseListarOrdemBancariaMatrix,
  parseListarOrdemBancariaMatrixAsync,
  OrdemBancariaItem,
} from '../../utils/ordemBancariaParser';
import { formatBRL } from '../../utils/excelParser';
import { InfoHelpButton } from '../InfoHelpButton';

interface ImportacoesViewProps {
  files: FileData[];
  onFileUpload: (file: File, stage?: StageKey) => void;
  onRemoveFile: (fileId: string) => void;
  onChangeFileStage?: (fileId: string, stage: StageKey) => void;
  onProcessData: () => void;
  onSaveData?: () => Promise<void> | void;
  onOpenColumnMapModal: (fileId: string) => void;
  isProcessing: boolean;
  metrics: DashboardMetrics;
  onNavigateToTab: (tab: string) => void;

  // Quebra de Ordem (Pagamentos Emitidos)
  rawPagamentosList?: RawPagamentoEmitido[] | null;
  pagamentosFileName?: string;
  onUpdatePagamentosList?: (list: RawPagamentoEmitido[] | null, fileName?: string) => void;

  // Ordens Bancárias / Pagos (Listar Ordem Bancária)
  ordensBancariasList?: OrdemBancariaItem[] | null;
  ordensBancariasFileName?: string;
  onUpdateOrdensBancariasList?: (list: OrdemBancariaItem[] | null, fileName?: string) => void;
}

export const ImportacoesView: React.FC<ImportacoesViewProps> = ({
  files,
  onFileUpload,
  onRemoveFile,
  onChangeFileStage,
  onProcessData,
  onSaveData,
  onOpenColumnMapModal,
  isProcessing,
  metrics,
  onNavigateToTab,
  rawPagamentosList,
  pagamentosFileName,
  onUpdatePagamentosList,
  ordensBancariasList,
  ordensBancariasFileName,
  onUpdateOrdensBancariasList,
}) => {
  const [activeImportCategory, setActiveImportCategory] = useState<'fases' | 'ordens_bancarias' | 'pagamentos_emitidos'>('fases');
  const [selectedFileId, setSelectedFileId] = useState<string | null>(null);

  // Loading states for local parsers
  const [isLoadingPagamentos, setIsLoadingPagamentos] = useState(false);
  const [isLoadingOrdensBancarias, setIsLoadingOrdensBancarias] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [justSaved, setJustSaved] = useState(false);
  const [uploadError, setUploadError] = useState<string | null>(null);

  // Hidden refs
  const pagamentosInputRef = useRef<HTMLInputElement>(null);
  const ordensBancariasInputRef = useRef<HTMLInputElement>(null);

  const activePreviewFile = files.find((f) => f.id === selectedFileId) || files[0];

  const handleSaveClick = async () => {
    if (!onSaveData || isSaving) return;
    setIsSaving(true);
    try {
      await onSaveData();
      setJustSaved(true);
      setTimeout(() => setJustSaved(false), 2500);
    } finally {
      setIsSaving(false);
    }
  };

  const workspaceAoData = React.useMemo(() => {
    return extractAORowsFromWorkspaceFiles(files);
  }, [files]);

  // Handler: Pagamentos Emitidos Upload
  const handlePagamentosUpload = async (e: React.ChangeEvent<HTMLInputElement> | React.DragEvent) => {
    let file: File | undefined;
    if ('dataTransfer' in e) {
      e.preventDefault();
      file = e.dataTransfer.files[0];
    } else {
      file = e.target.files?.[0];
    }

    if (!file) return;

    try {
      setIsLoadingPagamentos(true);
      setUploadError(null);
      const matrix = await readSpreadsheetAsMatrix(file);
      const parsed = parsePagamentosEmitidosMatrix(matrix);

      if (parsed.length === 0) {
        setUploadError('Não foi possível identificar registros válidos no arquivo selecionado. Verifique o formato.');
        return;
      }

      if (onUpdatePagamentosList) {
        onUpdatePagamentosList(parsed, file.name);
      }
    } catch (err: any) {
      setUploadError(`Erro ao processar Pagamentos Emitidos: ${err?.message || err}`);
    } finally {
      setIsLoadingPagamentos(false);
      if (pagamentosInputRef.current) pagamentosInputRef.current.value = '';
    }
  };

  const handleClearPagamentos = () => {
    if (onUpdatePagamentosList) {
      onUpdatePagamentosList(null, '');
    }
  };

  // Handler: Listar Ordem Bancária Upload (2 rows per record)
  const handleOrdensBancariasUpload = async (e: React.ChangeEvent<HTMLInputElement> | React.DragEvent) => {
    let file: File | undefined;
    if ('dataTransfer' in e) {
      e.preventDefault();
      file = e.dataTransfer.files[0];
    } else {
      file = e.target.files?.[0];
    }

    if (!file) return;

    try {
      setIsLoadingOrdensBancarias(true);
      setUploadError(null);
      const matrix = await readSpreadsheetAsMatrix(file);
      const parsed = await parseListarOrdemBancariaMatrixAsync(matrix);

      if (parsed.length === 0) {
        setUploadError(
          'Não foi possível identificar ordens bancárias conclusivas (ex: Situação CB) no arquivo selecionado. Verifique se o relatório é o "Listar Ordem Bancária" com estrutura de duas linhas por registro.'
        );
        return;
      }

      if (onUpdateOrdensBancariasList) {
        onUpdateOrdensBancariasList(parsed, file.name);
      }
    } catch (err: any) {
      setUploadError(`Erro ao processar relatório de Ordens Bancárias: ${err?.message || err}`);
    } finally {
      setIsLoadingOrdensBancarias(false);
      if (ordensBancariasInputRef.current) ordensBancariasInputRef.current.value = '';
    }
  };

  const handleClearOrdensBancarias = () => {
    if (onUpdateOrdensBancariasList) {
      onUpdateOrdensBancariasList(null, '');
    }
  };

  // Ordens Bancárias stats
  const ordensBancariasCount = ordensBancariasList?.length || 0;
  const ordensBancariasTotalValor = React.useMemo(() => {
    return (ordensBancariasList || []).reduce((acc, it) => acc + (it.valor || 0), 0);
  }, [ordensBancariasList]);

  // Quebra de Ordem counts
  const pagamentosCount = rawPagamentosList?.length || 0;
  const pagamentosAoCount = rawPagamentosList?.filter((p) => p.isAO && !p.isIgnoradaFonte600).length || 0;

  return (
    <div className="space-y-8 pb-12">
      {/* Banner / Title Header */}
      <div className="bg-gradient-to-r from-slate-900 via-emerald-950 to-slate-900 rounded-2xl p-6 md:p-8 text-white shadow-xl relative overflow-hidden border border-slate-800">
        <div className="relative z-10 flex flex-col md:flex-row md:items-center justify-between gap-6">
          <div className="flex items-center gap-3">
            <h2 className="text-2xl font-black text-white tracking-tight">
              Central Única de Importações
            </h2>
            <InfoHelpButton
              title="Central de Importações"
              variant="dark"
              content={
                <p>
                  Todas as entradas de arquivos do sistema concentradas em um só lugar: carregue planilhas das etapas de liquidação (PPs Prontas, Liquidação Obedece/Não Obedece) e o Relatório SIGEF de Pagamentos Emitidos para consolidação e auditoria das métricas.
                </p>
              }
            />
          </div>

          <div className="flex flex-wrap items-center gap-2.5 shrink-0">
            {files.length > 0 && onSaveData && (
              <button
                onClick={handleSaveClick}
                disabled={isSaving || isProcessing}
                className={`px-4 py-2 font-bold text-xs rounded-xl transition flex items-center gap-2 shadow-md shadow-emerald-950/20 active:scale-95 ${
                  justSaved
                    ? 'bg-emerald-600 text-white'
                    : isSaving
                    ? 'bg-emerald-900 text-emerald-200 cursor-wait'
                    : 'bg-emerald-700 hover:bg-emerald-600 text-white cursor-pointer'
                }`}
                title="Salvar importações e verificar alertas"
              >
                {isSaving ? (
                  <>
                    <Loader2 className="w-4 h-4 animate-spin text-emerald-200" />
                    <span>Salvando Importações...</span>
                  </>
                ) : justSaved ? (
                  <>
                    <Check className="w-4 h-4 text-emerald-300" />
                    <span>Importações Salvas!</span>
                  </>
                ) : (
                  <>
                    <Save className="w-4 h-4" />
                    <span>Salvar Importações</span>
                  </>
                )}
              </button>
            )}

            {metrics.isProcessed && (
              <button
                onClick={() => onNavigateToTab('painel')}
                className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-white font-bold text-xs rounded-xl transition flex items-center gap-2 cursor-pointer shadow-md"
              >
                <span>Ver Painel Geral</span>
                <ArrowRight className="w-4 h-4" />
              </button>
            )}
          </div>
        </div>
      </div>

      {/* Category Tabs Switcher */}
      <div className="flex flex-wrap gap-2 border-b border-slate-200 pb-3">
        <button
          onClick={() => setActiveImportCategory('fases')}
          className={`px-4 py-2.5 rounded-xl text-xs font-black transition flex items-center gap-2 cursor-pointer ${
            activeImportCategory === 'fases'
              ? 'bg-emerald-800 text-white shadow-xs'
              : 'bg-white text-slate-700 hover:bg-slate-100 border border-slate-200'
          }`}
        >
          <Layers className="w-4 h-4" />
          <span>Fases da Despesa (PPs e Liquidações)</span>
          <span className={`px-2 py-0.5 rounded-full text-[10px] ${
            activeImportCategory === 'fases' ? 'bg-emerald-900 text-emerald-100' : 'bg-slate-200 text-slate-700'
          }`}>
            {files.length}
          </span>
        </button>

        <button
          onClick={() => setActiveImportCategory('ordens_bancarias')}
          className={`px-4 py-2.5 rounded-xl text-xs font-black transition flex items-center gap-2 cursor-pointer ${
            activeImportCategory === 'ordens_bancarias'
              ? 'bg-emerald-800 text-white shadow-xs'
              : 'bg-white text-slate-700 hover:bg-slate-100 border border-slate-200'
          }`}
        >
          <FileCheck2 className="w-4 h-4" />
          <span>Detalhamento de Pagamentos</span>
          {ordensBancariasCount > 0 && (
            <span className="px-2 py-0.5 rounded-full text-[10px] bg-emerald-500 text-white">
              {ordensBancariasCount} reg.
            </span>
          )}
        </button>

        <button
          onClick={() => setActiveImportCategory('pagamentos_emitidos')}
          className={`px-4 py-2.5 rounded-xl text-xs font-black transition flex items-center gap-2 cursor-pointer ${
            activeImportCategory === 'pagamentos_emitidos'
              ? 'bg-emerald-800 text-white shadow-xs'
              : 'bg-white text-slate-700 hover:bg-slate-100 border border-slate-200'
          }`}
        >
          <FileSpreadsheet className="w-4 h-4" />
          <span>Pagamentos Emitidos (Quebra de Ordem)</span>
          {pagamentosCount > 0 && (
            <span className="px-2 py-0.5 rounded-full text-[10px] bg-emerald-500 text-white">
              {pagamentosCount} reg.
            </span>
          )}
        </button>
      </div>

      {/* VIEW 1: Fases da Despesa */}
      {activeImportCategory === 'fases' && (
        <div className="space-y-8">
          <UploadSection
            files={files}
            onFileUpload={onFileUpload}
            onRemoveFile={onRemoveFile}
            onChangeFileStage={onChangeFileStage}
            onProcessData={() => {
              onProcessData();
              setTimeout(() => onNavigateToTab('painel'), 300);
            }}
            onOpenColumnMapModal={onOpenColumnMapModal}
            isProcessing={isProcessing}
          />

          {/* Sheet Row Data Preview Inspector */}
          <div className="bg-white rounded-2xl p-6 shadow-xs border border-slate-200/90 space-y-4">
            <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-slate-100 pb-4">
              <div>
                <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
                  <TableIcon className="w-5 h-5 text-emerald-600" />
                  <span>Inspeção Prévia de Planilhas</span>
                </h3>
              </div>

              {/* File selector tabs for preview */}
              {files.length > 0 && (
                <div className="flex items-center gap-1.5 overflow-x-auto bg-slate-100 p-1.5 rounded-xl">
                  {files.map((f) => {
                    const isSelected = activePreviewFile?.id === f.id;
                    return (
                      <button
                        key={f.id}
                        onClick={() => setSelectedFileId(f.id)}
                        className={`px-3 py-1.5 rounded-lg text-xs font-bold transition flex items-center gap-1.5 cursor-pointer shrink-0 ${
                          isSelected
                            ? 'bg-white text-slate-900 shadow-xs'
                            : 'text-slate-600 hover:text-slate-900'
                        }`}
                      >
                        <span className="w-2 h-2 rounded-full bg-emerald-500" />
                        <span className="truncate max-w-[120px]">{f.fileName}</span>
                      </button>
                    );
                  })}
                </div>
              )}
            </div>

            {/* Preview Table */}
            {activePreviewFile ? (
              <div className="space-y-3">
                <div className="flex flex-wrap items-center justify-between gap-2 text-xs text-slate-600 bg-slate-50 p-3 rounded-xl border border-slate-200 font-medium">
                  <div>
                    <strong>Arquivo:</strong> {activePreviewFile.fileName}
                  </div>
                  <div>
                    <strong>Registros:</strong> {activePreviewFile.rows.length}
                  </div>
                  <div>
                    <strong>Coluna Favorecido:</strong>{' '}
                    <span className="text-emerald-700 font-bold">{activePreviewFile.detectedFavorecidoCol}</span>
                  </div>
                  <div>
                    <strong>Coluna Valor R$:</strong>{' '}
                    <span className="text-emerald-700 font-bold">{activePreviewFile.detectedValorCol}</span>
                  </div>
                  <button
                    onClick={() => onOpenColumnMapModal(activePreviewFile.id)}
                    className="text-xs text-emerald-700 font-bold underline hover:text-emerald-900 cursor-pointer"
                  >
                    Ajustar Mapeamento
                  </button>
                </div>

                <div className="overflow-x-auto border border-slate-200 rounded-xl max-h-80">
                  <table className="w-full text-left text-xs border-collapse">
                    <thead className="bg-slate-100 text-slate-700 font-bold uppercase tracking-wider sticky top-0 border-b border-slate-200">
                      <tr>
                        <th className="py-2.5 px-3 border-r border-slate-200 w-12 text-center">#</th>
                        {activePreviewFile.headers.slice(0, 8).map((col, cIdx) => {
                          const isFavCol = col === activePreviewFile.detectedFavorecidoCol;
                          const isValCol = col === activePreviewFile.detectedValorCol;

                          return (
                            <th
                              key={`th_${col}_${cIdx}`}
                              className={`py-2.5 px-3 border-r border-slate-200 ${
                                isFavCol
                                  ? 'bg-emerald-100/80 text-emerald-900 font-black'
                                  : isValCol
                                  ? 'bg-emerald-100/80 text-emerald-900 font-black'
                                  : ''
                              }`}
                            >
                              <div className="flex items-center gap-1">
                                <span>{col}</span>
                                {isFavCol && <span className="text-[9px] bg-emerald-700 text-white px-1 rounded">Credor</span>}
                                {isValCol && <span className="text-[9px] bg-emerald-800 text-white px-1 rounded">Valor</span>}
                              </div>
                            </th>
                          );
                        })}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 font-mono">
                      {activePreviewFile.rows.slice(0, 15).map((row, idx) => (
                        <tr key={idx} className="hover:bg-slate-50">
                          <td className="py-2 px-3 text-slate-400 border-r border-slate-100 text-[10px] text-center">
                            {idx + 1}
                          </td>
                          {activePreviewFile.headers.slice(0, 8).map((col, cIdx) => (
                            <td
                              key={`td_${col}_${cIdx}`}
                              className="py-2 px-3 border-r border-slate-100 truncate max-w-xs text-slate-700"
                            >
                              {String(row[col] ?? '')}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <p className="text-[11px] text-slate-400 italic">
                  Exibindo as primeiras 15 linhas de {activePreviewFile.rows.length} registros no total.
                </p>
              </div>
            ) : (
              <div className="py-12 text-center text-slate-400 space-y-2">
                <HelpCircle className="w-10 h-10 mx-auto text-slate-300" />
                <p className="text-xs font-semibold text-slate-500">
                  Nenhuma planilha carregada para inspeção.
                </p>
                <p className="text-[11px] text-slate-400">
                  Utilize o campo de importação acima para selecionar seus arquivos Excel ou CSV.
                </p>
              </div>
            )}
          </div>
        </div>
      )}

      {/* VIEW 2: Detalhamento de Pagamentos (Listar Ordem Bancária) */}
      {activeImportCategory === 'ordens_bancarias' && (
        <div className="bg-white rounded-2xl p-6 shadow-xs border border-slate-200/90 space-y-6">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-100 pb-4">
            <div>
              <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
                <FileCheck2 className="w-5 h-5 text-emerald-600" />
                <span>Detalhamento de Pagamentos</span>
              </h3>
              <p className="text-xs text-slate-500 mt-1">
                Relatório SIGEF em que cada registro ocupa duas linhas consecutivas (Linha 1: OB, Data, Valor, Situação; Linha 2: PP, Fonte, Favorecido). Alimenta a coluna "Detalhamento de Pagamentos" com detalhamento por Fonte de Recurso e filtro por Mês de Pagamento.
              </p>
            </div>

            {ordensBancariasList && ordensBancariasList.length > 0 && (
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={handleClearOrdensBancarias}
                  className="p-2 text-rose-600 hover:bg-rose-50 border border-rose-200 rounded-xl transition cursor-pointer"
                  title="Remover arquivo de ordens bancárias"
                >
                  <Trash2 className="w-4 h-4" />
                </button>
              </div>
            )}
          </div>

          <input
            type="file"
            ref={ordensBancariasInputRef}
            onChange={handleOrdensBancariasUpload}
            accept=".xlsx, .xls, .csv"
            className="hidden"
          />

          {/* Status if file is already loaded */}
          {ordensBancariasList && ordensBancariasList.length > 0 ? (
            <div className="p-5 bg-emerald-50/60 border border-emerald-200 rounded-2xl flex flex-col md:flex-row md:items-center justify-between gap-4">
              <div className="flex items-start space-x-3.5">
                <div className="p-2.5 bg-emerald-600 text-white rounded-xl shadow-xs">
                  <FileCheck2 className="w-6 h-6" />
                </div>
                <div>
                  <div className="flex items-center gap-2">
                    <span className="font-bold text-slate-900 text-sm">
                      {ordensBancariasFileName || 'Relatório Listar Ordem Bancária'}
                    </span>
                    <span className="px-2 py-0.5 rounded-full text-[10px] font-extrabold bg-emerald-200 text-emerald-900">
                      Ativo
                    </span>
                  </div>
                  <p className="text-xs text-slate-600 mt-1">
                    <strong>{ordensBancariasCount}</strong> ordens bancárias conclusivas (ex: Situação CB) lidas • Montante Total: <strong>{formatBRL(ordensBancariasTotalValor)}</strong>
                  </p>
                </div>
              </div>

              <div className="flex flex-wrap items-center gap-2">
                <button
                  type="button"
                  onClick={() => onNavigateToTab('pagamentos_fonte')}
                  className="px-4 py-2 bg-teal-700 hover:bg-teal-800 text-white text-xs font-bold rounded-xl transition flex items-center gap-1.5 cursor-pointer shadow-xs"
                >
                  <span>Abrir Consulta de Pagamentos por Fonte</span>
                  <ExternalLink className="w-3.5 h-3.5" />
                </button>
                <button
                  type="button"
                  onClick={() => ordensBancariasInputRef.current?.click()}
                  className="px-3 py-2 bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 text-xs font-bold rounded-xl transition cursor-pointer"
                >
                  Substituir
                </button>
                <button
                  type="button"
                  onClick={handleClearOrdensBancarias}
                  className="p-2 text-rose-600 hover:bg-rose-50 border border-rose-200 rounded-xl transition cursor-pointer"
                  title="Remover arquivo"
                >
                  <Trash2 className="w-4 h-4" />
                </button>
              </div>
            </div>
          ) : (
            <div className="space-y-4">
              {/* Upload Dropzone */}
              <div
                onDragOver={(e) => e.preventDefault()}
                onDrop={handleOrdensBancariasUpload}
                onClick={() => ordensBancariasInputRef.current?.click()}
                className="border-2 border-dashed border-slate-300 hover:border-emerald-500 bg-slate-50/60 hover:bg-emerald-50/20 rounded-2xl p-8 text-center cursor-pointer transition flex flex-col items-center justify-center space-y-3"
              >
                <div className="p-3 bg-emerald-100 text-emerald-800 rounded-2xl">
                  <Upload className="w-6 h-6" />
                </div>
                <div className="space-y-1">
                  <p className="text-xs font-bold text-slate-800">
                    {isLoadingOrdensBancarias
                      ? 'Lendo relatório e combinando linhas pares...'
                      : 'Arraste e solte o Relatório "Listar Ordem Bancária" aqui'}
                  </p>
                  <p className="text-[11px] text-slate-400">
                    ou clique para selecionar do computador (.xlsx, .xls, .csv)
                  </p>
                </div>
              </div>
            </div>
          )}
        </div>
      )}

      {/* VIEW 3: Pagamentos Emitidos (Quebra de Ordem) */}
      {activeImportCategory === 'pagamentos_emitidos' && (
        <div className="bg-white rounded-2xl p-6 shadow-xs border border-slate-200/90 space-y-6">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-100 pb-4">
            <div>
              <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
                <FileSpreadsheet className="w-5 h-5 text-emerald-600" />
                <span>Importação: Relatório de Pagamentos Emitidos</span>
              </h3>
              <p className="text-xs text-slate-500 mt-1">
                Planilha contendo as PPs/OBs emitidas. O sistema filtrará situação <strong>"AO"</strong> e ignorará Fontes <strong>.600</strong>.
              </p>
            </div>

            {rawPagamentosList && rawPagamentosList.length > 0 && (
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={handleClearPagamentos}
                  className="p-2 text-rose-600 hover:bg-rose-50 border border-rose-200 rounded-xl transition cursor-pointer"
                  title="Remover arquivo de pagamentos"
                >
                  <Trash2 className="w-4 h-4" />
                </button>
              </div>
            )}
          </div>

          <input
            type="file"
            ref={pagamentosInputRef}
            onChange={handlePagamentosUpload}
            accept=".xlsx, .xls, .csv"
            className="hidden"
          />

          {/* Status if file is already loaded */}
          {rawPagamentosList && rawPagamentosList.length > 0 ? (
            <div className="p-5 bg-emerald-50/60 border border-emerald-200 rounded-2xl flex flex-col md:flex-row md:items-center justify-between gap-4">
              <div className="flex items-start space-x-3.5">
                <div className="p-2.5 bg-emerald-600 text-white rounded-xl shadow-xs">
                  <FileCheck2 className="w-6 h-6" />
                </div>
                <div>
                  <div className="flex items-center gap-2">
                    <span className="font-bold text-slate-900 text-sm">
                      {pagamentosFileName || 'Relatório de Pagamentos Emitidos'}
                    </span>
                    <span className="px-2 py-0.5 rounded-full text-[10px] font-extrabold bg-emerald-200 text-emerald-900">
                      Ativo
                    </span>
                  </div>
                  <p className="text-xs text-slate-600 mt-1">
                    <strong>{pagamentosCount}</strong> registros lidos • <strong>{pagamentosAoCount}</strong> Ordens Bancárias em situação AO aptas para cruzamento.
                  </p>
                </div>
              </div>

              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => onNavigateToTab('quebra_ordem')}
                  className="px-4 py-2 bg-emerald-700 hover:bg-emerald-800 text-white text-xs font-bold rounded-xl transition flex items-center gap-1.5 cursor-pointer shadow-xs"
                >
                  <span>Ver Cruzamento em Quebra de Ordem</span>
                  <ExternalLink className="w-3.5 h-3.5" />
                </button>
                <button
                  type="button"
                  onClick={() => pagamentosInputRef.current?.click()}
                  className="px-3 py-2 bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 text-xs font-bold rounded-xl transition cursor-pointer"
                >
                  Substituir
                </button>
                <button
                  type="button"
                  onClick={handleClearPagamentos}
                  className="p-2 text-rose-600 hover:bg-rose-50 border border-rose-200 rounded-xl transition cursor-pointer"
                  title="Remover arquivo"
                >
                  <Trash2 className="w-4 h-4" />
                </button>
              </div>
            </div>
          ) : (
            <div className="space-y-4">
              {workspaceAoData.records.length > 0 && (
                <div className="bg-emerald-50 border border-emerald-300 rounded-xl p-4 flex flex-col sm:flex-row items-center justify-between gap-3 shadow-2xs">
                  <div className="flex items-center gap-3">
                    <div className="p-2 bg-emerald-600 text-white rounded-lg shrink-0">
                      <Sparkles className="w-4 h-4" />
                    </div>
                    <div>
                      <h4 className="text-xs font-bold text-slate-900">
                        Detectamos {workspaceAoData.records.length} PPs Prontas (AO) no Relatório SIGEF
                      </h4>
                      <p className="text-[11px] text-slate-600 mt-0.5">
                        Origem: <strong>{workspaceAoData.sourceFileName}</strong>. Você pode carregar essas PPs Prontas diretamente para a análise de Quebra de Ordem.
                      </p>
                    </div>
                  </div>
                  <button
                    type="button"
                    onClick={() => {
                      if (onUpdatePagamentosList) {
                        onUpdatePagamentosList(workspaceAoData.records, workspaceAoData.sourceFileName);
                      }
                    }}
                    className="px-3.5 py-2 bg-emerald-700 hover:bg-emerald-600 text-white font-bold text-xs rounded-xl transition flex items-center gap-1.5 cursor-pointer shrink-0 shadow-xs"
                  >
                    <CheckCircle2 className="w-3.5 h-3.5" />
                    <span>Usar PPs do SIGEF</span>
                  </button>
                </div>
              )}

              {/* Upload Dropzone */}
              <div
                onDragOver={(e) => e.preventDefault()}
                onDrop={handlePagamentosUpload}
                onClick={() => pagamentosInputRef.current?.click()}
                className="border-2 border-dashed border-slate-300 hover:border-emerald-500 bg-slate-50/60 hover:bg-emerald-50/20 rounded-2xl p-8 text-center cursor-pointer transition flex flex-col items-center justify-center space-y-3"
              >
                <div className="p-3 bg-emerald-100 text-emerald-800 rounded-2xl">
                  <Upload className="w-6 h-6" />
                </div>
                <div className="space-y-1">
                  <p className="text-xs font-bold text-slate-800">
                    {isLoadingPagamentos
                      ? 'Lendo e filtrando relatório de pagamentos...'
                      : 'Arraste e solte o Relatório de Pagamentos Emitidos aqui'}
                  </p>
                  <p className="text-[11px] text-slate-400">
                    ou clique para selecionar do computador (.xlsx, .xls, .csv)
                  </p>
                </div>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
};
