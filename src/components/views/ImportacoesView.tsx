import React, { useState, useMemo } from 'react';
import {
  Table as TableIcon,
  ArrowRight,
  HelpCircle,
  Save,
  Loader2,
  Check,
  Settings2,
  ChevronLeft,
  ChevronRight,
  Search,
  X,
} from 'lucide-react';
import { FileData, DashboardMetrics, StageKey } from '../../types';
import { UploadSection } from '../UploadSection';
import { OrdemBancariaItem } from '../../utils/ordemBancariaParser';

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
}) => {
  const [isSaving, setIsSaving] = useState(false);
  const [justSaved, setJustSaved] = useState(false);

  // Estados do Seletor e Paginação da Pré-visualização
  const [selectedPreviewFileId, setSelectedPreviewFileId] = useState<string>('');
  const [previewSearch, setPreviewSearch] = useState<string>('');
  const [previewPage, setPreviewPage] = useState<number>(1);
  const [previewPageSize, setPreviewPageSize] = useState<number>(50);

  const activePreviewFile = useMemo(() => {
    if (!files || files.length === 0) return null;
    return files.find((f) => f.id === selectedPreviewFileId) || files[0] || null;
  }, [files, selectedPreviewFileId]);

  // Linhas filtradas pela busca rápida
  const filteredRows = useMemo(() => {
    if (!activePreviewFile || !activePreviewFile.rows) return [];
    if (!previewSearch.trim()) return activePreviewFile.rows;
    const term = previewSearch.trim().toLowerCase();
    return activePreviewFile.rows.filter((row) => {
      return Object.values(row).some(
        (val) => val !== null && val !== undefined && String(val).toLowerCase().includes(term)
      );
    });
  }, [activePreviewFile, previewSearch]);

  // Cálculo de paginação
  const totalRows = filteredRows.length;
  const totalPages = Math.max(1, Math.ceil(totalRows / previewPageSize));
  const currentPage = Math.min(Math.max(1, previewPage), totalPages);
  const startIndex = (currentPage - 1) * previewPageSize;
  const endIndex = Math.min(startIndex + previewPageSize, totalRows);
  const paginatedRows = useMemo(() => {
    return filteredRows.slice(startIndex, endIndex);
  }, [filteredRows, startIndex, endIndex]);

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

  return (
    <div className="space-y-8 pb-12">
      {/* Banner / Title Header */}
      <div className="bg-gradient-to-r from-slate-900 via-emerald-950 to-slate-900 rounded-2xl p-6 md:p-8 text-white shadow-xl relative overflow-hidden border border-slate-800">
        <div className="relative z-10 flex flex-col md:flex-row md:items-center justify-between gap-6">
          <div className="flex items-center gap-3">
            <h2 className="text-2xl font-black text-white tracking-tight">
              Central Única de Importações
            </h2>
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

      {/* Central Unified Upload Section */}
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
                <span>Inspeção Prévia e Mapeamento de Colunas</span>
              </h3>
            </div>

            {/* Seletor dinâmico de planilha / abas importadas */}
            {files.length > 1 ? (
              <div className="flex items-center gap-2 bg-slate-50 border border-slate-200 p-1.5 rounded-xl text-xs">
                <span className="w-2 h-2 rounded-full bg-emerald-500 shrink-0 ml-1.5" />
                <label htmlFor="preview-file-select" className="font-bold text-slate-700 shrink-0">
                  Planilha:
                </label>
                <select
                  id="preview-file-select"
                  value={activePreviewFile?.id || ''}
                  onChange={(e) => {
                    setSelectedPreviewFileId(e.target.value);
                    setPreviewPage(1);
                    setPreviewSearch('');
                  }}
                  className="font-semibold bg-white border border-slate-300 rounded-lg px-2.5 py-1 text-slate-800 focus:outline-none focus:ring-2 focus:ring-emerald-500 max-w-[260px] md:max-w-[360px] truncate cursor-pointer shadow-2xs"
                >
                  {files.map((f, idx) => (
                    <option key={f.id} value={f.id}>
                      {idx + 1}. {f.fileName} ({f.rows.length} registros)
                    </option>
                  ))}
                </select>
              </div>
            ) : activePreviewFile ? (
              <div className="flex items-center gap-2 bg-slate-100 px-3 py-1.5 rounded-xl border border-slate-200 text-xs">
                <span className="w-2 h-2 rounded-full bg-emerald-500 shrink-0" />
                <span className="font-bold text-slate-700">Arquivo de Referência:</span>
                <span className="text-slate-900 font-medium truncate max-w-[240px]" title={activePreviewFile.fileName}>
                  {activePreviewFile.fileName}
                </span>
              </div>
            ) : null}
          </div>

          {/* Preview Table e Controles */}
          {activePreviewFile ? (
            <div className="space-y-4">
              {/* Badges de Colunas Mapeadas */}
              <div className="flex flex-wrap items-center justify-between gap-2 text-xs text-slate-600 bg-slate-50 p-3 rounded-xl border border-slate-200 font-medium">
                <div>
                  <strong>Arquivo:</strong> {activePreviewFile.fileName}
                </div>
                <div>
                  <strong>Total na Planilha:</strong> {activePreviewFile.rows.length}
                </div>
                {activePreviewFile.detectedObCol && (
                  <div>
                    <strong>Coluna OB:</strong>{' '}
                    <span className="text-slate-900 font-bold">{activePreviewFile.detectedObCol}</span>
                  </div>
                )}
                {activePreviewFile.detectedFonteCol && (
                  <div>
                    <strong>Coluna Fonte:</strong>{' '}
                    <span className="text-teal-700 font-bold">{activePreviewFile.detectedFonteCol}</span>
                  </div>
                )}
                <div>
                  <strong>Coluna Favorecido:</strong>{' '}
                  <span className="text-emerald-700 font-bold">{activePreviewFile.detectedFavorecidoCol || 'Automático'}</span>
                </div>
                <div>
                  <strong>Coluna Valor R$:</strong>{' '}
                  <span className="text-emerald-700 font-bold">{activePreviewFile.detectedValorCol || 'Automático'}</span>
                </div>
                {activePreviewFile.detectedPpCol && (
                  <div>
                    <strong>Coluna PP:</strong>{' '}
                    <span className="text-blue-700 font-bold">{activePreviewFile.detectedPpCol}</span>
                  </div>
                )}
                {activePreviewFile.detectedSituacaoCol && (
                  <div>
                    <strong>Coluna Situação:</strong>{' '}
                    <span className="text-amber-700 font-bold">{activePreviewFile.detectedSituacaoCol}</span>
                  </div>
                )}
                <button
                  onClick={() => onOpenColumnMapModal(activePreviewFile.id)}
                  className="px-3 py-1 bg-emerald-700 hover:bg-emerald-800 text-white rounded-lg text-xs font-bold flex items-center gap-1 cursor-pointer transition shadow-xs"
                >
                  <Settings2 className="w-3.5 h-3.5" />
                  <span>Ajustar Mapeamento</span>
                </button>
              </div>

              {/* Barra de Busca Rápida e Paginação Superior */}
              <div className="flex flex-col md:flex-row md:items-center justify-between gap-3 text-xs bg-slate-50/80 p-3 rounded-xl border border-slate-200">
                <div className="flex items-center gap-2 flex-1 max-w-md relative">
                  <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none" />
                  <input
                    type="text"
                    value={previewSearch}
                    onChange={(e) => {
                      setPreviewSearch(e.target.value);
                      setPreviewPage(1);
                    }}
                    placeholder="Filtrar dados nesta planilha..."
                    className="w-full pl-9 pr-8 py-1.5 text-xs bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-emerald-500 text-slate-800 font-medium"
                  />
                  {previewSearch && (
                    <button
                      onClick={() => {
                        setPreviewSearch('');
                        setPreviewPage(1);
                      }}
                      className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 cursor-pointer"
                      title="Limpar busca"
                    >
                      <X className="w-3.5 h-3.5" />
                    </button>
                  )}
                </div>

                <div className="flex flex-wrap items-center gap-3 shrink-0">
                  <div className="flex items-center gap-1.5 text-slate-600">
                    <span className="text-slate-500 font-medium">Por página:</span>
                    <select
                      value={previewPageSize}
                      onChange={(e) => {
                        setPreviewPageSize(Number(e.target.value));
                        setPreviewPage(1);
                      }}
                      className="bg-white border border-slate-300 rounded-lg px-2 py-1 text-xs font-bold text-slate-700 focus:outline-none focus:ring-2 focus:ring-emerald-500 cursor-pointer"
                    >
                      <option value={50}>50</option>
                      <option value={100}>100</option>
                      <option value={250}>250</option>
                      <option value={500}>500</option>
                    </select>
                  </div>

                  <div className="text-slate-600 font-semibold">
                    {totalRows > 0 ? (
                      <span>
                        Exibindo <span className="font-mono font-bold text-slate-900">{startIndex + 1}–{endIndex}</span> de <span className="font-mono font-bold text-slate-900">{totalRows}</span>
                        {previewSearch && ' (filtrados)'}
                      </span>
                    ) : (
                      <span className="text-slate-400">0 registros</span>
                    )}
                  </div>

                  <div className="flex items-center gap-1">
                    <button
                      onClick={() => setPreviewPage((p) => Math.max(1, p - 1))}
                      disabled={currentPage <= 1}
                      className="p-1 rounded-lg border border-slate-300 bg-white hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed text-slate-700 transition"
                      title="Página Anterior"
                    >
                      <ChevronLeft className="w-4 h-4" />
                    </button>
                    <span className="px-2 font-mono font-bold text-slate-800">
                      {currentPage} / {totalPages}
                    </span>
                    <button
                      onClick={() => setPreviewPage((p) => Math.min(totalPages, p + 1))}
                      disabled={currentPage >= totalPages}
                      className="p-1 rounded-lg border border-slate-300 bg-white hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed text-slate-700 transition"
                      title="Próxima Página"
                    >
                      <ChevronRight className="w-4 h-4" />
                    </button>
                  </div>
                </div>
              </div>

              {/* Tabela com TODAS as colunas renderizadas e scroll horizontal fluido */}
              <div className="overflow-x-auto border border-slate-200 rounded-xl max-h-[460px] shadow-2xs">
                <table className="w-full text-left text-xs border-collapse">
                  <thead className="bg-slate-100 text-slate-700 font-bold uppercase tracking-wider sticky top-0 z-10 border-b border-slate-200">
                    <tr>
                      <th className="py-2.5 px-3 border-r border-slate-200 w-14 text-center sticky left-0 bg-slate-100 z-20">#</th>
                      {(activePreviewFile.headers || []).map((col, cIdx) => {
                        const isObCol = col === activePreviewFile.detectedObCol;
                        const isFonteCol = col === activePreviewFile.detectedFonteCol;
                        const isFavCol = col === activePreviewFile.detectedFavorecidoCol;
                        const isValCol = col === activePreviewFile.detectedValorCol;
                        const isPpCol = col === activePreviewFile.detectedPpCol;
                        const isSitCol = col === activePreviewFile.detectedSituacaoCol;
                        const isHighlighted = isObCol || isFonteCol || isFavCol || isValCol || isPpCol || isSitCol;

                        return (
                          <th
                            key={`th_${col}_${cIdx}`}
                            className={`py-2.5 px-3 border-r border-slate-200 whitespace-nowrap ${
                              isHighlighted ? 'bg-emerald-100/90 text-emerald-950 font-black' : ''
                            }`}
                          >
                            <div className="flex items-center gap-1.5">
                              <span>{col}</span>
                              {isObCol && <span className="text-[9px] bg-slate-800 text-white px-1.5 py-0.5 rounded font-mono">OB</span>}
                              {isFonteCol && <span className="text-[9px] bg-teal-800 text-white px-1.5 py-0.5 rounded font-mono">Fonte</span>}
                              {isFavCol && <span className="text-[9px] bg-emerald-700 text-white px-1.5 py-0.5 rounded font-mono">Credor</span>}
                              {isValCol && <span className="text-[9px] bg-emerald-800 text-white px-1.5 py-0.5 rounded font-mono">Valor</span>}
                              {isPpCol && <span className="text-[9px] bg-blue-700 text-white px-1.5 py-0.5 rounded font-mono">PP</span>}
                              {isSitCol && <span className="text-[9px] bg-amber-700 text-white px-1.5 py-0.5 rounded font-mono">Situação</span>}
                            </div>
                          </th>
                        );
                      })}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 font-mono">
                    {paginatedRows.length === 0 ? (
                      <tr>
                        <td
                          colSpan={(activePreviewFile.headers?.length || 0) + 1}
                          className="py-8 text-center text-slate-400 italic font-sans"
                        >
                          Nenhum registro encontrado para o termo pesquisado.
                        </td>
                      </tr>
                    ) : (
                      paginatedRows.map((row, idx) => {
                        const rowNum = startIndex + idx + 1;
                        return (
                          <tr key={`row_${rowNum}_${idx}`} className="hover:bg-slate-50/80 transition-colors">
                            <td className="py-2 px-3 text-slate-400 border-r border-slate-100 text-[10px] text-center sticky left-0 bg-white hover:bg-slate-50 font-sans">
                              {rowNum}
                            </td>
                            {(activePreviewFile.headers || []).map((col, cIdx) => (
                              <td
                                key={`td_${col}_${cIdx}`}
                                className="py-2 px-3 border-r border-slate-100 whitespace-nowrap text-slate-700 max-w-sm truncate"
                                title={String(row[col] ?? '')}
                              >
                                {String(row[col] ?? '')}
                              </td>
                            ))}
                          </tr>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>

              {/* Paginação Inferior e Resumo */}
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs text-slate-500 pt-1">
                <p className="italic">
                  Visualizando todas as {activePreviewFile.headers?.length || 0} colunas da planilha selecionada ({activePreviewFile.rows.length} linhas totais).
                </p>

                {totalPages > 1 && (
                  <div className="flex items-center gap-2 self-end sm:self-auto">
                    <button
                      onClick={() => setPreviewPage((p) => Math.max(1, p - 1))}
                      disabled={currentPage <= 1}
                      className="px-3 py-1 rounded-lg border border-slate-300 bg-white hover:bg-slate-50 disabled:opacity-40 disabled:cursor-not-allowed text-slate-700 font-semibold transition"
                    >
                      Anterior
                    </button>
                    <span className="font-mono font-bold text-slate-700">
                      Página {currentPage} de {totalPages}
                    </span>
                    <button
                      onClick={() => setPreviewPage((p) => Math.min(totalPages, p + 1))}
                      disabled={currentPage >= totalPages}
                      className="px-3 py-1 rounded-lg border border-slate-300 bg-white hover:bg-slate-50 disabled:opacity-40 disabled:cursor-not-allowed text-slate-700 font-semibold transition"
                    >
                      Próxima
                    </button>
                  </div>
                )}
              </div>
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
    </div>
  );
};
