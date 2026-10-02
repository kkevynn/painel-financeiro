import React, { useState, useEffect, useRef, useCallback, startTransition } from 'react';
import { FileSpreadsheet, ArrowRight } from 'lucide-react';
import { FileData, DashboardMetrics, StageKey, ColumnMappingConfig } from './types';
import {
  parseExcelFile,
  processUploadedFiles,
  generateSampleFiles,
  detectStageFromFileName,
} from './utils/excelParser';
import { downloadStandaloneHtmlFile } from './utils/standaloneExporter';
import { loadAppState, saveAppState, clearAppState } from './utils/storageService';

import { Sidebar } from './components/Sidebar';
import { Header } from './components/Header';
import { KpiCards } from './components/KpiCards';
import { DetailedConsolidationTable } from './components/DetailedConsolidationTable';
import { ColumnMappingModal } from './components/ColumnMappingModal';

import { ImportacoesView } from './components/views/ImportacoesView';
import { CategoriasView } from './components/views/CategoriasView';
import { RecorrenciaView } from './components/views/RecorrenciaView';
import { FornecedoresMesView } from './components/views/FornecedoresMesView';
import { QuebraOrdemView } from './components/views/QuebraOrdemView';
import { PagamentosPorFonteView } from './components/views/PagamentosPorFonteView';
import { ExtratoAuditoriaEmpresaModal } from './components/ExtratoAuditoriaEmpresaModal';
import { ToastAlertContainer } from './components/ToastAlertContainer';
import { ModalAuditoriaML } from './components/ModalAuditoriaML';
import { ModalVerificacaoNotificacao } from './components/ModalVerificacaoNotificacao';
import { ModalHistoricoNotificacoes } from './components/ModalHistoricoNotificacoes';
import { toastAlertService } from './utils/toastAlertService';
import { mlInferenceService, getPendingInferences } from './utils/machineLearningInferenceService';
import { RawPagamentoEmitido, ToastAlert } from './types';
import {
  OrdemBancariaItem,
  enrichSuppliersWithOrdensBancarias,
  saveOrdensBancariasToStorage,
  loadOrdensBancariasFromStorage,
  loadOrdensBancariasAsync,
} from './utils/ordemBancariaParser';

const initialMetrics: DashboardMetrics = {
  grandTotalValue: 0,
  grandTotalCount: 0,
  totalPpProntas: 0,
  totalLiqObedece: 0,
  totalLiqNaoObedece: 0,
  totalOutros: 0,
  consolidatedSuppliers: [],
  isProcessed: false,
};

export function App() {
  const [activeTab, setActiveTab] = useState('painel');
  const [visitedTabs, setVisitedTabs] = useState<Record<string, boolean>>({ painel: true });
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);
  const [lastProcessedAt, setLastProcessedAt] = useState<Date | null>(null);

  const handleTabChange = useCallback((newTab: string) => {
    setVisitedTabs((prev) => (prev[newTab] ? prev : { ...prev, [newTab]: true }));
    startTransition(() => {
      setActiveTab(newTab);
    });
  }, []);

  // Array of loaded files
  const [files, setFiles] = useState<FileData[]>([]);

  // Consolidated dashboard metrics
  const [metrics, setMetrics] = useState<DashboardMetrics>(initialMetrics);

  const [isProcessing, setIsProcessing] = useState(false);
  const [processingMessage, setProcessingMessage] = useState<string>('Processando dados...');
  const [activeMappingFileId, setActiveMappingFileId] = useState<string | null>(null);
  const [isStorageLoaded, setIsStorageLoaded] = useState(false);

  // State for Relatório de Pagamentos Emitidos (Quebra de Ordem)
  const [rawPagamentosList, setRawPagamentosList] = useState<RawPagamentoEmitido[] | null>(null);
  const [pagamentosFileName, setPagamentosFileName] = useState<string>('');

  // State for Relatório "Listar Ordem Bancária" (Pagos / Detalhamento por Fonte e Mês)
  const [ordensBancariasList, setOrdensBancariasList] = useState<OrdemBancariaItem[] | null>(null);
  const [ordensBancariasFileName, setOrdensBancariasFileName] = useState<string>('');

  // Helper to re-enrich metrics whenever files or OB list changes
  const applyEnrichment = useCallback((baseMetrics: DashboardMetrics, obList: OrdemBancariaItem[] | null): DashboardMetrics => {
    if (!baseMetrics || !baseMetrics.consolidatedSuppliers || baseMetrics.consolidatedSuppliers.length === 0) {
      return baseMetrics;
    }
    if (!obList || obList.length === 0) {
      return baseMetrics;
    }
    enrichSuppliersWithOrdensBancarias(baseMetrics.consolidatedSuppliers, obList);
    return { ...baseMetrics };
  }, []);

  // Update ordens bancarias handler
  const handleUpdateOrdensBancariasList = (list: OrdemBancariaItem[] | null, fileName?: string) => {
    setOrdensBancariasList(list);
    if (fileName !== undefined) setOrdensBancariasFileName(fileName);
    saveOrdensBancariasToStorage(list, fileName);

    // Immediately re-enrich current metrics
    setMetrics((prev) => {
      const cloned = { ...prev, consolidatedSuppliers: [...prev.consolidatedSuppliers] };
      return applyEnrichment(cloned, list);
    });
    triggerAlertAnalysis(files, list);
  };

  // State for Extrato Geral de Auditoria da Empresa
  const [auditedSupplier, setAuditedSupplier] = useState<{ name: string; cnpj: string } | null>(null);

  // State for Machine Learning & AI Inferences Audit Modal
  const [isModalMLOpen, setIsModalMLOpen] = useState(false);
  const [pendingMLCount, setPendingMLCount] = useState(0);

  // State for Toast Notification Verification / Alteration Modal
  const [selectedToastForVerification, setSelectedToastForVerification] = useState<ToastAlert | null>(null);
  const [isModalNotificacoesOpen, setIsModalNotificacoesOpen] = useState(false);
  const [alertsCount, setAlertsCount] = useState(0);

  useEffect(() => {
    const updateAlerts = () => {
      // Unread alerts count from alerts history
      const history = toastAlertService.getAlertsHistory();
      const unresolved = history.filter((a) => !a.resolvido).length;
      setAlertsCount((prev) => (prev !== unresolved ? unresolved : prev));
    };

    updateAlerts();
    const unsubAlerts = toastAlertService.subscribe(updateAlerts);

    const updateML = () => {
      const count = mlInferenceService.getPendingMatchesCount();
      setPendingMLCount((prev) => (prev !== count ? count : prev));
    };
    updateML();
    const unsubML = mlInferenceService.subscribe(updateML);

    return () => {
      unsubAlerts();
      unsubML();
    };
  }, []);

  // Save alteration made directly via notification verification modal
  const handleSaveToastAlteration = (
    toast: ToastAlert,
    changes: { situacao: string; valor?: number; observacao?: string }
  ) => {
    const ppTarget = (toast.pp || '').trim().toLowerCase();
    const obTarget = (toast.ob || '').trim().toLowerCase();

    // Update all occurrences in loaded files
    const updatedFiles = files.map((file) => {
      const numCol = file.detectedNumeroCol || file.detectedPpCol;
      const obCol = file.detectedObCol;
      const valCol = file.detectedValorCol;
      const sitCol = file.detectedSituacaoCol;

      let fileModified = false;
      const updatedRows = file.rows.map((row) => {
        const rawNum = numCol ? String(row[numCol] || '').trim().toLowerCase() : '';
        const rawOb = obCol ? String(row[obCol] || '').trim().toLowerCase() : '';

        const matchesPP = ppTarget && rawNum.includes(ppTarget);
        const matchesOB = obTarget && rawOb.includes(obTarget);

        if (matchesPP || matchesOB) {
          fileModified = true;
          const newRow = { ...row };
          if (sitCol) {
            newRow[sitCol] = changes.situacao;
          }
          if (valCol && typeof changes.valor === 'number' && !isNaN(changes.valor)) {
            newRow[valCol] = changes.valor;
          }
          return newRow;
        }
        return row;
      });

      if (fileModified) {
        return {
          ...file,
          rows: updatedRows,
        };
      }
      return file;
    });

    hasUserChangesRef.current = true;
    setFiles(updatedFiles);
    const newMetrics = processUploadedFiles(updatedFiles);
    setMetrics(newMetrics);
    saveAppState(updatedFiles, lastProcessedAt || new Date());

    // Update alert status in service
    toastAlertService.updateToast(toast.id, {
      situacaoNova: changes.situacao,
      valor: typeof changes.valor === 'number' && !isNaN(changes.valor) ? changes.valor : toast.valor,
      observacao: changes.observacao,
      resolvido: true,
      dataHora: new Date(),
    });

    setSelectedToastForVerification((prev) =>
      prev && prev.id === toast.id
        ? {
            ...prev,
            situacaoNova: changes.situacao,
            valor: typeof changes.valor === 'number' && !isNaN(changes.valor) ? changes.valor : toast.valor,
            observacao: changes.observacao,
            resolvido: true,
          }
        : prev
    );
  };

  // Track if user performed real changes to avoid auto-saving on initial load
  const hasUserChangesRef = useRef(false);

  // Load saved state on mount
  useEffect(() => {
    async function initStorage() {
      try {
        const stored = await loadAppState();
        const obStored = await loadOrdensBancariasAsync();
        if (obStored.items && obStored.items.length > 0) {
          setOrdensBancariasList(obStored.items);
          setOrdensBancariasFileName(obStored.fileName);
        }

        const realFiles = (stored?.files || []).filter(
          (f) => f.id !== 'sample_file_1' && !f.fileName?.includes('Relatorio_Credores_Valor_Total.xlsx')
        );

        if (realFiles.length > 0) {
          setFiles(realFiles);
          const ts = stored?.lastProcessedAt ? new Date(stored.lastProcessedAt) : new Date();
          setLastProcessedAt(ts);
          const processed = processUploadedFiles(realFiles);
          const enriched = applyEnrichment(processed, obStored.items);
          setMetrics(enriched);
          toastAlertService.syncConfirmations(realFiles, obStored.items, false);
          mlInferenceService.processFilesForInference(realFiles, obStored.items);
        } else if (obStored.items && obStored.items.length > 0) {
          toastAlertService.syncConfirmations([], obStored.items, false);
          mlInferenceService.processFilesForInference([], obStored.items);
        } else {
          setFiles([]);
          const initial = {
            grandTotalValue: 0,
            grandTotalCount: 0,
            totalPpProntas: 0,
            totalLiqObedece: 0,
            totalLiqNaoObedece: 0,
            totalOutros: 0,
            consolidatedSuppliers: [],
            isProcessed: false,
          };
          setMetrics(initial);
        }
      } catch (err) {
        console.error('Erro ao carregar do armazenamento:', err);
      } finally {
        setIsStorageLoaded(true);
      }
    }
    initStorage();
  }, []);

  // Auto-save to storage ONLY when user actually makes changes (debounced to keep UI fluid)
  useEffect(() => {
    if (!isStorageLoaded) return;
    if (!hasUserChangesRef.current) return;

    const timer = setTimeout(() => {
      saveAppState(files, lastProcessedAt);
      hasUserChangesRef.current = false;
    }, 1000);
    return () => clearTimeout(timer);
  }, [files, lastProcessedAt, isStorageLoaded]);

  // Trigger alert and ML inference analysis whenever files are uploaded, changed, or processed
  const alertAnalysisTimeoutRef = useRef<any>(null);
  const triggerAlertAnalysis = (currentFiles: FileData[], currentOBs?: OrdemBancariaItem[] | null) => {
    if (alertAnalysisTimeoutRef.current) {
      clearTimeout(alertAnalysisTimeoutRef.current);
    }
    alertAnalysisTimeoutRef.current = setTimeout(() => {
      const targetOBs = currentOBs !== undefined ? currentOBs : ordensBancariasList;
      toastAlertService.analyzeFileTransitions(currentFiles, targetOBs);
      mlInferenceService.processFilesForInference(currentFiles, targetOBs);
    }, 250);
  };

  // Upload file action
  const handleFileUpload = async (file: File, stage?: StageKey) => {
    try {
      setIsProcessing(true);
      setProcessingMessage(`Lendo e processando "${file.name}"...`);
      // Brief pause to allow the animated spinner to appear cleanly
      await new Promise((res) => setTimeout(res, 80));
      const parsedFileData = await parseExcelFile(file);
      parsedFileData.stage = stage || detectStageFromFileName(file.name);

      hasUserChangesRef.current = true;
      setFiles((prev) => {
        const updated = [...prev, parsedFileData];
        const newMetrics = processUploadedFiles(updated);
        const enriched = applyEnrichment(newMetrics, ordensBancariasList);
        setMetrics(enriched);
        setLastProcessedAt(new Date());
        triggerAlertAnalysis(updated, ordensBancariasList);
        return updated;
      });
    } catch (error: any) {
      toastAlertService.addToast({
        tipo: 'critico',
        titulo: 'Erro ao carregar arquivo',
        mensagem: error.message || 'Arquivo corrompido ou formato inválido.',
      });
    } finally {
      setIsProcessing(false);
    }
  };

  // Change file stage tag
  const handleChangeFileStage = (fileId: string, stage: StageKey) => {
    hasUserChangesRef.current = true;
    setFiles((prev) => {
      const updated = prev.map((f) => (f.id === fileId ? { ...f, stage } : f));
      const newMetrics = processUploadedFiles(updated);
      const enriched = applyEnrichment(newMetrics, ordensBancariasList);
      setMetrics(enriched);
      triggerAlertAnalysis(updated, ordensBancariasList);
      return updated;
    });
  };

  // Remove file action
  const handleRemoveFile = (fileId: string) => {
    hasUserChangesRef.current = true;
    setFiles((prev) => {
      const updated = prev.filter((f) => f.id !== fileId);
      const newMetrics = processUploadedFiles(updated);
      const enriched = applyEnrichment(newMetrics, ordensBancariasList);
      setMetrics(enriched);
      return updated;
    });
  };

  // Process data button handler
  const handleProcessData = () => {
    setIsProcessing(true);
    setProcessingMessage('Consolidando planilhas e métricas gerais...');
    setTimeout(() => {
      try {
        const newMetrics = processUploadedFiles(files);
        const enriched = applyEnrichment(newMetrics, ordensBancariasList);
        setMetrics(enriched);
        const now = new Date();
        setLastProcessedAt(now);
        saveAppState(files, now);
        toastAlertService.analyzeFileTransitions(files, ordensBancariasList);
        mlInferenceService.processFilesForInference(files, ordensBancariasList);
      } catch (err: any) {
        toastAlertService.addToast({
          tipo: 'critico',
          titulo: 'Erro ao processar dados',
          mensagem: err.message || 'Falha ao processar e consolidar arquivos.',
        });
      } finally {
        setIsProcessing(false);
      }
    }, 450);
  };

  // Explicit Save Data handler
  const handleSaveData = async () => {
    setIsProcessing(true);
    setProcessingMessage('Salvando alterações e verificando transições...');
    try {
      const now = new Date();
      setLastProcessedAt(now);
      const newMetrics = processUploadedFiles(files);
      const enriched = applyEnrichment(newMetrics, ordensBancariasList);
      setMetrics(enriched);
      await saveAppState(files, now);
      // Scans for alerts strictly when user clicks on save button
      toastAlertService.analyzeFileTransitions(files, ordensBancariasList);
      mlInferenceService.processFilesForInference(files, ordensBancariasList);
      // Brief pause so the user visualizes the confirmation animation
      await new Promise((res) => setTimeout(res, 450));
    } finally {
      setIsProcessing(false);
    }
  };

  // Clear all data
  const handleClearData = async () => {
    setIsProcessing(true);
    setProcessingMessage('Limpando dados e redefinindo painel...');
    try {
      setFiles([]);
      setMetrics(initialMetrics);
      setLastProcessedAt(null);
      await clearAppState();
      await new Promise((res) => setTimeout(res, 350));
    } finally {
      setIsProcessing(false);
    }
  };

  // Save custom column mapping from modal
  const handleSaveColumnMapping = (
    fileId: string,
    mappingOrFav: ColumnMappingConfig | string,
    maybeValor?: string
  ) => {
    const mapping: ColumnMappingConfig = typeof mappingOrFav === 'string'
      ? { favorecidoCol: mappingOrFav, valorCol: maybeValor || '' }
      : mappingOrFav;

    hasUserChangesRef.current = true;
    setFiles((prev) => {
      const updated = prev.map((f) => {
        if (f.id === fileId) {
          return {
            ...f,
            detectedFavorecidoCol: mapping.favorecidoCol,
            detectedValorCol: mapping.valorCol,
            stage: mapping.stage || f.stage,
            detectedCnpjCol: mapping.cnpjCol || undefined,
            detectedSituacaoCol: mapping.situacaoCol || undefined,
            detectedNumeroCol: mapping.numeroCol || undefined,
            detectedPpCol: mapping.numeroCol || undefined,
            detectedObCol: mapping.obCol || undefined,
            detectedNotaEmpenhoCol: mapping.notaEmpenhoCol || undefined,
            detectedFavorecidoNECol: mapping.favorecidoNECol || undefined,
            detectedProcessoCol: mapping.processoCol || undefined,
            detectedFonteCol: mapping.fonteCol || undefined,
          };
        }
        return f;
      });
      const newMetrics = processUploadedFiles(updated);
      setMetrics(newMetrics);
      saveAppState(updated, lastProcessedAt || new Date());
      return updated;
    });
  };

  const activeMappingFile = files.find((f) => f.id === activeMappingFileId) || null;

  return (
    <div className="min-h-screen bg-[#f3f4f6] text-slate-800 font-sans flex flex-col overflow-x-clip antialiased">
      
      {/* Global Processing Animated Indicator */}
      {isProcessing && (
        <div className="fixed top-4 left-1/2 -translate-x-1/2 z-[99999] pointer-events-none animate-in fade-in slide-in-from-top-3 duration-200">
          <div className="flex items-center gap-3 px-5 py-2.5 bg-slate-900/95 text-white rounded-2xl shadow-2xl border border-emerald-500/40 backdrop-blur-md text-xs font-bold">
            <div className="w-4 h-4 border-2 border-emerald-400/30 border-t-emerald-400 rounded-full animate-spin shrink-0" />
            <span className="text-slate-100">{processingMessage || 'Processando dados...'}</span>
            <span className="flex h-2 w-2 relative">
              <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
              <span className="relative inline-flex rounded-full h-2 w-2 bg-emerald-500"></span>
            </span>
          </div>
        </div>
      )}

      {/* 1. Retractable Left Sidebar (Drawer) */}
      <Sidebar
        isOpen={isSidebarOpen}
        onClose={() => setIsSidebarOpen(false)}
        activeTab={activeTab}
        setActiveTab={handleTabChange}
      />

      {/* 2. Main Content Area (100% Width) */}
      <div className="flex-1 flex flex-col min-w-0 w-full min-h-screen">
        
        {/* Header */}
        <Header
          activeTab={activeTab}
          onToggleSidebar={() => setIsSidebarOpen((prev) => !prev)}
          onExportStandaloneHtml={downloadStandaloneHtmlFile}
          onClearData={handleClearData}
          onSaveData={handleSaveData}
          onOpenAuditoriaML={() => setIsModalMLOpen(true)}
          pendingInferencesCount={pendingMLCount}
          onOpenNotifications={() => setIsModalNotificacoesOpen(true)}
          unreadAlertsCount={alertsCount}
          isProcessed={metrics.isProcessed}
          lastProcessedAt={lastProcessedAt}
        />

        {/* Dynamic Workspace (Expanded Layout to 98% Screen Width) */}
        <main className="p-4 sm:p-6 md:p-8 space-y-8 w-full max-w-[98%] mx-auto flex-1">
          <div className="space-y-8">
            {visitedTabs['painel'] && (
              <div className={activeTab === 'painel' ? 'space-y-8 animate-in fade-in-50 duration-150' : 'hidden'}>
                {!metrics.isProcessed ? (
                  <div className="bg-[#061d15] border border-emerald-900/80 text-white rounded-3xl p-8 md:p-12 shadow-xl text-center space-y-6 max-w-3xl mx-auto my-8">
                    <div className="w-16 h-16 bg-emerald-600/20 text-emerald-400 rounded-2xl flex items-center justify-center mx-auto border border-emerald-500/30">
                      <FileSpreadsheet className="w-8 h-8" />
                    </div>
                    <div className="space-y-2">
                      <h2 className="text-2xl font-black tracking-tight text-white">
                        Nenhuma Planilha Importada
                      </h2>
                    </div>
                    <div className="flex flex-col sm:flex-row items-center justify-center gap-3 pt-2">
                      <button
                        onClick={() => handleTabChange('importacoes')}
                        className="w-full sm:w-auto px-6 py-3 bg-emerald-600 hover:bg-emerald-500 text-white font-bold text-sm rounded-xl transition flex items-center justify-center gap-2 shadow-lg shadow-emerald-600/20 cursor-pointer active:scale-98"
                      >
                        <span>Ir para Importação de Planilhas</span>
                        <ArrowRight className="w-4 h-4" />
                      </button>
                    </div>
                  </div>
                ) : (
                  <>
                    {/* KPI Cards */}
                    <KpiCards metrics={metrics} />

                    {/* Clean Main Creditors Table */}
                    <DetailedConsolidationTable
                      metrics={metrics}
                      files={files}
                      onOpenCompanyAudit={(s) => setAuditedSupplier(s)}
                    />
                  </>
                )}
              </div>
            )}

            {visitedTabs['pagamentos_fonte'] && (
              <div className={activeTab === 'pagamentos_fonte' ? 'space-y-8 animate-in fade-in-50 duration-150' : 'hidden'}>
                <PagamentosPorFonteView
                  files={files}
                  ordensBancariasList={ordensBancariasList}
                  ordensBancariasFileName={ordensBancariasFileName}
                  onUpdateOrdensBancariasList={handleUpdateOrdensBancariasList}
                  onOpenCompanyAudit={(s) => setAuditedSupplier(s)}
                  onNavigateToImportacoes={() => handleTabChange('importacoes')}
                />
              </div>
            )}

            {visitedTabs['quebra_ordem'] && (
              <div className={activeTab === 'quebra_ordem' ? 'space-y-8 animate-in fade-in-50 duration-150' : 'hidden'}>
                <QuebraOrdemView
                  files={files}
                  rawPagamentosList={rawPagamentosList}
                  onUpdatePagamentosList={(list, fn) => {
                    setRawPagamentosList(list);
                    if (fn) setPagamentosFileName(fn);
                  }}
                  onOpenCompanyAudit={(s) => setAuditedSupplier(s)}
                  onNavigateToImportacoes={() => handleTabChange('importacoes')}
                />
              </div>
            )}

            {visitedTabs['recorrencia'] && (
              <div className={activeTab === 'recorrencia' ? 'space-y-8 animate-in fade-in-50 duration-150' : 'hidden'}>
                <RecorrenciaView metrics={metrics} files={files} />
              </div>
            )}

            {visitedTabs['fornecedores_mes'] && (
              <div className={activeTab === 'fornecedores_mes' ? 'space-y-8 animate-in fade-in-50 duration-150' : 'hidden'}>
                <FornecedoresMesView
                  metrics={metrics}
                  files={files}
                  onOpenCompanyAudit={(s) => setAuditedSupplier(s)}
                  onNavigateToTab={(tab) => handleTabChange(tab)}
                />
              </div>
            )}

            {visitedTabs['categorias'] && (
              <div className={activeTab === 'categorias' ? 'space-y-8 animate-in fade-in-50 duration-150' : 'hidden'}>
                <CategoriasView metrics={metrics} />
              </div>
            )}

            {visitedTabs['importacoes'] && (
              <div className={activeTab === 'importacoes' ? 'space-y-8 animate-in fade-in-50 duration-150' : 'hidden'}>
                <ImportacoesView
                  files={files}
                  onFileUpload={handleFileUpload}
                  onRemoveFile={handleRemoveFile}
                  onChangeFileStage={handleChangeFileStage}
                  onProcessData={handleProcessData}
                  onSaveData={handleSaveData}
                  onOpenColumnMapModal={(fileId) => setActiveMappingFileId(fileId)}
                  isProcessing={isProcessing}
                  metrics={metrics}
                  onNavigateToTab={(tab) => handleTabChange(tab)}
                  rawPagamentosList={rawPagamentosList}
                  pagamentosFileName={pagamentosFileName}
                  onUpdatePagamentosList={(list, fn) => {
                    setRawPagamentosList(list);
                    if (fn) setPagamentosFileName(fn);
                  }}
                  ordensBancariasList={ordensBancariasList}
                  ordensBancariasFileName={ordensBancariasFileName}
                  onUpdateOrdensBancariasList={handleUpdateOrdensBancariasList}
                />
              </div>
            )}
          </div>
        </main>

        {/* Footer */}
        <footer className="bg-white border-t border-slate-200 py-4 px-8 text-xs text-slate-500 flex flex-col sm:flex-row items-center justify-between gap-2">
          <span>
            <strong>Painel Financeiro - FES/SESAP</strong> — Governo do Estado do Rio Grande do Norte
          </span>
        </footer>

      </div>

      {/* Column Mapping Modal */}
      {activeMappingFileId && (
        <ColumnMappingModal
          fileData={activeMappingFile}
          onClose={() => setActiveMappingFileId(null)}
          onSaveMapping={handleSaveColumnMapping}
        />
      )}

      {/* Extrato Geral de Auditoria da Empresa Modal */}
      {auditedSupplier && (
        <ExtratoAuditoriaEmpresaModal
          supplier={auditedSupplier}
          onClose={() => setAuditedSupplier(null)}
          files={files}
        />
      )}

      {/* Toast Alert Notification System */}
      <ToastAlertContainer onSelectToast={(t) => setSelectedToastForVerification(t)} />

      {/* Modal de Verificação / Alteração do Registro Notificado */}
      {selectedToastForVerification && (
        <ModalVerificacaoNotificacao
          isOpen={!!selectedToastForVerification}
          toast={selectedToastForVerification}
          files={files}
          onClose={() => setSelectedToastForVerification(null)}
          onSaveAlteration={handleSaveToastAlteration}
          onOpenCompanyAudit={(supplier) => setAuditedSupplier(supplier)}
        />
      )}

      {/* Central de Notificações / Histórico Modal */}
      {isModalNotificacoesOpen && (
        <ModalHistoricoNotificacoes
          isOpen={isModalNotificacoesOpen}
          onClose={() => setIsModalNotificacoesOpen(false)}
          onSelectToast={(t) => setSelectedToastForVerification(t)}
          files={files}
          ordensBancariasList={ordensBancariasList}
          onSyncAlerts={() => {
            toastAlertService.syncConfirmations(files, ordensBancariasList, true);
          }}
        />
      )}

      {/* Machine Learning / IA Audit Modal */}
      {isModalMLOpen && (
        <ModalAuditoriaML
          isOpen={isModalMLOpen}
          onClose={() => setIsModalMLOpen(false)}
          files={files}
          ordensBancariasList={ordensBancariasList}
        />
      )}

    </div>
  );
}

export default App;
