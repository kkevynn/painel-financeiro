import React, { useState, useEffect, useRef, useCallback, startTransition } from 'react';
import { FileSpreadsheet, ArrowRight } from 'lucide-react';
import { FileData, DashboardMetrics, StageKey, ColumnMappingConfig } from './types';
import {
  parseExcelFile,
  parseWorkbookIntoFileDatas,
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
import { ErrorBoundary } from './components/ErrorBoundary';

import { ImportacoesView } from './components/views/ImportacoesView';
import { CategoriasView } from './components/views/CategoriasView';
import { RecorrenciaView } from './components/views/RecorrenciaView';
import { FornecedoresMesView } from './components/views/FornecedoresMesView';
import { PagamentosPorFonteView } from './components/views/PagamentosPorFonteView';
import { PublicacaoView } from './components/views/PublicacaoView';
import { ParaPublicarView } from './components/views/ParaPublicarView';
import { ExtratoAuditoriaEmpresaModal } from './components/ExtratoAuditoriaEmpresaModal';
import { ToastAlertContainer } from './components/ToastAlertContainer';
import { ModalAuditoriaML } from './components/ModalAuditoriaML';
import { ModalVerificacaoNotificacao } from './components/ModalVerificacaoNotificacao';
import { ModalHistoricoNotificacoes } from './components/ModalHistoricoNotificacoes';
import { toastAlertService } from './utils/toastAlertService';
import { mlInferenceService, getPendingInferences } from './utils/machineLearningInferenceService';
import { ToastAlert } from './types';
import {
  OrdemBancariaItem,
  enrichSuppliersWithOrdensBancarias,
  saveOrdensBancariasToStorage,
  loadOrdensBancariasFromStorage,
  loadOrdensBancariasAsync,
  isObFileName,
  parseListarOrdemBancariaMatrixAsync,
  extractOrdensBancariasFromFileData,
  isValidFonteRecursoCode,
} from './utils/ordemBancariaParser';
import { readSpreadsheetAsMatrix } from './utils/excelParser';
import { googleDriveService, GoogleDriveState } from './utils/googleDriveService';
import { cloudSyncService, CloudSyncState } from './utils/cloudSyncService';
import { indexAppKnowledge } from './utils/appKnowledgeDirectory';
import { paraPublicarService } from './utils/paraPublicarService';
import { auth } from './firebase';
import { authService, getCurrentSession, AuthSession } from './utils/authService';
import { LoginScreen } from './components/LoginScreen';

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
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);
  const [lastProcessedAt, setLastProcessedAt] = useState<Date | null>(null);
  const lastProcessedAtRef = useRef<Date | null>(lastProcessedAt);
  useEffect(() => {
    lastProcessedAtRef.current = lastProcessedAt;
  }, [lastProcessedAt]);

  const [session, setSession] = useState<AuthSession | null>(() => getCurrentSession());

  const handleLogout = useCallback(() => {
    authService.logout();
    setSession(null);
  }, []);

  // Viewer cannot stay on importacoes tab: redirect to painel
  useEffect(() => {
    if (session?.role === 'viewer' && activeTab === 'importacoes') {
      setActiveTab('painel');
    }
  }, [session?.role, activeTab]);

  const handleTabChange = useCallback((newTab: string) => {
    if (session?.role === 'viewer' && newTab === 'importacoes') {
      return;
    }
    startTransition(() => {
      setActiveTab(newTab);
    });
  }, [session?.role]);

  // Array of loaded files
  const [files, setFiles] = useState<FileData[]>([]);

  // Consolidated dashboard metrics
  const [metrics, setMetrics] = useState<DashboardMetrics>(initialMetrics);

  const [isProcessing, setIsProcessing] = useState(false);
  const [processingProgress, setProcessingProgress] = useState<number>(0);
  const [processingMessage, setProcessingMessage] = useState<string>('Processando dados...');
  const [activeMappingFileId, setActiveMappingFileId] = useState<string | null>(null);
  const [isStorageLoaded, setIsStorageLoaded] = useState(false);

  // State for Relatório "Listar Ordem Bancária" (Pagos / Detalhamento por Fonte e Mês)
  const [ordensBancariasList, setOrdensBancariasList] = useState<OrdemBancariaItem[] | null>(null);
  const [ordensBancariasFileName, setOrdensBancariasFileName] = useState<string>('');

  // Google Drive & Sheets Integration State
  const [driveState, setDriveState] = useState<GoogleDriveState>(googleDriveService.getState());
  const [isDriveSyncing, setIsDriveSyncing] = useState(false);

  // Cloud Multi-Device Synchronization State (Tablet & PC)
  const [cloudSyncState, setCloudSyncState] = useState<CloudSyncState>(cloudSyncService.getState());



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

  // Ref to always access latest ordensBancariasList and fileName inside async handlers
  const ordensBancariasListRef = useRef<OrdemBancariaItem[] | null>(ordensBancariasList);
  useEffect(() => {
    ordensBancariasListRef.current = ordensBancariasList;
  }, [ordensBancariasList]);

  const ordensBancariasFileNameRef = useRef<string>(ordensBancariasFileName);
  useEffect(() => {
    ordensBancariasFileNameRef.current = ordensBancariasFileName;
  }, [ordensBancariasFileName]);

  // Update ordens bancarias handler
  const handleUpdateOrdensBancariasList = (list: OrdemBancariaItem[] | null, fileName?: string) => {
    // Proteção: nunca sobrescrever um ordensBancariasList com fontes identificadas por uma lista 100% sem fontes válidas
    if (
      list !== null &&
      list.length > 0 &&
      ordensBancariasListRef.current &&
      ordensBancariasListRef.current.some((it) => isValidFonteRecursoCode(it.fonteRecurso)) &&
      !list.some((it) => isValidFonteRecursoCode(it.fonteRecurso))
    ) {
      return;
    }

    ordensBancariasListRef.current = list;
    setOrdensBancariasList(list);
    if (fileName !== undefined) {
      ordensBancariasFileNameRef.current = fileName;
      setOrdensBancariasFileName(fileName);
    }
    saveOrdensBancariasToStorage(list, fileName);

    // Immediately re-enrich current metrics
    setMetrics((prev) => {
      const cloned = { ...prev, consolidatedSuppliers: [...prev.consolidatedSuppliers] };
      return applyEnrichment(cloned, list);
    });
    triggerAlertAnalysis(filesRef.current, list);
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
    if (session?.role !== 'admin') {
      toastAlertService.addToast({
        tipo: 'aviso',
        titulo: 'Acesso Restrito',
        mensagem: 'Modo Consulta: apenas Administradores podem alterar registros.',
      });
      return;
    }

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
      let stored: any = null;
      let obStored: { items: OrdemBancariaItem[] | null; fileName: string } = { items: null, fileName: '' };
      try {

        try {
          stored = await loadAppState();
        } catch (e) {
          console.warn('Falha ao ler loadAppState:', e);
        }

        try {
          obStored = await loadOrdensBancariasAsync();
        } catch (e) {
          console.warn('Falha ao ler loadOrdensBancariasAsync:', e);
        }

        const storedHasValidFonte = Boolean(
          obStored?.items &&
            obStored.items.length > 0 &&
            obStored.items.some((it) => isValidFonteRecursoCode(it.fonteRecurso))
        );

        if (obStored && obStored.items && obStored.items.length > 0 && storedHasValidFonte) {
          ordensBancariasListRef.current = obStored.items;
          setOrdensBancariasList(obStored.items);
          ordensBancariasFileNameRef.current = obStored.fileName || '';
          setOrdensBancariasFileName(obStored.fileName || '');
        }

        const realFiles = (stored?.files || []).filter(
          (f: any) => {
            if (!f || f.id === 'sample_file_1' || f.fileName?.includes('Relatorio_Credores_Valor_Total.xlsx')) return false;
            // Purge extraneous auxiliary tabs from "Processamento de dados": only keep 'obedece' and 'nao_obedece'
            const lower = (f.fileName || '').toLowerCase();
            if (lower.includes('processamento') && lower.includes('dados')) {
              const isOb = f.stage === 'obedece' || (lower.includes('obedece') && !lower.includes('nao') && !lower.includes('não'));
              const isNaoOb = f.stage === 'nao_obedece' || lower.includes('nao obedece') || lower.includes('não obedece');
              if (!isOb && !isNaoOb) return false;
            }
            return true;
          }
        );

        // Auto-extract OBs EXCLUSIVAMENTE de planilhas de Ordem Bancária se ordensBancariasList não tiver fontes válidas
        if ((!obStored?.items || obStored.items.length === 0 || !storedHasValidFonte) && realFiles.length > 0) {
          const obMap = new Map<string, OrdemBancariaItem>();
          let extractedObFileName = '';
          realFiles.forEach((f: any) => {
            if (f.stage !== 'base_pps' && (f.stage === 'ordem_bancaria' || isObFileName(f.fileName))) {
              const extracted = extractOrdensBancariasFromFileData(f).filter((item) =>
                isValidFonteRecursoCode(item.fonteRecurso)
              );
              if (extracted.length > 0) {
                extractedObFileName = f.fileName || extractedObFileName;
                extracted.forEach((item) => {
                  const k = item.ob ? ('OB_' + item.ob.trim().toUpperCase()) : (item.pp ? ('PP_' + item.pp.trim().toUpperCase()) : item.id);
                  obMap.set(k, item);
                });
              }
            }
          });
          const allExtracted = Array.from(obMap.values());
          if (allExtracted.length > 0) {
            const finalObFileName = extractedObFileName || 'Consolidado de Arquivos';
            ordensBancariasListRef.current = allExtracted;
            setOrdensBancariasList(allExtracted);
            ordensBancariasFileNameRef.current = finalObFileName;
            setOrdensBancariasFileName(finalObFileName);
            saveOrdensBancariasToStorage(allExtracted, finalObFileName);
            obStored = { items: allExtracted, fileName: finalObFileName };
          } else if (obStored?.items && obStored.items.length > 0) {
            ordensBancariasListRef.current = obStored.items;
            setOrdensBancariasList(obStored.items);
            ordensBancariasFileNameRef.current = obStored.fileName || '';
            setOrdensBancariasFileName(obStored.fileName || '');
          }
        }

        if (realFiles.length > 0) {
          setFiles(realFiles);
          const ts = stored?.lastProcessedAt ? new Date(stored.lastProcessedAt) : new Date();
          setLastProcessedAt(ts);
          try {
            const hasPartialFiles = realFiles.some(
              (f: any) =>
                f.isSummaryOnly === true ||
                (typeof f.rowCount === 'number' && (f.rows?.length || 0) < f.rowCount)
            );
            let restoredMetricsFromCache = false;
            if (hasPartialFiles) {
              try {
                const cachedRaw = localStorage.getItem('painel_fes_metrics_cache_v1');
                if (cachedRaw) {
                  const cached = JSON.parse(cachedRaw);
                  if (cached?.metrics?.isProcessed) {
                    setMetrics(cached.metrics);
                    restoredMetricsFromCache = true;
                  }
                }
              } catch (_) {}
            }
            if (!restoredMetricsFromCache) {
              const processed = processUploadedFiles(realFiles);
              const enriched = applyEnrichment(processed, obStored?.items || null);
              setMetrics(enriched);
            }
          } catch (procErr) {
            console.error('Erro ao reprocessar dados salvos:', procErr);
          }

          try {
            toastAlertService.forceUpdateBaseline(realFiles, obStored?.items || null);
          } catch (tErr) {
            console.warn('Erro ao atualizar baseline no boot:', tErr);
          }

          try {
            mlInferenceService.processFilesForInference(realFiles, obStored?.items || null);
          } catch (mlErr) {
            console.warn('Erro ao inferir ML no boot:', mlErr);
          }
        } else if (obStored?.items && obStored.items.length > 0) {
          try {
            toastAlertService.forceUpdateBaseline([], obStored.items);
            mlInferenceService.processFilesForInference([], obStored.items);
          } catch (_) {}
        } else {
          let restoredFromCache = false;
          try {
            const cachedRaw = localStorage.getItem('painel_fes_metrics_cache_v1');
            if (cachedRaw) {
              const cached = JSON.parse(cachedRaw);
              if (cached?.metrics?.isProcessed) {
                setMetrics(cached.metrics);
                if (cached.lastSyncAt) {
                  setLastProcessedAt(new Date(cached.lastSyncAt));
                }
                if (cached.obFileName) {
                  ordensBancariasFileNameRef.current = cached.obFileName;
                  setOrdensBancariasFileName(cached.obFileName);
                }
                restoredFromCache = true;
              }
            }
          } catch (_) {}
          if (!restoredFromCache) {
            setFiles([]);
            setMetrics(initialMetrics);
          }
        }
      } catch (err) {
        console.error('Erro ao carregar do armazenamento:', err);
      } finally {
        setIsStorageLoaded(true);
        // Index app knowledge across loaded files and OBs
        try {
          indexAppKnowledge(stored?.files || [], obStored?.items || null);
        } catch (_) {}
        // Automatic cloud sync verification: pulls updates directly from the cloud on startup
        cloudSyncService.checkAndAutoSync().catch((e) => {
          console.warn('Auto-sync na inicialização:', e);
        });
      }
    }
    initStorage();
  }, []);

  // Auto-save to local storage (IndexedDB) ONLY when user actually makes changes (debounced to keep UI fluid)
  useEffect(() => {
    if (!isStorageLoaded) return;
    if (!hasUserChangesRef.current) return;

    const timer = setTimeout(() => {
      saveAppState(files, lastProcessedAt);
      hasUserChangesRef.current = false;
    }, 1000);
    return () => clearTimeout(timer);
  }, [files, lastProcessedAt, isStorageLoaded]);

  const filesRef = useRef<FileData[]>(files);
  useEffect(() => {
    filesRef.current = files;
  }, [files]);
  const activeUploadsCountRef = useRef(0);

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
    if (session?.role !== 'admin') {
      toastAlertService.addToast({
        tipo: 'aviso',
        titulo: 'Acesso Restrito',
        mensagem: 'Apenas Administradores podem importar planilhas.',
      });
      return;
    }

    activeUploadsCountRef.current++;
    setIsProcessing(true);
    setProcessingProgress(20);
    setProcessingMessage(`Lendo e processando "${file.name}"...`);

    try {
      await new Promise((res) => setTimeout(res, 60));
      const parsedFiles = await parseWorkbookIntoFileDatas(file, stage);
      setProcessingProgress(55);

      hasUserChangesRef.current = true;

      // Ensure baseline captures current state before updating/replacing files
      if (filesRef.current.length > 0) {
        toastAlertService.ensureBaseline(filesRef.current, ordensBancariasList);
      }

      let updated = [...filesRef.current];
      let targetOBs = ordensBancariasList;

      for (const parsedFileData of parsedFiles) {
        // If a file with the same name already exists, replace it instead of duplicating
        const existingIdx = updated.findIndex((f) => f.fileName === parsedFileData.fileName);
        if (existingIdx >= 0) {
          updated = updated.map((f, i) => (i === existingIdx ? parsedFileData : f));
        } else {
          updated = [...updated, parsedFileData];
        }

        // Extrai ordensBancariasList EXCLUSIVAMENTE de planilhas de Ordem Bancária (jamais de base_pps)
        if (
          parsedFileData.stage !== 'base_pps' &&
          (parsedFileData.stage === 'ordem_bancaria' || isObFileName(parsedFileData.fileName))
        ) {
          try {
            const matrix = await readSpreadsheetAsMatrix(file);
            const parsedOBs = await parseListarOrdemBancariaMatrixAsync(matrix);
            if (parsedOBs.length > 0) {
              targetOBs = parsedOBs;
              handleUpdateOrdensBancariasList(parsedOBs, parsedFileData.fileName);
            } else {
              const extracted = extractOrdensBancariasFromFileData(parsedFileData);
              if (extracted.length > 0) {
                targetOBs = extracted;
                handleUpdateOrdensBancariasList(extracted, parsedFileData.fileName);
              }
            }
          } catch (obErr) {
            console.warn('Erro ao auto-processar planilha de OBs:', obErr);
            const extracted = extractOrdensBancariasFromFileData(parsedFileData);
            if (extracted.length > 0) {
              targetOBs = extracted;
              handleUpdateOrdensBancariasList(extracted, parsedFileData.fileName);
            }
          }
        }
      }

      filesRef.current = updated;
      setFiles(updated);

      if (activeUploadsCountRef.current <= 1) {
        setProcessingMessage('Consolidando dados...');
        setProcessingProgress(78);
        await new Promise((res) => setTimeout(res, 30));
        const newMetrics = processUploadedFiles(updated);
        const enriched = applyEnrichment(newMetrics, targetOBs);
        setMetrics(enriched);
        setLastProcessedAt(new Date());
        setProcessingProgress(92);
        triggerAlertAnalysis(updated, targetOBs);
      }
    } catch (error: any) {
      toastAlertService.addToast({
        tipo: 'critico',
        titulo: 'Erro ao carregar arquivo',
        mensagem: error.message || 'Arquivo corrompido ou formato inválido.',
      });
    } finally {
      activeUploadsCountRef.current = Math.max(0, activeUploadsCountRef.current - 1);
      if (activeUploadsCountRef.current === 0) {
        const finalFiles = filesRef.current;
        const finalOBs = ordensBancariasListRef.current;
        setTimeout(() => {
          const finalMetrics = processUploadedFiles(finalFiles);
          const finalEnriched = applyEnrichment(finalMetrics, finalOBs);
          setMetrics(finalEnriched);
          triggerAlertAnalysis(finalFiles, finalOBs);
          saveAppState(finalFiles, lastProcessedAt);
          // Auto-save automatically to Cloud Drive (apenas Administrador)
          if (session?.role === 'admin') {
            cloudSyncService.pushToCloud(finalFiles, finalOBs, ordensBancariasFileNameRef.current, finalEnriched).catch((e) => {
              console.warn('Auto-save na nuvem pós-upload:', e);
            });
          }
          setProcessingProgress(100);
          setTimeout(() => {
            setIsProcessing(false);
            setProcessingProgress(0);
          }, 350);
        }, 50);
      }
    }
  };

  // Google Drive Integration: register file delivery callback and subscription
  useEffect(() => {
    const unsubscribe = googleDriveService.subscribe((state) => {
      setDriveState(state);
    });

    // Provide currently loaded file names so auto-polling never downloads unrelated files from Drive
    googleDriveService.registerExistingFilesProvider(() => {
      const names = filesRef.current.map((f) => f.fileName).filter(Boolean);
      if (ordensBancariasFileNameRef.current && !names.includes(ordensBancariasFileNameRef.current)) {
        names.push(ordensBancariasFileNameRef.current);
      }
      return { names, lastSyncAt: lastProcessedAtRef.current };
    });

    // When Google Drive delivers a file (either manual sync or 20-minute background auto-polling)
    googleDriveService.registerNewFileCallback(async (file: File, isAuto: boolean) => {
      if (isAuto) {
        toastAlertService.addToast({
          tipo: 'info',
          titulo: 'Atualização Automática do Google Drive',
          mensagem: `Nova versão da planilha "${file.name}" detectada e sincronizada.`,
        });
      }
      await handleFileUpload(file);
    });

    return () => {
      unsubscribe();
      googleDriveService.stopAutoPolling();
    };
  }, []);

  // Cloud Multi-Device Sync: subscribe and register dataset reload handler
  useEffect(() => {
    const unsubscribe = cloudSyncService.subscribe((state) => {
      setCloudSyncState(state);
    });

    cloudSyncService.registerDatasetReloadCallback((reloadedFiles, reloadedOBs, syncDate, reloadedMetrics) => {
      if (reloadedFiles && reloadedFiles.length > 0) {
        filesRef.current = reloadedFiles;
        setFiles(reloadedFiles);

        if (reloadedOBs && reloadedOBs.length > 0) {
          ordensBancariasListRef.current = reloadedOBs;
          setOrdensBancariasList(reloadedOBs);
        } else if (reloadedOBs === null) {
          ordensBancariasListRef.current = null;
          setOrdensBancariasList(null);
        }
        if (syncDate) {
          setLastProcessedAt(syncDate);
        }

        // Se reloadedMetrics vier da nuvem, aplica DIRETAMENTE sem recalcular
        if (reloadedMetrics && reloadedMetrics.isProcessed) {
          setMetrics(reloadedMetrics);
        } else {
          try {
            const processed = processUploadedFiles(reloadedFiles);
            const enriched = applyEnrichment(processed, reloadedOBs || ordensBancariasListRef.current);
            setMetrics(enriched);
          } catch (e) {
            console.warn('Erro ao processar dados recarregados da nuvem:', e);
          }
        }
      } else if (Array.isArray(reloadedFiles) && reloadedFiles.length === 0 && !reloadedMetrics?.isProcessed) {
        filesRef.current = [];
        setFiles([]);
        ordensBancariasListRef.current = null;
        setOrdensBancariasList(null);
        ordensBancariasFileNameRef.current = '';
        setOrdensBancariasFileName('');
        setMetrics(initialMetrics);
        setLastProcessedAt(null);
      }
    });

    return () => {
      unsubscribe();
    };
  }, [applyEnrichment]);

  const handleConnectDrive = async () => {
    try {
      const ok = await googleDriveService.connectGoogleDrive();
      if (ok) {
        // Automatically link Cloud Sync (Firestore) with Google account as well
        try {
          if (!cloudSyncService.getState().isAuthenticated) {
            await cloudSyncService.loginWithGoogle();
          } else {
            await cloudSyncService.checkAndAutoSync();
          }
        } catch (cloudErr) {
          console.warn('Sincronização em nuvem automática:', cloudErr);
        }

        toastAlertService.addToast({
          tipo: 'sucesso',
          titulo: 'Conta Google Conectada',
          mensagem: 'Sincronização em nuvem (Tablet/PC) e Google Drive ativadas com sucesso.',
        });
        // Immediately sync only currently loaded financial reports from Drive (incremental) and sync cloud state
        await handleSyncDriveNow(false);
      }
    } catch (err: any) {
      toastAlertService.addToast({
        tipo: 'critico',
        titulo: 'Falha ao Conectar Conta Google',
        mensagem: err?.message || 'Verifique as permissões de acesso ao Google Cloud.',
      });
    }
  };

  const handleDisconnectDrive = async () => {
    googleDriveService.disconnectGoogleDrive();
    try {
      await cloudSyncService.logout();
    } catch {
      // ignore
    }
    toastAlertService.addToast({
      tipo: 'info',
      titulo: 'Conta Google Desconectada',
      mensagem: 'A sincronização com o Google Drive e a nuvem foi encerrada.',
    });
  };

  const handleSyncDriveNow = async (forceParam: boolean | unknown = true) => {
    if (isDriveSyncing) return;
    const force = typeof forceParam === 'boolean' ? forceParam : true;
    setIsDriveSyncing(true);
    setIsProcessing(true);
    setProcessingProgress(20);
    setProcessingMessage('Sincronizando com Google Drive e Nuvem...');
    try {
      setProcessingProgress(40);
      const currentFileNames = filesRef.current.map((f) => f.fileName).filter(Boolean);
      if (ordensBancariasFileNameRef.current && !currentFileNames.includes(ordensBancariasFileNameRef.current)) {
        currentFileNames.push(ordensBancariasFileNameRef.current);
      }
      const result = await googleDriveService.syncNow(currentFileNames, lastProcessedAt, force);
      
      // Also sync Firestore cloud state across Tablet & PC
      try {
        setProcessingProgress(70);
        await cloudSyncService.checkAndAutoSync();
      } catch (cloudErr) {
        console.warn('Sincronização em nuvem:', cloudErr);
      }

      setProcessingProgress(90);
      if (result.success) {
        toastAlertService.addToast({
          tipo: 'sucesso',
          titulo: 'Sincronização Concluída',
          mensagem: `${result.message} Dados em nuvem atualizados para todos os seus dispositivos.`,
        });
      } else {
        toastAlertService.addToast({
          tipo: 'sucesso',
          titulo: 'Nuvem e Dispositivos Atualizados',
          mensagem: 'Planilhas e dados sincronizados com sucesso entre Tablet e PC.',
        });
      }
    } catch (err: any) {
      toastAlertService.addToast({
        tipo: 'critico',
        titulo: 'Erro ao Sincronizar',
        mensagem: err?.message || 'Falha ao buscar planilhas no Google Drive.',
      });
    } finally {
      setIsDriveSyncing(false);
      setProcessingProgress(100);
      setTimeout(() => {
        setIsProcessing(false);
        setProcessingProgress(0);
      }, 400);
    }
  };

  // Change file stage tag
  const handleChangeFileStage = (fileId: string, stage: StageKey) => {
    if (session?.role !== 'admin') return;

    hasUserChangesRef.current = true;
    let targetUpdatedFile: FileData | null = null;
    const updated = filesRef.current.map((f) => {
      if (f.id === fileId) {
        const up = { ...f, stage };
        targetUpdatedFile = up;
        return up;
      }
      return f;
    });
    filesRef.current = updated;
    setFiles(updated);

    let targetOBs = ordensBancariasListRef.current;
    if (stage === 'ordem_bancaria' && targetUpdatedFile) {
      const extracted = extractOrdensBancariasFromFileData(targetUpdatedFile);
      if (extracted.length > 0) {
        targetOBs = extracted;
        handleUpdateOrdensBancariasList(extracted, (targetUpdatedFile as FileData).fileName);
      }
    }

    const newMetrics = processUploadedFiles(updated);
    const enriched = applyEnrichment(newMetrics, targetOBs);
    setMetrics(enriched);
    triggerAlertAnalysis(updated, targetOBs);
    cloudSyncService.pushToCloud(updated, targetOBs, ordensBancariasFileNameRef.current, enriched).catch((e) => {
      console.warn('Auto-save na nuvem alteração de etapa:', e);
    });
  };

  // Remove file action
  const handleRemoveFile = (fileId: string) => {
    if (session?.role !== 'admin') {
      toastAlertService.addToast({
        tipo: 'aviso',
        titulo: 'Acesso Restrito',
        mensagem: 'Modo Consulta: apenas Administradores podem remover arquivos.',
      });
      return;
    }

    hasUserChangesRef.current = true;
    const removed = filesRef.current.find((f) => f.id === fileId);
    const updated = filesRef.current.filter((f) => f.id !== fileId);
    filesRef.current = updated;
    setFiles(updated);

    // Delete directly from Cloud Firestore to keep all devices in sync
    cloudSyncService.deleteFileFromCloud(fileId);

    let targetOBs = ordensBancariasList;
    if (
      removed &&
      (removed.stage === 'ordem_bancaria' ||
        removed.fileName === ordensBancariasFileName ||
        isObFileName(removed.fileName))
    ) {
      const remainingOb = updated.find((f) => f.stage === 'ordem_bancaria' || isObFileName(f.fileName));
      if (remainingOb) {
        const extracted = extractOrdensBancariasFromFileData(remainingOb);
        targetOBs = extracted.length > 0 ? extracted : null;
        handleUpdateOrdensBancariasList(targetOBs, remainingOb.fileName);
      } else {
        targetOBs = null;
        handleUpdateOrdensBancariasList(null, '');
      }
    }

    // Use non-blocking schedule for consolidation so UI never stutters or freezes
    setTimeout(() => {
      const newMetrics = processUploadedFiles(updated);
      const enriched = applyEnrichment(newMetrics, targetOBs);
      setMetrics(enriched);
      saveAppState(updated, lastProcessedAt);
      if (session?.role === 'admin') {
        cloudSyncService.pushToCloud(updated, targetOBs, ordensBancariasFileNameRef.current, enriched).catch((e) => {
          console.warn('Auto-save na nuvem remoção de arquivo:', e);
        });
      }
    }, 10);
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
        if (session?.role === 'admin') {
          cloudSyncService.pushToCloud(files, ordensBancariasList, ordensBancariasFileNameRef.current, enriched).catch((e) => {
            console.warn('Auto-save na nuvem processamento de dados:', e);
          });
        }
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
    if (session?.role !== 'admin') {
      toastAlertService.addToast({
        tipo: 'aviso',
        titulo: 'Acesso Restrito',
        mensagem: 'Modo Consulta: alterações não podem ser salvas na nuvem.',
      });
      return;
    }

    setIsProcessing(true);
    setProcessingMessage('Salvando alterações e sincronizando com o Google Drive...');
    try {
      const now = new Date();
      setLastProcessedAt(now);
      const newMetrics = processUploadedFiles(files);
      const enriched = applyEnrichment(newMetrics, ordensBancariasList);
      setMetrics(enriched);

      // 1. Salva no armazenamento local primeiro (garantia imediata)
      await saveAppState(files, now);

      // 2. Analisa transições de alertas e aprendizado de máquina
      toastAlertService.analyzeFileTransitions(files, ordensBancariasList);
      mlInferenceService.processFilesForInference(files, ordensBancariasList);

      // 3. Aguarda sincronização completa no Google Drive antes de notificar sucesso
      const ok = await cloudSyncService.pushToCloud(
        files,
        ordensBancariasList,
        ordensBancariasFileName,
        enriched
      );

      if (ok) {
        toastAlertService.addToast({
          tipo: 'sucesso',
          titulo: 'Salvo e Sincronizado!',
          mensagem: 'Dados gravados localmente e sincronizados com o Google Drive.',
        });
      } else {
        toastAlertService.addToast({
          tipo: 'aviso',
          titulo: 'Salvo localmente',
          mensagem: 'Dados gravados no navegador, mas não foi possível confirmar o envio ao Google Drive.',
        });
      }

      // Breve pausa para visualização da confirmação
      await new Promise((res) => setTimeout(res, 200));
    } finally {
      setIsProcessing(false);
    }
  };

  // Clear all data
  const handleClearData = async () => {
    if (session?.role !== 'admin') {
      toastAlertService.addToast({
        tipo: 'aviso',
        titulo: 'Acesso Restrito',
        mensagem: 'Modo Consulta: apenas Administradores podem limpar o painel.',
      });
      return;
    }

    setIsProcessing(true);
    setProcessingMessage('Limpando dados e redefinindo painel...');
    try {
      filesRef.current = [];
      setFiles([]);
      ordensBancariasListRef.current = null;
      setOrdensBancariasList(null);
      ordensBancariasFileNameRef.current = '';
      setOrdensBancariasFileName('');
      saveOrdensBancariasToStorage(null, '');
      setMetrics(initialMetrics);
      setLastProcessedAt(null);
      await clearAppState();
      try {
        localStorage.removeItem('painel_fes_metrics_cache_v1');
      } catch (_) {}
      await cloudSyncService.clearCloudWorkspace();
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
    let targetUpdatedFile: FileData | null = null;
    const updated = filesRef.current.map((f) => {
      if (f.id === fileId) {
        const up: FileData = {
          ...f,
          detectedFavorecidoCol: mapping.favorecidoCol,
          detectedValorCol: mapping.valorCol,
          stage: mapping.stage || f.stage,
          detectedCnpjCol: mapping.cnpjCol || undefined,
          detectedSituacaoCol: mapping.situacaoCol || undefined,
          detectedSituacaoObCol: mapping.situacaoObCol || (f as any).detectedSituacaoObCol || undefined,
          detectedNumeroCol: mapping.numeroCol || undefined,
          detectedPpCol: mapping.numeroCol || undefined,
          detectedObCol: mapping.obCol || undefined,
          detectedNotaEmpenhoCol: mapping.notaEmpenhoCol || undefined,
          detectedFavorecidoNECol: mapping.favorecidoNECol || undefined,
          detectedProcessoCol: mapping.processoCol || undefined,
          detectedFonteCol: mapping.fonteCol || undefined,
          detectedNaturezaCol: mapping.naturezaCol || undefined,
          detectedDataCol: mapping.dataCol || undefined,
        };
        targetUpdatedFile = up;
        return up;
      }
      return f;
    });
    filesRef.current = updated;
    setFiles(updated);

    // Re-extrai OB items EXCLUSIVAMENTE se o arquivo atualizado for planilha de Ordem Bancária (nunca base_pps)
    let targetOBs = ordensBancariasList;
    if (targetUpdatedFile) {
      const uf: FileData = targetUpdatedFile;
      if (
        uf.stage !== 'base_pps' &&
        (uf.stage === 'ordem_bancaria' || isObFileName(uf.fileName))
      ) {
        const reExtracted = extractOrdensBancariasFromFileData(uf);
        if (reExtracted.length > 0) {
          targetOBs = reExtracted;
          handleUpdateOrdensBancariasList(reExtracted, uf.fileName);
        }
      }
    }

    const newMetrics = processUploadedFiles(updated);
    const enriched = applyEnrichment(newMetrics, targetOBs);
    setMetrics(enriched);
    saveAppState(updated, lastProcessedAt || new Date());
    triggerAlertAnalysis(updated, targetOBs);
    if (session?.role === 'admin') {
      cloudSyncService.pushToCloud(updated, targetOBs, ordensBancariasFileNameRef.current, enriched).catch((e) => {
        console.warn('Auto-save na nuvem mapeamento de colunas:', e);
      });
    }
  };

  const activeMappingFile = files.find((f) => f.id === activeMappingFileId) || null;

  // Se não houver sessão ativa, renderiza tela institucional de login
  if (!session) {
    return (
      <LoginScreen
        onLoginSuccess={async (newSession) => {
          setSession(newSession);
          setIsProcessing(true);
          setProcessingMessage('Carregando dados atualizados...');
          try {
            await cloudSyncService.pullFromCloud(false);
          } catch (err) {
            console.warn('Erro ao sincronizar dados após login:', err);
          } finally {
            setIsProcessing(false);
          }
        }}
      />
    );
  }

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
        userRole={session?.role}
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
          driveState={driveState}
          onConnectDrive={handleConnectDrive}
          onDisconnectDrive={handleDisconnectDrive}
          onSyncDriveNow={handleSyncDriveNow}
          isDriveSyncing={isDriveSyncing}
          cloudSyncState={cloudSyncState}
          isProcessing={isProcessing}
          processingProgress={processingProgress}
          processingMessage={processingMessage}
          userRole={session?.role}
          userName={session?.displayName}
          onLogout={handleLogout}
        />

        {/* Dynamic Workspace (Expanded Layout to 98% Screen Width) */}
        <main className="p-4 sm:p-6 md:p-8 space-y-8 w-full max-w-[98%] mx-auto flex-1">
          <ErrorBoundary fallbackTitle="Falha ao carregar o módulo selecionado">
            <div className="space-y-8">
              {activeTab === 'painel' && (
                <div className="space-y-8 animate-in fade-in-50 duration-150">
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
                        {session?.role === 'admin' ? (
                          <button
                            onClick={() => handleTabChange('importacoes')}
                            className="w-full sm:w-auto px-6 py-3 bg-emerald-600 hover:bg-emerald-500 text-white font-bold text-sm rounded-xl transition flex items-center justify-center gap-2 shadow-lg shadow-emerald-600/20 cursor-pointer active:scale-98"
                          >
                            <span>Ir para Importação de Planilhas</span>
                            <ArrowRight className="w-4 h-4" />
                          </button>
                        ) : (
                          <span className="text-xs text-emerald-300/80 font-medium">
                            Aguardando publicação de dados pelo Administrador.
                          </span>
                        )}
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
                        ordensBancariasList={ordensBancariasList}
                        onOpenCompanyAudit={(s) => setAuditedSupplier(s)}
                      />
                    </>
                  )}
                </div>
              )}

              {activeTab === 'pagamentos_fonte' && (
                <div className="space-y-8 animate-in fade-in-50 duration-150">
                  <PagamentosPorFonteView
                    files={files}
                    ordensBancariasList={ordensBancariasList}
                    ordensBancariasFileName={ordensBancariasFileName}
                    onUpdateOrdensBancariasList={handleUpdateOrdensBancariasList}
                    onOpenCompanyAudit={(s) => setAuditedSupplier(s)}
                    onNavigateToImportacoes={() => handleTabChange('importacoes')}
                    onOpenColumnMapModal={(fileId) => setActiveMappingFileId(fileId)}
                    onFileUpload={handleFileUpload}
                  />
                </div>
              )}

              {activeTab === 'publicacao' && (
                <div className="space-y-8 animate-in fade-in-50 duration-150">
                  <PublicacaoView
                    files={files}
                    ordensBancariasList={ordensBancariasList}
                    onFileUpload={handleFileUpload}
                    onNavigateToImportacoes={() => handleTabChange('importacoes')}
                  />
                </div>
              )}

              {activeTab === 'para_publicar' && (
                <div className="space-y-8 animate-in fade-in-50 duration-150">
                  <ParaPublicarView />
                </div>
              )}

              {activeTab === 'recorrencia' && (
                <div className="space-y-8 animate-in fade-in-50 duration-150">
                  <RecorrenciaView metrics={metrics} files={files} />
                </div>
              )}

              {activeTab === 'fornecedores_mes' && (
                <div className="space-y-8 animate-in fade-in-50 duration-150">
                  <FornecedoresMesView
                    metrics={metrics}
                    files={files}
                    onOpenCompanyAudit={(s) => setAuditedSupplier(s)}
                    onNavigateToTab={(tab) => handleTabChange(tab)}
                  />
                </div>
              )}

              {activeTab === 'categorias' && (
                <div className="space-y-8 animate-in fade-in-50 duration-150">
                  <CategoriasView metrics={metrics} />
                </div>
              )}

              {activeTab === 'importacoes' && (
                <div className="space-y-8 animate-in fade-in-50 duration-150">
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
                    ordensBancariasList={ordensBancariasList}
                    ordensBancariasFileName={ordensBancariasFileName}
                    onUpdateOrdensBancariasList={handleUpdateOrdensBancariasList}
                  />
                </div>
              )}
            </div>
          </ErrorBoundary>
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
          ordensBancariasList={ordensBancariasList}
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
