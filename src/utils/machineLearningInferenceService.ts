import { InferredPaymentMatch, LearnedCreditorPattern, FileData } from '../types';
import { parseCurrencyValue, parseFavorecidoField, isSituacaoAO, isSituacaoEnviadaBanco } from './excelParser';
import { OrdemBancariaItem } from './ordemBancariaParser';

const STORAGE_KEY_LEARNED_PATTERNS = 'painel_learned_creditor_patterns';
const STORAGE_KEY_INFERRED_MATCHES = 'painel_inferred_matches_cache';
const STORAGE_KEY_KNOWN_NLS = 'painel_known_nls_registry';

type MLChangeListener = () => void;

class MachineLearningInferenceService {
  private learnedPatterns: Record<string, LearnedCreditorPattern> = {};
  private inferredMatches: InferredPaymentMatch[] = [];
  private listeners: Set<MLChangeListener> = new Set();

  constructor() {
    this.loadFromStorage();
  }

  public subscribe(listener: MLChangeListener): () => void {
    this.listeners.add(listener);
    return () => {
      this.listeners.delete(listener);
    };
  }

  private notify() {
    this.saveToStorage();
    setTimeout(() => {
      this.listeners.forEach((fn) => {
        try {
          fn();
        } catch (e) {
          console.error('ML listener notification error:', e);
        }
      });
    }, 0);
  }

  private loadFromStorage() {
    try {
      const patternsRaw = localStorage.getItem(STORAGE_KEY_LEARNED_PATTERNS);
      if (patternsRaw) {
        this.learnedPatterns = JSON.parse(patternsRaw);
      }
    } catch (e) {
      this.learnedPatterns = {};
    }

    try {
      const matchesRaw = localStorage.getItem(STORAGE_KEY_INFERRED_MATCHES);
      if (matchesRaw) {
        this.inferredMatches = JSON.parse(matchesRaw);
      }
    } catch (e) {
      this.inferredMatches = [];
    }
  }

  private saveToStorage() {
    try {
      localStorage.setItem(STORAGE_KEY_LEARNED_PATTERNS, JSON.stringify(this.learnedPatterns));
      localStorage.setItem(STORAGE_KEY_INFERRED_MATCHES, JSON.stringify(this.inferredMatches));
    } catch (e) {
      // LocalStorage quota protection
    }
  }

  public getLearnedPatterns(): LearnedCreditorPattern[] {
    return Object.values(this.learnedPatterns);
  }

  public getInferredMatches(): InferredPaymentMatch[] {
    return [...this.inferredMatches];
  }

  public getPendingMatchesCount(): number {
    return this.inferredMatches.filter((m) => m.statusValidacao === 'pendente').length;
  }

  /**
   * Confirms an inferred match, memorizing the creditor tax retention pattern into local machine learning model
   */
  public confirmMatch(id: string) {
    const match = this.inferredMatches.find((m) => m.id === id);
    if (!match) return;

    match.statusValidacao = 'confirmado';

    const cnpjClean = (match.cnpj || '').replace(/\D/g, '');
    const patternKey = cnpjClean.length >= 11 ? cnpjClean : match.credor.toUpperCase().trim();

    const existingPattern = this.learnedPatterns[patternKey];
    const newRetention = match.percentualRetencao;

    if (existingPattern) {
      // Weighted average of retention rate
      const totalCount = existingPattern.totalConfirmacoes + 1;
      const updatedRate =
        (existingPattern.aliquotaImpostoRetidoMedia * existingPattern.totalConfirmacoes + newRetention) / totalCount;

      this.learnedPatterns[patternKey] = {
        cnpj: match.cnpj,
        credor: match.credor,
        aliquotaImpostoRetidoMedia: Number(updatedRate.toFixed(4)),
        toleranciaPercentual: 0.03, // 3% tolerance around learned tax rate
        totalConfirmacoes: totalCount,
        ultimaConfirmacao: new Date().toISOString(),
        sampleNL: match.nlNumero,
        samplePP: match.ppNumero,
      };
    } else {
      this.learnedPatterns[patternKey] = {
        cnpj: match.cnpj,
        credor: match.credor,
        aliquotaImpostoRetidoMedia: Number(newRetention.toFixed(4)),
        toleranciaPercentual: 0.03,
        totalConfirmacoes: 1,
        ultimaConfirmacao: new Date().toISOString(),
        sampleNL: match.nlNumero,
        samplePP: match.ppNumero,
      };
    }

    this.notify();
  }

  /**
   * Rejects an inferred match
   */
  public rejectMatch(id: string) {
    const match = this.inferredMatches.find((m) => m.id === id);
    if (!match) return;
    match.statusValidacao = 'rejeitado';
    this.notify();
  }

  /**
   * Confirms all pending matches that have high machine learning confidence (>85%)
   */
  public confirmAllHighConfidence() {
    let changed = false;
    this.inferredMatches.forEach((m) => {
      if (m.statusValidacao === 'pendente' && m.confianca >= 85) {
        this.confirmMatch(m.id);
        changed = true;
      }
    });
    if (changed) {
      this.notify();
    }
  }

  /**
   * Removes a learned pattern
   */
  public deleteLearnedPattern(key: string) {
    const cnpjClean = key.replace(/\D/g, '');
    const patternKey = cnpjClean.length >= 11 ? cnpjClean : key.toUpperCase().trim();
    if (this.learnedPatterns[patternKey]) {
      delete this.learnedPatterns[patternKey];
      this.notify();
    }
  }

  /**
   * Clear inference log
   */
  public clearInferences() {
    this.inferredMatches = [];
    this.notify();
  }

  /**
   * Detects NL transitions and cross-references liquidations with PPs/OBs:
   * 1. Direct cross-file match: matches liquidations (Obedece/Não Obedece) with PPs and OBs for the same creditor
   * 2. Temporal transition: detects NLs that disappeared from liquidation queue between uploads and appeared as PPs
   * 3. Tax withholding inference: learns creditor-specific tax deduction rates and auto-confirms high confidence matches
   */
  public processFilesForInference(files: FileData[], ordensBancariasList?: OrdemBancariaItem[] | null) {
    if ((!files || files.length === 0) && (!ordensBancariasList || ordensBancariasList.length === 0)) return;

    // 1. Extract currently active NLs from Liquidation files
    const currentActiveNLs: Record<
      string,
      {
        nlNumero: string;
        processo: string;
        credor: string;
        cnpj: string;
        cnpjDigits: string;
        valor: number;
        stage: 'obedece' | 'nao_obedece';
        fileName: string;
      }
    > = {};

    // 2. Extract current PPs and OBs
    const currentPPsList: Array<{
      ppNumero: string;
      obNumero: string;
      processo: string;
      credor: string;
      cnpj: string;
      cnpjDigits: string;
      valor: number;
      situacao: string;
      fileName: string;
    }> = [];

    // Process uploaded files
    (files || []).forEach((file) => {
      const stage = file.stage || 'outros';
      const isLiquidationFile =
        stage === 'obedece' ||
        stage === 'nao_obedece' ||
        (file.fileName && file.fileName.toLowerCase().includes('liquida')) ||
        (file.fileName && file.fileName.toLowerCase().includes('obedece'));

      const favCol = file.detectedFavorecidoCol;
      const valCol = file.detectedValorCol;
      const numCol = file.detectedNumeroCol || file.detectedPpCol;
      const obCol = file.detectedObCol;
      const sitCol = file.detectedSituacaoCol;
      const procCol = file.detectedProcessoCol;

      file.rows.forEach((row) => {
        const rawFav = favCol ? row[favCol] : '';
        const { name, cnpj } = parseFavorecidoField(rawFav);
        const cnpjDigits = (cnpj || '').replace(/\D/g, '');
        const val = valCol ? parseCurrencyValue(row[valCol]) : 0;
        let rawNum = numCol ? String(row[numCol] || '').trim() : '';
        let rawOB = obCol ? String(row[obCol] || '').trim() : '';
        const rawProc = procCol ? String(row[procCol] || '').trim() : '';
        const rawSit = sitCol ? String(row[sitCol] || '').trim().toUpperCase() : '';

        if (val <= 0) return;

        if (isLiquidationFile) {
          // Identify NL document number
          let docNum = rawNum || rawOB;
          if (!docNum || docNum === '-' || docNum === '0') {
            for (const [k, v] of Object.entries(row)) {
              const lk = k.toLowerCase();
              if (
                (lk.includes('documento') ||
                  lk.includes('nl') ||
                  lk.includes('nota de lancamento') ||
                  lk.includes('numero') ||
                  lk.includes('número')) &&
                v
              ) {
                const s = String(v).trim();
                if (s && s !== '-' && s !== '0') {
                  docNum = s;
                  break;
                }
              }
            }
          }

          if (docNum) {
            const key = `${cnpjDigits || name}_${docNum}`;
            currentActiveNLs[key] = {
              nlNumero: docNum,
              processo: rawProc,
              credor: name,
              cnpj,
              cnpjDigits,
              valor: val,
              stage: stage === 'nao_obedece' ? 'nao_obedece' : 'obedece',
              fileName: file.fileName,
            };
          }
        } else {
          // It's a PP file or general report
          let ppNum = rawNum || rawOB;
          if (!ppNum || ppNum === '-' || ppNum === '0') {
            for (const [k, v] of Object.entries(row)) {
              const lk = k.toLowerCase();
              if (
                (lk.includes('pp') ||
                  lk.includes('documento') ||
                  lk.includes('ordem bancaria') ||
                  lk.includes('ob') ||
                  lk.includes('numero') ||
                  lk.includes('número')) &&
                v
              ) {
                const s = String(v).trim();
                if (s && s !== '-' && s !== '0') {
                  ppNum = s;
                  break;
                }
              }
            }
          }
          if (ppNum && ppNum !== '-' && ppNum !== '0') {
            currentPPsList.push({
              ppNumero: ppNum,
              obNumero: rawOB,
              processo: rawProc,
              credor: name,
              cnpj,
              cnpjDigits,
              valor: val,
              situacao: rawSit || 'AO',
              fileName: file.fileName,
            });
          }
        }
      });
    });

    // Also include OBs from ordensBancariasList into target PPs
    if (ordensBancariasList && ordensBancariasList.length > 0) {
      ordensBancariasList.forEach((obItem) => {
        const cnpjDigits = (obItem.favorecidoCnpj || '').replace(/\D/g, '');
        currentPPsList.push({
          ppNumero: obItem.pp || obItem.ob,
          obNumero: obItem.ob,
          processo: '',
          credor: obItem.favorecidoName,
          cnpj: obItem.favorecidoCnpj,
          cnpjDigits,
          valor: obItem.valor,
          situacao: obItem.situacao || 'CB',
          fileName: 'Listar Ordem Bancária (SIGEF)',
        });
      });
    }

    // 3. Load previously known NLs for temporal transitions
    let previouslyKnownNLs: Record<
      string,
      {
        nlNumero: string;
        processo: string;
        credor: string;
        cnpj: string;
        cnpjDigits: string;
        valor: number;
        stage: 'obedece' | 'nao_obedece';
        registeredAt: string;
      }
    > = {};

    try {
      const raw = localStorage.getItem(STORAGE_KEY_KNOWN_NLS);
      if (raw) {
        previouslyKnownNLs = JSON.parse(raw);
      }
    } catch (e) {
      previouslyKnownNLs = {};
    }

    // Pre-index currentPPsList by CNPJ and Credor for O(1) matching
    const ppsByCnpj = new Map<string, typeof currentPPsList>();
    const ppsByCredor = new Map<string, typeof currentPPsList>();
    currentPPsList.forEach((pp) => {
      if (pp.cnpjDigits) {
        let arr = ppsByCnpj.get(pp.cnpjDigits);
        if (!arr) {
          arr = [];
          ppsByCnpj.set(pp.cnpjDigits, arr);
        }
        arr.push(pp);
      }
      if (pp.credor && pp.credor.length >= 4) {
        const norm = pp.credor.toUpperCase().trim();
        let arr = ppsByCredor.get(norm);
        if (!arr) {
          arr = [];
          ppsByCredor.set(norm, arr);
        }
        arr.push(pp);
      }
    });

    const existingMatchIds = new Set(this.inferredMatches.map((m) => m.id));

    // Helper: evaluate match between an NL and candidate PPs
    const evaluateAndAddMatch = (
      nl: {
        nlNumero: string;
        processo: string;
        credor: string;
        cnpj: string;
        cnpjDigits: string;
        valor: number;
        stage: 'obedece' | 'nao_obedece';
      },
      isTemporalTransition: boolean = false
    ) => {
      const matchingPPs = (nl.cnpjDigits ? ppsByCnpj.get(nl.cnpjDigits) : null) ||
        (nl.credor ? ppsByCredor.get(nl.credor.toUpperCase().trim()) : null) || [];

      matchingPPs.forEach((pp) => {
        const valorNL = nl.valor;
        const valorPP = pp.valor;
        const dif = valorNL - valorPP;
        const pctRetention = dif >= 0 && valorNL > 0 ? dif / valorNL : 0;

        // Plausible match: exact value OR net value with tax retention up to 45%
        const isExact = Math.abs(valorNL - valorPP) < 0.05;
        const isRetention = valorPP < valorNL && pctRetention > 0 && pctRetention <= 0.45;

        if (isExact || isRetention) {
          const matchId = `inf_${nl.nlNumero}_${pp.ppNumero}`;
          if (!existingMatchIds.has(matchId)) {
            existingMatchIds.add(matchId);
            const patternKey = nl.cnpjDigits || nl.credor.toUpperCase().trim();
            const learnedPattern = this.learnedPatterns[patternKey];

            let confianca = isExact ? 99 : 88;
            let autoConfirmado = false;
            let statusValidacao: 'pendente' | 'confirmado' = 'pendente';
            let motivo = isExact
              ? `Valor idêntico exato (R$ ${valorPP.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}). Transição direta da NL para PP/OB.`
              : `Retenção de impostos de ${(pctRetention * 100).toFixed(1)}% deduzida entre a NL (R$ ${valorNL.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}) e a PP (R$ ${valorPP.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}).`;

            if (learnedPattern) {
              const difFromLearned = Math.abs(pctRetention - learnedPattern.aliquotaImpostoRetidoMedia);
              if (difFromLearned <= learnedPattern.toleranciaPercentual) {
                confianca = 98;
                autoConfirmado = true;
                statusValidacao = 'confirmado';
                motivo = `Auto-confirmado via IA Local: taxa de retenção ${(pctRetention * 100).toFixed(1)}% compatível com padrão memorizado (${(learnedPattern.aliquotaImpostoRetidoMedia * 100).toFixed(1)}%).`;
              } else {
                confianca = 92;
                motivo = `Alíquota de retenção ${(pctRetention * 100).toFixed(1)}% próxima do histórico memorizado do credor.`;
              }
            } else if (isTemporalTransition) {
              motivo = `Transição entre uploads: NL ${nl.nlNumero} saiu da fila de ${nl.stage} e gerou a PP ${pp.ppNumero} (${pp.situacao}).`;
            }

            const newInference: InferredPaymentMatch = {
              id: matchId,
              nlNumero: nl.nlNumero,
              processo: nl.processo || pp.processo,
              credor: nl.credor,
              cnpj: nl.cnpj,
              valorNL,
              ppNumero: pp.ppNumero,
              obNumero: pp.obNumero,
              valorPP,
              diferencaImpostos: dif > 0 ? dif : 0,
              percentualRetencao: pctRetention,
              statusSituacaoPP: pp.situacao,
              dataInferido: new Date().toISOString(),
              statusValidacao,
              confianca,
              motivo,
              baseOrigem: nl.stage,
              automaticamenteConfirmado: autoConfirmado,
            };

            this.inferredMatches.unshift(newInference);
          }
        }
      });
    };

    // 4. CROSS-MATCHING: Match current active NLs with current PPs/OBs
    Object.values(currentActiveNLs).forEach((nl) => {
      evaluateAndAddMatch(nl, false);
    });

    // 4.1 CROSS-MATCHING: Match PPs from files with OBs from ordensBancariasList using Indexed Maps
    if (ordensBancariasList && ordensBancariasList.length > 0 && currentPPsList.length > 0) {
      const obsByPP = new Map<string, OrdemBancariaItem[]>();
      const obsByCnpj = new Map<string, OrdemBancariaItem[]>();
      ordensBancariasList.forEach((ob) => {
        const ppKey = (ob.pp || '').trim().toUpperCase().replace(/[\s\.\-\/]/g, '');
        if (ppKey) {
          let arr = obsByPP.get(ppKey);
          if (!arr) {
            arr = [];
            obsByPP.set(ppKey, arr);
          }
          arr.push(ob);
        }
        const cnpjKey = (ob.favorecidoCnpj || '').replace(/\D/g, '');
        if (cnpjKey) {
          let arr = obsByCnpj.get(cnpjKey);
          if (!arr) {
            arr = [];
            obsByCnpj.set(cnpjKey, arr);
          }
          arr.push(ob);
        }
      });

      for (const pp of currentPPsList) {
        if (pp.fileName === 'Listar Ordem Bancária (SIGEF)') continue;
        const ppClean = (pp.ppNumero || '').trim().toUpperCase().replace(/[\s\.\-\/]/g, '');
        let matchingOBs: OrdemBancariaItem[] = [];
        if (ppClean && obsByPP.has(ppClean)) {
          matchingOBs = obsByPP.get(ppClean)!;
        } else if (pp.cnpjDigits && obsByCnpj.has(pp.cnpjDigits)) {
          matchingOBs = obsByCnpj.get(pp.cnpjDigits)!.slice(0, 5);
        }

        matchingOBs.forEach((ob) => {
          const valorPP = pp.valor;
          const valorOB = ob.valor;
          const dif = valorPP - valorOB;
          const pctRetention = dif >= 0 && valorPP > 0 ? dif / valorPP : 0;
          const isExact = Math.abs(valorPP - valorOB) < 0.05;
          const isRetention = valorOB < valorPP && pctRetention > 0 && pctRetention <= 0.45;

          if (isExact || isRetention) {
            const matchId = `inf_pp_${pp.ppNumero}_ob_${ob.ob}`;
            if (!existingMatchIds.has(matchId)) {
              existingMatchIds.add(matchId);
              const patternKey = pp.cnpjDigits || pp.credor.toUpperCase().trim();
              const learnedPattern = this.learnedPatterns[patternKey];
              let confianca = isExact ? 99 : 90;
              let autoConfirmado = false;
              let statusValidacao: 'pendente' | 'confirmado' = 'pendente';
              let motivo = isExact
                ? `PP ${pp.ppNumero} convertida em OB ${ob.ob} com valor idêntico.`
                : `Retenção tributária de ${(pctRetention * 100).toFixed(1)}% deduzida entre a PP e a OB.`;

              if (learnedPattern) {
                const difFromLearned = Math.abs(pctRetention - learnedPattern.aliquotaImpostoRetidoMedia);
                if (difFromLearned <= learnedPattern.toleranciaPercentual) {
                  confianca = 98;
                  autoConfirmado = true;
                  statusValidacao = 'confirmado';
                  motivo = `Auto-confirmado via padrão memorizado (${(learnedPattern.aliquotaImpostoRetidoMedia * 100).toFixed(1)}%).`;
                }
              }

              this.inferredMatches.unshift({
                id: matchId,
                nlNumero: pp.ppNumero,
                processo: pp.processo,
                credor: pp.credor,
                cnpj: pp.cnpj,
                valorNL: valorPP,
                ppNumero: pp.ppNumero,
                obNumero: ob.ob,
                valorPP: valorOB,
                diferencaImpostos: dif > 0 ? dif : 0,
                percentualRetencao: pctRetention,
                statusSituacaoPP: ob.situacao || 'CB',
                dataInferido: new Date().toISOString(),
                statusValidacao,
                confianca,
                motivo,
                baseOrigem: 'pp_prontas',
                automaticamenteConfirmado: autoConfirmado,
              });
            }
          }
        });
      }
    }

    // 5. TEMPORAL TRANSITIONS: Check previously known NLs that are no longer in active liquidation queue
    Object.keys(previouslyKnownNLs).forEach((prevKey) => {
      const oldNL = previouslyKnownNLs[prevKey];
      const isStillActive = !!currentActiveNLs[prevKey];
      if (!isStillActive) {
        evaluateAndAddMatch(oldNL, true);
      }
    });

    // 6. Update known NLs registry for next comparative upload
    const updatedKnownNLs: Record<string, any> = { ...previouslyKnownNLs };
    Object.keys(currentActiveNLs).forEach((k) => {
      updatedKnownNLs[k] = {
        ...currentActiveNLs[k],
        registeredAt: new Date().toISOString(),
      };
    });

    try {
      localStorage.setItem(STORAGE_KEY_KNOWN_NLS, JSON.stringify(updatedKnownNLs));
    } catch (e) {
      // Storage quota safe
    }

    // Limit inference matches list to 250 most relevant
    if (this.inferredMatches.length > 250) {
      this.inferredMatches = this.inferredMatches.slice(0, 250);
    }

    this.notify();
  }
}

export const mlInferenceService = new MachineLearningInferenceService();

export function getPendingInferences(): InferredPaymentMatch[] {
  return mlInferenceService.getInferredMatches().filter((m) => m.statusValidacao === 'pendente');
}

export function getPendingMatchesCount(): number {
  return mlInferenceService.getPendingMatchesCount();
}
