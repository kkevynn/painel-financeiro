import { InferredPaymentMatch, LearnedCreditorPattern, FileData } from '../types';
import { parseCurrencyValue, parseFavorecidoField, formatBRL } from './excelParser';
import { OrdemBancariaItem, isFonte600OuDerivada, findFonteRecursoInRow } from './ordemBancariaParser';
import {
  extractRealNlFromRow,
  resolveCnpj,
  resolveSupplierName,
  formatCleanNL,
  isValidNL,
} from './appKnowledgeDirectory';

export interface ExactNlMappingItem {
  nl: string;
  pp: string;
  ob: string;
  valor?: number;
  manual: boolean;
  updatedAt: string;
}

const STORAGE_KEY_LEARNED_PATTERNS = 'painel_learned_creditor_patterns';
const STORAGE_KEY_INFERRED_MATCHES = 'painel_inferred_matches_cache';
const STORAGE_KEY_KNOWN_NLS = 'painel_known_nls_registry';
const STORAGE_KEY_EXACT_MAPPINGS = 'painel_exact_nl_mappings';

type MLChangeListener = () => void;

class MachineLearningInferenceService {
  private learnedPatterns: Record<string, LearnedCreditorPattern> = {};
  private exactNlMappings: Record<string, ExactNlMappingItem> = {};
  private inferredMatches: InferredPaymentMatch[] = [];
  private listeners: Set<MLChangeListener> = new Set();
  private lastProcessedFingerprint: string = '';

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
        const parsed = JSON.parse(patternsRaw);
        if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
          this.learnedPatterns = parsed;
        } else {
          this.learnedPatterns = {};
        }
      } else {
        this.learnedPatterns = {};
      }
    } catch (e) {
      this.learnedPatterns = {};
    }

    try {
      const mappingsRaw = localStorage.getItem(STORAGE_KEY_EXACT_MAPPINGS);
      if (mappingsRaw) {
        this.exactNlMappings = JSON.parse(mappingsRaw) || {};
      }
    } catch (e) {
      this.exactNlMappings = {};
    }

    try {
      const matchesRaw = localStorage.getItem(STORAGE_KEY_INFERRED_MATCHES);
      if (matchesRaw) {
        const parsed = JSON.parse(matchesRaw);
        if (Array.isArray(parsed)) {
          // Strictly keep only liquidation transitions (Obedece / Não Obedece -> A Pagar) and CNPJ only (exclude CPF)
          this.inferredMatches = parsed.filter((m) => {
            if (!m || m.baseOrigem === 'pp_prontas' || m.id?.startsWith('inf_pp_')) return false;
            const digits = (m.cnpj || '').replace(/\D/g, '');
            // Reject CPFs (11 digits)
            if (digits.length === 11) return false;
            return true;
          });
        } else {
          this.inferredMatches = [];
        }
      } else {
        this.inferredMatches = [];
      }
    } catch (e) {
      this.inferredMatches = [];
    }
  }

  private saveToStorage() {
    try {
      localStorage.setItem(STORAGE_KEY_LEARNED_PATTERNS, JSON.stringify(this.learnedPatterns || {}));
      localStorage.setItem(STORAGE_KEY_EXACT_MAPPINGS, JSON.stringify(this.exactNlMappings || {}));
      localStorage.setItem(STORAGE_KEY_INFERRED_MATCHES, JSON.stringify(this.inferredMatches || []));
    } catch (e) {
      // LocalStorage quota protection
    }
  }

  public getExactNlMappings(): Record<string, ExactNlMappingItem> {
    return { ...this.exactNlMappings };
  }

  public getAllLearnedPatterns(): Record<string, LearnedCreditorPattern> {
    return { ...this.learnedPatterns };
  }

  public restoreLearnedPatterns(
    patterns?: Record<string, LearnedCreditorPattern>,
    mappings?: Record<string, ExactNlMappingItem>
  ) {
    if (patterns && typeof patterns === 'object') {
      this.learnedPatterns = { ...this.learnedPatterns, ...patterns };
    }
    if (mappings && typeof mappings === 'object') {
      this.exactNlMappings = { ...this.exactNlMappings, ...mappings };
    }
    this.saveToStorage();
    this.notify();
  }

  public getLearnedPatterns(): LearnedCreditorPattern[] {
    return Object.values(this.learnedPatterns || {}).filter((p) => {
      const digits = (p.cnpj || '').replace(/\D/g, '');
      return digits.length !== 11;
    });
  }

  public getInferredMatches(): InferredPaymentMatch[] {
    if (!Array.isArray(this.inferredMatches)) return [];
    return this.inferredMatches.filter((m) => {
      if (!m) return false;
      const digits = (m.cnpj || '').replace(/\D/g, '');
      return digits.length !== 11;
    });
  }

  public getPendingMatchesCount(): number {
    return this.getInferredMatches().filter((m) => m && m.statusValidacao === 'pendente').length;
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

  public reopenMatch(id: string) {
    const match = this.inferredMatches.find((m) => m.id === id);
    if (!match) return;
    match.statusValidacao = 'pendente';
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
   * Manually adds an NL -> PP/OB match and immediately memorizes the pattern in the model
   */
  public addManualInferredMatch(params: {
    nlNumero: string;
    credor: string;
    cnpj?: string;
    valorNL: number;
    ppNumero: string;
    obNumero?: string;
    situacao?: string;
    valorPP?: number;
    fonteRecurso?: string;
  }): InferredPaymentMatch {
    const valorNL = Math.max(0, Number(params.valorNL) || 0);
    const valorPP = typeof params.valorPP === 'number' && params.valorPP > 0 ? params.valorPP : valorNL;
    const diferencaImpostos = Math.max(0, valorNL - valorPP);
    const percentualRetencao = valorNL > 0 ? diferencaImpostos / valorNL : 0;

    const newMatch: InferredPaymentMatch = {
      id: `manual_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
      nlNumero: params.nlNumero.trim(),
      credor: params.credor.trim(),
      cnpj: params.cnpj?.trim() || '',
      fonteRecurso: params.fonteRecurso || '0.6.00',
      valorNL,
      baseOrigem: 'obedece',
      ppNumero: params.ppNumero.trim(),
      obNumero: params.obNumero?.trim() || undefined,
      valorPP,
      statusSituacaoPP: params.situacao?.trim() || 'AO',
      diferencaImpostos,
      percentualRetencao,
      confianca: 100,
      motivo: 'Ligação manual registrada pelo usuário (NL virou PP/OB)',
      statusValidacao: 'confirmado',
      dataInferido: new Date().toISOString(),
      isFonte600: isFonte600OuDerivada(params.fonteRecurso || '0.6.00'),
    };

    // Add to matches
    this.inferredMatches = [newMatch, ...this.inferredMatches];

    // Train the creditor pattern
    this.recordManualPattern({
      credor: newMatch.credor,
      cnpj: newMatch.cnpj,
      valorNL: newMatch.valorNL,
      valorPP: newMatch.valorPP,
      nlNumero: newMatch.nlNumero,
      ppNumero: newMatch.ppNumero,
      obNumero: newMatch.obNumero,
    });

    this.notify();
    return newMatch;
  }

  /**
   * Updates the PP and OB information for an existing match and updates the learning pattern
   */
  public updateInferredMatch(
    id: string,
    updates: {
      ppNumero?: string;
      obNumero?: string;
      situacao?: string;
      valorPP?: number;
    }
  ): boolean {
    const match = this.inferredMatches.find((m) => m.id === id);
    if (!match) return false;

    if (updates.ppNumero !== undefined) {
      match.ppNumero = updates.ppNumero.trim();
    }
    if (updates.obNumero !== undefined) {
      match.obNumero = updates.obNumero.trim() || undefined;
    }
    if (updates.situacao !== undefined) {
      match.statusSituacaoPP = updates.situacao.trim() || 'AO';
    }
    if (updates.valorPP !== undefined && !isNaN(updates.valorPP)) {
      match.valorPP = Math.max(0, updates.valorPP);
      match.diferencaImpostos = Math.max(0, match.valorNL - match.valorPP);
      match.percentualRetencao = match.valorNL > 0 ? match.diferencaImpostos / match.valorNL : 0;
    }

    match.motivo = `${match.motivo} (Editado pelo usuário)`;
    match.statusValidacao = 'confirmado';

    // Train or update pattern with the edited data
    this.recordManualPattern({
      credor: match.credor,
      cnpj: match.cnpj,
      valorNL: match.valorNL,
      valorPP: match.valorPP,
      nlNumero: match.nlNumero,
      ppNumero: match.ppNumero,
      obNumero: match.obNumero,
    });

    this.notify();
    return true;
  }

  /**
   * Records a manual PP/OB association pattern for a creditor (from Publicação view)
   * so the AI pattern model learns the tax retention rate and creditor mapping
   */
  public recordManualPattern(params: {
    credor: string;
    cnpj?: string;
    valorNL: number;
    valorPP: number;
    nlNumero: string;
    ppNumero?: string;
    obNumero?: string;
  }) {
    const valorNL = params.valorNL || 0;
    const valorPP = params.valorPP || 0;
    const dif = valorNL - valorPP;
    const pctRetention = dif > 0 && valorNL > 0 ? dif / valorNL : 0;
    const cnpjClean = (params.cnpj || '').replace(/\D/g, '');
    const patternKey = cnpjClean.length >= 11 ? cnpjClean : (params.credor || '').toUpperCase().trim();
    if (!patternKey) return;

    const existingPattern = this.learnedPatterns[patternKey];
    if (existingPattern) {
      const totalCount = existingPattern.totalConfirmacoes + 1;
      const updatedRate =
        (existingPattern.aliquotaImpostoRetidoMedia * existingPattern.totalConfirmacoes + pctRetention) / totalCount;

      this.learnedPatterns[patternKey] = {
        cnpj: params.cnpj || existingPattern.cnpj,
        credor: params.credor || existingPattern.credor,
        aliquotaImpostoRetidoMedia: Number(updatedRate.toFixed(4)),
        toleranciaPercentual: 0.03,
        totalConfirmacoes: totalCount,
        ultimaConfirmacao: new Date().toISOString(),
        sampleNL: params.nlNumero,
        samplePP: params.ppNumero || (params.obNumero ? `OB ${params.obNumero}` : 'Manual'),
      };
      this.learnedPatterns[patternKey] = {
        cnpj: params.cnpj,
        credor: params.credor,
        aliquotaImpostoRetidoMedia: Number(pctRetention.toFixed(4)),
        toleranciaPercentual: 0.03,
        totalConfirmacoes: 1,
        ultimaConfirmacao: new Date().toISOString(),
        sampleNL: params.nlNumero,
        samplePP: params.ppNumero || (params.obNumero ? `OB ${params.obNumero}` : 'Manual'),
      };
    }

    const cleanNl = (params.nlNumero || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
    if (cleanNl) {
      this.exactNlMappings[cleanNl] = {
        nl: params.nlNumero,
        pp: params.ppNumero || '',
        ob: params.obNumero || '',
        valor: params.valorPP,
        manual: true,
        updatedAt: new Date().toISOString(),
      };
    }

    this.saveToStorage();
    this.notify();
  }

  /**
   * Cruzamento de Liquidações (Obedece e Não Obedece) para saber se passaram para A Pagar (PPs):
   * 1. Cruzamento direto: identifica se uma liquidação (NL Obedece ou Não Obedece) passou para "A Pagar" (PP Pronta / Base de PPs).
   * 2. Transição temporal: detecta NLs que saíram da fila de liquidação entre uploads e entraram na fila de PPs a pagar.
   * 3. Retenção tributária: calcula a alíquota de impostos retidos entre o valor bruto da liquidação e o valor líquido da PP.
   */
  public processFilesForInference(
    files: FileData[],
    _ordensBancariasList?: OrdemBancariaItem[] | null,
    force: boolean = false
  ) {
    try {
      if (!files || files.length === 0) return;

      // Fingerprint to avoid redundant heavy recalculations on modal open
      const fp = `${files.map((f) => `${f.fileName}_${f.rows?.length || 0}`).join('|')}_OB_${_ordensBancariasList?.length || 0}`;
      if (!force && fp === this.lastProcessedFingerprint && this.inferredMatches.length > 0) {
        return;
      }
      this.lastProcessedFingerprint = fp;

      // 1. Extract currently active NLs from Liquidation files (Obedece e Não Obedece)
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
          fonteRecurso: string;
          isFonte600: boolean;
        }
      > = {};

      // 2. Extract current PPs ("A Pagar" / Base de PPs / PPs Prontas) & OBs
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

      // Ingest OBs from ordensBancariasList (from Detalhamento de Pagamentos report)
      if (_ordensBancariasList && _ordensBancariasList.length > 0) {
        _ordensBancariasList.forEach((it) => {
          const cDigits = (it.favorecidoCnpj || '').replace(/\D/g, '');
          currentPPsList.push({
            ppNumero: it.pp || (it.ob ? `OB ${it.ob}` : 'S/N'),
            obNumero: it.ob || '',
            processo: '',
            credor: it.favorecidoName,
            cnpj: it.favorecidoCnpj || '',
            cnpjDigits: cDigits,
            valor: it.valor,
            situacao: it.situacao || 'CB',
            fileName: 'Listar Ordem Bancária',
          });
        });
      }

      // Process uploaded files
      (files || []).forEach((file) => {
        const stage = file.stage || 'outros';
        const isLiquidationFile =
          stage === 'obedece' ||
          stage === 'nao_obedece' ||
          (file.fileName && file.fileName.toLowerCase().includes('liquida')) ||
          (file.fileName && file.fileName.toLowerCase().includes('obedece'));

        const isPPFile =
          stage === 'base_pps' ||
          stage === 'pp_prontas' ||
          (!isLiquidationFile &&
            !file.fileName.toLowerCase().includes('ordem bancaria') &&
            !file.fileName.toLowerCase().includes('ordens bancarias') &&
            !file.fileName.toLowerCase().includes('listar ordem') &&
            !file.fileName.toLowerCase().includes('pagamentos emitidos'));

        const favCol = file.detectedFavorecidoCol;
        const valCol = file.detectedValorCol;
        const numCol = file.detectedNumeroCol || file.detectedPpCol;
        const obCol = file.detectedObCol;
        const sitCol = file.detectedSituacaoCol;
        const procCol = file.detectedProcessoCol;
        const fonteCol = file.detectedFonteCol;

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

          // RULE: "Na análise de auditoria, está me mostrando liquidações de cpf, quero que me mostre apenas de cnpj"
          // Strictly exclude CPFs (11 digits, pessoa física) from audit analysis
          if (cnpjDigits.length === 11 || (cnpjDigits.length > 0 && cnpjDigits.length !== 14)) {
            return;
          }

          // Extract real NL (strictly rejecting line numbers)
          let docNum = extractRealNlFromRow(row, numCol);
          if (!docNum && rawNum && isValidNL(rawNum)) {
            docNum = formatCleanNL(rawNum);
          }
          if (!docNum) {
            for (const [k, v] of Object.entries(row)) {
              if (!v) continue;
              const lk = k.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
              if (lk.includes('empenho') || lk.includes(' ne ')) continue;
              if (
                lk.includes('nl') ||
                lk.includes('liquidacao') ||
                lk.includes('liquidac') ||
                lk === 'documento' ||
                lk === 'doc'
              ) {
                const s = String(v).trim();
                if (s && s !== '-' && s !== '0' && isValidNL(s)) {
                  docNum = formatCleanNL(s);
                  break;
                }
              }
            }
          }

          // Auto-resolve CNPJ from central directory if missing
          const resolvedCnpj = (cnpj && cnpj !== 'N/I' && cnpj !== '-') ? cnpj : resolveCnpj(name, docNum, cnpj);
          const effectiveDigits = (resolvedCnpj || '').replace(/\D/g, '');

          // Strictly exclude CPFs (11 digits, pessoa física) from audit analysis
          if (effectiveDigits.length === 11 || (effectiveDigits.length > 0 && effectiveDigits.length !== 14)) {
            return;
          }

          const effectiveCredor = resolveSupplierName(resolvedCnpj, docNum, name);

          if (isLiquidationFile) {
            // Extract Fonte de Recursos and check if it is F600 or derivative
            let rawFonte = fonteCol ? String(row[fonteCol] || '').trim() : '';
            if (!rawFonte || rawFonte === '-' || rawFonte === '0') {
              rawFonte = findFonteRecursoInRow(Object.values(row));
            }
            const isF600 = isFonte600OuDerivada(rawFonte);

            if (docNum) {
              const key = `${effectiveDigits || effectiveCredor}_${docNum}`;
              currentActiveNLs[key] = {
                nlNumero: docNum,
                processo: rawProc,
                credor: effectiveCredor,
                cnpj: resolvedCnpj,
                cnpjDigits: effectiveDigits,
                valor: val,
                stage: stage === 'nao_obedece' ? 'nao_obedece' : 'obedece',
                fileName: file.fileName,
                fonteRecurso: rawFonte || (isF600 ? '0.6.00' : ''),
                isFonte600: isF600,
              };
            }
          } else if (isPPFile) {
            // It's a PP file ("A Pagar" / Base de PPs)
            let ppNum = rawNum || rawOB;
            if (!ppNum || ppNum === '-' || ppNum === '0') {
              for (const [k, v] of Object.entries(row)) {
                const lk = k.toLowerCase();
                if (
                  (lk.includes('pp') ||
                    lk.includes('documento') ||
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
                situacao: rawSit || 'AO (A Pagar)',
                fileName: file.fileName,
              });
            }
          }
        });
      });

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

      // Helper: evaluate match between a Liquidation (NL) and candidate PPs ("A Pagar")
      const evaluateAndAddMatch = (
        nl: {
          nlNumero: string;
          processo: string;
          credor: string;
          cnpj: string;
          cnpjDigits: string;
          valor: number;
          stage: 'obedece' | 'nao_obedece';
          fonteRecurso?: string;
          isFonte600?: boolean;
        },
        isTemporalTransition: boolean = false
      ) => {
        const cleanNl = (nl.nlNumero || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
        const learnedExact = this.exactNlMappings[cleanNl];
        if (learnedExact) {
          const matchId = `inf_nl_${nl.nlNumero}_manual`;
          if (!existingMatchIds.has(matchId)) {
            existingMatchIds.add(matchId);
            const valPP = learnedExact.valor !== undefined && learnedExact.valor > 0 ? learnedExact.valor : nl.valor;
            const dif = Math.max(0, nl.valor - valPP);
            const pct = nl.valor > 0 ? dif / nl.valor : 0;
            this.inferredMatches.unshift({
              id: matchId,
              nlNumero: nl.nlNumero,
              processo: nl.processo,
              credor: nl.credor,
              cnpj: nl.cnpj,
              valorNL: nl.valor,
              ppNumero: learnedExact.pp || 'PP Confirmada',
              obNumero: learnedExact.ob || undefined,
              valorPP: valPP,
              diferencaImpostos: dif,
              percentualRetencao: pct,
              statusSituacaoPP: 'AO (Confirmado Manualmente)',
              confianca: 100,
              motivo: `Vínculo manual aprendido: Esta NL foi vinculada à PP ${learnedExact.pp || '-'} ${learnedExact.ob ? `/ OB ${learnedExact.ob}` : ''}.`,
              baseOrigem: nl.stage,
              automaticamenteConfirmado: true,
              statusValidacao: 'confirmado',
              dataInferido: learnedExact.updatedAt || new Date().toISOString(),
            });
            return;
          }
        }

        const matchingPPs =
          (nl.cnpjDigits ? ppsByCnpj.get(nl.cnpjDigits) : null) ||
          (nl.credor ? ppsByCredor.get(nl.credor.toUpperCase().trim()) : null) ||
          [];

        matchingPPs.forEach((pp) => {
          const valorNL = nl.valor;
          const valorPP = pp.valor;
          const dif = valorNL - valorPP;
          const pctRetention = dif >= 0 && valorNL > 0 ? dif / valorNL : 0;

          // Plausible match: exact value OR net value with tax retention up to 45%
          const isExact = Math.abs(valorNL - valorPP) < 0.05;
          const isRetention = valorPP < valorNL && pctRetention > 0 && pctRetention <= 0.45;
          const matchesProcesso = !!(
            nl.processo &&
            pp.processo &&
            nl.processo.replace(/\D/g, '') === pp.processo.replace(/\D/g, '') &&
            nl.processo.length >= 4
          );

          if (isExact || isRetention || matchesProcesso) {
            const matchId = `inf_nl_${nl.nlNumero}_pp_${pp.ppNumero}_ob_${pp.obNumero || 'none'}`;
            if (!existingMatchIds.has(matchId)) {
              existingMatchIds.add(matchId);
              const patternKey = nl.cnpjDigits || nl.credor.toUpperCase().trim();
              const learnedPattern = this.learnedPatterns[patternKey];

              let confianca = isExact ? 99 : 88;
              let motivo = isExact
                ? `Despesa equivalente de valor idêntico (${formatBRL(valorPP)}) do fornecedor ${nl.credor}. Ligação da Liquidação NL ${nl.nlNumero} (${nl.stage === 'obedece' ? 'Obedece' : 'Não Obedece'}${nl.isFonte600 ? ' - Fonte 600' : ''}) à PP ${pp.ppNumero}${pp.obNumero ? ` / OB ${pp.obNumero}` : ''}.`
                : `Despesa equivalente do fornecedor ${nl.credor} com dedução tributária de ${(pctRetention * 100).toFixed(1).replace('.', ',')}% (NL ${formatBRL(valorNL)} → PP ${formatBRL(valorPP)}${pp.obNumero ? ` / OB ${pp.obNumero}` : ''}).`;

              if (learnedPattern) {
                const difFromLearned = Math.abs(pctRetention - learnedPattern.aliquotaImpostoRetidoMedia);
                if (difFromLearned <= learnedPattern.toleranciaPercentual) {
                  confianca = 98;
                  motivo = `Ligação baseada em eventos passados deste fornecedor (${learnedPattern.credor}): retenção aprendida de ${(learnedPattern.aliquotaImpostoRetidoMedia * 100).toFixed(1).replace('.', ',')}% (amostra anterior NL ${learnedPattern.sampleNL || 'anterior'} → PP ${learnedPattern.samplePP || 'anterior'}). Confirma esta ligação?`;
                } else {
                  confianca = 92;
                  motivo = `Ligação com retenção de ${(pctRetention * 100).toFixed(1).replace('.', ',')}% próxima do histórico de eventos passados deste fornecedor (${(learnedPattern.aliquotaImpostoRetidoMedia * 100).toFixed(1).replace('.', ',')}%). Requer validação.`;
                }
              } else if (isTemporalTransition) {
                motivo = `Transição temporal: NL ${nl.nlNumero} da Liquidação (${nl.stage === 'obedece' ? 'Obedece' : 'Não Obedece'}${nl.isFonte600 ? ' - Fonte 600' : ''}) vinculada à PP ${pp.ppNumero}${pp.obNumero ? ` / OB ${pp.obNumero}` : ''} de ${nl.credor}.`;
              }

              // Inferences are always 'pendente' so the user is asked to validate
              const statusValidacao: 'pendente' | 'confirmado' = 'pendente';
              const autoConfirmado = false;

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
                statusSituacaoPP: pp.situacao || 'AO (A Pagar)',
                dataInferido: new Date().toISOString(),
                statusValidacao,
                confianca,
                motivo,
                baseOrigem: nl.stage,
                automaticamenteConfirmado: autoConfirmado,
                isFonte600: nl.isFonte600,
                fonteRecurso: nl.fonteRecurso,
              };

              this.inferredMatches.unshift(newInference);
            }
          }
        });
      };

      // 4. CROSS-MATCHING: Match current active NLs (Liquidações Obedece e Não Obedece) with PPs a Pagar
      Object.values(currentActiveNLs).forEach((nl) => {
        evaluateAndAddMatch(nl, false);
      });

      // 5. TEMPORAL TRANSITIONS: Check previously known NLs that are no longer in active liquidation queue
      Object.keys(previouslyKnownNLs).forEach((prevKey) => {
        const oldNL = previouslyKnownNLs[prevKey];
        const isStillActive = !!currentActiveNLs[prevKey];
        if (!isStillActive) {
          evaluateAndAddMatch(oldNL, true);
        }
      });

    // 6. Update known NLs registry for next comparative upload (compacted to prevent UI lag)
    const updatedKnownNLs: Record<string, any> = {};
    const activeKeys = Object.keys(currentActiveNLs).slice(0, 300);
    activeKeys.forEach((k) => {
      const it = currentActiveNLs[k];
      updatedKnownNLs[k] = {
        nlNumero: it.nlNumero,
        valor: it.valor,
        stage: it.stage,
        cnpjDigits: it.cnpjDigits,
        credor: it.credor,
        isFonte600: it.isFonte600,
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
    } catch (err) {
      console.warn('Erro ao processar inferências ML:', err);
    }
  }
}

export const mlInferenceService = new MachineLearningInferenceService();

export function getPendingInferences(): InferredPaymentMatch[] {
  return mlInferenceService.getInferredMatches().filter((m) => m.statusValidacao === 'pendente');
}

export function getPendingMatchesCount(): number {
  return mlInferenceService.getPendingMatchesCount();
}
