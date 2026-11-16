import React, { useState, useMemo, useEffect } from 'react';
import {
  X,
  Check,
  Settings2,
  Hash,
  CreditCard,
  Building2,
  FileText,
  UserCheck,
  DollarSign,
  PieChart,
  Layers,
  Calendar,
} from 'lucide-react';
import { FileData, StageKey, ColumnMappingConfig } from '../types';
import { formatBRL, parseCurrencyValue, analyzeColumnsForStage } from '../utils/excelParser';

interface ColumnMappingModalProps {
  fileData: FileData | null;
  onClose: () => void;
  onSaveMapping: (fileId: string, mapping: ColumnMappingConfig) => void;
}

export const ColumnMappingModal: React.FC<ColumnMappingModalProps> = ({
  fileData,
  onClose,
  onSaveMapping,
}) => {
  const headers = fileData?.headers || [];

  const [selectedStage, setSelectedStage] = useState<StageKey>(
    fileData?.stage || 'outros'
  );

  // 6 Primary Mapping Columns requested by user
  const [selectedNumero, setSelectedNumero] = useState<string>(() => {
    return fileData?.detectedNumeroCol || fileData?.detectedPpCol || '';
  });

  const [selectedOb, setSelectedOb] = useState<string>(() => {
    return fileData?.detectedObCol || '';
  });

  const [selectedFavorecido, setSelectedFavorecido] = useState<string>(() => {
    return fileData?.detectedFavorecidoCol || headers[0] || '';
  });

  const [selectedNotaEmpenho, setSelectedNotaEmpenho] = useState<string>(() => {
    if (fileData?.detectedNotaEmpenhoCol) return fileData.detectedNotaEmpenhoCol;
    const found = headers.find((h) => {
      const norm = h.toLowerCase().trim();
      if (norm.includes('favorecido')) return false;
      return (
        norm.includes('nota de empenho') ||
        norm.includes('nota empenho') ||
        norm === 'ne' ||
        norm === 'empenho'
      );
    });
    return found || '';
  });

  const [selectedFavorecidoNE, setSelectedFavorecidoNE] = useState<string>(() => {
    if (fileData?.detectedFavorecidoNECol) return fileData.detectedFavorecidoNECol;
    const found = headers.find((h) => {
      const norm = h.toLowerCase().trim();
      return (
        norm.includes('favorecido nota empenho') ||
        norm.includes('favorecido da ne') ||
        norm.includes('favorecido ne') ||
        norm.includes('credor ne')
      );
    });
    return found || '';
  });

  const [selectedValor, setSelectedValor] = useState<string>(() => {
    return fileData?.detectedValorCol || headers[1] || '';
  });

  // Additional secondary columns
  const [selectedCnpj, setSelectedCnpj] = useState<string>(
    fileData?.detectedCnpjCol || ''
  );
  const [selectedSituacao, setSelectedSituacao] = useState<string>(
    fileData?.detectedSituacaoCol || ''
  );
  const [selectedNatureza, setSelectedNatureza] = useState<string>(
    fileData?.detectedNaturezaCol || ''
  );
  const [selectedFonte, setSelectedFonte] = useState<string>(
    fileData?.detectedFonteCol || ''
  );
  const [selectedData, setSelectedData] = useState<string>(() => {
    if (fileData?.detectedDataCol) return fileData.detectedDataCol;
    const found = headers.find((h) => {
      const norm = h.toLowerCase().trim();
      return norm.includes('data') || norm.includes('dt');
    });
    return found || '';
  });

  const previewRow = useMemo(() => {
    if (!fileData?.rows || fileData.rows.length === 0) return null;
    const row = fileData.rows[0] || {};
    const rawVal = selectedValor ? row[selectedValor] : null;
    const parsedVal = parseCurrencyValue(rawVal);

    return {
      numero: selectedNumero ? String(row[selectedNumero] || '-') : '-',
      ob: selectedOb ? String(row[selectedOb] || '-') : '-',
      favorecido: selectedFavorecido ? String(row[selectedFavorecido] || '-') : '-',
      empenho: selectedNotaEmpenho ? String(row[selectedNotaEmpenho] || '-') : '-',
      favorecidoNE: selectedFavorecidoNE ? String(row[selectedFavorecidoNE] || '-') : '-',
      valor: parsedVal,
      cnpj: selectedCnpj ? String(row[selectedCnpj] || '-') : '-',
      situacao: selectedSituacao ? String(row[selectedSituacao] || '-') : '-',
      fonte: selectedFonte ? String(row[selectedFonte] || '-') : '-',
      data: selectedData ? String(row[selectedData] || '-') : '-',
    };
  }, [
    fileData?.rows,
    selectedNumero,
    selectedOb,
    selectedFavorecido,
    selectedNotaEmpenho,
    selectedFavorecidoNE,
    selectedValor,
    selectedCnpj,
    selectedSituacao,
    selectedNatureza,
    selectedFonte,
    selectedData,
  ]);

  const handleStageChange = (newStage: StageKey) => {
    setSelectedStage(newStage);

    // Dynamic analysis of file contents (rows & headers) for the newly selected stage
    const analyzed = analyzeColumnsForStage(fileData.headers, fileData.rows, newStage);
    if (analyzed.numeroCol !== undefined) setSelectedNumero(analyzed.numeroCol);
    if (analyzed.obCol !== undefined) setSelectedOb(analyzed.obCol);
    if (analyzed.favorecidoCol) setSelectedFavorecido(analyzed.favorecidoCol);
    if (analyzed.valorCol) setSelectedValor(analyzed.valorCol);
    if (analyzed.notaEmpenhoCol !== undefined) setSelectedNotaEmpenho(analyzed.notaEmpenhoCol);
    if (analyzed.favorecidoNECol !== undefined) setSelectedFavorecidoNE(analyzed.favorecidoNECol);
    if (analyzed.cnpjCol !== undefined) setSelectedCnpj(analyzed.cnpjCol);
    if (analyzed.situacaoObCol !== undefined) {
      setSelectedSituacao(analyzed.situacaoObCol);
    } else if (analyzed.situacaoCol !== undefined) {
      setSelectedSituacao(analyzed.situacaoCol);
    }
    if (analyzed.naturezaCol !== undefined) setSelectedNatureza(analyzed.naturezaCol);
    if (analyzed.fonteCol !== undefined) setSelectedFonte(analyzed.fonteCol);
    if (analyzed.dataCol !== undefined) setSelectedData(analyzed.dataCol);
  };

  const handleSave = () => {
    onSaveMapping(fileData.id, {
      favorecidoCol: selectedFavorecido,
      valorCol: selectedValor,
      stage: selectedStage,
      numeroCol: selectedNumero || undefined,
      obCol: selectedOb || undefined,
      notaEmpenhoCol: selectedNotaEmpenho || undefined,
      favorecidoNECol: selectedFavorecidoNE || undefined,
      cnpjCol: selectedCnpj || undefined,
      situacaoCol: selectedSituacao || undefined,
      situacaoObCol: fileData.detectedSituacaoObCol || undefined,
      naturezaCol: selectedNatureza || undefined,
      fonteCol: selectedFonte || undefined,
      dataCol: selectedData || undefined,
    });
    onClose();
  };

  // Lock background scroll when open
  useEffect(() => {
    if (!fileData) return;
    const orig = document.body.style.overflow;
    document.body.style.overflow = 'hidden';
    return () => {
      document.body.style.overflow = orig;
    };
  }, [fileData]);

  if (!fileData) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-4 bg-slate-900/60 backdrop-blur-xs overscroll-contain animate-in fade-in duration-200">
      <div className="bg-white rounded-2xl max-w-3xl w-full max-h-[92vh] shadow-2xl border border-slate-200 overflow-hidden flex flex-col overscroll-contain">
        
        {/* Modal Header */}
        <div className="p-4 sm:p-5 bg-slate-50 border-b border-slate-200 flex items-center justify-between shrink-0">
          <div className="flex items-center space-x-3">
            <div className="w-9 h-9 rounded-xl bg-emerald-100 text-emerald-800 flex items-center justify-center shrink-0">
              <Settings2 className="w-5 h-5" />
            </div>
            <div>
              <h3 className="text-base font-bold text-slate-900">
                Ajustar Mapeamento
              </h3>
              <p className="text-xs text-slate-500 font-medium truncate max-w-xs sm:max-w-md">
                {fileData.fileName} • {fileData.rows?.length || 0} registros
              </p>
            </div>
          </div>

          <button
            onClick={onClose}
            className="text-slate-400 hover:text-slate-700 p-1.5 rounded-lg hover:bg-slate-200/60 transition cursor-pointer"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Modal Body */}
        <div className="p-4 sm:p-6 overflow-y-auto space-y-5 text-xs text-slate-700">
          
          {/* Classificação no Painel */}
          <div className="space-y-2">
            <label className="block font-bold text-slate-900 text-xs">
              Classificação no Painel:
            </label>
            <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
              <button
                type="button"
                onClick={() => handleStageChange('nao_obedece')}
                className={`py-2 px-3 rounded-xl border text-xs font-bold transition flex items-center gap-2 cursor-pointer ${
                  selectedStage === 'nao_obedece'
                    ? 'border-amber-500 bg-amber-50 text-amber-950 ring-2 ring-amber-400/30'
                    : 'border-slate-200 bg-slate-50 text-slate-700 hover:bg-slate-100'
                }`}
              >
                <span className="w-2.5 h-2.5 rounded-full bg-amber-500 shrink-0" />
                <span className="truncate">Liquidado - Não Obedece</span>
              </button>

              <button
                type="button"
                onClick={() => handleStageChange('obedece')}
                className={`py-2 px-3 rounded-xl border text-xs font-bold transition flex items-center gap-2 cursor-pointer ${
                  selectedStage === 'obedece'
                    ? 'border-blue-500 bg-blue-50 text-blue-950 ring-2 ring-blue-400/30'
                    : 'border-slate-200 bg-slate-50 text-slate-700 hover:bg-slate-100'
                }`}
              >
                <span className="w-2.5 h-2.5 rounded-full bg-blue-500 shrink-0" />
                <span className="truncate">Liquidado - Obedece</span>
              </button>

              <button
                type="button"
                onClick={() => handleStageChange('base_pps')}
                className={`py-2 px-3 rounded-xl border text-xs font-bold transition flex items-center gap-2 cursor-pointer ${
                  selectedStage === 'base_pps' || selectedStage === 'pp_prontas'
                    ? 'border-emerald-500 bg-emerald-50 text-emerald-950 ring-2 ring-emerald-400/30'
                    : 'border-slate-200 bg-slate-50 text-slate-700 hover:bg-slate-100'
                }`}
              >
                <span className="w-2.5 h-2.5 rounded-full bg-emerald-500 shrink-0" />
                <span className="truncate">Base de PPs</span>
              </button>

              <button
                type="button"
                onClick={() => handleStageChange('ordem_bancaria')}
                className={`py-2 px-3 rounded-xl border text-xs font-bold transition flex items-center gap-2 cursor-pointer ${
                  selectedStage === 'ordem_bancaria'
                    ? 'border-teal-500 bg-teal-50 text-teal-950 ring-2 ring-teal-400/30'
                    : 'border-slate-200 bg-slate-50 text-slate-700 hover:bg-slate-100'
                }`}
              >
                <span className="w-2.5 h-2.5 rounded-full bg-teal-500 shrink-0" />
                <span className="truncate">Detalhamento de Pagamentos (OB)</span>
              </button>

              <button
                type="button"
                onClick={() => handleStageChange('natureza_despesa')}
                className={`py-2 px-3 rounded-xl border text-xs font-bold transition flex items-center gap-2 cursor-pointer ${
                  selectedStage === 'natureza_despesa'
                    ? 'border-violet-500 bg-violet-50 text-violet-950 ring-2 ring-violet-400/30'
                    : 'border-slate-200 bg-slate-50 text-slate-700 hover:bg-slate-100'
                }`}
              >
                <span className="w-2.5 h-2.5 rounded-full bg-violet-500 shrink-0" />
                <span className="truncate">Natureza Despesa</span>
              </button>

              <button
                type="button"
                onClick={() => handleStageChange('publicacao')}
                className={`py-2 px-3 rounded-xl border text-xs font-bold transition flex items-center gap-2 cursor-pointer ${
                  selectedStage === 'publicacao'
                    ? 'border-indigo-500 bg-indigo-50 text-indigo-950 ring-2 ring-indigo-400/30'
                    : 'border-slate-200 bg-slate-50 text-slate-700 hover:bg-slate-100'
                }`}
              >
                <span className="w-2.5 h-2.5 rounded-full bg-indigo-500 shrink-0" />
                <span className="truncate">Para Publicação</span>
              </button>

              <button
                type="button"
                onClick={() => handleStageChange('outros')}
                className={`py-2 px-3 rounded-xl border text-xs font-bold transition flex items-center gap-2 cursor-pointer ${
                  selectedStage === 'outros'
                    ? 'border-slate-500 bg-slate-100 text-slate-900 ring-2 ring-slate-400/30'
                    : 'border-slate-200 bg-slate-50 text-slate-700 hover:bg-slate-100'
                }`}
              >
                <span className="w-2.5 h-2.5 rounded-full bg-slate-400 shrink-0" />
                <span className="truncate">Outras Planilhas</span>
              </button>
            </div>
          </div>

          {/* 6 Requested Columns */}
          <div className="pt-2 border-t border-slate-200 space-y-3">
            <h4 className="font-bold text-slate-900 text-xs uppercase tracking-wider text-slate-600">
              Colunas de Mapeamento:
            </h4>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
              
              {/* 1. Número */}
              <div className="space-y-1">
                <label className="block font-bold text-slate-900 text-xs flex items-center gap-1.5">
                  <Hash className="w-3.5 h-3.5 text-slate-600" />
                  Número
                </label>
                <select
                  value={selectedNumero}
                  onChange={(e) => setSelectedNumero(e.target.value)}
                  className="w-full p-2 bg-slate-50 border border-slate-300 rounded-xl font-medium text-xs text-slate-900 focus:bg-white focus:outline-hidden focus:border-emerald-500"
                >
                  <option value="">Não mapeada</option>
                  {fileData.headers.map((h) => (
                    <option key={h} value={h}>
                      {h}
                    </option>
                  ))}
                </select>
              </div>

              {/* 2. OB */}
              <div className="space-y-1">
                <label className="block font-bold text-slate-900 text-xs flex items-center gap-1.5">
                  <CreditCard className="w-3.5 h-3.5 text-slate-600" />
                  OB
                </label>
                <select
                  value={selectedOb}
                  onChange={(e) => setSelectedOb(e.target.value)}
                  className="w-full p-2 bg-slate-50 border border-slate-300 rounded-xl font-medium text-xs text-slate-900 focus:bg-white focus:outline-hidden focus:border-emerald-500"
                >
                  <option value="">Não mapeada</option>
                  {fileData.headers.map((h) => (
                    <option key={h} value={h}>
                      {h}
                    </option>
                  ))}
                </select>
              </div>

              {/* 3. Favorecido */}
              <div className="space-y-1">
                <label className="block font-bold text-slate-900 text-xs flex items-center gap-1.5">
                  <Building2 className="w-3.5 h-3.5 text-slate-600" />
                  Favorecido
                </label>
                <select
                  value={selectedFavorecido}
                  onChange={(e) => setSelectedFavorecido(e.target.value)}
                  className="w-full p-2 bg-slate-50 border border-slate-300 rounded-xl font-medium text-xs text-slate-900 focus:bg-white focus:outline-hidden focus:border-emerald-500"
                >
                  <option value="">Não mapeada</option>
                  {fileData.headers.map((h) => (
                    <option key={h} value={h}>
                      {h}
                    </option>
                  ))}
                </select>
              </div>

              {/* 4. Nota Empenho */}
              <div className="space-y-1">
                <label className="block font-bold text-slate-900 text-xs flex items-center gap-1.5">
                  <FileText className="w-3.5 h-3.5 text-slate-600" />
                  Nota Empenho
                </label>
                <select
                  value={selectedNotaEmpenho}
                  onChange={(e) => setSelectedNotaEmpenho(e.target.value)}
                  className="w-full p-2 bg-slate-50 border border-slate-300 rounded-xl font-medium text-xs text-slate-900 focus:bg-white focus:outline-hidden focus:border-emerald-500"
                >
                  <option value="">Não mapeada</option>
                  {fileData.headers.map((h) => (
                    <option key={h} value={h}>
                      {h}
                    </option>
                  ))}
                </select>
              </div>

              {/* 5. Favorecido Nota Empenho */}
              <div className="space-y-1">
                <label className="block font-bold text-slate-900 text-xs flex items-center gap-1.5">
                  <UserCheck className="w-3.5 h-3.5 text-slate-600" />
                  Favorecido Nota Empenho
                </label>
                <select
                  value={selectedFavorecidoNE}
                  onChange={(e) => setSelectedFavorecidoNE(e.target.value)}
                  className="w-full p-2 bg-slate-50 border border-slate-300 rounded-xl font-medium text-xs text-slate-900 focus:bg-white focus:outline-hidden focus:border-emerald-500"
                >
                  <option value="">Não mapeada</option>
                  {fileData.headers.map((h) => (
                    <option key={h} value={h}>
                      {h}
                    </option>
                  ))}
                </select>
              </div>

              {/* 6. Valor */}
              <div className="space-y-1">
                <label className="block font-bold text-slate-900 text-xs flex items-center gap-1.5">
                  <DollarSign className="w-3.5 h-3.5 text-slate-600" />
                  Valor
                </label>
                <select
                  value={selectedValor}
                  onChange={(e) => setSelectedValor(e.target.value)}
                  className="w-full p-2 bg-slate-50 border border-slate-300 rounded-xl font-medium text-xs text-slate-900 focus:bg-white focus:outline-hidden focus:border-emerald-500"
                >
                  <option value="">Não mapeada</option>
                  {fileData.headers.map((h) => (
                    <option key={h} value={h}>
                      {h}
                    </option>
                  ))}
                </select>
              </div>

              {/* CNPJ */}
              <div className="space-y-1">
                <label className="block font-semibold text-slate-700 text-[11px] flex items-center gap-1.5">
                  <Building2 className="w-3 h-3 text-slate-400" />
                  CNPJ
                </label>
                <select
                  value={selectedCnpj}
                  onChange={(e) => setSelectedCnpj(e.target.value)}
                  className="w-full p-2 bg-slate-50 border border-slate-300 rounded-xl font-medium text-xs text-slate-900 focus:bg-white focus:outline-hidden focus:border-emerald-500"
                >
                  <option value="">Não mapeada</option>
                  {fileData.headers.map((h) => (
                    <option key={h} value={h}>
                      {h}
                    </option>
                  ))}
                </select>
              </div>

              {/* Situação / Status */}
              <div className="space-y-1">
                <label className="block font-semibold text-slate-700 text-[11px] flex items-center gap-1.5">
                  <PieChart className="w-3 h-3 text-slate-400" />
                  Situação / Status
                </label>
                <select
                  value={selectedSituacao}
                  onChange={(e) => setSelectedSituacao(e.target.value)}
                  className="w-full p-2 bg-slate-50 border border-slate-300 rounded-xl font-medium text-xs text-slate-900 focus:bg-white focus:outline-hidden focus:border-emerald-500"
                >
                  <option value="">Não mapeada</option>
                  {fileData.headers.map((h) => (
                    <option key={h} value={h}>
                      {h}
                    </option>
                  ))}
                </select>
              </div>

              {/* Natureza da Despesa */}
              <div className="space-y-1">
                <label className="block font-semibold text-slate-700 text-[11px] flex items-center gap-1.5">
                  <FileText className="w-3 h-3 text-slate-400" />
                  Natureza da Despesa
                </label>
                <select
                  value={selectedNatureza}
                  onChange={(e) => setSelectedNatureza(e.target.value)}
                  className="w-full p-2 bg-slate-50 border border-slate-300 rounded-xl font-medium text-xs text-slate-900 focus:bg-white focus:outline-hidden focus:border-violet-500"
                >
                  <option value="">Não mapeada</option>
                  {fileData.headers.map((h) => (
                    <option key={h} value={h}>
                      {h}
                    </option>
                  ))}
                </select>
              </div>

              {/* Fonte de Recursos */}
              <div className="space-y-1">
                <label className="block font-semibold text-slate-700 text-[11px] flex items-center gap-1.5">
                  <Layers className="w-3 h-3 text-slate-400" />
                  Fonte de Recursos
                </label>
                <select
                  value={selectedFonte}
                  onChange={(e) => setSelectedFonte(e.target.value)}
                  className="w-full p-2 bg-slate-50 border border-slate-300 rounded-xl font-medium text-xs text-slate-900 focus:bg-white focus:outline-hidden focus:border-emerald-500"
                >
                  <option value="">Não mapeada</option>
                  {fileData.headers.map((h) => (
                    <option key={h} value={h}>
                      {h}
                    </option>
                  ))}
                </select>
              </div>

              {/* Data de Referência / Pagamento */}
              <div className="space-y-1">
                <label className="block font-semibold text-slate-700 text-[11px] flex items-center gap-1.5">
                  <Calendar className="w-3 h-3 text-slate-400" />
                  Data de Referência / Pagamento
                </label>
                <select
                  value={selectedData}
                  onChange={(e) => setSelectedData(e.target.value)}
                  className="w-full p-2 bg-slate-50 border border-slate-300 rounded-xl font-medium text-xs text-slate-900 focus:bg-white focus:outline-hidden focus:border-emerald-500"
                >
                  <option value="">Não mapeada</option>
                  {fileData.headers.map((h) => (
                    <option key={h} value={h}>
                      {h}
                    </option>
                  ))}
                </select>
              </div>

            </div>
          </div>

          {/* Live Simulation Preview */}
          {previewRow && (
            <div className="p-3.5 bg-slate-50 border border-slate-200 rounded-xl space-y-2">
              <div className="font-bold text-slate-800 text-xs">
                Pré-visualização (1ª linha):
              </div>
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 text-xs bg-white p-2.5 rounded-lg border border-slate-200">
                <div>
                  <span className="text-[10px] text-slate-400 font-bold block uppercase">Número:</span>
                  <span className="font-medium text-slate-800 truncate block">{previewRow.numero}</span>
                </div>
                <div>
                  <span className="text-[10px] text-slate-400 font-bold block uppercase">OB:</span>
                  <span className="font-medium text-slate-800 truncate block">{previewRow.ob}</span>
                </div>
                <div>
                  <span className="text-[10px] text-slate-400 font-bold block uppercase">Favorecido:</span>
                  <span className="font-medium text-slate-800 truncate block">{previewRow.favorecido}</span>
                </div>
                <div>
                  <span className="text-[10px] text-slate-400 font-bold block uppercase">Fonte:</span>
                  <span className="font-medium text-slate-800 truncate block">{previewRow.fonte}</span>
                </div>
                <div>
                  <span className="text-[10px] text-slate-400 font-bold block uppercase">Data:</span>
                  <span className="font-medium text-slate-800 truncate block">{previewRow.data}</span>
                </div>
                <div>
                  <span className="text-[10px] text-slate-400 font-bold block uppercase">Nota Empenho:</span>
                  <span className="font-medium text-slate-800 truncate block">{previewRow.empenho}</span>
                </div>
                <div>
                  <span className="text-[10px] text-slate-400 font-bold block uppercase">Favorecido NE:</span>
                  <span className="font-medium text-slate-800 truncate block">{previewRow.favorecidoNE}</span>
                </div>
                <div>
                  <span className="text-[10px] text-slate-400 font-bold block uppercase">Valor:</span>
                  <span className="font-bold text-emerald-800 block">{formatBRL(previewRow.valor)}</span>
                </div>
              </div>
            </div>
          )}

        </div>

        {/* Modal Footer */}
        <div className="p-4 bg-slate-50 border-t border-slate-200 flex items-center justify-end space-x-2 shrink-0">
          <button
            onClick={onClose}
            className="px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold text-xs rounded-xl transition cursor-pointer"
          >
            Cancelar
          </button>
          <button
            onClick={handleSave}
            className="px-5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs rounded-xl shadow-xs transition flex items-center gap-1.5 cursor-pointer"
          >
            <Check className="w-4 h-4" />
            <span>Confirmar Mapeamento</span>
          </button>
        </div>

      </div>
    </div>
  );
};
