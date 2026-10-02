import React, { useRef } from 'react';
import {
  Upload,
  FileSpreadsheet,
  CheckCircle2,
  Trash2,
  Settings2,
  Play,
  Database,
  Plus,
  Check,
  Tag,
  Clock,
  AlertTriangle,
  FileCheck2,
} from 'lucide-react';
import { FileData, StageKey, STAGES_CONFIG } from '../types';

interface UploadSectionProps {
  files: FileData[];
  onFileUpload: (file: File, stage?: StageKey) => void;
  onRemoveFile: (fileId: string) => void;
  onChangeFileStage?: (fileId: string, stage: StageKey) => void;
  onProcessData: () => void;
  onOpenColumnMapModal: (fileId: string) => void;
  isProcessing: boolean;
}

export const UploadSection: React.FC<UploadSectionProps> = ({
  files,
  onFileUpload,
  onRemoveFile,
  onChangeFileStage,
  onProcessData,
  onOpenColumnMapModal,
  isProcessing,
}) => {
  const generalInputRef = useRef<HTMLInputElement>(null);
  const ppInputRef = useRef<HTMLInputElement>(null);
  const obedeceInputRef = useRef<HTMLInputElement>(null);
  const naoObedeceInputRef = useRef<HTMLInputElement>(null);

  const handleGeneralFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      Array.from(e.target.files).forEach((file) => onFileUpload(file));
      e.target.value = '';
    }
  };

  const handleStageFileChange = (e: React.ChangeEvent<HTMLInputElement>, stage: StageKey) => {
    if (e.target.files) {
      Array.from(e.target.files).forEach((file) => onFileUpload(file, stage));
      e.target.value = '';
    }
  };

  const handleDropGeneral = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.dataTransfer.files) {
      Array.from(e.dataTransfer.files).forEach((file) => onFileUpload(file));
    }
  };

  const handleDropStage = (e: React.DragEvent, stage: StageKey) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.dataTransfer.files) {
      Array.from(e.dataTransfer.files).forEach((file) => onFileUpload(file, stage));
    }
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
  };

  // Count files per stage
  const ppCount = files.filter((f) => f.stage === 'base_pps' || f.stage === 'pp_prontas').length;
  const obedeceCount = files.filter((f) => f.stage === 'obedece').length;
  const naoObedeceCount = files.filter((f) => f.stage === 'nao_obedece').length;

  return (
    <section className="bg-white rounded-2xl p-6 shadow-xs border border-slate-200/90 space-y-6">
      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-100 pb-4">
        <div>
          <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
            <FileSpreadsheet className="w-5 h-5 text-emerald-600" />
            <span>Central de Upload e Importação de Planilhas</span>
          </h3>
        </div>
      </div>

      {/* Hidden Inputs for Stage Uploads */}
      <input
        type="file"
        ref={generalInputRef}
        accept=".xlsx, .xls, .csv"
        multiple
        className="hidden"
        onChange={handleGeneralFileChange}
      />
      <input
        type="file"
        ref={ppInputRef}
        accept=".xlsx, .xls, .csv"
        multiple
        className="hidden"
        onChange={(e) => handleStageFileChange(e, 'base_pps')}
      />
      <input
        type="file"
        ref={obedeceInputRef}
        accept=".xlsx, .xls, .csv"
        multiple
        className="hidden"
        onChange={(e) => handleStageFileChange(e, 'obedece')}
      />
      <input
        type="file"
        ref={naoObedeceInputRef}
        accept=".xlsx, .xls, .csv"
        multiple
        className="hidden"
        onChange={(e) => handleStageFileChange(e, 'nao_obedece')}
      />

      {/* 3 DEDICATED STAGE UPLOAD CARDS */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
        {/* Card 1: Base de PPs (Relatório SIGEF) */}
        <div
          onDragOver={handleDragOver}
          onDrop={(e) => handleDropStage(e, 'base_pps')}
          onClick={() => ppInputRef.current?.click()}
          className="bg-emerald-50/60 hover:bg-emerald-100/60 border-2 border-dashed border-emerald-300 hover:border-emerald-500 rounded-2xl p-5 text-center cursor-pointer transition-all space-y-3 relative group"
        >
          <div className="flex items-center justify-between">
            <span className="px-2.5 py-1 bg-emerald-100 text-emerald-900 border border-emerald-300 rounded-lg text-[10px] font-black uppercase tracking-wider flex items-center gap-1">
              <FileCheck2 className="w-3.5 h-3.5 text-emerald-700" />
              Base de PPs
            </span>
            {ppCount > 0 && (
              <span className="bg-emerald-600 text-white text-[10px] font-bold px-2 py-0.5 rounded-full">
                {ppCount} {ppCount === 1 ? 'arquivo' : 'arquivos'}
              </span>
            )}
          </div>

          <div className="w-12 h-12 rounded-xl bg-emerald-600 text-white flex items-center justify-center mx-auto shadow-md group-hover:scale-105 transition-transform">
            <Upload className="w-6 h-6" />
          </div>

          <div>
            <h4 className="text-xs font-black text-emerald-950 uppercase tracking-tight">
              Arquivo Base de PPs
            </h4>
            <p className="text-[11px] text-slate-500 mt-0.5">
              Relatório SIGEF (ex: Relatorio_*.xls)
            </p>
          </div>
        </div>

        {/* Card 2: Liquidação - Obedece */}
        <div
          onDragOver={handleDragOver}
          onDrop={(e) => handleDropStage(e, 'obedece')}
          onClick={() => obedeceInputRef.current?.click()}
          className="bg-blue-50/60 hover:bg-blue-100/60 border-2 border-dashed border-blue-300 hover:border-blue-500 rounded-2xl p-5 text-center cursor-pointer transition-all space-y-3 relative group"
        >
          <div className="flex items-center justify-between">
            <span className="px-2.5 py-1 bg-blue-100 text-blue-900 border border-blue-300 rounded-lg text-[10px] font-black uppercase tracking-wider flex items-center gap-1">
              <Clock className="w-3.5 h-3.5 text-blue-700" />
              Liquidado - Obedece
            </span>
            {obedeceCount > 0 && (
              <span className="bg-blue-600 text-white text-[10px] font-bold px-2 py-0.5 rounded-full">
                {obedeceCount} {obedeceCount === 1 ? 'arquivo' : 'arquivos'}
              </span>
            )}
          </div>

          <div className="w-12 h-12 rounded-xl bg-blue-600 text-white flex items-center justify-center mx-auto shadow-md group-hover:scale-105 transition-transform">
            <Upload className="w-6 h-6" />
          </div>

          <div>
            <h4 className="text-xs font-black text-blue-950 uppercase tracking-tight">
              Liquidado - Obedece
            </h4>
          </div>
        </div>

        {/* Card 3: Liquidação - Não Obedece */}
        <div
          onDragOver={handleDragOver}
          onDrop={(e) => handleDropStage(e, 'nao_obedece')}
          onClick={() => naoObedeceInputRef.current?.click()}
          className="bg-amber-50/60 hover:bg-amber-100/60 border-2 border-dashed border-amber-300 hover:border-amber-500 rounded-2xl p-5 text-center cursor-pointer transition-all space-y-3 relative group"
        >
          <div className="flex items-center justify-between">
            <span className="px-2.5 py-1 bg-amber-100 text-amber-900 border border-amber-300 rounded-lg text-[10px] font-black uppercase tracking-wider flex items-center gap-1">
              <AlertTriangle className="w-3.5 h-3.5 text-amber-700" />
              Liquidado - Não Obedece
            </span>
            {naoObedeceCount > 0 && (
              <span className="bg-amber-600 text-white text-[10px] font-bold px-2 py-0.5 rounded-full">
                {naoObedeceCount} {naoObedeceCount === 1 ? 'arquivo' : 'arquivos'}
              </span>
            )}
          </div>

          <div className="w-12 h-12 rounded-xl bg-amber-600 text-white flex items-center justify-center mx-auto shadow-md group-hover:scale-105 transition-transform">
            <Upload className="w-6 h-6" />
          </div>

          <div>
            <h4 className="text-xs font-black text-amber-950 uppercase tracking-tight">
              Liquidado - Não Obedece
            </h4>
          </div>
        </div>
      </div>

      {/* General Drag & Drop Zone */}
      <div
        onDragOver={handleDragOver}
        onDrop={handleDropGeneral}
        onClick={() => generalInputRef.current?.click()}
        className="border-2 border-dashed border-slate-300 hover:border-slate-400 bg-slate-50/80 hover:bg-slate-100/80 rounded-2xl p-4 text-center cursor-pointer transition-all space-y-1.5"
      >
        <p className="text-xs font-extrabold text-slate-700">
          📁 Soltar arquivos aqui para detecção automática
        </p>
      </div>

      {/* List of Uploaded Files */}
      {files.length > 0 && (
        <div className="space-y-3 pt-2">
          <div className="flex items-center justify-between text-xs font-bold text-slate-700">
            <span>Planilhas Importadas ({files.length})</span>
            <button
              onClick={() => generalInputRef.current?.click()}
              className="text-emerald-700 hover:text-emerald-900 flex items-center gap-1 cursor-pointer font-bold text-xs"
            >
              <Plus className="w-3.5 h-3.5" />
              <span>Adicionar Mais Planilhas</span>
            </button>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
            {files.map((fileData) => {
              const currentStage = fileData.stage || 'outros';
              const stageConfig = STAGES_CONFIG[currentStage];

              return (
                <div
                  key={fileData.id}
                  className="bg-slate-50 rounded-xl p-4 border border-slate-200 flex flex-col justify-between space-y-3 shadow-2xs"
                >
                  <div className="flex items-start justify-between gap-2">
                    <div className="flex items-center space-x-2.5 min-w-0">
                      <CheckCircle2 className="w-5 h-5 text-emerald-600 shrink-0" />
                      <div className="min-w-0">
                        <p className="text-xs font-bold text-slate-900 truncate" title={fileData.fileName}>
                          {fileData.fileName}
                        </p>
                        <p className="text-[11px] text-slate-500 font-medium">
                          {fileData.rows.length} registros no arquivo
                        </p>
                      </div>
                    </div>

                    <button
                      onClick={(e) => {
                        e.stopPropagation();
                        onRemoveFile(fileData.id);
                      }}
                      className="text-slate-400 hover:text-rose-600 p-1 rounded-md hover:bg-rose-50 transition cursor-pointer"
                      title="Remover planilha"
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  </div>

                  {/* Stage Tag Selector & Mapping button */}
                  <div className="pt-2 border-t border-slate-200/60 flex flex-col sm:flex-row sm:items-center justify-between gap-2 text-[11px]">
                    <div className="flex items-center gap-1.5">
                      <Tag className="w-3.5 h-3.5 text-slate-400 shrink-0" />
                      <select
                        value={currentStage === 'pp_prontas' ? 'base_pps' : currentStage}
                        onChange={(e) =>
                          onChangeFileStage && onChangeFileStage(fileData.id, e.target.value as StageKey)
                        }
                        className={`text-[10px] font-bold px-2 py-1 rounded-lg border cursor-pointer ${stageConfig.badgeText}`}
                      >
                        <option value="base_pps">📁 Base de PPs (Relatório SIGEF)</option>
                        <option value="obedece">🔵 Liquidação - Obedece</option>
                        <option value="nao_obedece">🟠 Liquidação - Não Obedece</option>
                        <option value="outros">⚪ Outras / Geral</option>
                      </select>
                    </div>

                    <button
                      onClick={() => onOpenColumnMapModal(fileData.id)}
                      className="text-emerald-700 font-bold hover:underline flex items-center gap-1 cursor-pointer shrink-0 self-end sm:self-auto"
                    >
                      <Settings2 className="w-3 h-3" />
                      <span>Ajustar Colunas</span>
                    </button>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* Action Button */}
      <div className="pt-2 flex flex-col sm:flex-row items-center justify-between gap-4">
        <p className="text-xs text-slate-500 font-medium">
          Status: <strong>{files.length} planilha(s)</strong> prontas para consolidação.
        </p>

        <button
          onClick={onProcessData}
          disabled={files.length === 0 || isProcessing}
          className={`w-full sm:w-auto px-8 py-3 rounded-xl font-bold text-sm shadow-md transition-all flex items-center justify-center gap-2 ${
            files.length > 0 && !isProcessing
              ? 'bg-emerald-600 hover:bg-emerald-700 text-white shadow-emerald-600/25 active:scale-98 cursor-pointer'
              : 'bg-slate-200 text-slate-400 cursor-not-allowed shadow-none'
          }`}
        >
          {isProcessing ? (
            <>
              <div className="w-4 h-4 border-2 border-white/30 border-t-white rounded-full animate-spin" />
              <span>Processando...</span>
            </>
          ) : (
            <>
              <Play className="w-4 h-4 fill-white" />
              <span>Processar e Consolidar Dados</span>
            </>
          )}
        </button>
      </div>
    </section>
  );
};
