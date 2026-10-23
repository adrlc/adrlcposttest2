/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React from 'react';
import { Play, RotateCcw, Sparkles, Download, CheckCircle2 } from 'lucide-react';

interface ActionToolbarProps {
  canGenerate: boolean;
  isProcessing: boolean;
  onGenerate: () => void;
  onReset: () => void;
  onLoadSamples: () => void;
  onDownloadSamples: () => void;
  validationMessage?: string | null;
  hasResult: boolean;
}

export const ActionToolbar: React.FC<ActionToolbarProps> = ({
  canGenerate,
  isProcessing,
  onGenerate,
  onReset,
  onLoadSamples,
  onDownloadSamples,
  validationMessage,
  hasResult,
}) => {
  return (
    <div className="bg-white rounded-3xl p-5 sm:p-6 space-y-4">
      <div className="flex flex-col md:flex-row items-stretch md:items-center justify-between gap-4">
        {/* Left Side: Test / Quick Actions with pastel flat buttons */}
        <div className="flex items-center gap-2.5 flex-wrap">
          <button
            type="button"
            onClick={onLoadSamples}
            disabled={isProcessing}
            className="inline-flex items-center gap-2 px-4 py-2.5 text-xs font-bold text-[#1E1B4B] bg-[#D6D6FD] hover:bg-[#C7C7FE] rounded-2xl transition-colors cursor-pointer disabled:opacity-50"
            title="Muat data contoh otomatis untuk pengujian cepat"
          >
            <Sparkles className="w-3.5 h-3.5 text-[#1E1B4B]" />
            <span>Muat Contoh Data Uji Coba</span>
          </button>

          <button
            type="button"
            onClick={onDownloadSamples}
            disabled={isProcessing}
            className="inline-flex items-center gap-2 px-4 py-2.5 text-xs font-bold text-[#78350F] bg-[#FEEBB6] hover:bg-[#FEDC8B] rounded-2xl transition-colors cursor-pointer disabled:opacity-50"
            title="Unduh file template contoh (.xlsx) ke komputer"
          >
            <Download className="w-3.5 h-3.5 text-[#78350F]" />
            <span>Unduh File Sampel (.xlsx)</span>
          </button>
        </div>

        {/* Right Side: Main Actions (Generate & Reset) */}
        <div className="flex items-center gap-3 justify-end">
          <button
            type="button"
            onClick={onReset}
            disabled={isProcessing}
            className="inline-flex items-center gap-1.5 px-5 py-3 text-xs font-bold text-slate-700 bg-[#F5F5FA] hover:bg-[#EAEAF4] rounded-2xl transition-colors cursor-pointer disabled:opacity-50"
          >
            <RotateCcw className="w-3.5 h-3.5 text-slate-500" />
            <span>Reset</span>
          </button>

          <button
            type="button"
            onClick={onGenerate}
            disabled={!canGenerate || isProcessing}
            className={`inline-flex items-center justify-center gap-2 px-8 py-3.5 text-sm font-bold rounded-2xl transition-all cursor-pointer ${
              canGenerate && !isProcessing
                ? 'bg-[#121212] hover:bg-[#262626] text-white active:scale-[0.99]'
                : 'bg-[#E5E7EB] text-slate-400 cursor-not-allowed'
            }`}
          >
            {isProcessing ? (
              <>
                <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin" />
                <span>Memproses Data Excel...</span>
              </>
            ) : hasResult ? (
              <>
                <CheckCircle2 className="w-4 h-4 text-[#D6D6FD]" />
                <span>Generate Ulang Excel</span>
              </>
            ) : (
              <>
                <Play className="w-4 h-4 fill-white" />
                <span>Generate Excel</span>
              </>
            )}
          </button>
        </div>
      </div>

      {/* Validation hint if cannot generate */}
      {validationMessage && !canGenerate && (
        <div className="p-3 bg-[#F5F5FA] rounded-2xl flex items-center gap-2">
          <span className="w-2 h-2 rounded-full bg-[#FF7170]" />
          <span className="text-xs font-semibold text-slate-600">{validationMessage}</span>
        </div>
      )}
    </div>
  );
};
