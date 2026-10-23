/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useRef, useState } from 'react';
import {
  UploadCloud,
  FileSpreadsheet,
  Check,
  AlertCircle,
  X,
  Eye,
  Settings2,
} from 'lucide-react';
import type { OverwriteMode, WorkbookMeta } from '../types/excel';

interface FileUploadCardProps {
  title: string;
  badgeLabel: string;
  accentColor: 'lilac' | 'butter';
  description: string;
  fileMeta: WorkbookMeta | null;
  isLoading: boolean;
  error?: string | null;
  onFileSelect: (file: File) => void;
  onRemoveFile: () => void;
  onSheetChange: (sheetName: string) => void;
  onNipColumnChange: (colName: string) => void;
  onNameColumnChange?: (colName: string) => void;
  onPreview: () => void;
  showNameColumnSelector?: boolean;
  overwriteMode?: OverwriteMode;
  onOverwriteModeChange?: (mode: OverwriteMode) => void;
  existingNameColumnWarning?: boolean;
}

export const FileUploadCard: React.FC<FileUploadCardProps> = ({
  title,
  badgeLabel,
  accentColor,
  description,
  fileMeta,
  isLoading,
  error,
  onFileSelect,
  onRemoveFile,
  onSheetChange,
  onNipColumnChange,
  onNameColumnChange,
  onPreview,
  showNameColumnSelector = false,
  overwriteMode = 'update_existing',
  onOverwriteModeChange,
  existingNameColumnWarning = false,
}) => {
  const [isDragOver, setIsDragOver] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const formatFileSize = (bytes: number): string => {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragOver(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragOver(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragOver(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const file = e.dataTransfer.files[0];
      if (
        file.name.endsWith('.xlsx') ||
        file.name.endsWith('.xls') ||
        file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      ) {
        onFileSelect(file);
      }
    }
  };

  const handleFileInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      onFileSelect(e.target.files[0]);
    }
  };

  const activeSheet =
    fileMeta?.sheets.find((s) => s.name === fileMeta.selectedSheet) ||
    fileMeta?.sheets[0];

  const badgeBg =
    accentColor === 'lilac'
      ? 'bg-[#D6D6FD] text-[#1E1B4B]'
      : 'bg-[#FEEBB6] text-[#78350F]';

  return (
    <div className="bg-white rounded-3xl p-6 sm:p-7 flex flex-col justify-between transition-all">
      {/* Card Header */}
      <div>
        <div className="flex items-center justify-between gap-3 mb-2">
          <div className="flex items-center gap-2.5">
            <span
              className={`text-xs font-bold px-3 py-1 rounded-full tracking-wide ${badgeBg}`}
            >
              {badgeLabel}
            </span>
            <h2 className="text-lg font-bold text-[#121212] tracking-tight">{title}</h2>
          </div>
          {fileMeta && (
            <button
              type="button"
              onClick={onRemoveFile}
              className="w-8 h-8 rounded-full bg-[#F5F5FA] hover:bg-rose-50 text-slate-400 hover:text-rose-600 transition-colors flex items-center justify-center cursor-pointer"
              title="Ganti file"
            >
              <X className="w-4 h-4" />
            </button>
          )}
        </div>
        <p className="text-xs text-slate-400 mb-5 font-normal leading-relaxed">
          {description}
        </p>

        {/* Upload Zone or File Details */}
        {!fileMeta ? (
          <div
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
            onClick={() => fileInputRef.current?.click()}
            className={`rounded-2xl p-8 text-center cursor-pointer transition-all duration-200 flex flex-col items-center justify-center min-h-[200px] ${
              isDragOver
                ? 'bg-[#E8E8FF]'
                : 'bg-[#F5F5FA] hover:bg-[#EFEFF8]'
            }`}
          >
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls"
              onChange={handleFileInputChange}
              className="hidden"
            />
            {isLoading ? (
              <div className="flex flex-col items-center gap-3">
                <div className="w-8 h-8 border-3 border-[#121212] border-t-transparent rounded-full animate-spin" />
                <p className="text-xs font-medium text-slate-600">
                  Membaca file Excel...
                </p>
              </div>
            ) : (
              <>
                <div className="w-13 h-13 rounded-2xl bg-white text-[#121212] flex items-center justify-center mb-3.5">
                  <UploadCloud className="w-6 h-6 text-[#121212]" />
                </div>
                <p className="text-sm font-bold text-[#121212]">
                  Tarik & lepas file Excel di sini
                </p>
                <p className="text-xs text-slate-400 mt-1 font-medium">
                  atau <span className="text-[#121212] font-semibold underline underline-offset-2">pilih file dari perangkat</span> (.xlsx)
                </p>
              </>
            )}
          </div>
        ) : (
          <div className="space-y-4">
            {/* File Info Bar */}
            <div className="flex items-center justify-between p-3.5 bg-[#F5F5FA] rounded-2xl">
              <div className="flex items-center gap-3 min-w-0">
                <div
                  className={`w-10 h-10 rounded-xl flex items-center justify-center shrink-0 ${
                    accentColor === 'lilac'
                      ? 'bg-[#D6D6FD] text-[#1E1B4B]'
                      : 'bg-[#FEEBB6] text-[#78350F]'
                  }`}
                >
                  <FileSpreadsheet className="w-5 h-5" />
                </div>
                <div className="min-w-0">
                  <p
                    className="text-xs font-bold text-[#121212] truncate"
                    title={fileMeta.fileName}
                  >
                    {fileMeta.fileName}
                  </p>
                  <p className="text-[11px] text-slate-400 font-mono mt-0.5">
                    {formatFileSize(fileMeta.fileSize)} · {activeSheet?.rowCount || 0} baris
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={onPreview}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold text-[#121212] bg-white hover:bg-slate-100 rounded-full transition-colors shrink-0 cursor-pointer"
                title="Lihat isi data sheet"
              >
                <Eye className="w-3.5 h-3.5 text-slate-600" />
                <span>Cuplikan</span>
              </button>
            </div>

            {/* Column & Sheet Configuration */}
            <div className="p-4 bg-[#F5F5FA] rounded-2xl space-y-3.5">
              <div className="flex items-center gap-1.5 text-xs font-bold text-slate-700">
                <Settings2 className="w-4 h-4 text-[#121212]" />
                <span>Konfigurasi Kolom & Sheet</span>
              </div>

              {/* Sheet Selector */}
              <div>
                <label className="block text-[11px] font-semibold text-slate-500 mb-1.5">
                  Pilih Sheet:
                </label>
                <select
                  value={fileMeta.selectedSheet}
                  onChange={(e) => onSheetChange(e.target.value)}
                  className="w-full text-xs bg-white rounded-xl px-3.5 py-2.5 text-[#121212] font-semibold border-0 focus:ring-2 focus:ring-[#D6D6FD] outline-hidden cursor-pointer"
                >
                  {fileMeta.sheets.map((s) => (
                    <option key={s.name} value={s.name}>
                      {s.name} ({s.rowCount} baris, {s.columnCount} kolom)
                    </option>
                  ))}
                </select>
                {accentColor === 'lilac' &&
                  !fileMeta.sheets.some(
                    (s) => s.name.trim().toLowerCase() === 'data peserta'
                  ) && (
                    <p className="text-[10px] text-amber-700 mt-1 font-medium">
                      Note: Sheet &quot;Data Peserta&quot; tidak ditemukan. Menggunakan sheet pertama (&quot;{activeSheet?.name}&quot;) sebagai fallback.
                    </p>
                  )}
              </div>

              {/* NIP Column Selector */}
              <div>
                <div className="flex items-center justify-between mb-1.5">
                  <label className="block text-[11px] font-semibold text-slate-500">
                    Kolom NIP / Peserta (Player Name):
                  </label>
                  {fileMeta.nipColumn && (
                    <span className="text-[10px] text-emerald-700 bg-emerald-100/70 px-2 py-0.5 rounded-full flex items-center gap-1 font-bold">
                      <Check className="w-3 h-3 stroke-[3]" /> Terpilih
                    </span>
                  )}
                </div>
                <select
                  value={fileMeta.nipColumn}
                  onChange={(e) => onNipColumnChange(e.target.value)}
                  className="w-full text-xs bg-white rounded-xl px-3.5 py-2.5 text-[#121212] font-semibold border-0 focus:ring-2 focus:ring-[#D6D6FD] outline-hidden cursor-pointer"
                >
                  <option value="">-- Pilih Kolom NIP / Peserta --</option>
                  {activeSheet?.headers.map((h, i) => (
                    <option key={`${h}-${i}`} value={h}>
                      {h}
                    </option>
                  ))}
                </select>
                {accentColor === 'lilac' && !fileMeta.nipColumn && (
                  <div className="p-3 bg-amber-50 rounded-xl border border-amber-200/80 text-[11px] text-amber-900 flex items-start gap-2 mt-2">
                    <AlertCircle className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
                    <div>
                      <p className="font-bold">Kolom Default &quot;Player Name&quot; Tidak Ditemukan</p>
                      <p className="mt-0.5 font-normal text-amber-800 leading-relaxed">
                        Kolom <strong>Player Name</strong> tidak ditemukan secara otomatis pada sheet &quot;{activeSheet?.name}&quot;. Silakan pilih kolom NIP / Peserta secara manual dari dropdown di atas.
                      </p>
                    </div>
                  </div>
                )}
              </div>

              {/* Name Column Selector (for Master file) */}
              {showNameColumnSelector && onNameColumnChange && (
                <div>
                  <div className="flex items-center justify-between mb-1.5">
                    <label className="block text-[11px] font-semibold text-slate-500">
                      Kolom Nama Karyawan:
                    </label>
                    {fileMeta.nameColumn && (
                      <span className="text-[10px] text-emerald-700 bg-emerald-100/70 px-2 py-0.5 rounded-full flex items-center gap-1 font-bold">
                        <Check className="w-3 h-3 stroke-[3]" /> Terpilih
                      </span>
                    )}
                  </div>
                  <select
                    value={fileMeta.nameColumn || ''}
                    onChange={(e) => onNameColumnChange(e.target.value)}
                    className="w-full text-xs bg-white rounded-xl px-3.5 py-2.5 text-[#121212] font-semibold border-0 focus:ring-2 focus:ring-[#D6D6FD] outline-hidden cursor-pointer"
                  >
                    <option value="">-- Pilih Kolom Nama --</option>
                    {activeSheet?.headers.map((h, i) => (
                      <option key={`${h}-${i}`} value={h}>
                        {h}
                      </option>
                    ))}
                  </select>
                </div>
              )}

              {/* Post Test Existing Column Notice */}
              {existingNameColumnWarning && onOverwriteModeChange && (
                <div className="p-3.5 bg-[#FEF3C7] rounded-xl text-[11px] text-amber-950">
                  <div className="flex items-center gap-1.5 font-bold mb-1.5">
                    <AlertCircle className="w-4 h-4 text-amber-700 shrink-0" />
                    <span>Kolom &quot;NIP&quot; &amp; &quot;Nama Lengkap&quot; sudah ada di file ini</span>
                  </div>
                  <div className="mt-2 space-y-1.5 font-medium">
                    <label className="flex items-center gap-2 cursor-pointer">
                      <input
                        type="radio"
                        name="overwriteMode"
                        value="update_existing"
                        checked={overwriteMode === 'update_existing'}
                        onChange={() => onOverwriteModeChange('update_existing')}
                        className="accent-[#121212]"
                      />
                      <span>Perbarui kolom &quot;NIP&quot; &amp; &quot;Nama Lengkap&quot; yang ada</span>
                    </label>
                    <label className="flex items-center gap-2 cursor-pointer">
                      <input
                        type="radio"
                        name="overwriteMode"
                        value="create_new_column"
                        checked={overwriteMode === 'create_new_column'}
                        onChange={() => onOverwriteModeChange('create_new_column')}
                        className="accent-[#121212]"
                      />
                      <span>Tambahkan sepasang kolom baru di sebelah paling kanan</span>
                    </label>
                  </div>
                </div>
              )}
            </div>
          </div>
        )}
      </div>

      {/* Error notification if any */}
      {error && (
        <div className="mt-4 p-3 bg-rose-100 rounded-2xl text-xs text-rose-900 flex items-start gap-2">
          <AlertCircle className="w-4 h-4 text-rose-700 shrink-0 mt-0.5" />
          <span className="font-medium">{error}</span>
        </div>
      )}
    </div>
  );
};
