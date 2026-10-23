/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { useState } from 'react';
import { Header } from './components/Header';
import { FileUploadCard } from './components/FileUploadCard';
import { ActionToolbar } from './components/ActionToolbar';
import { ProcessingSummary } from './components/ProcessingSummary';
import { DataPreviewModal } from './components/DataPreviewModal';
import { HowItWorksModal } from './components/HowItWorksModal';
import { parseWorkbook } from './utils/excelReader';
import { enrichPostTestWorkbook, generateEditedFileName } from './utils/excelEnricher';
import {
  createSamplePostTestWorkbook,
  createSampleMasterKaryawanWorkbook,
} from './utils/sampleGenerator';
import type { EnrichmentSummary, OverwriteMode, WorkbookMeta } from './types/excel';
import { AlertCircle, Sparkles, Zap } from 'lucide-react';

export default function App() {
  const [postTestMeta, setPostTestMeta] = useState<WorkbookMeta | null>(null);
  const [masterMeta, setMasterMeta] = useState<WorkbookMeta | null>(null);

  const [postTestLoading, setPostTestLoading] = useState(false);
  const [masterLoading, setMasterLoading] = useState(false);
  const [isGenerating, setIsGenerating] = useState(false);

  const [postTestError, setPostTestError] = useState<string | null>(null);
  const [masterError, setMasterError] = useState<string | null>(null);
  const [globalError, setGlobalError] = useState<string | null>(null);

  const [overwriteMode, setOverwriteMode] = useState<OverwriteMode>('update_existing');
  const [summary, setSummary] = useState<EnrichmentSummary | null>(null);

  const [previewState, setPreviewState] = useState<{
    isOpen: boolean;
    title: string;
    buffer: ArrayBuffer | null;
    sheetName?: string;
  }>({
    isOpen: false,
    title: '',
    buffer: null,
  });

  const [isGuideOpen, setIsGuideOpen] = useState(false);

  // File selection handlers
  const handlePostTestFileSelect = async (file: File) => {
    setPostTestLoading(true);
    setPostTestError(null);
    setGlobalError(null);
    setSummary(null);

    try {
      const { meta } = await parseWorkbook(file, file.name, file.size);
      setPostTestMeta(meta);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      setPostTestError(`Gagal membaca file: ${msg}`);
      setPostTestMeta(null);
    } finally {
      setPostTestLoading(false);
    }
  };

  const handleMasterFileSelect = async (file: File) => {
    setMasterLoading(true);
    setMasterError(null);
    setGlobalError(null);
    setSummary(null);

    try {
      const { meta } = await parseWorkbook(file, file.name, file.size);
      setMasterMeta(meta);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      setMasterError(`Gagal membaca file: ${msg}`);
      setMasterMeta(null);
    } finally {
      setMasterLoading(false);
    }
  };

  // Sheet & column config handlers
  const handlePostTestSheetChange = (sheetName: string) => {
    if (!postTestMeta) return;
    const targetSheet = postTestMeta.sheets.find((s) => s.name === sheetName);
    const nipCol =
      targetSheet?.headers.find((h) => h.toLowerCase().includes('nip')) ||
      targetSheet?.headers[0] ||
      '';
    setPostTestMeta({
      ...postTestMeta,
      selectedSheet: sheetName,
      nipColumn: nipCol,
    });
  };

  const handleMasterSheetChange = (sheetName: string) => {
    if (!masterMeta) return;
    const targetSheet = masterMeta.sheets.find((s) => s.name === sheetName);
    const nipCol =
      targetSheet?.headers.find((h) => h.toLowerCase().includes('nip')) ||
      targetSheet?.headers[0] ||
      '';
    const nameCol =
      targetSheet?.headers.find((h) => h.toLowerCase().includes('nama')) ||
      targetSheet?.headers[1] ||
      '';
    setMasterMeta({
      ...masterMeta,
      selectedSheet: sheetName,
      nipColumn: nipCol,
      nameColumn: nameCol,
    });
  };

  // Check if Post Test sheet already has "NIP" & "Nama Lengkap" columns from previous enrichment
  const postTestActiveSheet =
    postTestMeta?.sheets.find((s) => s.name === postTestMeta.selectedSheet) ||
    postTestMeta?.sheets[0];
  const postTestHeaders = postTestActiveSheet?.headers || [];
  const lastCol = postTestHeaders.length;
  const postTestHasExistingEnrichedCols = Boolean(
    lastCol >= 2 &&
      (postTestHeaders[lastCol - 1]?.trim().toLowerCase() === 'nama lengkap' ||
        postTestHeaders[lastCol - 1]?.trim().toLowerCase() === 'nama karyawan') &&
      postTestHeaders[lastCol - 2]?.trim().toLowerCase() === 'nip'
  );

  // Validation check
  const canGenerate = Boolean(
    postTestMeta &&
      masterMeta &&
      postTestMeta.nipColumn &&
      masterMeta.nipColumn &&
      masterMeta.nameColumn
  );

  let validationMessage: string | null = null;
  if (!postTestMeta && !masterMeta) {
    validationMessage =
      'Silakan unggah File Data Post Test dan File Data Karyawan untuk memulai.';
  } else if (!postTestMeta) {
    validationMessage = 'File Data Post Test belum diunggah.';
  } else if (!masterMeta) {
    validationMessage = 'File Data Karyawan belum diunggah.';
  } else if (!postTestMeta.nipColumn) {
    validationMessage = 'Pilih kolom NIP pada File Data Post Test.';
  } else if (!masterMeta.nipColumn) {
    validationMessage = 'Pilih kolom NIP pada File Data Karyawan.';
  } else if (!masterMeta.nameColumn) {
    validationMessage = 'Pilih kolom Nama Karyawan pada File Data Karyawan.';
  }

  // Generate handler
  const handleGenerate = async () => {
    if (!canGenerate || !postTestMeta || !masterMeta) return;

    setIsGenerating(true);
    setGlobalError(null);
    setSummary(null);

    try {
      const result = await enrichPostTestWorkbook(
        postTestMeta.buffer,
        masterMeta.buffer,
        {
          postTestFileName: postTestMeta.fileName,
          postTestSheetName: postTestMeta.selectedSheet,
          postTestNipCol: postTestMeta.nipColumn,
          masterSheetName: masterMeta.selectedSheet,
          masterNipCol: masterMeta.nipColumn,
          masterNameCol: masterMeta.nameColumn,
          overwriteMode,
          targetNipColumnName: 'NIP',
          targetNameColumnName: 'Nama Lengkap',
        }
      );
      setSummary(result);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      setGlobalError(msg);
    } finally {
      setIsGenerating(false);
    }
  };

  // Reset handler: returns app to initial state
  const handleReset = () => {
    setPostTestMeta(null);
    setMasterMeta(null);
    setPostTestError(null);
    setMasterError(null);
    setGlobalError(null);
    setSummary(null);
    setOverwriteMode('update_existing');
  };

  // Sample data loader for quick demo / testing
  const handleLoadSamples = async () => {
    setPostTestLoading(true);
    setMasterLoading(true);
    setPostTestError(null);
    setMasterError(null);
    setGlobalError(null);
    setSummary(null);

    try {
      const postTestBuf = await createSamplePostTestWorkbook();
      const masterBuf = await createSampleMasterKaryawanWorkbook();

      const postTestParsed = await parseWorkbook(
        postTestBuf,
        'Sample_Hasil_Post_Test.xlsx',
        postTestBuf.byteLength
      );
      const masterParsed = await parseWorkbook(
        masterBuf,
        'Sample_Data_Master_Karyawan.xlsx',
        masterBuf.byteLength
      );

      setPostTestMeta(postTestParsed.meta);
      setMasterMeta(masterParsed.meta);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      setGlobalError(`Gagal memuat file contoh: ${msg}`);
    } finally {
      setPostTestLoading(false);
      setMasterLoading(false);
    }
  };

  // Download sample files directly to disk
  const handleDownloadSamples = async () => {
    try {
      const postTestBuf = await createSamplePostTestWorkbook();
      const masterBuf = await createSampleMasterKaryawanWorkbook();

      // Download Post Test sample
      const blob1 = new Blob([postTestBuf], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      const url1 = URL.createObjectURL(blob1);
      const a1 = document.createElement('a');
      a1.href = url1;
      a1.download = 'Sample_Post_Test.xlsx';
      document.body.appendChild(a1);
      a1.click();
      document.body.removeChild(a1);
      URL.revokeObjectURL(url1);

      // Download Master sample with small timeout
      setTimeout(() => {
        const blob2 = new Blob([masterBuf], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
        const url2 = URL.createObjectURL(blob2);
        const a2 = document.createElement('a');
        a2.href = url2;
        a2.download = 'Sample_Data_Karyawan.xlsx';
        document.body.appendChild(a2);
        a2.click();
        document.body.removeChild(a2);
        URL.revokeObjectURL(url2);
      }, 300);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      setGlobalError(`Gagal mengunduh file contoh: ${msg}`);
    }
  };

  // Download enriched Excel result
  const handleDownloadResult = () => {
    if (!summary) return;

    try {
      const blob = new Blob([summary.outputBuffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      const downloadFileName =
        summary.outputFileName || generateEditedFileName(postTestMeta?.fileName);
      a.download = downloadFileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      setGlobalError(`Gagal mendownload file hasil: ${msg}`);
    }
  };

  return (
    <div className="min-h-screen bg-[#F5F5FA] text-[#121212] flex flex-col font-sans selection:bg-[#D6D6FD]">
      {/* Top Header */}
      <Header
        onOpenGuide={() => setIsGuideOpen(true)}
        onReset={handleReset}
        hasFiles={Boolean(postTestMeta || masterMeta || summary)}
      />

      {/* Main Container */}
      <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-4 sm:py-6 space-y-6">
        {/* Flat Hero Banner (Cooper-inspired flat aesthetic) */}
        <div className="bg-white rounded-3xl p-7 sm:p-9">
          <div className="flex flex-col lg:flex-row items-start lg:items-center justify-between gap-6">
            <div className="max-w-2xl">
              <div className="flex items-center gap-2 mb-3">
                <span className="text-xs font-bold px-3 py-1 bg-[#121212] text-white rounded-full">
                  Instant Matcher
                </span>
                <span className="text-xs font-bold px-3 py-1 bg-[#FEEBB6] text-[#78350F] rounded-full">
                  Preserve Format
                </span>
                <span className="text-xs font-bold px-3 py-1 bg-[#D6D6FD] text-[#1E1B4B] rounded-full">
                  Client-Side
                </span>
              </div>
              <h2 className="text-2xl sm:text-4xl font-extrabold tracking-tight text-[#121212] mb-2.5">
                Penggabung NIP & Nama ke Data Post Test
              </h2>
              <p className="text-sm text-slate-400 font-medium leading-relaxed">
                Tambahkan dua kolom baru (NIP dan Nama Lengkap) secara akurat ke file hasil post test harian Anda berdasarkan data master karyawan. Seluruh rumus, layout, border, dan format asli workbook tetap terjaga seutuhnya.
              </p>
            </div>

            {/* Quick Feature Pill Badge */}
            <div className="flex items-center gap-3 bg-[#F5F5FA] rounded-2xl p-4 shrink-0">
              <div className="w-10 h-10 rounded-xl bg-[#D6D6FD] text-[#1E1B4B] flex items-center justify-center">
                <Zap className="w-5 h-5 fill-current" />
              </div>
              <div>
                <p className="text-xs font-bold text-[#121212]">Proses Cepat & Aman</p>
                <p className="text-[11px] text-slate-400 font-medium">100% di browser tanpa upload</p>
              </div>
            </div>
          </div>
        </div>

        {/* Global Error Notice if any */}
        {globalError && (
          <div className="p-4 bg-[#FFE4E6] rounded-2xl text-xs text-[#881337] flex items-start gap-3">
            <AlertCircle className="w-5 h-5 text-[#F43F5E] shrink-0 mt-0.5" />
            <div>
              <p className="font-bold text-sm">Terjadi Kesalahan</p>
              <p className="mt-1 font-medium">{globalError}</p>
            </div>
          </div>
        )}

        {/* Two Upload Cards Grid */}
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          {/* Card A: Post Test File (Lilac Accent) */}
          <FileUploadCard
            title="File Data Post Test"
            badgeLabel="File Utama"
            accentColor="lilac"
            description="File hasil ujian post test yang akan ditambahkan kolom NIP dan Nama Lengkap (.xlsx)"
            fileMeta={postTestMeta}
            isLoading={postTestLoading}
            error={postTestError}
            onFileSelect={handlePostTestFileSelect}
            onRemoveFile={() => {
              setPostTestMeta(null);
              setPostTestError(null);
              setSummary(null);
            }}
            onSheetChange={handlePostTestSheetChange}
            onNipColumnChange={(col) =>
              postTestMeta &&
              setPostTestMeta({ ...postTestMeta, nipColumn: col })
            }
            onPreview={() =>
              postTestMeta &&
              setPreviewState({
                isOpen: true,
                title: `Cuplikan Data: ${postTestMeta.fileName}`,
                buffer: postTestMeta.buffer,
                sheetName: postTestMeta.selectedSheet,
              })
            }
            existingNameColumnWarning={postTestHasExistingEnrichedCols}
            overwriteMode={overwriteMode}
            onOverwriteModeChange={setOverwriteMode}
          />

          {/* Card B: Master Employee File (Butter Yellow Accent) */}
          <FileUploadCard
            title="File Data Karyawan"
            badgeLabel="Master Referensi"
            accentColor="butter"
            description="File master karyawan berisi daftar NIP dan Nama Karyawan yang valid (.xlsx)"
            fileMeta={masterMeta}
            isLoading={masterLoading}
            error={masterError}
            onFileSelect={handleMasterFileSelect}
            onRemoveFile={() => {
              setMasterMeta(null);
              setMasterError(null);
              setSummary(null);
            }}
            onSheetChange={handleMasterSheetChange}
            onNipColumnChange={(col) =>
              masterMeta && setMasterMeta({ ...masterMeta, nipColumn: col })
            }
            showNameColumnSelector
            onNameColumnChange={(col) =>
              masterMeta && setMasterMeta({ ...masterMeta, nameColumn: col })
            }
            onPreview={() =>
              masterMeta &&
              setPreviewState({
                isOpen: true,
                title: `Cuplikan Data: ${masterMeta.fileName}`,
                buffer: masterMeta.buffer,
                sheetName: masterMeta.selectedSheet,
              })
            }
          />
        </div>

        {/* Action Toolbar (Generate, Reset, Sample Actions) */}
        <ActionToolbar
          canGenerate={canGenerate}
          isProcessing={isGenerating}
          onGenerate={handleGenerate}
          onReset={handleReset}
          onLoadSamples={handleLoadSamples}
          onDownloadSamples={handleDownloadSamples}
          validationMessage={validationMessage}
          hasResult={Boolean(summary)}
        />

        {/* Processing Summary & Result Area */}
        {summary && (
          <ProcessingSummary
            summary={summary}
            onDownload={handleDownloadResult}
            onPreviewOutput={() =>
              setPreviewState({
                isOpen: true,
                title: `Pratinjau Hasil Akhir: ${summary.outputFileName}`,
                buffer: summary.outputBuffer,
                sheetName: summary.targetSheetName,
              })
            }
          />
        )}
      </main>

      {/* Footer */}
      <footer className="py-8 text-center text-xs text-slate-400">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 flex flex-col sm:flex-row items-center justify-between gap-3">
          <p className="font-medium">
            Post Test Excel Enricher · Sistem Pengayaan Hasil Ujian Karyawan
          </p>
          <div className="flex items-center gap-3 text-[11px] font-semibold">
            <span className="px-3 py-1 bg-white rounded-full text-slate-600">
              Client-Side Browser
            </span>
            <span className="px-3 py-1 bg-white rounded-full text-slate-600">
              ExcelJS Native
            </span>
          </div>
        </div>
      </footer>

      {/* Modals */}
      <DataPreviewModal
        isOpen={previewState.isOpen}
        onClose={() => setPreviewState({ ...previewState, isOpen: false })}
        title={previewState.title}
        buffer={previewState.buffer}
        sheetName={previewState.sheetName}
      />

      <HowItWorksModal
        isOpen={isGuideOpen}
        onClose={() => setIsGuideOpen(false)}
      />
    </div>
  );
}
