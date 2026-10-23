/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState } from 'react';
import {
  Check,
  AlertTriangle,
  Download,
  Copy,
  Eye,
  FileCheck2,
  HelpCircle,
} from 'lucide-react';
import type { EnrichmentSummary } from '../types/excel';

interface ProcessingSummaryProps {
  summary: EnrichmentSummary;
  onDownload: () => void;
  onPreviewOutput: () => void;
}

export const ProcessingSummary: React.FC<ProcessingSummaryProps> = ({
  summary,
  onDownload,
  onPreviewOutput,
}) => {
  const [activeTab, setActiveTab] = useState<'unmatched' | 'ambiguous' | 'conflicts'>('unmatched');
  const [copied, setCopied] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');

  const formatFileSize = (bytes: number): string => {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
  };

  const handleCopyUnmatchedNips = () => {
    const textList = summary.unmatchedRows
      .map((r) => r.nip || r.rawText)
      .filter(Boolean)
      .join('\n');
    navigator.clipboard.writeText(textList);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const hasIssues =
    summary.unmatchedCount > 0 ||
    summary.ambiguousCount > 0 ||
    summary.conflictCount > 0;

  const matchRate =
    summary.totalRowsProcessed > 0
      ? Math.round((summary.matchedCount / summary.totalRowsProcessed) * 100)
      : 0;

  const filteredUnmatched = summary.unmatchedRows.filter(
    (r) =>
      r.nip.toLowerCase().includes(searchTerm.toLowerCase()) ||
      r.rawText.toLowerCase().includes(searchTerm.toLowerCase()) ||
      String(r.rowNumber).includes(searchTerm)
  );

  const filteredAmbiguous = summary.ambiguousRows.filter(
    (r) =>
      r.rawText.toLowerCase().includes(searchTerm.toLowerCase()) ||
      String(r.rowNumber).includes(searchTerm) ||
      r.candidateNips.some((c) => c.toLowerCase().includes(searchTerm.toLowerCase()))
  );

  return (
    <div className="bg-white rounded-3xl p-6 sm:p-8 space-y-6">
      {/* Top Banner Status */}
      <div
        className={`p-5 rounded-3xl flex flex-col sm:flex-row sm:items-center justify-between gap-4 ${
          !hasIssues
            ? 'bg-[#D1FAE5] text-[#065F46]'
            : summary.ambiguousCount > 0 || summary.conflictCount > 0
            ? 'bg-[#FFE4E6] text-[#881337]'
            : 'bg-[#D6D6FD] text-[#1E1B4B]'
        }`}
      >
        <div className="flex items-start sm:items-center gap-3.5">
          <div
            className={`w-11 h-11 rounded-2xl flex items-center justify-center shrink-0 ${
              !hasIssues
                ? 'bg-[#10B981] text-white'
                : summary.ambiguousCount > 0 || summary.conflictCount > 0
                ? 'bg-[#F43F5E] text-white'
                : 'bg-[#121212] text-white'
            }`}
          >
            {!hasIssues ? (
              <Check className="w-6 h-6 stroke-[3]" />
            ) : (
              <AlertTriangle className="w-5 h-5" />
            )}
          </div>
          <div>
            <h3 className="text-base font-bold tracking-tight">
              {!hasIssues
                ? 'Pemrosesan Selesai: 100% NIP Berhasil Dicocokkan!'
                : summary.ambiguousCount > 0
                ? `Pemrosesan Selesai: Terdapat ${summary.ambiguousCount} Data Ambigu & ${summary.unmatchedCount} Tidak Cocok`
                : summary.conflictCount > 0
                ? 'Pemrosesan Selesai: Terdapat Konflik Data NIP di Master Karyawan'
                : `Pemrosesan Selesai: ${summary.unmatchedCount} NIP Tidak Ditemukan`}
            </h3>
            <p className="text-xs opacity-80 mt-0.5 font-medium">
              Kolom &quot;{summary.targetNipColumnName}&quot; dan &quot;{summary.targetNameColumnName}&quot; telah ditambahkan ke posisi paling kanan lembar &quot;{summary.targetSheetName}&quot; menggunakan data resmi dari master. Isi asli sel post test tetap utuh.
            </p>
          </div>
        </div>

        {/* Action Group */}
        <div className="flex items-center gap-2.5 shrink-0">
          <button
            type="button"
            onClick={onPreviewOutput}
            className="inline-flex items-center gap-1.5 px-4 py-2.5 text-xs font-bold text-[#121212] bg-white hover:bg-slate-100 rounded-2xl transition-colors cursor-pointer"
          >
            <Eye className="w-4 h-4 text-slate-700" />
            <span>Pratinjau Hasil</span>
          </button>

          <button
            type="button"
            onClick={onDownload}
            className="inline-flex items-center gap-2 px-5 py-2.5 text-xs font-bold text-white bg-[#121212] hover:bg-[#262626] rounded-2xl transition-all cursor-pointer active:scale-[0.99]"
          >
            <Download className="w-4 h-4" />
            <span>Download Excel</span>
          </button>
        </div>
      </div>

      {/* Flat Pastel Metric Cards */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
        {/* Total Baris */}
        <div className="p-5 bg-[#F5F5FA] rounded-2xl flex flex-col justify-between min-h-[110px]">
          <p className="text-xs font-bold text-slate-400 uppercase tracking-wider">
            Total Baris
          </p>
          <div>
            <p className="text-3xl font-extrabold text-[#121212] font-mono tabular-nums tracking-tight">
              {summary.totalRowsProcessed.toLocaleString('id-ID')}
            </p>
            <p className="text-[11px] text-slate-500 font-medium mt-0.5">Baris post test</p>
          </div>
        </div>

        {/* Berhasil Ditemukan (Lilac) */}
        <div className="p-5 bg-[#D6D6FD] rounded-2xl flex flex-col justify-between min-h-[110px] text-[#1E1B4B]">
          <div className="flex items-center justify-between">
            <p className="text-xs font-bold uppercase tracking-wider opacity-80">
              NIP Cocok
            </p>
            <span className="text-xs font-extrabold px-2 py-0.5 bg-white/70 rounded-full font-mono">
              {matchRate}%
            </span>
          </div>
          <div>
            <p className="text-3xl font-extrabold font-mono tabular-nums tracking-tight">
              {summary.matchedCount.toLocaleString('id-ID')}
            </p>
            <p className="text-[11px] opacity-80 font-medium mt-0.5">Nama resmi terisi</p>
          </div>
        </div>

        {/* Tidak Ditemukan (Butter Yellow) */}
        <div className="p-5 bg-[#FEEBB6] rounded-2xl flex flex-col justify-between min-h-[110px] text-[#78350F]">
          <div className="flex items-center justify-between">
            <p className="text-xs font-bold uppercase tracking-wider opacity-80">
              Tidak Ditemukan
            </p>
            {summary.unmatchedCount > 0 && (
              <span className="text-[10px] font-bold bg-[#FF7170] text-white px-2 py-0.5 rounded-full">
                Periksa
              </span>
            )}
          </div>
          <div>
            <p className="text-3xl font-extrabold font-mono tabular-nums tracking-tight">
              {summary.unmatchedCount.toLocaleString('id-ID')}
            </p>
            <p className="text-[11px] opacity-80 font-medium mt-0.5">
              {summary.unmatchedCount > 0 ? 'Nama dikosongkan' : 'Nihil'}
            </p>
          </div>
        </div>

        {/* Data Ambigu atau Konflik (Rose / Pink) */}
        <div className="p-5 bg-[#FFE4E6] rounded-2xl flex flex-col justify-between min-h-[110px] text-[#881337]">
          <div className="flex items-center justify-between">
            <p className="text-xs font-bold uppercase tracking-wider opacity-80">
              Ambigu / Konflik
            </p>
            {(summary.ambiguousCount > 0 || summary.conflictCount > 0) && (
              <span className="text-[10px] font-bold bg-[#F43F5E] text-white px-2 py-0.5 rounded-full">
                Perlu Cek
              </span>
            )}
          </div>
          <div>
            <p className="text-3xl font-extrabold font-mono tabular-nums tracking-tight">
              {(summary.ambiguousCount + summary.conflictCount).toLocaleString('id-ID')}
            </p>
            <p className="text-[11px] opacity-80 font-medium mt-0.5">
              {summary.ambiguousCount > 0
                ? `${summary.ambiguousCount} ambigu, ${summary.conflictCount} konflik`
                : summary.conflictCount > 0
                ? `${summary.conflictCount} konflik master`
                : 'Nihil masalah'}
            </p>
          </div>
        </div>
      </div>

      {/* Warning Notice if Unmatched or Ambiguous exist */}
      {(summary.unmatchedCount > 0 || summary.ambiguousCount > 0) && (
        <div className="p-4 bg-[#FEF3C7] rounded-2xl text-xs text-[#78350F] flex items-start gap-3">
          <div className="w-2 h-2 rounded-full bg-[#D97706] mt-1.5 shrink-0" />
          <div className="space-y-1">
            <p className="font-bold text-[#451A03]">
              Catatan Validasi Data: Terdapat {summary.unmatchedCount + summary.ambiguousCount} baris yang memerlukan perhatian.
            </p>
            <p className="text-[11px] text-[#92400E] leading-relaxed font-medium">
              File Excel hasil tetap dapat diunduh sekarang. Kolom nama pada baris yang tidak ditemukan atau ambigu dibiarkan kosong agar tidak terjadi kesalahan penetapan nama.
            </p>
          </div>
        </div>
      )}

      {/* Discrepancy Breakdown Table if any issues exist */}
      {hasIssues && (
        <div className="bg-[#F5F5FA] rounded-3xl p-4 sm:p-5 space-y-3">
          {/* Tabs header */}
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div className="flex items-center gap-2 flex-wrap">
              <button
                type="button"
                onClick={() => setActiveTab('unmatched')}
                className={`px-4 py-2 text-xs font-bold rounded-xl transition-colors cursor-pointer ${
                  activeTab === 'unmatched'
                    ? 'bg-white text-[#121212]'
                    : 'text-slate-500 hover:text-slate-900'
                }`}
              >
                NIP Tidak Ditemukan ({summary.unmatchedCount})
              </button>

              {summary.ambiguousCount > 0 && (
                <button
                  type="button"
                  onClick={() => setActiveTab('ambiguous')}
                  className={`px-4 py-2 text-xs font-bold rounded-xl transition-colors cursor-pointer ${
                    activeTab === 'ambiguous'
                      ? 'bg-white text-[#881337]'
                      : 'text-slate-500 hover:text-slate-900'
                  }`}
                >
                  Data Ambigu ({summary.ambiguousCount})
                </button>
              )}

              {summary.conflictCount > 0 && (
                <button
                  type="button"
                  onClick={() => setActiveTab('conflicts')}
                  className={`px-4 py-2 text-xs font-bold rounded-xl transition-colors cursor-pointer ${
                    activeTab === 'conflicts'
                      ? 'bg-white text-[#881337]'
                      : 'text-slate-500 hover:text-slate-900'
                  }`}
                >
                  Konflik Master NIP ({summary.conflictCount})
                </button>
              )}
            </div>

            {/* Quick Actions for tab */}
            {activeTab === 'unmatched' && summary.unmatchedCount > 0 && (
              <div className="flex items-center gap-2">
                <input
                  type="text"
                  placeholder="Cari..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  className="text-xs bg-white rounded-xl px-3.5 py-1.5 text-slate-800 placeholder-slate-400 border-0 focus:ring-2 focus:ring-[#D6D6FD] outline-hidden w-36 font-medium"
                />
                <button
                  type="button"
                  onClick={handleCopyUnmatchedNips}
                  className="inline-flex items-center gap-1.5 text-xs font-bold px-3 py-1.5 bg-white text-slate-700 hover:text-[#121212] rounded-xl transition-colors cursor-pointer"
                  title="Salin daftar ke clipboard"
                >
                  {copied ? (
                    <>
                      <Check className="w-3.5 h-3.5 text-emerald-600" />
                      <span className="text-emerald-700">Tersalin</span>
                    </>
                  ) : (
                    <>
                      <Copy className="w-3.5 h-3.5 text-slate-400" />
                      <span>Salin</span>
                    </>
                  )}
                </button>
              </div>
            )}
          </div>

          {/* Table content */}
          <div className="max-h-64 overflow-y-auto rounded-2xl bg-white p-2">
            {activeTab === 'unmatched' && (
              <table className="w-full text-left text-xs">
                <thead className="text-slate-400 font-bold uppercase text-[10px] tracking-wider">
                  <tr>
                    <th className="py-2.5 px-4 w-24">Baris</th>
                    <th className="py-2.5 px-4">Teks Sel Post Test</th>
                    <th className="py-2.5 px-4">Keterangan</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#F5F5FA]">
                  {filteredUnmatched.map((item, idx) => (
                    <tr key={idx} className="hover:bg-[#F5F5FA] transition-colors">
                      <td className="py-2.5 px-4 font-mono font-medium text-slate-500">
                        Baris {item.rowNumber}
                      </td>
                      <td className="py-2.5 px-4 font-mono font-bold text-[#121212]">
                        {item.rawText || item.nip}
                      </td>
                      <td className="py-2.5 px-4 text-[#D97706] font-medium">
                        Tidak ada NIP master yang cocok
                      </td>
                    </tr>
                  ))}
                  {filteredUnmatched.length === 0 && (
                    <tr>
                      <td colSpan={3} className="py-4 text-center text-slate-400 font-medium">
                        Tidak ada data yang sesuai pencarian
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            )}

            {activeTab === 'ambiguous' && (
              <table className="w-full text-left text-xs">
                <thead className="text-slate-400 font-bold uppercase text-[10px] tracking-wider">
                  <tr>
                    <th className="py-2.5 px-4 w-24">Baris</th>
                    <th className="py-2.5 px-4">Teks Sel Post Test</th>
                    <th className="py-2.5 px-4">Kandidat NIP yang Cocok</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#F5F5FA]">
                  {filteredAmbiguous.map((item, idx) => (
                    <tr key={idx} className="hover:bg-[#F5F5FA] transition-colors">
                      <td className="py-2.5 px-4 font-mono font-medium text-slate-500">
                        Baris {item.rowNumber}
                      </td>
                      <td className="py-2.5 px-4 font-mono font-bold text-[#121212]">
                        {item.rawText}
                      </td>
                      <td className="py-2.5 px-4 text-[#F43F5E] font-medium font-mono">
                        {item.candidateNips.join(', ')}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}

            {activeTab === 'conflicts' && (
              <table className="w-full text-left text-xs">
                <thead className="text-slate-400 font-bold uppercase text-[10px] tracking-wider">
                  <tr>
                    <th className="py-2.5 px-4 w-36">NIP Bermasalah</th>
                    <th className="py-2.5 px-4">Variasi Nama Terdaftar di Master</th>
                    <th className="py-2.5 px-4">Baris Master</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#F5F5FA]">
                  {summary.duplicateConflicts.map((conf, idx) => (
                    <tr key={idx} className="hover:bg-[#F5F5FA] transition-colors">
                      <td className="py-2.5 px-4 font-mono font-bold text-[#F43F5E]">
                        {conf.nip}
                      </td>
                      <td className="py-2.5 px-4 text-[#121212]">
                        <ul className="list-disc list-inside space-y-0.5">
                          {conf.names.map((nm, nIdx) => (
                            <li key={nIdx} className="font-semibold">
                              &quot;{nm}&quot;
                            </li>
                          ))}
                        </ul>
                      </td>
                      <td className="py-2.5 px-4 font-mono text-slate-500 font-medium">
                        {conf.rows.join(', ')}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        </div>
      )}

      {/* Download Box Highlight */}
      <div className="p-5 bg-[#F5F5FA] rounded-3xl flex flex-col sm:flex-row items-center justify-between gap-4">
        <div className="flex items-center gap-3.5">
          <div className="w-11 h-11 rounded-2xl bg-[#D6D6FD] text-[#1E1B4B] flex items-center justify-center shrink-0">
            <FileCheck2 className="w-6 h-6" />
          </div>
          <div>
            <p className="text-sm font-bold text-[#121212]">
              {summary.outputFileName}
            </p>
            <p className="text-xs text-slate-400 font-mono mt-0.5 font-medium">
              {formatFileSize(summary.outputFileSize)} · Format Excel Asli (.xlsx)
            </p>
          </div>
        </div>

        <button
          type="button"
          onClick={onDownload}
          className="w-full sm:w-auto inline-flex items-center justify-center gap-2 px-7 py-3 text-xs font-bold text-white bg-[#121212] hover:bg-[#262626] rounded-2xl transition-colors cursor-pointer"
        >
          <Download className="w-4 h-4" />
          <span>Download Excel Sekarang</span>
        </button>
      </div>
    </div>
  );
};
