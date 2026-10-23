/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React from 'react';
import { X, Check, ShieldCheck, FileSpreadsheet, AlertCircle } from 'lucide-react';

interface HowItWorksModalProps {
  isOpen: boolean;
  onClose: () => void;
}

export const HowItWorksModal: React.FC<HowItWorksModalProps> = ({ isOpen, onClose }) => {
  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/40 backdrop-blur-xs">
      <div className="bg-white rounded-3xl max-w-2xl w-full max-h-[85vh] flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-150">
        <div className="flex items-center justify-between px-7 py-5 bg-[#F5F5FA]">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-2xl bg-[#121212] text-white flex items-center justify-center">
              <FileSpreadsheet className="w-5 h-5 text-white" />
            </div>
            <div>
              <h3 className="text-sm font-bold text-[#121212]">
                Panduan Format File & Cara Kerja
              </h3>
              <p className="text-xs text-slate-400 font-medium">
                Post Test Excel Enricher — Solusi Internal
              </p>
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="w-8 h-8 rounded-full bg-white hover:bg-slate-200 text-slate-500 hover:text-slate-900 flex items-center justify-center transition-colors cursor-pointer"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <div className="p-7 overflow-y-auto space-y-5 text-xs text-slate-700">
          {/* Privacy Guarantee */}
          <div className="p-4 bg-[#D1FAE5] rounded-2xl flex items-start gap-3">
            <ShieldCheck className="w-5 h-5 text-emerald-800 shrink-0 mt-0.5" />
            <div>
              <h4 className="font-bold text-[#065F46] text-xs">
                Keamanan & Privasi 100% Client-Side
              </h4>
              <p className="text-[11px] text-[#047857] mt-0.5 leading-relaxed font-medium">
                Aplikasi ini berjalan seluruhnya di dalam browser Anda. Tidak ada satu pun data peserta, NIP, atau file Excel yang dikirimkan ke server eksternal atau AI pihak ketiga.
              </p>
            </div>
          </div>

          {/* Section 1 */}
          <div className="p-4 bg-[#F5F5FA] rounded-2xl">
            <h4 className="font-bold text-[#121212] text-xs mb-2 flex items-center gap-2">
              <div className="w-5 h-5 rounded-full bg-[#D6D6FD] text-[#1E1B4B] flex items-center justify-center">
                <Check className="w-3 h-3 stroke-[3]" />
              </div>
              1. Persyaratan File Post Test
            </h4>
            <ul className="space-y-1.5 text-slate-600 pl-7 font-medium">
              <li>• Format file harus <strong>.xlsx</strong>.</li>
              <li>• Memuat kolom <strong>NIP</strong> atau <strong>Player Name / Peserta</strong> (baik NIP murni atau digabung nama seperti <code className="bg-white px-1 rounded font-mono text-[11px]">26090115 Estherina</code>).</li>
              <li>• Pemisah spasi, underscore (<code className="bg-white px-1 rounded font-mono text-[11px]">_</code>), tanda hubung (<code className="bg-white px-1 rounded font-mono text-[11px]">-</code>), atau tanpa pemisah didukung penuh.</li>
              <li>• Seluruh formula, border, style, dan nama sheet asli <strong>dipertahankan 100%</strong>.</li>
              <li>• Dua kolom baru (<strong>&quot;NIP&quot;</strong> lalu <strong>&quot;Nama Lengkap&quot;</strong>) ditambahkan secara berurutan di posisi paling kanan tabel dengan data resmi dari master data karyawan.</li>
            </ul>
          </div>

          {/* Section 2 */}
          <div className="p-4 bg-[#F5F5FA] rounded-2xl">
            <h4 className="font-bold text-[#121212] text-xs mb-2 flex items-center gap-2">
              <div className="w-5 h-5 rounded-full bg-[#FEEBB6] text-[#78350F] flex items-center justify-center">
                <Check className="w-3 h-3 stroke-[3]" />
              </div>
              2. Persyaratan File Master Karyawan
            </h4>
            <ul className="space-y-1.5 text-slate-600 pl-7 font-medium">
              <li>• File referensi format <strong>.xlsx</strong>.</li>
              <li>• Memuat kolom <strong>NIP</strong> dan <strong>Nama Karyawan</strong>.</li>
              <li>• NIP dicocokkan berdasarkan string identitas (bukan urutan baris).</li>
              <li>• Peringatan otomatis muncul bila ada duplikat NIP dengan nama berbeda.</li>
            </ul>
          </div>

          {/* Section 3 */}
          <div className="p-4 bg-[#F5F5FA] rounded-2xl">
            <h4 className="font-bold text-[#121212] text-xs mb-2 flex items-center gap-2">
              <div className="w-5 h-5 rounded-full bg-[#FFE4E6] text-[#881337] flex items-center justify-center">
                <AlertCircle className="w-3 h-3 stroke-[3]" />
              </div>
              3. Penanganan NIP & Angka Nol di Depan
            </h4>
            <p className="text-slate-600 pl-7 font-medium leading-relaxed">
              Sistem memperlakukan NIP murni sebagai teks string. Angka nol di depan (misal: <code className="bg-white px-1.5 py-0.5 rounded-md font-mono text-[11px] text-[#121212]">001201</code>) tidak akan terpotong atau hilang.
            </p>
          </div>
        </div>

        <div className="px-7 py-4 bg-[#F5F5FA] flex justify-end">
          <button
            type="button"
            onClick={onClose}
            className="px-6 py-2.5 text-xs font-bold text-white bg-[#121212] hover:bg-[#262626] rounded-2xl transition-colors cursor-pointer"
          >
            Mengerti & Tutup
          </button>
        </div>
      </div>
    </div>
  );
};
