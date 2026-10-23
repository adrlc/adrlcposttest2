/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React from 'react';
import { Sparkles, ShieldCheck, HelpCircle, RotateCcw } from 'lucide-react';

interface HeaderProps {
  onOpenGuide: () => void;
  onReset: () => void;
  hasFiles: boolean;
}

export const Header: React.FC<HeaderProps> = ({ onOpenGuide, onReset, hasFiles }) => {
  return (
    <header className="bg-transparent sticky top-0 z-30 pt-4 pb-2">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="bg-white rounded-3xl px-5 py-3.5 flex items-center justify-between gap-4">
          {/* Brand Pill */}
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-2xl bg-[#121212] text-white flex items-center justify-center shrink-0">
              <Sparkles className="w-5 h-5 text-white" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="text-base font-bold text-[#121212] tracking-tight">
                  Post Test Excel Enricher
                </span>
                <span className="hidden sm:inline-block text-[11px] font-semibold px-2.5 py-0.5 bg-[#FF6B6B] text-white rounded-full">
                  Pro
                </span>
              </div>
              <p className="text-xs text-slate-400 font-medium hidden sm:block">
                Penggabung Nama Karyawan ke Data Post Test
              </p>
            </div>
          </div>

          {/* Center Info Pills */}
          <div className="hidden md:flex items-center gap-2">
            <div className="flex items-center gap-2 px-3.5 py-1.5 rounded-full bg-[#F5F5FA] text-xs font-medium text-slate-700">
              <ShieldCheck className="w-3.5 h-3.5 text-emerald-600" />
              <span>100% Client-Side</span>
              <span className="text-slate-300">·</span>
              <span className="text-slate-500">Tanpa Kirim ke Server</span>
            </div>

            <button
              type="button"
              onClick={onOpenGuide}
              className="flex items-center gap-1.5 px-3.5 py-1.5 rounded-full bg-[#F5F5FA] hover:bg-[#EAEAF4] text-xs font-semibold text-slate-800 transition-colors cursor-pointer"
            >
              <HelpCircle className="w-3.5 h-3.5 text-slate-600" />
              <span>Panduan Format</span>
            </button>
          </div>

          {/* Right Action */}
          <div className="flex items-center gap-2.5">
            {hasFiles && (
              <button
                type="button"
                onClick={onReset}
                className="inline-flex items-center gap-1.5 px-3.5 py-1.5 text-xs font-semibold text-slate-700 bg-[#F5F5FA] hover:bg-[#EAEAF4] rounded-full transition-colors cursor-pointer"
                title="Reset formulir dan data"
              >
                <RotateCcw className="w-3.5 h-3.5 text-slate-500" />
                <span>Reset</span>
              </button>
            )}

            <div className="px-3.5 py-1.5 bg-[#D6D6FD] text-[#1E1B4B] rounded-full text-xs font-bold tracking-tight">
              Enterprise
            </div>
          </div>
        </div>
      </div>
    </header>
  );
};
