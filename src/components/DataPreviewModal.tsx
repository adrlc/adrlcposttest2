/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useEffect, useState } from 'react';
import ExcelJS from 'exceljs';
import { X, FileSpreadsheet, Loader2 } from 'lucide-react';
import { extractCellValueAsString } from '../utils/nipUtils';

interface DataPreviewModalProps {
  isOpen: boolean;
  onClose: () => void;
  title: string;
  buffer: ArrayBuffer | null;
  sheetName?: string;
}

interface PreviewRow {
  rowNumber: number;
  cells: string[];
}

export const DataPreviewModal: React.FC<DataPreviewModalProps> = ({
  isOpen,
  onClose,
  title,
  buffer,
  sheetName,
}) => {
  const [headers, setHeaders] = useState<string[]>([]);
  const [rows, setRows] = useState<PreviewRow[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (!isOpen || !buffer) return;

    let isMounted = true;
    setLoading(true);
    setError(null);

    const loadData = async () => {
      try {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buffer);

        const ws = sheetName ? wb.getWorksheet(sheetName) : wb.worksheets[0];
        if (!ws) {
          throw new Error('Lembar kerja tidak ditemukan');
        }

        // Find header row
        const maxHeaderRow = Math.min(ws.rowCount || 5, 5);
        let headerRowIdx = 1;
        for (let r = 1; r <= maxHeaderRow; r++) {
          const row = ws.getRow(r);
          if (row && row.hasValues) {
            headerRowIdx = r;
            break;
          }
        }

        const headerRow = ws.getRow(headerRowIdx);
        const colHeaders: string[] = [];
        const maxCols = Math.max(ws.columnCount, ws.actualColumnCount, 1);

        for (let c = 1; c <= maxCols; c++) {
          const val = extractCellValueAsString(headerRow.getCell(c).value).trim();
          colHeaders.push(val || `Kolom ${c}`);
        }

        // Read up to 15 data rows
        const dataRows: PreviewRow[] = [];
        const startRow = headerRowIdx + 1;
        const endRow = Math.min(ws.rowCount, startRow + 14);

        for (let r = startRow; r <= endRow; r++) {
          const row = ws.getRow(r);
          if (!row || !row.hasValues) continue;

          const cells: string[] = [];
          for (let c = 1; c <= maxCols; c++) {
            cells.push(extractCellValueAsString(row.getCell(c).value));
          }
          dataRows.push({ rowNumber: r, cells });
        }

        if (isMounted) {
          setHeaders(colHeaders);
          setRows(dataRows);
          setLoading(false);
        }
      } catch (err: unknown) {
        if (isMounted) {
          setError(err instanceof Error ? err.message : String(err));
          setLoading(false);
        }
      }
    };

    loadData();

    return () => {
      isMounted = false;
    };
  }, [isOpen, buffer, sheetName]);

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/40 backdrop-blur-xs">
      <div className="bg-white rounded-3xl max-w-5xl w-full max-h-[85vh] flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-150">
        {/* Modal Header */}
        <div className="flex items-center justify-between px-7 py-5 bg-[#F5F5FA]">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-2xl bg-[#D6D6FD] text-[#1E1B4B] flex items-center justify-center">
              <FileSpreadsheet className="w-5 h-5" />
            </div>
            <div>
              <h3 className="text-sm font-bold text-[#121212]">{title}</h3>
              <p className="text-xs text-slate-400 font-medium">
                Pratinjau 15 baris pertama data lembar kerja
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

        {/* Modal Body */}
        <div className="p-6 overflow-auto flex-1">
          {loading ? (
            <div className="py-16 flex flex-col items-center justify-center gap-2 text-slate-500">
              <Loader2 className="w-7 h-7 animate-spin text-[#121212]" />
              <p className="text-xs font-semibold">Membaca data lembar kerja...</p>
            </div>
          ) : error ? (
            <div className="p-4 bg-rose-100 rounded-2xl text-rose-900 text-xs font-medium">
              {error}
            </div>
          ) : (
            <div className="bg-[#F5F5FA] rounded-2xl p-2 overflow-x-auto">
              <table className="w-full text-left text-xs border-collapse">
                <thead className="text-slate-400 font-bold uppercase text-[10px] tracking-wider">
                  <tr>
                    <th className="py-3 px-3 text-center w-12">#</th>
                    {headers.map((h, i) => (
                      <th
                        key={i}
                        className="py-3 px-3 whitespace-nowrap text-[#121212] font-bold"
                      >
                        {h}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-white/60 bg-white rounded-xl">
                  {rows.map((row) => (
                    <tr key={row.rowNumber} className="hover:bg-[#F5F5FA] transition-colors">
                      <td className="py-2.5 px-3 text-center font-mono text-slate-400 text-[11px] font-semibold">
                        {row.rowNumber}
                      </td>
                      {row.cells.map((val, cellIdx) => (
                        <td
                          key={cellIdx}
                          className="py-2.5 px-3 whitespace-nowrap font-mono text-slate-800 text-[11px] font-medium"
                        >
                          {val || <span className="text-slate-300 italic">—</span>}
                        </td>
                      ))}
                    </tr>
                  ))}
                  {rows.length === 0 && (
                    <tr>
                      <td
                        colSpan={headers.length + 1}
                        className="py-8 text-center text-slate-400 font-medium"
                      >
                        Tidak ada baris data pada lembar kerja ini.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          )}
        </div>

        {/* Modal Footer */}
        <div className="px-7 py-4 bg-[#F5F5FA] flex justify-end">
          <button
            type="button"
            onClick={onClose}
            className="px-6 py-2.5 text-xs font-bold text-slate-700 bg-white hover:bg-slate-100 rounded-2xl transition-colors cursor-pointer"
          >
            Tutup
          </button>
        </div>
      </div>
    </div>
  );
};
