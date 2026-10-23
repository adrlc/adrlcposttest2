/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import ExcelJS from 'exceljs';
import type {
  AmbiguousRow,
  DuplicateConflict,
  EnrichmentSummary,
  OverwriteMode,
  UnmatchedRow,
} from '../types/excel';
import {
  createEmployeeIndex,
  extractCellValueAsString,
  matchEmployeeFromCell,
} from './nipUtils';
import { inspectWorksheet } from './excelReader';

export interface EnrichOptions {
  postTestFileName?: string;
  postTestSheetName?: string;
  postTestNipCol?: string;
  masterSheetName?: string;
  masterNipCol?: string;
  masterNameCol?: string;
  overwriteMode?: OverwriteMode;
  targetNipColumnName?: string;
  targetNameColumnName?: string;
}

/**
 * Generates download filename following user rules:
 * 1. Takes original Post Test file name.
 * 2. Retains original file name base.
 * 3. Appends `_edited` before `.xlsx` extension.
 * 4. Ensures extension is `.xlsx`.
 * 5. Does not produce duplicate suffixes like `_edited_edited.xlsx` if already edited.
 */
export function generateEditedFileName(originalFileName?: string): string {
  if (!originalFileName || !originalFileName.trim()) {
    return 'PostTest_edited.xlsx';
  }

  const trimmed = originalFileName.trim();

  // Strip extension (.xlsx, .xls) case-insensitively
  let baseName = trimmed.replace(/\.xlsx$/i, '').replace(/\.xls$/i, '');

  if (!baseName) {
    return 'PostTest_edited.xlsx';
  }

  // Prevent multiple _edited suffixes if previously processed
  while (baseName.endsWith('_edited')) {
    baseName = baseName.slice(0, -7);
  }

  if (!baseName) {
    return 'PostTest_edited.xlsx';
  }

  return `${baseName}_edited.xlsx`;
}

/**
 * Finds column index (1-based) by header name or fuzzy match.
 */
function findColumnIndex(headers: string[], targetName: string): number {
  if (!targetName) return -1;
  const targetClean = targetName.trim().toLowerCase();

  // 1. Exact match (case insensitive)
  for (let i = 0; i < headers.length; i++) {
    if ((headers[i] || '').trim().toLowerCase() === targetClean) {
      return i + 1; // 1-based index
    }
  }

  // 2. Starts with or includes
  for (let i = 0; i < headers.length; i++) {
    const h = (headers[i] || '').trim().toLowerCase();
    if (h && (h.includes(targetClean) || targetClean.includes(h))) {
      return i + 1;
    }
  }

  return -1;
}

/**
 * Core function to enrich Post Test Excel with TWO new columns:
 * 1. "NIP" (official NIP from master)
 * 2. "Nama Lengkap" (official full name from master)
 * Preserves original workbook formatting, formulas, styles, and sheets.
 * Handles combined NIP + Name cells via matchEmployeeFromCell.
 */
export async function enrichPostTestWorkbook(
  postTestBuffer: ArrayBuffer,
  masterBuffer: ArrayBuffer,
  options: EnrichOptions = {}
): Promise<EnrichmentSummary> {
  const targetNipHeader = options.targetNipColumnName || 'NIP';
  const targetNameHeader = options.targetNameColumnName || 'Nama Lengkap';

  // --- 1. Load Master Workbook & Parse Mapping ---
  const masterWorkbook = new ExcelJS.Workbook();
  try {
    await masterWorkbook.xlsx.load(masterBuffer);
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    throw new Error(`Gagal membaca file Data Karyawan: ${msg}`);
  }

  const masterSheet = options.masterSheetName
    ? masterWorkbook.getWorksheet(options.masterSheetName)
    : masterWorkbook.worksheets[0];

  if (!masterSheet) {
    throw new Error(
      `Sheet data karyawan "${options.masterSheetName || 'default'}" tidak ditemukan dalam file.`
    );
  }

  const masterSheetInfo = inspectWorksheet(masterSheet);
  const masterNipIndex = findColumnIndex(
    masterSheetInfo.headers,
    options.masterNipCol || 'NIP'
  );
  const masterNameIndex = findColumnIndex(
    masterSheetInfo.headers,
    options.masterNameCol || 'Nama'
  );

  if (masterNipIndex === -1) {
    throw new Error(
      `Kolom NIP (${options.masterNipCol || 'NIP'}) tidak ditemukan pada sheet "${masterSheet.name}". Kolom yang tersedia: ${masterSheetInfo.headers.join(', ')}`
    );
  }
  if (masterNameIndex === -1) {
    throw new Error(
      `Kolom Nama Karyawan (${options.masterNameCol || 'Nama'}) tidak ditemukan pada sheet "${masterSheet.name}". Kolom yang tersedia: ${masterSheetInfo.headers.join(', ')}`
    );
  }

  // Extract all entries from master
  const masterEntries: Array<{ nip: ExcelJS.CellValue; name: ExcelJS.CellValue; rowNumber: number }> = [];
  for (let r = masterSheetInfo.headerRowIndex + 1; r <= masterSheet.rowCount; r++) {
    const row = masterSheet.getRow(r);
    if (!row || !row.hasValues) continue;

    const rawNip = row.getCell(masterNipIndex).value;
    const rawName = row.getCell(masterNameIndex).value;
    if (rawNip !== null && rawNip !== undefined && rawName !== null && rawName !== undefined) {
      masterEntries.push({ nip: rawNip, name: rawName, rowNumber: r });
    }
  }

  // Build employee index
  const employeeIndex = createEmployeeIndex(masterEntries);

  // Extract duplicate conflicts in master
  const duplicateConflicts: DuplicateConflict[] = [];
  employeeIndex.nipMap.forEach((val, nip) => {
    if (val.names.length > 1) {
      duplicateConflicts.push({
        nip,
        names: val.names,
        rows: val.rows,
      });
    }
  });

  // --- 2. Load Post Test Workbook (Preserves all styles, sheets, formulas) ---
  const postTestWorkbook = new ExcelJS.Workbook();
  try {
    await postTestWorkbook.xlsx.load(postTestBuffer);
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    throw new Error(`Gagal membaca file Data Post Test: ${msg}`);
  }

  const postTestSheet = options.postTestSheetName
    ? postTestWorkbook.getWorksheet(options.postTestSheetName)
    : postTestWorkbook.worksheets[0];

  if (!postTestSheet) {
    throw new Error(
      `Sheet post test "${options.postTestSheetName || 'default'}" tidak ditemukan dalam file.`
    );
  }

  const postTestSheetInfo = inspectWorksheet(postTestSheet);
  const postTestNipIndex = findColumnIndex(
    postTestSheetInfo.headers,
    options.postTestNipCol || 'NIP'
  );

  if (postTestNipIndex === -1) {
    throw new Error(
      `Kolom NIP/Peserta (${options.postTestNipCol || 'NIP'}) tidak ditemukan pada sheet "${postTestSheet.name}". Kolom yang tersedia: ${postTestSheetInfo.headers.join(', ')}`
    );
  }

  // --- 3. Determine Target Column Indices for "NIP" and "Nama Lengkap" ---
  // Check if both columns already exist at the end from previous enrichment run
  const headers = postTestSheetInfo.headers;
  const lastCol = headers.length;
  let targetNipColIndex: number;
  let targetNameColIndex: number;
  let wasColumnUpdated = false;

  const isLastColName =
    lastCol >= 2 &&
    (headers[lastCol - 1]?.trim().toLowerCase() === targetNameHeader.toLowerCase() ||
      headers[lastCol - 1]?.trim().toLowerCase() === 'nama karyawan' ||
      headers[lastCol - 1]?.trim().toLowerCase() === 'nama');

  const isSecondLastColNip =
    lastCol >= 2 &&
    headers[lastCol - 2]?.trim().toLowerCase() === targetNipHeader.toLowerCase();

  if (isLastColName && isSecondLastColNip && options.overwriteMode !== 'create_new_column') {
    // Re-use existing columns at the end
    targetNipColIndex = lastCol - 1;
    targetNameColIndex = lastCol;
    wasColumnUpdated = true;
  } else {
    // Append two new columns at the rightmost position
    const headerRow = postTestSheet.getRow(postTestSheetInfo.headerRowIndex);
    let maxColInRow = headers.length;
    headerRow.eachCell((_, colNumber) => {
      if (colNumber > maxColInRow) maxColInRow = colNumber;
    });

    targetNipColIndex = maxColInRow + 1;
    targetNameColIndex = maxColInRow + 2;

    const nipHeaderCell = headerRow.getCell(targetNipColIndex);
    nipHeaderCell.value = targetNipHeader;

    const nameHeaderCell = headerRow.getCell(targetNameColIndex);
    nameHeaderCell.value = targetNameHeader;

    // Inherit header style from adjacent left column
    const adjacentCell = headerRow.getCell(targetNipColIndex - 1);
    const adjacentStyle = adjacentCell && adjacentCell.style ? JSON.parse(JSON.stringify(adjacentCell.style)) : null;

    if (adjacentStyle) {
      nipHeaderCell.style = JSON.parse(JSON.stringify(adjacentStyle));
      nameHeaderCell.style = JSON.parse(JSON.stringify(adjacentStyle));
    } else {
      const defaultHeaderStyle = {
        font: { bold: true, color: { argb: 'FFFFFFFF' }, size: 11, name: 'Calibri' },
        fill: {
          type: 'pattern' as const,
          pattern: 'solid' as const,
          fgColor: { argb: 'FF121212' },
        },
        alignment: { horizontal: 'center' as const, vertical: 'middle' as const },
      };
      nipHeaderCell.style = defaultHeaderStyle;
      nameHeaderCell.style = defaultHeaderStyle;
    }

    // Set appropriate column widths
    const colNipObj = postTestSheet.getColumn(targetNipColIndex);
    colNipObj.width = 18;
    const colNameObj = postTestSheet.getColumn(targetNameColIndex);
    colNameObj.width = 30;
  }

  // --- 4. Iterate Post Test Rows & Populate Both Columns ---
  let totalRowsProcessed = 0;
  let matchedCount = 0;
  let unmatchedCount = 0;
  let ambiguousCount = 0;
  const unmatchedRows: UnmatchedRow[] = [];
  const ambiguousRows: AmbiguousRow[] = [];

  for (
    let r = postTestSheetInfo.headerRowIndex + 1;
    r <= postTestSheet.rowCount;
    r++
  ) {
    const row = postTestSheet.getRow(r);
    if (!row || !row.hasValues) continue;

    let hasContent = false;
    row.eachCell((c) => {
      if (extractCellValueAsString(c.value).trim() !== '') {
        hasContent = true;
      }
    });
    if (!hasContent) continue;

    totalRowsProcessed++;

    const rawCellVal = row.getCell(postTestNipIndex).value;
    const matchResult = matchEmployeeFromCell(rawCellVal, employeeIndex);

    const nipCell = row.getCell(targetNipColIndex);
    const nameCell = row.getCell(targetNameColIndex);

    // Apply basic cell style from adjacent cell if newly created
    if (!wasColumnUpdated) {
      const adjacentCell = row.getCell(targetNipColIndex - 1);
      if (adjacentCell && adjacentCell.style) {
        const baseStyle = {
          font: adjacentCell.font ? { ...adjacentCell.font } : undefined,
          border: adjacentCell.border ? { ...adjacentCell.border } : undefined,
        };
        nipCell.style = {
          ...baseStyle,
          alignment: { vertical: 'middle', horizontal: 'center' },
        };
        nameCell.style = {
          ...baseStyle,
          alignment: { vertical: 'middle', horizontal: 'left' },
        };
      }
    }

    // Crucial: Format NIP as text to preserve leading zeros in Excel
    nipCell.numFmt = '@';

    if (matchResult.status === 'matched' && matchResult.nip && matchResult.name) {
      // Both columns populated with official data from Master
      nipCell.value = String(matchResult.nip);
      nameCell.value = String(matchResult.name);
      matchedCount++;
    } else if (matchResult.status === 'ambiguous') {
      nipCell.value = '';
      nameCell.value = '';
      ambiguousCount++;
      ambiguousRows.push({
        rowNumber: r,
        rawText: matchResult.rawCellText,
        candidateNips: matchResult.candidateNips || [],
      });
    } else if (matchResult.status === 'conflict_master') {
      nipCell.value = '';
      nameCell.value = '';
      unmatchedCount++;
      unmatchedRows.push({
        rowNumber: r,
        nip: matchResult.nip || matchResult.rawCellText,
        rawText: matchResult.rawCellText,
        contextData: { conflict: true, conflictingNames: matchResult.conflictingNames },
      });
    } else {
      // Unmatched
      nipCell.value = '';
      nameCell.value = '';
      unmatchedCount++;
      unmatchedRows.push({
        rowNumber: r,
        nip: matchResult.nip || matchResult.rawCellText || '(Kosong)',
        rawText: matchResult.rawCellText,
      });
    }
  }

  if (totalRowsProcessed === 0) {
    throw new Error(
      'File Post Test tidak memiliki baris data setelah baris judul/header.'
    );
  }

  // --- 5. Generate Output Buffer ---
  let outputBuffer: ArrayBuffer;
  try {
    const rawBuffer = await postTestWorkbook.xlsx.writeBuffer();
    if (rawBuffer instanceof ArrayBuffer) {
      outputBuffer = rawBuffer;
    } else {
      const u8 = new Uint8Array(rawBuffer as ArrayBufferLike);
      const ab = new ArrayBuffer(u8.byteLength);
      new Uint8Array(ab).set(u8);
      outputBuffer = ab;
    }
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    throw new Error(`Gagal menghasilkan file output Excel: ${msg}`);
  }

  const outputFileName = generateEditedFileName(options.postTestFileName);

  return {
    totalRowsProcessed,
    matchedCount,
    unmatchedCount,
    ambiguousCount,
    conflictCount: duplicateConflicts.length,
    unmatchedRows,
    ambiguousRows,
    duplicateConflicts,
    targetSheetName: postTestSheet.name,
    targetColumnName: `${targetNipHeader} & ${targetNameHeader}`,
    targetNipColumnName: targetNipHeader,
    targetNameColumnName: targetNameHeader,
    wasColumnUpdated,
    outputFileName,
    outputBuffer,
    outputFileSize: outputBuffer.byteLength,
    processedAt: new Date(),
  };
}
