/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import ExcelJS from 'exceljs';
import type { SheetInfo, WorkbookMeta } from '../types/excel';
import { extractCellValueAsString, isLikelyNipHeader, isLikelyNameHeader } from './nipUtils';

/**
 * Inspects a worksheet to find the most probable header row and column headers.
 */
export function inspectWorksheet(worksheet: ExcelJS.Worksheet): SheetInfo {
  let bestHeaderRowIndex = 1;
  let maxConfidence = -1;
  let bestHeaders: string[] = [];

  const maxScanRows = Math.min(worksheet.rowCount || 10, 10);

  // Scan first few rows to locate the real header
  for (let r = 1; r <= maxScanRows; r++) {
    const row = worksheet.getRow(r);
    if (!row || !row.hasValues) continue;

    const rowHeaders: string[] = [];
    let nonEmptyCount = 0;
    let keywordConfidence = 0;

    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const val = extractCellValueAsString(cell.value).trim();
      rowHeaders[colNumber] = val;
      if (val) {
        nonEmptyCount++;
        if (isLikelyNipHeader(val)) keywordConfidence += 5;
        if (isLikelyNameHeader(val)) keywordConfidence += 5;
      }
    });

    const confidence = nonEmptyCount + keywordConfidence;
    if (confidence > maxConfidence && nonEmptyCount >= 1) {
      maxConfidence = confidence;
      bestHeaderRowIndex = r;
      // Clean up sparse array to clean strings
      bestHeaders = [];
      const colCount = Math.max(worksheet.columnCount, rowHeaders.length - 1);
      for (let c = 1; c <= colCount; c++) {
        bestHeaders.push(rowHeaders[c] || `Kolom ${c}`);
      }
    }
  }

  // Fallback if no header was detected
  if (bestHeaders.length === 0) {
    const colCount = Math.max(worksheet.columnCount, 1);
    for (let c = 1; c <= colCount; c++) {
      bestHeaders.push(`Kolom ${c}`);
    }
  }

  return {
    name: worksheet.name,
    rowCount: worksheet.rowCount || 0,
    columnCount: bestHeaders.length,
    headers: bestHeaders,
    headerRowIndex: bestHeaderRowIndex,
  };
}

/**
 * Reads a File object or ArrayBuffer into an ExcelJS Workbook and parses workbook metadata.
 */
export async function parseWorkbook(
  source: File | ArrayBuffer,
  fileName: string,
  fileSize: number
): Promise<{ workbook: ExcelJS.Workbook; meta: WorkbookMeta }> {
  let buffer: ArrayBuffer;
  if (source instanceof File) {
    buffer = await source.arrayBuffer();
  } else {
    buffer = source;
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  if (!workbook.worksheets || workbook.worksheets.length === 0) {
    throw new Error('File Excel tidak memiliki lembar kerja (worksheet) yang valid.');
  }

  const sheets: SheetInfo[] = [];
  for (const ws of workbook.worksheets) {
    // Only process visible or normal sheets with data
    if (ws.state === 'visible' || !ws.state) {
      sheets.push(inspectWorksheet(ws));
    }
  }

  if (sheets.length === 0) {
    sheets.push(inspectWorksheet(workbook.worksheets[0]));
  }

  // Default fallback sheet
  const defaultSheet = sheets[0];

  const meta: WorkbookMeta = {
    fileName,
    fileSize,
    sheets,
    selectedSheet: defaultSheet.name,
    nipColumn: '',
    nameColumn: '',
    buffer,
  };

  return { workbook, meta };
}

/**
 * Resolves default sheet and column selection for Post Test workbook:
 * 1. Sheet: auto-select sheet named "Data Peserta" (case-insensitive, trimmed).
 *    If not found, fall back to the first sheet.
 * 2. Column: auto-select column named "Player Name" (case-insensitive, trimmed).
 *    If not found, set nipColumn to "" (empty) - DO NOT pick randomly.
 */
export function getPostTestDefaults(
  sheets: SheetInfo[],
  preferredSheetName?: string
): { selectedSheet: string; nipColumn: string; defaultColumnNotFound: boolean } {
  if (!sheets || sheets.length === 0) {
    return { selectedSheet: '', nipColumn: '', defaultColumnNotFound: false };
  }

  // Find sheet
  let targetSheet: SheetInfo | undefined;
  if (preferredSheetName) {
    targetSheet = sheets.find((s) => s.name === preferredSheetName);
  }
  if (!targetSheet) {
    targetSheet = sheets.find(
      (s) => s.name.trim().toLowerCase() === 'data peserta'
    ) || sheets[0];
  }

  // Find column "Player Name"
  const playerColHeader = targetSheet.headers.find(
    (h) => h.trim().toLowerCase() === 'player name'
  );

  if (playerColHeader) {
    return {
      selectedSheet: targetSheet.name,
      nipColumn: playerColHeader,
      defaultColumnNotFound: false,
    };
  } else {
    return {
      selectedSheet: targetSheet.name,
      nipColumn: '',
      defaultColumnNotFound: true,
    };
  }
}

/**
 * Resolves default sheet and columns for Master Employee workbook.
 */
export function getMasterDefaults(
  sheets: SheetInfo[],
  preferredSheetName?: string
): { selectedSheet: string; nipColumn: string; nameColumn: string } {
  if (!sheets || sheets.length === 0) {
    return { selectedSheet: '', nipColumn: '', nameColumn: '' };
  }

  let targetSheet: SheetInfo | undefined;
  if (preferredSheetName) {
    targetSheet = sheets.find((s) => s.name === preferredSheetName);
  }
  if (!targetSheet) {
    targetSheet = sheets.find(
      (s) =>
        s.name.trim().toLowerCase() === 'master karyawan' ||
        s.name.trim().toLowerCase() === 'master'
    ) || sheets[0];
  }

  let nipCol = '';
  for (const h of targetSheet.headers) {
    if (isLikelyNipHeader(h)) {
      nipCol = h;
      break;
    }
  }
  if (!nipCol && targetSheet.headers.length > 0) {
    nipCol =
      targetSheet.headers.find((h) => !h.startsWith('Kolom ')) ||
      targetSheet.headers[0];
  }

  let nameCol = '';
  for (const h of targetSheet.headers) {
    if (isLikelyNameHeader(h)) {
      nameCol = h;
      break;
    }
  }
  if (!nameCol && targetSheet.headers.length > 1) {
    nameCol = targetSheet.headers.find((h) => h !== nipCol) || '';
  }

  return {
    selectedSheet: targetSheet.name,
    nipColumn: nipCol,
    nameColumn: nameCol,
  };
}

