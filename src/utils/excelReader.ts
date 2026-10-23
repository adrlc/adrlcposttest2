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

  // Pick first sheet by default
  const defaultSheet = sheets[0];

  // Auto-detect NIP column
  let detectedNip = '';
  for (const header of defaultSheet.headers) {
    if (isLikelyNipHeader(header)) {
      detectedNip = header;
      break;
    }
  }
  if (!detectedNip && defaultSheet.headers.length > 0) {
    // If not detected, fallback to first non-empty header
    detectedNip = defaultSheet.headers.find((h) => !h.startsWith('Kolom ')) || defaultSheet.headers[0];
  }

  // Auto-detect Name column
  let detectedName = '';
  for (const header of defaultSheet.headers) {
    if (isLikelyNameHeader(header)) {
      detectedName = header;
      break;
    }
  }
  if (!detectedName && defaultSheet.headers.length > 1) {
    // Pick another column distinct from NIP
    detectedName = defaultSheet.headers.find((h) => h !== detectedNip) || '';
  }

  const meta: WorkbookMeta = {
    fileName,
    fileSize,
    sheets,
    selectedSheet: defaultSheet.name,
    nipColumn: detectedNip,
    nameColumn: detectedName,
    buffer,
  };

  return { workbook, meta };
}
