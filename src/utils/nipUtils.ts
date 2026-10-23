/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import type ExcelJS from 'exceljs';
import type { EmployeeIndex, MatchResult } from '../types/excel';

/**
 * Safely extracts a string representation of any ExcelJS cell value.
 * Handles rich text, formulas, numbers, dates, and objects.
 */
export function extractCellValueAsString(value: ExcelJS.CellValue): string {
  if (value === null || value === undefined) {
    return '';
  }

  // Primitive strings & numbers
  if (typeof value === 'string') {
    return value;
  }
  if (typeof value === 'number') {
    return value.toString();
  }
  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE';
  }
  if (value instanceof Date) {
    return value.toISOString();
  }

  // ExcelJS formula result: { formula: '...', result: '...' }
  if (typeof value === 'object') {
    if ('result' in value && value.result !== undefined && value.result !== null) {
      return extractCellValueAsString(value.result as ExcelJS.CellValue);
    }
    // ExcelJS rich text: { richText: [{ text: '...' }] }
    if ('richText' in value && Array.isArray(value.richText)) {
      return value.richText.map((chunk) => chunk.text || '').join('');
    }
    // ExcelJS hyperlink: { text: '...', hyperlink: '...' }
    if ('text' in value && typeof value.text === 'string') {
      return value.text;
    }
    // ExcelJS error: { error: '#N/A' }
    if ('error' in value && typeof value.error === 'string') {
      return value.error;
    }
  }

  return String(value);
}

/**
 * Normalizes NIP by trimming leading and trailing whitespace.
 * PRESERVES leading zeros and all characters. Never converts to numeric truncations.
 */
export function normalizeNip(rawVal: ExcelJS.CellValue): string {
  const str = extractCellValueAsString(rawVal);
  return str.trim();
}

/**
 * Normalizes Name by trimming leading and trailing whitespace.
 */
export function normalizeName(rawVal: ExcelJS.CellValue): string {
  const str = extractCellValueAsString(rawVal);
  return str.trim();
}

/**
 * Builds an EmployeeIndex from raw master data entries.
 * Stores NIP -> unique names & rows mapping, plus sorted NIP list.
 */
export function createEmployeeIndex(
  entries: Array<{ nip: ExcelJS.CellValue; name: ExcelJS.CellValue; rowNumber: number }>
): EmployeeIndex {
  const nipMap = new Map<string, { names: string[]; rows: number[] }>();

  for (const entry of entries) {
    const nip = normalizeNip(entry.nip);
    const name = normalizeName(entry.name);
    if (!nip || !name) continue;

    if (!nipMap.has(nip)) {
      nipMap.set(nip, {
        names: [name],
        rows: [entry.rowNumber],
      });
    } else {
      const existing = nipMap.get(nip)!;
      existing.rows.push(entry.rowNumber);
      // Check if this name is already recorded (case-insensitive check)
      const isAlreadyRecorded = existing.names.some(
        (n) => n.trim().toLowerCase() === name.toLowerCase()
      );
      if (!isAlreadyRecorded) {
        existing.names.push(name);
      }
    }
  }

  // Sort NIP list descending by length (longer NIPs first)
  const nipList = Array.from(nipMap.keys()).sort((a, b) => b.length - a.length);

  return { nipMap, nipList };
}

/**
 * Matches an employee from a single Post Test cell.
 * Handles combined NIP + Name cells with inconsistent delimiters:
 * - 123456 Andi Saputra
 * - 123456_Andi Saputra
 * - 123456-Andi Saputra
 * - 123456Andi Saputra
 * - 123456___Andi Saputra
 * - 123456 - Andi Saputra
 * - 123456 (pure NIP)
 *
 * Uses master NIP list as source of truth.
 */
export function matchEmployeeFromCell(
  cellValue: ExcelJS.CellValue | string,
  employeeIndex: EmployeeIndex
): MatchResult {
  const rawText = typeof cellValue === 'string' ? cellValue : extractCellValueAsString(cellValue);
  const cleanText = rawText.trim();

  if (!cleanText) {
    return {
      status: 'unmatched',
      rawCellText: rawText,
      reason: 'Sel kosong',
    };
  }

  // Find all NIPs from the master list that prefix cleanText
  const matchingNips = employeeIndex.nipList.filter((nip) => cleanText.startsWith(nip));

  if (matchingNips.length === 0) {
    return {
      status: 'unmatched',
      rawCellText: cleanText,
      reason: 'Tidak ada NIP master yang cocok dengan awalan teks',
    };
  }

  // Filter candidates based on suffix boundary:
  // After the candidate NIP, the remaining text must be either:
  // - empty string (exact match)
  // - non-digit character (space, underscore, hyphen, delimiter, or letter)
  const boundaryCandidates = matchingNips.filter((nip) => {
    const remainder = cleanText.slice(nip.length);
    if (remainder.length === 0) return true;
    const firstChar = remainder[0];
    return !/\d/.test(firstChar);
  });

  const candidatesToConsider = boundaryCandidates.length > 0 ? boundaryCandidates : matchingNips;

  if (candidatesToConsider.length === 1) {
    const chosenNip = candidatesToConsider[0];
    const masterData = employeeIndex.nipMap.get(chosenNip);

    if (!masterData || masterData.names.length === 0) {
      return {
        status: 'unmatched',
        nip: chosenNip,
        rawCellText: cleanText,
        reason: 'NIP tidak ditemukan dalam master data',
      };
    }

    // Check if duplicate with DIFFERENT names in master
    if (masterData.names.length > 1) {
      return {
        status: 'conflict_master',
        nip: chosenNip,
        conflictingNames: masterData.names,
        rawCellText: cleanText,
        reason: `Konflik: NIP "${chosenNip}" memiliki ${masterData.names.length} nama berbeda di master data`,
      };
    }

    // Exactly 1 unique official name
    return {
      status: 'matched',
      nip: chosenNip,
      name: masterData.names[0],
      rawCellText: cleanText,
    };
  }

  // Multiple candidates match! Check if one is an exact match
  const exactMatch = candidatesToConsider.find((nip) => nip === cleanText);
  if (exactMatch) {
    const masterData = employeeIndex.nipMap.get(exactMatch);
    if (masterData && masterData.names.length === 1) {
      return {
        status: 'matched',
        nip: exactMatch,
        name: masterData.names[0],
        rawCellText: cleanText,
      };
    }
  }

  // Truly ambiguous: multiple candidates match the prefix
  return {
    status: 'ambiguous',
    candidateNips: candidatesToConsider,
    rawCellText: cleanText,
    reason: `Ditemukan ${candidatesToConsider.length} kandidat NIP (${candidatesToConsider.join(', ')})`,
  };
}

/**
 * Heuristic auto-detector for NIP or Player Name column header
 */
export function isLikelyNipHeader(header: string): boolean {
  const clean = header.toLowerCase().replace(/[^a-z0-9]/g, '');
  const keywords = [
    'nip',
    'playername',
    'player',
    'peserta',
    'namapeserta',
    'idpeserta',
    'noinduk',
    'nomorinduk',
    'employeeid',
    'empid',
    'idkaryawan',
    'nik',
    'nomorpegawai',
    'nopegawai',
    'userid',
    'username',
    'idpegawai',
  ];
  return keywords.some((kw) => clean === kw || clean.startsWith(kw) || clean.endsWith(kw));
}

/**
 * Heuristic auto-detector for Employee Name column header
 */
export function isLikelyNameHeader(header: string): boolean {
  const clean = header.toLowerCase().replace(/[^a-z0-9]/g, '');
  const keywords = [
    'namakaryawan',
    'namapegawai',
    'namalengkap',
    'employeename',
    'fullname',
    'nama',
    'name',
  ];
  return keywords.some((kw) => clean === kw || clean.startsWith(kw));
}
