/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export interface SheetInfo {
  name: string;
  rowCount: number;
  columnCount: number;
  headers: string[];
  headerRowIndex: number;
}

export interface WorkbookMeta {
  fileName: string;
  fileSize: number;
  sheets: SheetInfo[];
  selectedSheet: string;
  nipColumn: string;
  nameColumn?: string;
  buffer: ArrayBuffer;
}

export interface DuplicateConflict {
  nip: string;
  names: string[];
  rows: number[];
}

export interface UnmatchedRow {
  rowNumber: number;
  nip: string;
  rawText: string;
  contextData?: Record<string, unknown>;
}

export interface AmbiguousRow {
  rowNumber: number;
  rawText: string;
  candidateNips: string[];
}

export interface EnrichmentSummary {
  totalRowsProcessed: number;
  matchedCount: number;
  unmatchedCount: number;
  ambiguousCount: number;
  conflictCount: number;
  unmatchedRows: UnmatchedRow[];
  ambiguousRows: AmbiguousRow[];
  duplicateConflicts: DuplicateConflict[];
  targetSheetName: string;
  targetColumnName: string;
  targetNipColumnName: string;
  targetNameColumnName: string;
  wasColumnUpdated: boolean;
  outputFileName: string;
  outputBuffer: ArrayBuffer;
  outputFileSize: number;
  processedAt: Date;
}

export type OverwriteMode = 'update_existing' | 'create_new_column';

export type MatchStatus = 'matched' | 'unmatched' | 'ambiguous' | 'conflict_master';

export interface MatchResult {
  status: MatchStatus;
  nip?: string;
  name?: string;
  candidateNips?: string[];
  conflictingNames?: string[];
  rawCellText: string;
  reason?: string;
}

export interface EmployeeIndex {
  nipMap: Map<string, { names: string[]; rows: number[] }>;
  nipList: string[];
}
