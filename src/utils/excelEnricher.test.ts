/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import test from 'node:test';
import assert from 'node:assert/strict';
import {
  normalizeNip,
  createEmployeeIndex,
  matchEmployeeFromCell,
  isLikelyNipHeader,
  isLikelyNameHeader,
} from './nipUtils';
import { enrichPostTestWorkbook, generateEditedFileName } from './excelEnricher';
import {
  createSamplePostTestWorkbook,
  createSampleMasterKaryawanWorkbook,
} from './sampleGenerator';
import ExcelJS from 'exceljs';

test('generateEditedFileName follows all naming rules without duplicate _edited', () => {
  // 1. Standard file
  assert.equal(generateEditedFileName('PostTest_K3.xlsx'), 'PostTest_K3_edited.xlsx');

  // 2. File with spaces
  assert.equal(
    generateEditedFileName('Hasil Post Test Oktober.xlsx'),
    'Hasil Post Test Oktober_edited.xlsx'
  );

  // 3. File with date in name
  assert.equal(
    generateEditedFileName('PostTest_2026-10-21.xlsx'),
    'PostTest_2026-10-21_edited.xlsx'
  );

  // 4. File already having _edited (must NOT produce _edited_edited.xlsx)
  assert.equal(
    generateEditedFileName('PostTest_K3_edited.xlsx'),
    'PostTest_K3_edited.xlsx'
  );

  // 5. File having duplicate _edited_edited
  assert.equal(
    generateEditedFileName('PostTest_K3_edited_edited.xlsx'),
    'PostTest_K3_edited.xlsx'
  );

  // 6. Uppercase extension
  assert.equal(
    generateEditedFileName('Laporan_PostTest.XLSX'),
    'Laporan_PostTest_edited.xlsx'
  );

  // 7. Fallback when undefined or empty
  assert.equal(generateEditedFileName(''), 'PostTest_edited.xlsx');
  assert.equal(generateEditedFileName(undefined), 'PostTest_edited.xlsx');
});

test('normalizeNip preserves leading zeros and trims whitespace', () => {
  assert.equal(normalizeNip('001234'), '001234');
  assert.equal(normalizeNip('  001234  '), '001234');
  assert.equal(normalizeNip('000987654321'), '000987654321');
  assert.equal(normalizeNip(1234), '1234');
  assert.equal(normalizeNip(''), '');
});

test('matchEmployeeFromCell extracts NIP with various delimiter styles', () => {
  const employeeIndex = createEmployeeIndex([
    { nip: '123456', name: 'Andi Saputra, S.Kom', rowNumber: 2 },
    { nip: '26090115', name: 'Estherina Ermi, S.E.', rowNumber: 3 },
    { nip: '26090119', name: 'Noval Ramadhan, S.Kom', rowNumber: 4 },
    { nip: '001201', name: 'Budi Prasetyo, S.Kom', rowNumber: 5 },
  ]);

  // Pattern 1: Space delimiter (123456 Andi Saputra)
  const res1 = matchEmployeeFromCell('123456 Andi Saputra', employeeIndex);
  assert.equal(res1.status, 'matched');
  assert.equal(res1.nip, '123456');
  assert.equal(res1.name, 'Andi Saputra, S.Kom');

  // Pattern 2: Underscore delimiter (123456_Andi Saputra)
  const res2 = matchEmployeeFromCell('123456_Andi Saputra', employeeIndex);
  assert.equal(res2.status, 'matched');
  assert.equal(res2.nip, '123456');

  // Pattern 3: Dash delimiter (123456-Andi Saputra)
  const res3 = matchEmployeeFromCell('123456-Andi Saputra', employeeIndex);
  assert.equal(res3.status, 'matched');
  assert.equal(res3.nip, '123456');

  // Pattern 4: No delimiter, directly followed by letters (123456Andi Saputra)
  const res4 = matchEmployeeFromCell('123456Andi Saputra', employeeIndex);
  assert.equal(res4.status, 'matched');
  assert.equal(res4.nip, '123456');

  // Pattern 5: Multiple underscores (123456___Andi Saputra)
  const res5 = matchEmployeeFromCell('123456___Andi Saputra', employeeIndex);
  assert.equal(res5.status, 'matched');
  assert.equal(res5.nip, '123456');

  // Pattern 6: Space dash space (123456 - Andi Saputra)
  const res6 = matchEmployeeFromCell('123456 - Andi Saputra', employeeIndex);
  assert.equal(res6.status, 'matched');
  assert.equal(res6.nip, '123456');

  // Pattern 7: Leading zero NIP (001201 Budi)
  const res7 = matchEmployeeFromCell('001201 Budi', employeeIndex);
  assert.equal(res7.status, 'matched');
  assert.equal(res7.nip, '001201');
  assert.equal(res7.name, 'Budi Prasetyo, S.Kom');

  // Pattern 8: Pure NIP only
  const res8 = matchEmployeeFromCell('26090115', employeeIndex);
  assert.equal(res8.status, 'matched');
  assert.equal(res8.nip, '26090115');

  // Pattern 9: Unmatched NIP
  const res9 = matchEmployeeFromCell('999999 John Doe', employeeIndex);
  assert.equal(res9.status, 'unmatched');
});

test('matchEmployeeFromCell detects duplicate conflicting names in master', () => {
  const employeeIndex = createEmployeeIndex([
    { nip: '001001', name: 'Alice Anderson', rowNumber: 2 },
    { nip: '001001', name: 'Alicia Smith', rowNumber: 3 },
  ]);

  const res = matchEmployeeFromCell('001001 Alice', employeeIndex);
  assert.equal(res.status, 'conflict_master');
  assert.equal(res.nip, '001001');
  assert.equal(res.conflictingNames?.length, 2);
});

test('matchEmployeeFromCell handles ambiguous prefix candidates without guessing', () => {
  const employeeIndex = createEmployeeIndex([
    { nip: '123', name: 'User A', rowNumber: 2 },
    { nip: '123A', name: 'User B', rowNumber: 3 },
  ]);

  const res = matchEmployeeFromCell('123Andi', employeeIndex);
  assert.equal(res.status, 'ambiguous');
  assert.ok(res.candidateNips && res.candidateNips.length >= 2);
});

test('header auto-detection recognizes PLAYER NAME and NIP variants', () => {
  assert.equal(isLikelyNipHeader('PLAYER NAME'), true);
  assert.equal(isLikelyNipHeader('Player'), true);
  assert.equal(isLikelyNipHeader('NIP'), true);
  assert.equal(isLikelyNipHeader('Nama Peserta'), true);
  assert.equal(isLikelyNameHeader('Nama Karyawan'), true);
  assert.equal(isLikelyNameHeader('Nama Lengkap'), true);
});

test('enrichPostTestWorkbook appends two columns in exact order: NIP then Nama Lengkap', async () => {
  const postTestBuffer = await createSamplePostTestWorkbook();
  const masterBuffer = await createSampleMasterKaryawanWorkbook();

  const summary = await enrichPostTestWorkbook(postTestBuffer, masterBuffer, {
    postTestFileName: 'PostTest_K3.xlsx',
    postTestSheetName: 'Hasil Post Test',
    masterSheetName: 'Master Karyawan',
    postTestNipCol: 'PLAYER NAME',
    masterNipCol: 'NIP',
    masterNameCol: 'Nama Karyawan',
  });

  assert.equal(summary.outputFileName, 'PostTest_K3_edited.xlsx');
  assert.equal(summary.totalRowsProcessed, 12);
  assert.equal(summary.matchedCount, 11);
  assert.equal(summary.unmatchedCount, 1);

  // Load enriched output workbook and check columns
  const resultWb = new ExcelJS.Workbook();
  await resultWb.xlsx.load(summary.outputBuffer);
  const resultWs = resultWb.getWorksheet('Hasil Post Test')!;
  assert.ok(resultWs);

  // Original columns were 7 (#, PERINGKAT, NAMA DEPAN, NAMA BELAKANG, PLAYER NAME, TOTAL POIN, STATUS)
  // Column 8 must be 'NIP'
  const headerNip = resultWs.getRow(1).getCell(8);
  assert.equal(headerNip.value, 'NIP');

  // Column 9 must be 'Nama Lengkap'
  const headerName = resultWs.getRow(1).getCell(9);
  assert.equal(headerName.value, 'Nama Lengkap');

  // Row 2: "26090115 Estherina Ermi SE"
  // Col 8 (NIP) should be "26090115"
  // Col 9 (Nama Lengkap) should be "Estherina Ermi, S.E."
  const row2 = resultWs.getRow(2);
  assert.equal(row2.getCell(8).value, '26090115');
  assert.equal(row2.getCell(9).value, 'Estherina Ermi, S.E.');

  // Row 12: "001201 Budi Prasetyo" (leading zero!)
  // Col 8 (NIP) should be "001201"
  // Col 9 (Nama Lengkap) should be "Budi Prasetyo, S.Kom."
  const row12 = resultWs.getRow(12);
  assert.equal(row12.getCell(8).value, '001201');
  assert.equal(row12.getCell(9).value, 'Budi Prasetyo, S.Kom.');

  // Row 13: "99999999 Peserta Tamu" (unmatched)
  // Both Col 8 and Col 9 must be empty
  const row13 = resultWs.getRow(13);
  assert.equal(row13.getCell(8).value, '');
  assert.equal(row13.getCell(9).value, '');

  // Now test re-enriching the already-enriched workbook:
  // Must NOT create duplicate columns (must keep column 8 as NIP and column 9 as Nama Lengkap)!
  const reEnrichSummary = await enrichPostTestWorkbook(summary.outputBuffer, masterBuffer, {
    postTestSheetName: 'Hasil Post Test',
    masterSheetName: 'Master Karyawan',
    postTestNipCol: 'PLAYER NAME',
    masterNipCol: 'NIP',
    masterNameCol: 'Nama Karyawan',
  });

  assert.equal(reEnrichSummary.wasColumnUpdated, true);
  const reResultWb = new ExcelJS.Workbook();
  await reResultWb.xlsx.load(reEnrichSummary.outputBuffer);
  const reResultWs = reResultWb.getWorksheet('Hasil Post Test')!;

  // Column 8 is NIP, Column 9 is Nama Lengkap, and Column 10 does NOT exist / is empty
  assert.equal(reResultWs.getRow(1).getCell(8).value, 'NIP');
  assert.equal(reResultWs.getRow(1).getCell(9).value, 'Nama Lengkap');
  assert.equal(reResultWs.getRow(1).getCell(10).value, null);
});
