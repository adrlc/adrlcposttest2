/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import ExcelJS from 'exceljs';

/**
 * Generates sample Post Test Excel workbook matching user's real format:
 * Combined NIP and name in a "PLAYER NAME" column with various delimiters.
 */
export async function createSamplePostTestWorkbook(): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Corporate Training Center';
  wb.created = new Date();

  const ws = wb.addWorksheet('Hasil Post Test', {
    views: [{ showGridLines: true }],
  });

  // Setup Columns matching user's quiz platform output
  ws.columns = [
    { header: '#', key: 'no', width: 6 },
    { header: 'PERINGKAT', key: 'rank', width: 14 },
    { header: 'NAMA DEPAN', key: 'firstName', width: 18 },
    { header: 'NAMA BELAKANG', key: 'lastName', width: 18 },
    { header: 'PLAYER NAME', key: 'playerName', width: 34 },
    { header: 'TOTAL POIN', key: 'totalPoints', width: 14 },
    { header: 'STATUS', key: 'status', width: 16 },
  ];

  // Header Styling
  const headerRow = ws.getRow(1);
  headerRow.height = 28;
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11, name: 'Calibri' };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF121212' },
    };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  });

  // Realistic sample rows with different delimiter styles as requested
  const rows = [
    { no: 1, rank: 1, firstName: 'Estherina', lastName: 'Estherina', playerName: '26090115 Estherina Ermi SE', totalPoints: 100, status: 'LULUS' },
    { no: 2, rank: 2, firstName: 'Noval', lastName: 'R', playerName: '26090119_Noval Ramadhan', totalPoints: 95, status: 'LULUS' },
    { no: 3, rank: 3, firstName: 'Aditya', lastName: 'Bakti', playerName: '26100134 Firdaus Aditya Bakti', totalPoints: 95, status: 'LULUS' },
    { no: 4, rank: 4, firstName: 'Khadlel', lastName: 'Romiz', playerName: '26100140 Khadlel Muhammad Romiz', totalPoints: 90, status: 'LULUS' },
    { no: 5, rank: 5, firstName: 'sheva', lastName: '237', playerName: '26100133 Sheva Yan Permana', totalPoints: 90, status: 'LULUS' },
    { no: 6, rank: 6, firstName: 'Cahyo', lastName: 'Wibowo', playerName: '26100138-Cahyo Adi Wibowo', totalPoints: 85, status: 'LULUS' },
    { no: 7, rank: 7, firstName: 'Ayyub', lastName: 'Darmawan', playerName: '26090117___Ayyub Darmawan', totalPoints: 85, status: 'LULUS' },
    { no: 8, rank: 8, firstName: 'Dessy', lastName: 'Putri0212', playerName: '26090121Dessy Putri Mutiarani', totalPoints: 80, status: 'LULUS' },
    { no: 9, rank: 9, firstName: 'Hendi', lastName: 'Sutisna', playerName: '26090123 - Hendi', totalPoints: 80, status: 'LULUS' },
    { no: 10, rank: 10, firstName: 'francisco', lastName: 'dhamma', playerName: '26100135 francisco dhamma', totalPoints: 75, status: 'LULUS' },
    { no: 11, rank: 11, firstName: 'Budi', lastName: 'Prasetyo', playerName: '001201 Budi Prasetyo', totalPoints: 70, status: 'LULUS' },
    { no: 12, rank: 12, firstName: 'Guest', lastName: 'User', playerName: '99999999 Peserta Tamu', totalPoints: 60, status: 'REMIDIAL' }, // Unmatched sample
  ];

  rows.forEach((r, idx) => {
    const row = ws.addRow(r);
    row.height = 22;
    row.eachCell((cell, colNum) => {
      cell.font = { size: 10, name: 'Calibri' };
      if (colNum === 1 || colNum === 2 || colNum === 6 || colNum === 7) {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      } else {
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
      }
    });

    if (idx % 2 === 1) {
      row.eachCell((cell) => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFF9FAFB' },
        };
      });
    }
  });

  const buffer = await wb.xlsx.writeBuffer();
  return buffer as ArrayBuffer;
}

/**
 * Generates sample Master Employee Excel workbook with official names.
 */
export async function createSampleMasterKaryawanWorkbook(): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Corporate Human Capital';
  wb.created = new Date();

  const ws = wb.addWorksheet('Master Karyawan', {
    views: [{ showGridLines: true }],
  });

  ws.columns = [
    { header: 'NIP', key: 'nip', width: 16 },
    { header: 'Nama Karyawan', key: 'nama', width: 32 },
    { header: 'Jabatan', key: 'jabatan', width: 24 },
    { header: 'Unit Kerja', key: 'unit', width: 24 },
  ];

  const headerRow = ws.getRow(1);
  headerRow.height = 28;
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11, name: 'Calibri' };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF1E1B4B' },
    };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  });

  const employees = [
    { nip: '26090115', nama: 'Estherina Ermi, S.E.', jabatan: 'Senior Analyst', unit: 'Operations' },
    { nip: '26090119', nama: 'Noval Ramadhan, S.Kom.', jabatan: 'Lead Developer', unit: 'Technology' },
    { nip: '26100134', nama: 'Firdaus Aditya Bakti, S.T.', jabatan: 'Systems Engineer', unit: 'Infrastructure' },
    { nip: '26100140', nama: 'Khadlel Muhammad Romiz, S.Psi.', jabatan: 'People Specialist', unit: 'Human Capital' },
    { nip: '26100133', nama: 'Sheva Yan Permana, S.M.', jabatan: 'Marketing Executive', unit: 'Commercial' },
    { nip: '26100138', nama: 'Cahyo Adi Wibowo, S.T.', jabatan: 'Quality Engineer', unit: 'Quality Assurance' },
    { nip: '26090117', nama: 'Ayyub Darmawan, S.Kom.', jabatan: 'Database Administrator', unit: 'Technology' },
    { nip: '26090121', nama: 'Dessy Putri Mutiarani, S.Ak.', jabatan: 'Internal Auditor', unit: 'Finance' },
    { nip: '26090123', nama: 'Hendi Sutisna, S.E.', jabatan: 'Account Manager', unit: 'Sales' },
    { nip: '26100135', nama: 'Francisco Dhamma, S.Ds.', jabatan: 'UI/UX Specialist', unit: 'Product' },
    { nip: '001201', nama: 'Budi Prasetyo, S.Kom.', jabatan: 'Principal Architect', unit: 'Architecture' },
  ];

  employees.forEach((emp, idx) => {
    const row = ws.addRow(emp);
    row.height = 22;
    row.eachCell((cell, colNum) => {
      cell.font = { size: 10, name: 'Calibri' };
      if (colNum === 1) {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.numFmt = '@'; // Force text format to preserve leading zeros
      } else {
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
      }
    });

    if (idx % 2 === 1) {
      row.eachCell((cell) => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFF9FAFB' },
        };
      });
    }
  });

  const buffer = await wb.xlsx.writeBuffer();
  return buffer as ArrayBuffer;
}
