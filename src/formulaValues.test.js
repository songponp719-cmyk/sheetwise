import assert from 'node:assert/strict';
import test from 'node:test';
import XLSX from '@e965/xlsx';
import { recalculateFormulaValues } from './formulaValues.js';

function createWorkbook() {
  const workbook = XLSX.utils.book_new();
  const staff = XLSX.utils.aoa_to_sheet([
    ['Role', 'Hours', 'Rate', 'Pay'],
    ['Cashier', 8, 20, ''],
    ['Manager', 5, 15, ''],
    ['Cashier', 2, 10, ''],
  ]);
  staff.D2 = { t: 'n', f: 'B2*C2', v: 999, w: '999' };
  staff.D3 = { t: 'n', f: 'B3*C3' };
  staff.D4 = { t: 'n', f: 'B4*C4' };
  const summary = XLSX.utils.aoa_to_sheet([['Cashier pay', 'With bonus']]);
  summary.A2 = { t: 'n', f: 'SUMIF(Staff!A2:A4,"Cashier",Staff!D2:D4)' };
  summary.B2 = { t: 'n', f: 'IF(A2>0,A2*1.1,0)' };
  summary['!ref'] = 'A1:B2';
  XLSX.utils.book_append_sheet(workbook, staff, 'Staff');
  XLSX.utils.book_append_sheet(workbook, summary, 'Summary');
  return workbook;
}

test('recalculates formulas, dependencies, and cross-sheet ranges', async () => {
  const workbook = createWorkbook();

  const issues = await recalculateFormulaValues(workbook, XLSX.utils);

  assert.deepEqual(issues, []);
  assert.equal(workbook.Sheets.Staff.D2.v, 160);
  assert.equal(workbook.Sheets.Staff.D3.v, 75);
  assert.equal(workbook.Sheets.Summary.A2.v, 180);
  assert.ok(Math.abs(workbook.Sheets.Summary.B2.v - 198) < 1e-10);
  assert.equal(
    XLSX.utils.sheet_to_json(workbook.Sheets.Summary, { header: 1, defval: '', raw: false })[1][1],
    '198',
  );
});

test('keeps cached values and reports formulas the evaluator cannot calculate', async () => {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([['Cached', 'Missing']]);
  sheet.A2 = { t: 'n', f: 'XLOOKUP(1,A3:A4,B3:B4)', v: 42, w: '42' };
  sheet.B2 = { t: 'n', f: 'UNSUPPORTED_FUNCTION(A1)' };
  sheet['!ref'] = 'A1:B2';
  XLSX.utils.book_append_sheet(workbook, sheet, 'Data');

  const issues = await recalculateFormulaValues(workbook, XLSX.utils);

  assert.equal(sheet.A2.v, 42);
  assert.equal(sheet.B2.v, '#NAME?');
  assert.equal(sheet.B2.t, 'e');
  assert.equal(issues.length, 2);
  assert.deepEqual(issues.map(({ cell }) => cell).sort(), ['A2', 'B2']);
});

test('writes formula errors as cell errors instead of preserving stale values', async () => {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([['Result']]);
  sheet.A2 = { t: 'n', f: '1/0', v: 12, w: '12' };
  XLSX.utils.book_append_sheet(workbook, sheet, 'Data');

  const issues = await recalculateFormulaValues(workbook, XLSX.utils);

  assert.deepEqual(issues, []);
  assert.equal(sheet.A2.v, '#DIV/0!');
  assert.equal(sheet.A2.t, 'e');
});
