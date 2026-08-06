/**
 * Excel Parser Number Format Tests
 *
 * 验证单元格数字格式的显示，特别是多段/含颜色括号等复杂格式。
 */

import { describe, it, expect } from 'vitest';
import ExcelJS from 'exceljs';
import { parseExcel } from '../parser/excel-parser';

async function parseWithFormats(cells: Array<{ row: number; col: number; value: number; numFmt: string }>) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');

  for (const { row, col, value, numFmt } of cells) {
    const cell = sheet.getCell(row, col);
    cell.value = value;
    cell.numFmt = numFmt;
  }

  const buffer = await workbook.xlsx.writeBuffer();
  const parsed = await parseExcel(buffer as ArrayBuffer);
  return parsed.sheets[0];
}

describe('formatNumberValue', () => {
  it('应用带括号与颜色段的单小数格式 0.0_);[Red](0.0)', async () => {
    const sheet = await parseWithFormats([
      { row: 1, col: 1, value: 3979.4210526315787, numFmt: '0.0_);[Red](0.0)' },
      { row: 2, col: 1, value: 13661.57894736842, numFmt: '0.0_);[Red](0.0)' },
    ]);
    expect(sheet.rows[0][0].text).toBe('3979.4');
    expect(sheet.rows[1][0].text).toBe('13661.6');
  });

  it('应用千分位与两位小数格式 #,##0.00', async () => {
    const sheet = await parseWithFormats([
      { row: 1, col: 1, value: 1234567.891, numFmt: '#,##0.00;[Red](#,##0.00)' },
    ]);
    expect(sheet.rows[0][0].text).toBe('1,234,567.89');
  });

  it('应用百分比格式 0.00%', async () => {
    const sheet = await parseWithFormats([
      { row: 1, col: 1, value: 0.375, numFmt: '0.00%' },
    ]);
    expect(sheet.rows[0][0].text).toBe('37.50%');
  });

  it('无格式时去除浮点误差（15 位有效数字）', async () => {
    const sheet = await parseWithFormats([
      { row: 1, col: 1, value: 37.057000000000002, numFmt: 'General' },
    ]);
    expect(sheet.rows[0][0].text).toBe('37.057');
  });
});
