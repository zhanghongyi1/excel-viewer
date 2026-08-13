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

describe('OOXML theme colors', () => {
  it('maps theme index 1 to the workbook dark text colour, not a light neutral', async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Sheet1');
    const cell = sheet.getCell('A1');
    cell.value = '主题文字';
    cell.font = { name: '微软雅黑', size: 11, color: { theme: 1 } };

    const parsed = await parseExcel(await workbook.xlsx.writeBuffer() as ArrayBuffer);
    expect(parsed.sheets[0].rows[0][0].style.font?.color).toBe('#000000');
  });

  it('resolves theme accents, OOXML tint, and indexed colours through the same path', async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Sheet1');
    sheet.getCell('A1').value = 'accent';
    sheet.getCell('A1').font = { color: { theme: 4 } };
    sheet.getCell('A2').value = 'tint';
    sheet.getCell('A2').font = { color: { theme: 4, tint: 0.5 } };
    sheet.getCell('A3').value = 'indexed';
    sheet.getCell('A3').font = { color: { indexed: 30 } };

    const parsed = await parseExcel(await workbook.xlsx.writeBuffer() as ArrayBuffer);
    expect(parsed.sheets[0].rows[0][0].style.font?.color).toBe('#4F81BD');
    expect(parsed.sheets[0].rows[1][0].style.font?.color).toBe('#a7c0de');
    expect(parsed.sheets[0].rows[2][0].style.font?.color).toBe('#0066CC');
  });
});

describe('font size fidelity', () => {
  it('preserves Excel font sizes instead of applying an extra pt-to-px multiplier', async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Sheet1');
    sheet.getCell('A1').value = '11pt';
    sheet.getCell('A1').font = { name: '微软雅黑', size: 11 };

    const parsed = await parseExcel(await workbook.xlsx.writeBuffer() as ArrayBuffer);
    expect(parsed.sheets[0].rows[0][0].style.font?.size).toBe(11);
  });
});

describe('row height fidelity', () => {
  it('preserves the workbook row height without adding a pt-to-px multiplier', async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Sheet1');
    sheet.getRow(1).height = 25;
    sheet.getCell('A1').value = '25 height';

    const parsed = await parseExcel(await workbook.xlsx.writeBuffer() as ArrayBuffer);
    expect(parsed.sheets[0].rowHeights[0]).toBe(25);
  });
});
