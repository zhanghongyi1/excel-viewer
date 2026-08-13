/**
 * Excel 表格数据解析器
 *
 * 使用 exceljs 解析 xlsx 文件，提取：
 * - 工作表元信息（名称、行高、列宽）
 * - 单元格值与类型（数字、字符串、日期、公式等）
 * - 单元格样式（字体、背景色、边框、对齐）
 * - 合并单元格区域
 */

import ExcelJS from 'exceljs';
import { colLetterToNumber } from '../utils/ooxml';
import { HyperFormula, DetailedCellError } from 'hyperformula';
import type { RawCellContent } from 'hyperformula';
import { parseThemeFromZip } from '../chart/theme-parser';
import type { ChartTheme } from '../chart/theme-parser';
import type {
  ParsedWorkbook,
  ParsedSheet,
  ParsedCell,
  CellType,
  CellStyle,
  MergeInfo,
  FreezePane,
  ConditionalFormatting,
  CfRule,
  RichTextRun,
} from '../types';
// ===== 常量定义 =====

/** 默认列宽 (像素) */
const DEFAULT_COL_WIDTH = 80;

/** 默认行高 (像素) */
const DEFAULT_ROW_HEIGHT = 20;

/** Excel 列宽转像素的近似比例 */
const COL_WIDTH_TO_PX = 7; // exceljs 列宽单位 ≈ 7px

/** Excel 网格线和边距额外像素 */
const COL_WIDTH_MARGIN = 5;

/** Excel 内置 indexed 色板（OOXML 标准色表的可携带部分）。 */
const INDEXED_COLORS: Record<number, string> = {
  0: '#000000', 1: '#FFFFFF', 2: '#FF0000', 3: '#00FF00', 4: '#0000FF',
  5: '#FFFF00', 6: '#FF00FF', 7: '#00FFFF', 8: '#000000', 9: '#FFFFFF',
  10: '#FF0000', 11: '#00FF00', 12: '#0000FF', 13: '#FFFF00', 14: '#FF00FF',
  15: '#00FFFF', 16: '#800000', 17: '#008000', 18: '#000080', 19: '#808000',
  20: '#800080', 21: '#008080', 22: '#C0C0C0', 23: '#808080',
  24: '#9999FF', 25: '#993366', 26: '#FFFFCC', 27: '#CCFFFF', 28: '#660066',
  29: '#FF8080', 30: '#0066CC', 31: '#CCCCFF', 32: '#000080', 33: '#FF00FF',
  34: '#FFFF00', 35: '#00FFFF', 36: '#800080', 37: '#800000', 38: '#008080',
  39: '#0000FF', 40: '#00CCFF', 41: '#CCFFFF', 42: '#CCFFCC', 43: '#FFFF99',
  44: '#99CCFF', 45: '#FF99CC', 46: '#CC99FF', 47: '#FFCC99', 48: '#3366FF',
  49: '#33CCCC', 50: '#99CC00', 51: '#FFCC00', 52: '#FF9900', 53: '#FF6600',
  54: '#666699', 55: '#969696', 56: '#003366', 57: '#339966', 58: '#003300',
  59: '#333300', 60: '#993300', 61: '#993366', 62: '#333399', 63: '#333333',
};

/**
 * OOXML theme color indices are positional, not a greyscale palette:
 * 0=lt1, 1=dk1, 2=lt2, 3=dk2, 4-9=accent1-6, 10=hlink, 11=folHlink.
 */
function themeColorsFromTheme(theme: ChartTheme): string[] {
  const { colors } = theme;
  return [
    colors.lt1, colors.dk1, colors.lt2, colors.dk2,
    colors.accent1, colors.accent2, colors.accent3,
    colors.accent4, colors.accent5, colors.accent6,
    colors.hlink, colors.folHlink,
  ];
}

// ===== 类型转换工具函数 =====

/**
 * 将 ARGB 颜色值转换为 RGB 十六进制
 * Excel 中颜色格式为 AARRGGBB 或 RRGGBB
 */
function argbToHex(color: string | undefined): string {
  if (!color) return '#000000';

  // 如果已经是 #RRGGBB 格式，直接返回
  if (color.startsWith('#') && color.length === 7) return color;

  const trimmed = color.trim();

  // 尝试匹配 ARGB 格式 (8位十六进制)
  const argbMatch = trimmed.match(/^#?([a-fA-F0-9]{2})([a-fA-F0-9]{2})([a-fA-F0-9]{2})([a-fA-F0-9]{2})$/);
  if (argbMatch) {
    // 忽略 alpha 通道，返回 RGB
    return `#${argbMatch[2]}${argbMatch[3]}${argbMatch[4]}`;
  }

  // 尝试匹配 RGB 格式 (6位十六进制)
  const rgbMatch = trimmed.match(/^#?([a-fA-F0-9]{6})$/);
  if (rgbMatch) {
    return `#${rgbMatch[1]}`;
  }

  // 无法识别的颜色
  console.warn(`[excel-preview] Unknown color format: ${color}`);
  return '#000000';
}

/**
 * 获取主题颜色
 */
function getThemeColor(themeIndex: number, tint: number | string | undefined, themeColors: string[]): string {
  if (themeIndex < 0 || themeIndex >= themeColors.length) {
    return '#C7C9CC';
  }

  let baseColor = themeColors[themeIndex];

  const numericTint = typeof tint === 'string' ? Number(tint) : tint;
  if (numericTint !== undefined && numericTint !== null && Number.isFinite(numericTint)) {
    baseColor = applyTint(baseColor, numericTint);
  }

  return baseColor;
}

/** Resolve ExcelJS colour objects using the workbook's actual OOXML theme. */
function resolveColor(color: any, themeColors: string[], fallback = '#000000'): string {
  if (!color) return fallback;
  if (typeof color === 'string') return argbToHex(color);
  if (color.argb) return argbToHex(color.argb);
  if (color.theme !== undefined) return getThemeColor(color.theme, color.tint, themeColors);
  if (color.indexed !== undefined) return INDEXED_COLORS[color.indexed] || fallback;
  return fallback;
}

/**
 * 应用色调调整
 */
function applyTint(hexColor: string, tint: number): string {
  // OOXML tint adjusts HSL luminosity, not each RGB channel independently.
  // Clamp malformed producer values so a damaged file cannot generate invalid CSS.
  const safeTint = Math.max(-1, Math.min(1, tint));
  const hex = hexColor.replace('#', '');
  const r = parseInt(hex.substring(0, 2), 16) / 255;
  const g = parseInt(hex.substring(2, 4), 16) / 255;
  const b = parseInt(hex.substring(4, 6), 16) / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  let h = 0;
  let s = 0;
  let l = (max + min) / 2;

  if (max !== min) {
    const delta = max - min;
    s = l > 0.5 ? delta / (2 - max - min) : delta / (max + min);
    if (max === r) h = (g - b) / delta + (g < b ? 6 : 0);
    else if (max === g) h = (b - r) / delta + 2;
    else h = (r - g) / delta + 4;
    h /= 6;
  }

  l = safeTint < 0 ? l * (1 + safeTint) : l * (1 - safeTint) + safeTint;
  const hueToRgb = (p: number, q: number, t: number): number => {
    let value = t;
    if (value < 0) value += 1;
    if (value > 1) value -= 1;
    if (value < 1 / 6) return p + (q - p) * 6 * value;
    if (value < 1 / 2) return q;
    if (value < 2 / 3) return p + (q - p) * (2 / 3 - value) * 6;
    return p;
  };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const toHex = (value: number) => Math.round(value * 255).toString(16).padStart(2, '0');
  const rgb = s === 0
    ? [l, l, l]
    : [hueToRgb(p, q, h + 1 / 3), hueToRgb(p, q, h), hueToRgb(p, q, h - 1 / 3)];
  return `#${toHex(rgb[0])}${toHex(rgb[1])}${toHex(rgb[2])}`;
}

// ===== 核心解析函数 =====

/**
 * 解析单元格样式
 */
function parseCellStyle(cell: ExcelJS.Cell, themeColors: string[]): CellStyle {
  const style: CellStyle = {};
  const cellStyle: any = (cell as any).style || {};

    // 字体样式
  if (cellStyle.font) {
    style.font = {};
    if (cellStyle.font.name) style.font.name = cellStyle.font.name;
    if (cellStyle.font.size !== undefined) {
      // Browser CSS uses a different text rasterisation path than Excel.
      // Using the physical pt→px multiplier made an 11pt Excel font render
      // as 15px, visibly larger than the source workbook. Keep the numeric
      // size so the preview matches Excel's on-screen visual scale.
      style.font.size = cellStyle.font.size;
    }
    if (cellStyle.font.bold !== undefined) style.font.bold = cellStyle.font.bold;
    if (cellStyle.font.italic !== undefined) style.font.italic = cellStyle.font.italic;
    if (cellStyle.font.underline !== undefined) style.font.underline = !!cellStyle.font.underline;
    if (cellStyle.font.strike !== undefined) style.font.strike = cellStyle.font.strike;

    // 字体颜色
    if (cellStyle.font.color) {
      style.font.color = resolveColor(cellStyle.font.color, themeColors);
    }
  }

    // 背景色 (填充)
  if ((cellStyle as any).fill && (cellStyle as any).fill?.fgColor) {
    const fgColor: any = (cellStyle as any).fill.fgColor;
    style.bgcolor = resolveColor(fgColor, themeColors, '#C7C9CC');
  }

  // 对齐方式
  if (cellStyle.alignment) {
    if (cellStyle.alignment.horizontal) {
      style.align = cellStyle.alignment.horizontal as 'left' | 'center' | 'right';
    }
    if (cellStyle.alignment.vertical) {
      style.valign = cellStyle.alignment.vertical as 'top' | 'middle' | 'bottom';
    }
    if (cellStyle.alignment.wrapText) {
      style.textwrap = true;
    }
  }

  // 边框
  if (cellStyle.border) {
    style.border = {};

    const borderPositions: Array<'top' | 'bottom' | 'left' | 'right'> = [
      'top', 'bottom', 'left', 'right'
    ];

    for (const pos of borderPositions) {
      const border = cellStyle.border[pos];
      if (border) {
        let borderColor = '#000000';
        if (border.color) {
          borderColor = resolveColor(border.color, themeColors);
        }
        style.border[pos] = [border.style || 'thin', borderColor];
      }
    }
  }

  // 数字格式
  if (cellStyle.numFmt) {
    style.numFmt = cellStyle.numFmt;
  }

  return style;
}

/**
 * 获取单元格类型
 */
/**
 * 归一化单元格 value：对公式单元格提取原始结果值，避免下游拿到公式对象
 */
function normalizeCellValue(cell: ExcelJS.Cell): any {
  if (cell.type !== ExcelJS.ValueType.Formula) {
    return cell.value;
  }
  const v = cell.value;
  if (typeof v !== 'object' || v === null) {
    return v;
  }
  const formulaVal = v as { result?: any; error?: string };
  if (formulaVal.error) return formulaVal.error;
  if (formulaVal.result !== undefined) {
    if (typeof formulaVal.result === 'object' && formulaVal.result !== null) {
      return formulaVal.result.error || null;
    }
    return formulaVal.result;
  }
  return v;
}

type FormulaCalculator = {
  getValue: (sheetName: string, row: number, column: number) => any;
  dispose: () => void;
};

function getFormula(cell: ExcelJS.Cell): string | undefined {
  if (cell.type !== ExcelJS.ValueType.Formula) return undefined;

  const value = cell.value;
  if (typeof value === 'object' && value !== null && 'formula' in value) {
    return String((value as { formula: string }).formula);
  }

  const formula = (cell as any).formula;
  return typeof formula === 'string' ? formula : undefined;
}

/** Excel 序列号与 Unix 纪元(1970-01-01)之间的天数差 (基于 1899-12-30, 兼容 Excel 1900 闰年 bug) */
const EXCEL_EPOCH_OFFSET_DAYS = 25569;

/**
 * 检测 Excel 数字格式是否为日期/时间格式
 * 通过识别 y/m/d/h/s 及 [h]/[m]/[s] 等时间令牌判断，
 * 忽略引号包裹的文本，避免 "#,##0.00"、"0 \"kg\"" 等误判
 */
function isDateFormat(numFmt: string | undefined): boolean {
  if (!numFmt) return false;
  const stripped = numFmt.replace(/"[^"]*"/g, '');
  return /(^|[^a-z])(y{1,4}|m{1,2}|d{1,2}|h{1,2}|s{1,2}|\[h\]|\[m\]|\[s\])([^a-z]|$)/i.test(stripped);
}

/** Excel 日期序列号转 JS Date */
function excelSerialToDate(serial: number): Date {
  return new Date(Math.round((serial - EXCEL_EPOCH_OFFSET_DAYS) * 86400000));
}

/**
 * 以「文本」语义写入 HyperFormula, 匹配 Excel 的文本处理规则:
 * - 加单引号前缀强制所有字符串视为文本(否则 "41128.367" 会被自动转为数字、"TRUE" 转为布尔)
 * - 空字符串写成 "''", 使 HyperFormula 存为非数值文本; 若只写单个 "'" 会被剥掉存为空串,
 *   参与除法等算术时被强转为 0 (0/0 → #DIV/0!), 与 Excel 的 #VALUE! 不一致
 */
function toTextContent(text: string): string {
  return text.length === 0 ? "''" : `'${text}`;
}

/**
 * 将 exceljs 单元格值归一化为 HyperFormula 可接受的原始内容
 * (公式/字符串/数字/布尔/日期/错误, 对象类型需提取纯文本)
 */
function toRawCellContent(cell: ExcelJS.Cell): RawCellContent {
  if (cell.type === ExcelJS.ValueType.Formula) {
    const formula = getFormula(cell);
    return formula ? `=${formula}` : null;
  }

  const value = cell.value;
  if (value === null || value === undefined) return null;

  if (cell.type === ExcelJS.ValueType.RichText) {
    const richText = value as any;
    const text = richText.text ?? richText.richText?.map((r: any) => r.text).join('') ?? '';
    return toTextContent(text);
  }

  if (cell.type === ExcelJS.ValueType.Hyperlink) {
    const hyperlink = value as ExcelJS.CellHyperlinkValue;
    const text = hyperlink.text ?? hyperlink.hyperlink ?? '';
    return toTextContent(text);
  }

  if (cell.type === ExcelJS.ValueType.Error) {
    return (value as { error?: string })?.error ?? String(value);
  }

  if (value instanceof Date) return value;

  if (typeof value === 'object') {
    const error = (value as { error?: string }).error;
    return typeof error === 'string' ? error : null;
  }

  return typeof value === 'string' ? toTextContent(value) : (value as number | boolean);
}

function createFormulaCalculator(workbook: ExcelJS.Workbook): FormulaCalculator | undefined {
  try {
    const hf = HyperFormula.buildEmpty({
      licenseKey: 'gpl-v3',
      maxRows: 1_048_576,
      maxColumns: 16_384,
    });
    const sheetIdMap = new Map<string, number>();

    for (const worksheet of workbook.worksheets) {
      let sheetId: number;
      try {
        hf.addSheet(worksheet.name);
        const id = hf.getSheetId(worksheet.name);
        if (id === undefined) {
          console.warn(`[excel-preview] Skip sheet "${worksheet.name}" for formula engine: unknown id`);
          continue;
        }
        sheetId = id;
      } catch (error) {
        console.warn(`[excel-preview] Skip sheet "${worksheet.name}" for formula engine:`, error);
        continue;
      }
      sheetIdMap.set(worksheet.name, sheetId);

      // 按行批量写入, 降低逐单元格调用开销
      worksheet.eachRow((row, rowNumber) => {
        const cells: RawCellContent[] = [];
        let maxCol = -1;
        row.eachCell((cell, colNumber) => {
          const idx = colNumber - 1;
          cells[idx] = toRawCellContent(cell);
          if (idx > maxCol) maxCol = idx;
        });
        if (maxCol < 0) return;

        cells.length = maxCol + 1;
        for (let i = 0; i < cells.length; i++) {
          if (cells[i] === undefined) cells[i] = null;
        }

        try {
          hf.setCellContents({ sheet: sheetId, row: rowNumber - 1, col: 0 }, [cells]);
        } catch (error) {
          console.warn(`[excel-preview] Skip row ${rowNumber} of "${worksheet.name}" for formula engine:`, error);
        }
      });
    }

    return {
      getValue: (sheetName, row, column) => {
        const sheetId = sheetIdMap.get(sheetName);
        if (sheetId === undefined) return undefined;

        let value: any;
        try {
          value = hf.getCellValue({ sheet: sheetId, row: row - 1, col: column - 1 });
        } catch {
          return undefined;
        }

        if (value instanceof DetailedCellError) return value.value;

        if (typeof value === 'number') {
          // HyperFormula 日期返回序列号, 依据单元格数字格式还原为 Date
          const cell = workbook.getWorksheet(sheetName)?.getCell(row, column);
          if (cell && isDateFormat(cell.style?.numFmt)) {
            return excelSerialToDate(value);
          }
        }

        return value;
      },
      dispose: () => undefined,
    };
  } catch (error) {
    console.warn('[excel-preview] Formula calculation unavailable, using cached results:', error);
    return undefined;
  }
}

function getParsedFormulaValue(
  cell: ExcelJS.Cell,
  calculator: FormulaCalculator | undefined,
  sheetName: string,
  row: number,
  column: number
): any {
  if (!calculator || !getFormula(cell)) return normalizeCellValue(cell);
  const calculatedValue = calculator.getValue(sheetName, row, column);
  if (calculatedValue !== undefined) return calculatedValue;

  const cachedValue = normalizeCellValue(cell);
  if (isUsableFormulaCache(cachedValue, cell.value)) {
    return cachedValue;
  }

  return `=${getFormula(cell)}`;
}

function isUsableFormulaCache(value: any, originalValue: any): boolean {
  return value !== originalValue
    && value !== undefined
    && value !== null
    && (typeof value !== 'object' || value instanceof Date);
}

function getParsedFormulaText(
  cell: ExcelJS.Cell,
  calculator: FormulaCalculator | undefined,
  sheetName: string,
  row: number,
  column: number
): string {
  if (!calculator || !getFormula(cell)) return formatCellText(cell);

  const calculatedValue = calculator.getValue(sheetName, row, column);
  if (calculatedValue !== undefined) {
    return formatCellText({
      ...cell,
      value: calculatedValue,
      type: getValueType(calculatedValue),
    } as ExcelJS.Cell);
  }

  const cachedValue = normalizeCellValue(cell);
  if (isUsableFormulaCache(cachedValue, cell.value)) {
    return formatCellText({
      ...cell,
      value: cachedValue,
      type: getValueType(cachedValue),
    } as ExcelJS.Cell);
  }

  return `=${getFormula(cell)}`;
}

function getCellType(cell: ExcelJS.Cell): CellType {
  switch (cell.type) {
    case ExcelJS.ValueType.Number:
      return 'number';
    case ExcelJS.ValueType.String:
      return 'string';
    case ExcelJS.ValueType.Date:
      return 'date';
    case ExcelJS.ValueType.Boolean:
      return 'boolean';
    case ExcelJS.ValueType.Formula:
      return 'formula';
    case ExcelJS.ValueType.Hyperlink:
      return 'hyperlink';
    case ExcelJS.ValueType.RichText:
      return 'richText';
    case ExcelJS.ValueType.Null:
    default:
      return 'empty';
  }
}

/**
 * 格式化单元格显示文本
 */
function formatCellText(cell: ExcelJS.Cell): string {
  const value = cell.value;

  // 空值
  if (value === null || value === undefined) {
    return '';
  }

  // 公式 - exceljs 可能返回公式对象或计算结果
  if (cell.type === ExcelJS.ValueType.Formula) {
    // 如果 value 是对象，尝试提取 result
    if (typeof value === 'object' && value !== null) {
      const formulaResult = value as { result?: any; error?: string; formula?: string };

      // 检查 result 是否包含错误
      if (formulaResult.result && typeof formulaResult.result === 'object') {
        const resultObj = formulaResult.result as { error?: string };
        if (resultObj.error) {
          return resultObj.error; // 返回 #VALUE! 等错误
        }
      }

      if (formulaResult.error) {
        return formulaResult.error;
      }
      if (formulaResult.result !== undefined && formulaResult.result !== null) {
        // 递归格式化结果
        const resultCell = { ...cell, value: formulaResult.result, type: getValueType(formulaResult.result) } as ExcelJS.Cell;
        return formatCellText(resultCell);
      }
    }
    // 如果直接是字符串或数字
    if (typeof value === 'string' || typeof value === 'number') {
      return String(value);
    }
    return '';
  }

  // 数字格式化
  if (cell.type === ExcelJS.ValueType.Number && typeof value === 'number') {
    return formatNumberValue(value, cell);
  }

  // 日期格式化
  if (cell.type === ExcelJS.ValueType.Date && value instanceof Date) {
    return formatDateValue(value, cell);
  }

  // 富文本
  if (cell.type === ExcelJS.ValueType.RichText) {
    const richText = value as any;
    return richText.text || richText.richText?.map((r: any) => r.text).join('') || '';
  }

  // 超链接
  if (cell.type === ExcelJS.ValueType.Hyperlink) {
    const hyperlink = value as ExcelJS.CellHyperlinkValue;
    return hyperlink.text || hyperlink.hyperlink || '';
  }

  // 其他情况直接转字符串
  return String(value);
}

/**
 * 根据值推断单元格类型
 */
function getValueType(value: any): ExcelJS.ValueType {
  if (value === null || value === undefined) return ExcelJS.ValueType.Null;
  if (typeof value === 'number') return ExcelJS.ValueType.Number;
  if (typeof value === 'boolean') return ExcelJS.ValueType.Boolean;
  if (value instanceof Date) return ExcelJS.ValueType.Date;
  return ExcelJS.ValueType.String;
}

/**
 * 智能格式化数字，去除浮点数精度误差
 * 采用与 Excel 一致的 15 位有效数字显示精度：
 * 0.1 + 0.2 => 0.3; 37.057000000000002 => 37.057; 1/3 => 0.333333333333333
 */
function smartFormatNumber(value: number): string {
  if (!Number.isFinite(value)) return String(value);
  return String(parseFloat(value.toPrecision(15)));
}

/**
 * 解析 Excel 数字格式的「首个数值段」，提取百分比/货币/小数位/千分位等特征。
 * 兼容带颜色、条件、括号、转义字符的多段格式，如 "0.0_);[Red](0.0)"、"#,##0.00;[Red](#,##0.00)"。
 */
function parseNumFmt(numFmt: string): {
  percent: boolean;
  currency: string;
  grouped: boolean;
  precision: number;
  isNumeric: boolean;
} {
  // 取第一个 ';' 分段（正数部分）
  let section = numFmt.split(';')[0];
  // 去掉引号包裹的文本
  section = section.replace(/"[^"]*"/g, '');
  // 去掉转义字符（如 \_ \） ）  以及 [_...] 占位段
  section = section.replace(/\\[^\\]*/g, '').replace(/\[[^\]]*\]/g, '');

  const percent = section.includes('%');

  let currency = '';
  if (section.startsWith('"¥"')) currency = '¥';
  else if (section.startsWith('$')) currency = '$';
  else if (section.includes('¥')) currency = '¥';

  const grouped = section.includes('#,##');

  // 仅保留 0 # . 以判断是否为纯数字格式
  const digits = section.replace(/[^0#.]/g, '');

  const dot = digits.lastIndexOf('.');
  const precision = dot >= 0 ? (digits.slice(dot + 1).match(/0/g) || []).length : 0;

  return { percent, currency, grouped, precision, isNumeric: /^[0#.]+$/.test(digits) };
}

/** 整数部分添加千分位分隔符 */
function addThousandsSeparator(result: string): string {
  const parts = result.split('.');
  const intPart = parts[0].split('').reverse();
  const newIntPart: string[] = [];
  for (let i = 0; i < intPart.length; i++) {
    newIntPart.push(intPart[i]);
    if ((i + 1) % 3 === 0 && i < intPart.length - 1 && intPart[i + 1] !== '-') {
      newIntPart.push(',');
    }
  }
  let grouped = newIntPart.reverse().join('');
  if (parts.length > 1) grouped += '.' + parts[1];
  return grouped;
}

/**
 * 数字值格式化
 */
function formatNumberValue(value: number, cell: ExcelJS.Cell): string {
  const numFmt = cell.style?.numFmt;

  if (!numFmt) {
    return smartFormatNumber(value);
  }

  const fmt = parseNumFmt(numFmt);

  // 百分比格式
  if (fmt.percent) {
    return (value * 100).toFixed(fmt.precision) + '%';
  }

  // 货币格式
  if (fmt.currency) {
    if (value === 0 && numFmt.startsWith('_')) {
      return '-';
    }
    let result = value.toFixed(fmt.precision);
    if (fmt.grouped) result = addThousandsSeparator(result);
    return fmt.currency + result;
  }

  // 普通数字
  if (fmt.isNumeric) {
    let result = value.toFixed(fmt.precision);
    if (fmt.grouped) result = addThousandsSeparator(result);
    return result;
  }

  return smartFormatNumber(value);
}

/**
 * 日期值格式化
 */
function formatDateValue(value: Date, cell: ExcelJS.Cell): string {
  const numFmt = cell.style?.numFmt;
  const pad = (n: number) => String(n).padStart(2, '0');
  const year = value.getFullYear();
  const month = value.getMonth() + 1;
  const day = value.getDate();
  const hour = pad(value.getHours());
  const minute = pad(value.getMinutes());

  switch (numFmt) {
    case 'mm-dd-yy': return `${year}/${pad(month)}/${pad(day)}`;
    case '[$-F800]dddd, mmmm dd, yyyy': {
      const weekday = new Intl.DateTimeFormat('zh-CN', { weekday: 'short' }).format(value);
      return `${year}年${month}月${day}日 ${weekday}`;
    }
    case 'm"月"d"日";@': return `${month}月${day}日`;
    case 'yyyy/m/d h:mm;@': return `${year}/${month}/${pad(day)} ${hour}:${minute}`;
    case 'm/d/yy "h":mm': return `${year}/${pad(month)}/${pad(day)} ${hour}:${minute}`;
    case 'h:mm;@': return `${hour}:${minute}`;
    case 'yyyy-mm-dd;@':
    default: return `${year}-${pad(month)}-${pad(day)}`;
  }
}

/**
 * 解析单个工作表
 */
function parseSheet(
  worksheet: ExcelJS.Worksheet,
  sheetIndex: number,
  calculator: FormulaCalculator | undefined,
  themeColors: string[],
): ParsedSheet {
  const rows: ParsedCell[][] = [];
  const merges: MergeInfo[] = [];
  const colWidths: number[] = [];
  const rowHeights: number[] = [];

  // 收集合并单元格
  // model.merges 可能是字符串数组 (如 "A1:B1") 或对象数组
  const model = (worksheet as any).model;
  if (model && model.merges) {
    for (const merge of model.merges) {
      if (typeof merge === 'string') {
        // 解析 "A1:B1" 或 "F3:H3" 格式
        const match = merge.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
        if (match) {
          const startCol = colLetterToNumber(match[1]);
          const startRow = parseInt(match[2]) - 1; // 转为 0-based
          const endCol = colLetterToNumber(match[3]);
          const endRow = parseInt(match[4]) - 1;
          merges.push({ startRow, startCol, endRow, endCol });
        }
      } else {
        // 对象格式
        const m = merge as any;
        const startRow = (m.start?.row ?? m.s?.r ?? 1) - 1;
        const startCol = (m.start?.column ?? m.s?.c ?? 1) - 1;
        const endRow = (m.end?.row ?? m.e?.r ?? 1) - 1;
        const endCol = (m.end?.column ?? m.e?.c ?? 1) - 1;
        if (startRow >= 0 && startCol >= 0 && endRow >= startRow && endCol >= startCol) {
          merges.push({ startRow, startCol, endRow, endCol });
        }
      }
    }
  }

  // 收集列宽
  if (worksheet.columns) {
    for (let i = 0; i < worksheet.columns.length; i++) {
      const col = worksheet.columns[i];
      colWidths[i] = col.width ? Math.round(col.width * COL_WIDTH_TO_PX + COL_WIDTH_MARGIN) : DEFAULT_COL_WIDTH;
    }
  }

  // 收集隐藏行/列
  const hiddenRows = new Set<number>();
  const hiddenCols = new Set<number>();

  if (worksheet.columns) {
    for (let i = 0; i < worksheet.columns.length; i++) {
      if (worksheet.columns[i]?.hidden) {
        hiddenCols.add(i);
      }
    }
  }

  // 遍历每一行
  worksheet.eachRow((row, rowNumber) => {
    // 记录隐藏行
    if (row.hidden) {
      hiddenRows.add(rowNumber - 1);
    }
    // Excel 的行高数值若按物理 pt→px 转换会使 25 高度变成 33px，
    // 在浏览器预览中明显比 Excel 网格更松。保留该数值以匹配屏幕视觉高度。
    rowHeights[rowNumber - 1] = row.height ? Math.round(row.height) : DEFAULT_ROW_HEIGHT;

    const cells: ParsedCell[] = [];

    // 获取该行的实际列数
    const colCount = Math.max(row.cellCount, worksheet.columnCount || 10);

    for (let colNumber = 1; colNumber <= colCount; colNumber++) {
      const cell = row.getCell(colNumber);

      // 检查是否是合并区域的起始单元格
      let mergeInfo: MergeInfo | undefined;
      for (const merge of merges) {
        if (
          merge.startRow === rowNumber - 1 &&
          merge.startCol === colNumber - 1
        ) {
          mergeInfo = merge;
          break;
        }
      }

      // 检查是否在合并区域内但不是起始单元格（跳过）
      const isInMergeButNotStart = merges.some(
        m =>
          rowNumber - 1 >= m.startRow &&
          rowNumber - 1 <= m.endRow &&
          colNumber - 1 >= m.startCol &&
          colNumber - 1 <= m.endCol &&
          !(m.startRow === rowNumber - 1 && m.startCol === colNumber - 1)
      );

      if (isInMergeButNotStart) {
        // 合并区域内的非起始单元格用空对象占位
        cells.push({
          value: null,
          text: '',
          type: 'empty',
          style: {},
        });
        continue;
      }

      const parsedCell: ParsedCell = {
        value: getParsedFormulaValue(cell, calculator, worksheet.name, rowNumber, colNumber),
        text: getParsedFormulaText(cell, calculator, worksheet.name, rowNumber, colNumber),
        type: getCellType(cell),
        style: parseCellStyle(cell, themeColors),
      };

      // 提取富文本片段
      if (cell.type === ExcelJS.ValueType.RichText && cell.value) {
        const richValue = cell.value as any;
        const runs = richValue.richText;
        if (Array.isArray(runs) && runs.length > 0) {
          parsedCell.richTextRuns = runs.map((r: any) => {
            const run: RichTextRun = { text: r.text ?? '' };
            if (r.font) {
              run.font = {};
              if (r.font.name) run.font.name = r.font.name;
              if (r.font.size !== undefined) run.font.size = r.font.size;
              if (r.font.bold) run.font.bold = true;
              if (r.font.italic) run.font.italic = true;
              if (r.font.underline) run.font.underline = !!r.font.underline;
              if (r.font.strike) run.font.strike = true;
              if (r.font.color) {
                const fc = r.font.color;
                run.font.color = resolveColor(fc, themeColors);
              }
            }
            return run;
          });
        }
      }

      // 提取批注
      const commentData = (cell as any).comment ?? (cell as any).note;
      if (commentData) {
        const texts = commentData.texts ?? commentData.text ?? [];
        const text = Array.isArray(texts)
          ? texts.map((t: any) => (typeof t === 'object' ? t.text ?? '' : t)).join('')
          : String(texts);
        if (text) {
          parsedCell.comment = {
            text,
            author: commentData.authors?.[0]?.name ?? undefined,
          };
        }
      }

      if (mergeInfo) {
        parsedCell.merge = mergeInfo;
      }

      cells.push(parsedCell);
    }

    rows[rowNumber - 1] = cells;
  });

  // 解析冻结窗格
  let freezePane: FreezePane | undefined;
  try {
    const wsViews = (worksheet as any).views;
    if (wsViews && wsViews.length > 0) {
      const view = wsViews[0];
      if (view.state === 'frozen' || view.state === 'frozenSplit') {
        const xSplit = parseInt(view.xSplit, 10) || 0;
        const ySplit = parseInt(view.ySplit, 10) || 0;
        if (xSplit > 0 || ySplit > 0) {
          freezePane = { xSplit, ySplit };
        }
      }
    }
  } catch {
    // ignore freeze pane parse errors
  }

  // 解析条件格式
  let conditionalFormatting: ConditionalFormatting[] | undefined;
  try {
    const cfData = (worksheet as any).conditionalFormatting;
    if (cfData && cfData.length > 0) {
      conditionalFormatting = [];
      for (const cf of cfData) {
        const rules: CfRule[] = [];
        for (const rule of cf.rules || []) {
          const parsedRule: CfRule = {
            type: rule.type || 'expression',
            priority: rule.priority ?? 0,
          };
          if (rule.formula) {
            parsedRule.formula = Array.isArray(rule.formula) ? rule.formula : [rule.formula];
          }
          if (rule.operator) parsedRule.operator = rule.operator;
          if (rule.text) parsedRule.text = rule.text;

          // 提取样式
          if (rule.fill) {
            const fillColor = rule.fill.fgColor ?? rule.fill.bgColor;
            if (fillColor) {
              parsedRule.fill = resolveColor(fillColor, themeColors, '#C7C9CC');
            }
          }
          if (rule.font) {
            parsedRule.font = {};
            if (rule.font.color) {
              const fc = rule.font.color;
              parsedRule.font.color = resolveColor(fc, themeColors);
            }
            if (rule.font.bold !== undefined) parsedRule.font.bold = rule.font.bold;
            if (rule.font.italic !== undefined) parsedRule.font.italic = rule.font.italic;
            if (rule.font.strike !== undefined) parsedRule.font.strike = rule.font.strike;
            if (rule.font.underline !== undefined) parsedRule.font.underline = !!rule.font.underline;
          }

          // 色阶
          if (rule.type === 'colorScale' && rule.colorScale) {
            const cs = rule.colorScale;
            parsedRule.colorScale = {
              type: 'colorScale',
              cfvo: (cs.cfvo || []).map((v: any) => ({
                type: v.type as any,
                value: v.value !== undefined ? v.value : undefined,
              })),
              colors: (cs.colors || []).map((c: any) => resolveColor(c, themeColors, '#C7C9CC')),
            };
          }

          // 数据条
          if (rule.type === 'dataBar' && rule.dataBar) {
            const db = rule.dataBar;
            parsedRule.dataBar = {
              type: 'dataBar',
              cfvo: (db.cfvo || []).map((v: any) => ({
                type: v.type as any,
                value: v.value !== undefined ? v.value : undefined,
              })),
              color: resolveColor(db.color, themeColors, '#C7C9CC'),
              showValue: db.showValue !== false,
            };
          }

          // 图标集
          if (rule.type === 'iconSet' && rule.iconSet) {
            const is = rule.iconSet;
            parsedRule.iconSet = {
              type: 'iconSet',
              iconSet: is.iconSet,
              cfvo: (is.cfvo || []).map((v: any) => ({
                type: v.type as any,
                value: v.value !== undefined ? v.value : undefined,
              })),
              showValue: is.showValue !== false,
            };
          }

          rules.push(parsedRule);
        }
        conditionalFormatting.push({ ref: cf.ref, rules });
      }
    }
  } catch {
    // ignore conditional formatting parse errors
  }

  return {
    name: worksheet.name || `Sheet${sheetIndex + 1}`,
    id: `sheet_${sheetIndex}`,
    rows,
    merges,
    colWidths,
    rowHeights,
    freezePane,
    conditionalFormatting,
    hiddenRows: hiddenRows.size > 0 ? hiddenRows : undefined,
    hiddenCols: hiddenCols.size > 0 ? hiddenCols : undefined,
  };
}

// ===== 主入口函数 =====

/**
 * 解析 Excel 文件 (ArrayBuffer)
 *
 * @param buffer - Excel 文件的 ArrayBuffer 数据
 * @param options - 解析选项
 * @returns Promise<ParsedWorkbook> 解析后的工作簿数据
 */
export async function parseExcel(
  buffer: ArrayBuffer
): Promise<ParsedWorkbook> {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    // ExcelJS preserves theme references (e.g. { theme: 1 }) on cell styles.
    // Resolve them from xl/theme/theme1.xml before producing browser CSS.
    const themeColors = themeColorsFromTheme(await parseThemeFromZip(buffer));

    if (!workbook.worksheets || workbook.worksheets.length === 0) {
      throw new Error('未获取到任何工作表，可能文件格式不正确或文件已损坏');
    }

    const sheets: ParsedSheet[] = [];
    const calculator = createFormulaCalculator(workbook);

    try {
      workbook.worksheets.forEach((worksheet, index) => {
        // 跳过隐藏的工作表
        if (worksheet.state === 'hidden' || worksheet.state === 'veryHidden') {
          return;
        }

        sheets.push(parseSheet(worksheet, index, calculator, themeColors));
      });
    } finally {
      calculator?.dispose();
    }

    if (sheets.length === 0) {
      throw new Error('没有可见的工作表');
    }

    return { sheets };
  } catch (error) {
    console.error('[excel-preview] Failed to parse Excel file:', error);
    throw error;
  }
}

/**
 * 获取原始 Workbook 对象（用于高级操作）
 */
export async function loadRawWorkbook(buffer: ArrayBuffer): Promise<ExcelJS.Workbook> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  return workbook;
}
