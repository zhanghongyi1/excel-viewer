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
import dayjs from 'dayjs';
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

/** 行高点转像素系数 (1pt = 1.333px @ 96dpi) */
const PT_TO_PX = 1.333;

/** 索引颜色表 (Excel 内置) */
const INDEXED_COLORS: Record<number, string> = {
  0: '#000000', 1: '#FFFFFF', 2: '#FF0000', 3: '#00FF00', 4: '#0000FF',
  5: '#FFFF00', 6: '#FF00FF', 7: '#00FFFF', 8: '#000000', 9: '#FFFFFF',
  10: '#FF0000', 11: '#00FF00', 12: '#0000FF', 13: '#FFFF00', 14: '#FF00FF',
  15: '#00FFFF', 16: '#800000', 17: '#008000', 18: '#000080', 19: '#808000',
  20: '#800080', 21: '#008080', 22: '#C0C0C0', 23: '#808080',
};

/** 主题颜色表 (简化版) */
const THEME_COLORS: string[] = [
  '#FFFFFF', '#F2F2F2', '#E7E6E6', '#D9D9D9', '#BFBCBB',
  '#A6A5A5', '#848484', '#474744', '#44546A', '#ED7D31',
  '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#7030A0',
  '#C00000', '#BF8F00', '#305497', '#375623', '#264478',
];

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
function getThemeColor(themeIndex: number, tint?: number): string {
  if (themeIndex < 0 || themeIndex >= THEME_COLORS.length) {
    return '#C7C9CC';
  }

  let baseColor = THEME_COLORS[themeIndex];

  if (tint !== undefined && tint !== null) {
    baseColor = applyTint(baseColor, tint);
  }

  return baseColor;
}

/**
 * 应用色调调整
 */
function applyTint(hexColor: string, tint: number): string {
  // 移除 # 号
  let hex = hexColor.replace('#', '');

  let r = parseInt(hex.substring(0, 2), 16);
  let g = parseInt(hex.substring(2, 4), 16);
  let b = parseInt(hex.substring(4, 6), 16);

  if (tint > 0) {
    // 向白色混合
    r = Math.round(r + (255 - r) * tint);
    g = Math.round(g + (255 - g) * tint);
    b = Math.round(b + (255 - b) * tint);
  } else {
    // 向黑色混合
    const absTint = Math.abs(tint);
    r = Math.round(r * (1 - absTint));
    g = Math.round(g * (1 - absTint));
    b = Math.round(b * (1 - absTint));
  }

  return `#${r.toString(16).padStart(2, '0')}${g.toString(16).padStart(2, '0')}${b.toString(16).padStart(2, '0')}`;
}

// ===== 核心解析函数 =====

/**
 * 解析单元格样式
 */
function parseCellStyle(cell: ExcelJS.Cell): CellStyle {
  const style: CellStyle = {};
  const cellStyle: any = (cell as any).style || {};

    // 字体样式
  if (cellStyle.font) {
    style.font = {};
    if (cellStyle.font.name) style.font.name = cellStyle.font.name;
    if (cellStyle.font.size !== undefined) {
      // Excel 字号是 points，转换为 px (1pt ≈ 1.333px，但浏览器渲染通常需要调整)
      // Excel 8pt 字体在浏览器中应该显示为约 11px
      style.font.size = Math.round(cellStyle.font.size * 1.333);
    }
    if (cellStyle.font.bold !== undefined) style.font.bold = cellStyle.font.bold;
    if (cellStyle.font.italic !== undefined) style.font.italic = cellStyle.font.italic;
    if (cellStyle.font.underline !== undefined) style.font.underline = !!cellStyle.font.underline;
    if (cellStyle.font.strike !== undefined) style.font.strike = cellStyle.font.strike;

    // 字体颜色
    if (cellStyle.font.color) {
      let fontColor = '#000000';
      if (typeof cellStyle.font.color === 'string') {
        fontColor = argbToHex(cellStyle.font.color as string);
      } else if (cellStyle.font.color.argb) {
        fontColor = argbToHex(cellStyle.font.color.argb);
      } else if (cellStyle.font.color.theme !== undefined) {
        fontColor = getThemeColor(cellStyle.font.color.theme, cellStyle.font.color.tint);
      } else if (cellStyle.font.color.indexed !== undefined) {
        fontColor = INDEXED_COLORS[cellStyle.font.color.indexed] || '#000000';
      }
      style.font.color = fontColor;
    }
  }

    // 背景色 (填充)
  if ((cellStyle as any).fill && (cellStyle as any).fill?.fgColor) {
    const fgColor: any = (cellStyle as any).fill.fgColor;
    if (fgColor.argb) {
      style.bgcolor = argbToHex(fgColor.argb);
    } else if (fgColor.theme !== undefined) {
      style.bgcolor = getThemeColor(fgColor.theme, fgColor.tint);
    } else if (fgColor.indexed !== undefined) {
      style.bgcolor = INDEXED_COLORS[fgColor.indexed] || '#C7C9CC';
    }
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
          if (typeof border.color === 'string') {
            borderColor = border.color;
          } else if (border.color.argb) {
            borderColor = argbToHex(border.color.argb);
          } else if (border.color.theme !== undefined) {
            borderColor = getThemeColor(border.color.theme, border.color.tint);
          } else if (border.color.indexed !== undefined) {
            borderColor = INDEXED_COLORS[border.color.indexed] || '#000000';
          }
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
 * 例如：37.057000000000002 -> 37.057
 */
function smartFormatNumber(value: number): string {
  // 处理科学计数法
  if (Math.abs(value) < 0.0001 || Math.abs(value) > 1e9) {
    return String(value);
  }

  // 转换为字符串并去除末尾的 0
  let str = value.toFixed(10);

  // 去除末尾的 0
  while (str.includes('.') && (str.endsWith('0') || str.endsWith('.'))) {
    if (str.endsWith('.')) {
      str = str.slice(0, -1);
      break;
    }
    str = str.slice(0, -1);
  }

  // 如果结果是整数，直接返回
  if (!str.includes('.')) {
    return str;
  }

  // 进一步清理：检测并修复类似 37.0569999999 -> 37.057 的情况
  const match = str.match(/^(-?\d+)\.(\d+)$/);
  if (match) {
    const intPart = match[1];
    let decimalPart = match[2];

    // 如果小数部分很长且末尾有多个 9 或 0，尝试四舍五入
    if (decimalPart.length >= 8) {
      // 检查末尾是否是 999... 或 000...
      const lastChar = decimalPart[decimalPart.length - 1];
      let trailingCount = 0;
      for (let i = decimalPart.length - 1; i >= 0; i--) {
        if (decimalPart[i] === lastChar && (lastChar === '9' || lastChar === '0')) {
          trailingCount++;
        } else {
          break;
        }
      }

      // 如果末尾有超过 5 个 9 或 0，进行四舍五入
      if (trailingCount >= 5) {
        const precision = Math.max(1, decimalPart.length - trailingCount);
        return String(parseFloat(value.toFixed(precision)));
      }
    }
  }

  return str;
}

/**
 * 数字值格式化
 */
function formatNumberValue(value: number, cell: ExcelJS.Cell): string {
  const numFmt = cell.style?.numFmt;

  if (!numFmt) {
    return smartFormatNumber(value);
  }

  // 百分比格式
  if (numFmt.includes('%')) {
    const precisionMatch = numFmt.match(/\.(\d+)%/);
    if (precisionMatch) {
      return (value * 100).toFixed(precisionMatch[1].length) + '%';
    }
    return (value * 100).toFixed(0) + '%';
  }

  // 货币格式
  if (numFmt.startsWith('$') || numFmt.startsWith('"¥"') || numFmt.includes('¥')) {
    const prefix = numFmt.startsWith('"¥"') ? '¥' : (numFmt.startsWith('$') ? '$' : '');
    const precisionMatch = numFmt.match(/0\.(0+)/);
    const precision = precisionMatch ? precisionMatch[1].length : 0;

    if (value === 0 && numFmt.startsWith('_')) {
      return '-';
    }

    let result = value.toFixed(precision);

    // 千分位分隔符
    if (numFmt.includes('#,##')) {
      const parts = result.split('.');
      const intPart = parts[0].split('').reverse();
      const newIntPart: string[] = [];
      for (let i = 0; i < intPart.length; i++) {
        newIntPart.push(intPart[i]);
        if ((i + 1) % 3 === 0 && i < intPart.length - 1 && intPart[i + 1] !== '-') {
          newIntPart.push(',');
        }
      }
      result = newIntPart.reverse().join('');
      if (parts.length > 1) result += '.' + parts[1];
    }

    return prefix + result;
  }

  // 普通数字
  if (/^0+(\.0+)?$/.test(numFmt.replace(/[^0.]/g, ''))) {
    const precisionMatch = numFmt.match(/0\.(0+)/);
    const precision = precisionMatch ? precisionMatch[1].length : 0;
    return value.toFixed(precision);
  }

  return String(value);
}

/**
 * 日期值格式化
 */
function formatDateValue(value: Date, cell: ExcelJS.Cell): string {
  const numFmt = cell.style?.numFmt;

  // 常见日期格式映射
  const dateFormats: Record<string, string> = {
    'yyyy-mm-dd;@': 'YYYY-MM-DD',
    'mm-dd-yy': 'YYYY/MM/DD',
    '[$-F800]dddd, mmmm dd, yyyy': 'YYYY年M月D日 ddd',
    'm"月"d"日";@': 'M月D日',
    'yyyy/m/d h:mm;@': 'YYYY/M/DD HH:mm',
    'm/d/yy "h":mm': 'YYYY/MM/DD HH:mm',
    'h:mm;@': 'HH:mm',
  };

  if (numFmt && dateFormats[numFmt]) {
    return dayjs(value).format(dateFormats[numFmt]);
  }

  // 默认日期格式
  return dayjs(value).format('YYYY-MM-DD');
}

/**
 * 列字母转数字索引 (A=0, B=1, ..., Z=25, AA=26)
 */
function colLetterToNumber(letters: string): number {
  let result = 0;
  for (let i = 0; i < letters.length; i++) {
    result = result * 26 + (letters.charCodeAt(i) - 64);
  }
  return result - 1; // 转为 0-based
}

/**
 * 解析单个工作表
 */
function parseSheet(worksheet: ExcelJS.Worksheet, sheetIndex: number): ParsedSheet {
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
    // 记录行高 (exceljs 返回的行高单位为 pt，需转为 px)
    rowHeights[rowNumber - 1] = row.height ? Math.round(row.height * PT_TO_PX) : DEFAULT_ROW_HEIGHT;

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
        value: normalizeCellValue(cell),
        text: formatCellText(cell),
        type: getCellType(cell),
        style: parseCellStyle(cell),
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
              if (r.font.size !== undefined) run.font.size = Math.round(r.font.size * 1.333);
              if (r.font.bold) run.font.bold = true;
              if (r.font.italic) run.font.italic = true;
              if (r.font.underline) run.font.underline = !!r.font.underline;
              if (r.font.strike) run.font.strike = true;
              if (r.font.color) {
                const fc = r.font.color;
                run.font.color = fc.argb ? argbToHex(fc.argb) : (fc.theme !== undefined ? getThemeColor(fc.theme, fc.tint) : '#000000');
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
              parsedRule.fill = argbToHex(fillColor.argb || fillColor);
            }
          }
          if (rule.font) {
            parsedRule.font = {};
            if (rule.font.color) {
              const fc = rule.font.color;
              parsedRule.font.color = argbToHex(fc.argb || fc.theme || fc);
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
              colors: (cs.colors || []).map((c: any) => argbToHex(c.argb || c)),
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
              color: argbToHex(db.color?.argb || db.color),
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
    charts: [], // 图表由 chart-parser 单独填充
  };
}

// ===== 主入口函数 =====

/**
 * 解析 Excel 文件 (ArrayBuffer)
 *
 * @param buffer - Excel 文件的 ArrayBuffer 数据
 * @returns Promise<ParsedWorkbook> 解析后的工作簿数据
 */
export async function parseExcel(buffer: ArrayBuffer): Promise<ParsedWorkbook> {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);

    if (!workbook.worksheets || workbook.worksheets.length === 0) {
      throw new Error('未获取到任何工作表，可能文件格式不正确或文件已损坏');
    }

    const sheets: ParsedSheet[] = [];

    workbook.worksheets.forEach((worksheet, index) => {
      // 跳过隐藏的工作表
      if (worksheet.state === 'hidden' || worksheet.state === 'veryHidden') {
        return;
      }

      sheets.push(parseSheet(worksheet, index));
    });

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
