/**
 * OOXML Chart Parser — OpenXML 图表解析器
 *
 * 将 xl/charts/chart*.xml 解析为 ChartModel 领域模型。
 *
 * 支持的 OOXML 图表元素:
 *   c:barChart / c:bar3DChart        → bar (3D 标志)
 *   c:lineChart / c:line3DChart      → line
 *   c:areaChart / c:area3DChart      → area
 *   c:pieChart / c:pie3DChart        → pie
 *   c:doughnutChart                  → doughnut
 *   c:scatterChart                   → scatter
 *   c:bubbleChart                    → bubble
 *   c:radarChart                     → radar
 *   c:stockChart                     → stock
 *   c:surfaceChart / c:surface3DChart→ surface
 *   多个上述元素共存                  → combo（组合图）
 *
 * 解析流程:
 *   1. 提取 plotArea，扫描所有图表类型元素
 *   2. 若仅一个类型 → 单一图表；若多个 → combo 组合图
 *   3. 对每个图表类型元素解析 series / categories / 样式
 *   4. 解析标题、图例、坐标轴
 *   5. 组装 ChartModel
 */

import ExcelJS from 'exceljs';
import type { ChartAnchor } from '../types';
import type {
  ChartModel,
  ChartType,
  ChartSeriesModel,
  ChartAxisModel,
  ChartTitleModel,
  ChartLegendModel,
  PlotAreaModel,
  PlotGroup,
  BarDirection,
  Grouping,
  MarkerSymbol,
  ChartDataPoint,
  StockOhlcData,
  StockChartType,
  RadarStyle,
} from './chart-model';
import {
  createDefaultAxis,
  createDefaultLegend,
} from './chart-model';
import { DEFAULT_THEME, themeColorsToMap } from './theme-parser';
import type { ChartTheme } from './theme-parser';

// ===== OOXML 图表类型元素映射 =====

interface ChartElementInfo {
  ooxmlKey: string; // XML 标签名 (如 'c:barChart')
  chartType: ChartType;
  is3D: boolean;
}

const CHART_ELEMENT_MAP: ChartElementInfo[] = [
  { ooxmlKey: 'c:barChart', chartType: 'bar', is3D: false },
  { ooxmlKey: 'c:bar3DChart', chartType: 'bar', is3D: true },
  { ooxmlKey: 'c:lineChart', chartType: 'line', is3D: false },
  { ooxmlKey: 'c:line3DChart', chartType: 'line', is3D: true },
  { ooxmlKey: 'c:areaChart', chartType: 'area', is3D: false },
  { ooxmlKey: 'c:area3DChart', chartType: 'area', is3D: true },
  { ooxmlKey: 'c:pieChart', chartType: 'pie', is3D: false },
  { ooxmlKey: 'c:pie3DChart', chartType: 'pie', is3D: true },
  { ooxmlKey: 'c:doughnutChart', chartType: 'doughnut', is3D: false },
  { ooxmlKey: 'c:scatterChart', chartType: 'scatter', is3D: false },
  { ooxmlKey: 'c:bubbleChart', chartType: 'bubble', is3D: false },
  { ooxmlKey: 'c:radarChart', chartType: 'radar', is3D: false },
  { ooxmlKey: 'c:stockChart', chartType: 'stock', is3D: false },
  { ooxmlKey: 'c:surfaceChart', chartType: 'surface', is3D: false },
  { ooxmlKey: 'c:surface3DChart', chartType: 'surface', is3D: true },
];

/** 主题颜色映射 (默认 Office 主题，可被 parseChartXmlToModel 的 theme 参数覆盖) */
let THEME_COLOR_MAP: Record<string, string> = themeColorsToMap(DEFAULT_THEME);

/** 设置主题颜色映射（由 parseCharts 在解析前调用） */
export function setThemeColorMap(theme: ChartTheme): void {
  THEME_COLOR_MAP = themeColorsToMap(theme);
}

/** 默认系列颜色序列 */
const DEFAULT_SERIES_COLORS = [
  '#5B9BD5', '#ED7D31', '#A5A5A5', '#FFC000',
  '#4472C4', '#70AD47', '#264478', '#9B59B6',
];

// ===== XML 工具函数 =====

/** 安全获取数值 */
function getNum(val: any): number {
  if (val === undefined || val === null) return 0;
  if (typeof val === 'number') return val;
  if (typeof val === 'string') return parseFloat(val) || 0;
  if (typeof val === 'object') {
    if (val['#text'] !== undefined) return parseFloat(val['#text']) || 0;
  }
  return 0;
}

/** 获取数组形式节点（兼容单值） */
function toArray(val: any): any[] {
  if (val === undefined || val === null) return [];
  return Array.isArray(val) ? val : [val];
}

/** 提取文本节点值 */
function extractText(node: any): string {
  if (node === undefined || node === null) return '';
  if (typeof node === 'string' || typeof node === 'number') return String(node);
  if (Array.isArray(node)) return node.map(extractText).join('');
  if (node['#text']) return String(node['#text']);
  if (node['c:v']) return Array.isArray(node['c:v']) ? String(node['c:v'][0]) : String(node['c:v']);
  return '';
}

/** 获取 plotArea 节点 */
function getPlotArea(chartXmlObj: any): any {
  return chartXmlObj?.['c:chartSpace']?.['c:chart']?.['c:plotArea'];
}

/** 列字母转数字索引 (A=0) */
function colLetterToNumber(letters: string): number {
  let result = 0;
  for (let i = 0; i < letters.length; i++) {
    result = result * 26 + (letters.charCodeAt(i) - 64);
  }
  return result - 1;
}

/** 解析单元格范围引用 */
function parseCellRangeRef(refStr: string): {
  sheetName?: string;
  startCol: number;
  startRow: number;
  endCol: number;
  endRow: number;
} | null {
  const cleaned = refStr.replace(/\$/g, '');
  let rangePart = cleaned;
  let sheetName: string | undefined;
  const exclIndex = cleaned.lastIndexOf('!');
  if (exclIndex > -1) {
    sheetName = cleaned.substring(0, exclIndex);
    if (sheetName.startsWith("'") && sheetName.endsWith("'")) {
      sheetName = sheetName.substring(1, sheetName.length - 1);
    }
    rangePart = cleaned.substring(exclIndex + 1);
  }
  const match = rangePart.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (!match) {
    const singleMatch = rangePart.match(/^([A-Z]+)(\d+)$/);
    if (singleMatch) {
      const col = colLetterToNumber(singleMatch[1]);
      const row = parseInt(singleMatch[2]) - 1;
      return { sheetName, startCol: col, startRow: row, endCol: col, endRow: row };
    }
    return null;
  }
  return {
    sheetName,
    startCol: colLetterToNumber(match[1]),
    startRow: parseInt(match[2]) - 1,
    endCol: colLetterToNumber(match[3]),
    endRow: parseInt(match[4]) - 1,
  };
}

/** 解析 numCache / numLit 中的数值数组 */
function parseNumPtValues(container: any): number[] {
  if (!container) return [];
  const pts = toArray(container['c:pt']);
  if (!pts.length) return [];
  return pts
    .sort((a: any, b: any) => getNum(a['@_idx']) - getNum(b['@_idx']))
    .map((pt: any) => {
      const v = pt['c:v'];
      if (typeof v === 'number') return v;
      return parseFloat(typeof v === 'string' ? v : v?.['#text'] || '0') || 0;
    });
}

/** 解析 strCache / strLit 中的字符串数组 */
function parseStrPtValues(container: any): string[] {
  if (!container) return [];
  const pts = toArray(container['c:pt']);
  if (!pts.length) return [];
  return pts
    .sort((a: any, b: any) => getNum(a['@_idx']) - getNum(b['@_idx']))
    .map((pt: any) => {
      const v = pt['c:v'];
      return typeof v === 'string' ? v : v?.['#text'] || '';
    });
}

/** 从 workbook 读取单元格范围值 */
function readCellValues(
  workbook: ExcelJS.Workbook,
  sheetName: string | undefined,
  startCol: number,
  startRow: number,
  endCol: number,
  endRow: number
): (string | number | null)[][] {
  let worksheet: ExcelJS.Worksheet | undefined;
  if (sheetName) worksheet = workbook.getWorksheet(sheetName);
  if (!worksheet && workbook.worksheets.length > 0) worksheet = workbook.worksheets[0];
  if (!worksheet) return [];

  const values: (string | number | null)[][] = [];
  for (let r = startRow + 1; r <= endRow + 1; r++) {
    const rowValues: (string | number | null)[] = [];
    for (let c = startCol + 1; c <= endCol + 1; c++) {
      try {
        const cell = worksheet.getCell(r, c);
        if (cell.value === null || cell.value === undefined) {
          rowValues.push(null);
        } else if (typeof cell.value === 'number') {
          rowValues.push(cell.value);
        } else if (cell.value instanceof Date) {
          rowValues.push(cell.value.toISOString());
        } else if (typeof cell.value === 'object' && cell.value !== null) {
          const result = (cell.value as any).result;
          if (result && typeof result === 'object' && result.error) {
            rowValues.push(null);
          } else {
            rowValues.push(result !== undefined ? String(result) : null);
          }
        } else {
          rowValues.push(String(cell.value));
        }
      } catch {
        rowValues.push(null);
      }
    }
    values.push(rowValues);
  }
  return values;
}

/** 从引用中读取数值数组（优先用 workbook 实时数据，回退到 numCache） */
function readNumRef(valNode: any, workbook: ExcelJS.Workbook): number[] {
  if (!valNode) return [];
  const numRef = valNode['c:numRef']?.[0] || valNode['c:numRef'];
  if (numRef) {
    const f = toArray(numRef['c:f'])[0];
    const refText = typeof f === 'string' ? f : extractText(f);
    const parsed = parseCellRangeRef(refText);
    if (parsed) {
      const values = readCellValues(workbook, parsed.sheetName, parsed.startCol, parsed.startRow, parsed.endCol, parsed.endRow);
      const flat = values.flat().map(v => (typeof v === 'number' ? v : parseFloat(String(v)) || 0));
      if (flat.length > 0) return flat;
    }
    // 回退: numCache
    const numCache = toArray(numRef['c:numCache'])[0];
    if (numCache) return parseNumPtValues(numCache);
  }
  // 直接数值 (c:numLit)
  const numLit = valNode['c:numLit']?.[0] || valNode['c:numLit'];
  if (numLit) return parseNumPtValues(numLit);
  return [];
}

/** 从引用中读取字符串数组 */
function readStrRef(valNode: any, workbook: ExcelJS.Workbook): string[] {
  if (!valNode) return [];
  const strRef = valNode['c:strRef']?.[0] || valNode['c:strRef'];
  if (strRef) {
    const f = toArray(strRef['c:f'])[0];
    const refText = typeof f === 'string' ? f : extractText(f);
    const parsed = parseCellRangeRef(refText);
    if (parsed) {
      const values = readCellValues(workbook, parsed.sheetName, parsed.startCol, parsed.startRow, parsed.endCol, parsed.endRow);
      const flat = values.flat().map(v => String(v || ''));
      if (flat.length > 0) return flat;
    }
    const strCache = toArray(strRef['c:strCache'])[0];
    if (strCache) return parseStrPtValues(strCache);
  }
  const strLit = valNode['c:strLit']?.[0] || valNode['c:strLit'];
  if (strLit) return parseStrPtValues(strLit);
  return [];
}

// ===== 颜色解析 =====

/** 从 spPr 节点提取颜色 */
function parseColorFromSpPr(spPrRaw: any): string | undefined {
  const spPr = Array.isArray(spPrRaw) ? spPrRaw[0] : spPrRaw;
  if (!spPr) return undefined;

  // solidFill → 填充色（柱状/面积图）
  const solidFill = toArray(spPr['a:solidFill'])[0];
  if (solidFill) {
    const color = extractColorFromNode(solidFill);
    if (color) return color;
  }

  // ln → 线条颜色（折线图）
  const ln = toArray(spPr['a:ln'])[0];
  if (ln) {
    const lnSolidFill = toArray(ln['a:solidFill'])[0];
    if (lnSolidFill) {
      const color = extractColorFromNode(lnSolidFill);
      if (color) return color;
    }
  }

  return undefined;
}

/** 从颜色节点提取 hex 颜色 */
function extractColorFromNode(node: any): string | undefined {
  const schemeClr = toArray(node['a:schemeClr'])[0];
  if (schemeClr) {
    const val = schemeClr['@_val'];
    if (val && THEME_COLOR_MAP[val]) {
      const base = THEME_COLOR_MAP[val];
      const tint = schemeClr['@_tint'];
      return tint !== undefined ? applyTint(base, parseFloat(tint)) : base;
    }
  }
  const srgbClr = toArray(node['a:srgbClr'])[0];
  if (srgbClr) {
    const val = srgbClr['@_val'];
    if (val) {
      const tint = srgbClr['@_tint'];
      return tint !== undefined ? applyTint(`#${val}`, parseFloat(tint)) : `#${val}`;
    }
  }
  return undefined;
}

/** 应用色调调整 */
function applyTint(hexColor: string, tint: number): string {
  let hex = hexColor.replace('#', '');
  let r = parseInt(hex.substring(0, 2), 16);
  let g = parseInt(hex.substring(2, 4), 16);
  let b = parseInt(hex.substring(4, 6), 16);
  if (tint > 0) {
    r = Math.round(r + (255 - r) * tint);
    g = Math.round(g + (255 - g) * tint);
    b = Math.round(b + (255 - b) * tint);
  } else {
    const abs = Math.abs(tint);
    r = Math.round(r * (1 - abs));
    g = Math.round(g * (1 - abs));
    b = Math.round(b * (1 - abs));
  }
  return `#${r.toString(16).padStart(2, '0')}${g.toString(16).padStart(2, '0')}${b.toString(16).padStart(2, '0')}`;
}

// ===== 标题解析 =====

/** 提取图表标题 */
function extractTitle(chartXmlObj: any, workbook: ExcelJS.Workbook): ChartTitleModel | undefined {
  const titleNode = chartXmlObj?.['c:chartSpace']?.['c:chart']?.['c:title'];
  const titleArr = toArray(titleNode);
  if (!titleArr.length) return undefined;
  const tn = titleArr[0];
  if (!tn) return undefined;

  // overlay 标志 (c:overlay 子元素，val 属性)
  const overlayNode = toArray(tn['c:overlay'])[0];
  const overlay = overlayNode?.['@_val'] === '1';

  // 富文本标题
  const tx = toArray(tn['c:tx'])[0];
  if (tx) {
    const rich = toArray(tx['c:rich'])[0];
    if (rich) {
      const paras = toArray(rich['a:p']);
      const texts: string[] = [];
      for (const para of paras) {
        const runs = toArray(para['a:r']);
        for (const run of runs) {
          const t = run['a:t'];
          if (t) texts.push(typeof t === 'string' ? t : extractText(t));
        }
      }
      if (texts.length > 0) return { text: texts.join(''), overlay };
    }
    // 引用单元格标题
    const strRef = toArray(tx['c:strRef'])[0];
    if (strRef) {
      const f = toArray(strRef['c:f'])[0];
      const refText = typeof f === 'string' ? f : extractText(f);
      const parsed = parseCellRangeRef(refText);
      if (parsed) {
        const values = readCellValues(workbook, parsed.sheetName, parsed.startCol, parsed.startRow, parsed.endCol, parsed.endRow);
        if (values.length > 0 && values[0].length > 0) {
          return { text: String(values[0][0] || ''), overlay };
        }
      }
      const strCache = toArray(strRef['c:strCache'])[0];
      if (strCache) {
        const cached = parseStrPtValues(strCache);
        if (cached.length > 0) return { text: cached[0], overlay };
      }
    }
  }
  return undefined;
}

// ===== 系列解析 =====

/** 解析系列名称 */
function parseSeriesName(ser: any, workbook: ExcelJS.Workbook): string {
  const txNode = toArray(ser['c:tx'])[0];
  if (!txNode) return '';
  const strRef = toArray(txNode['c:strRef'])[0];
  if (strRef) {
    const f = toArray(strRef['c:f'])[0];
    const refText = typeof f === 'string' ? f : extractText(f);
    const parsed = parseCellRangeRef(refText);
    if (parsed) {
      const values = readCellValues(workbook, parsed.sheetName, parsed.startCol, parsed.startRow, parsed.endCol, parsed.endRow);
      if (values.length > 0 && values[0].length > 0) {
        return String(values[0][0] || '');
      }
    }
    const strCache = toArray(strRef['c:strCache'])[0];
    if (strCache) {
      const cached = parseStrPtValues(strCache);
      if (cached.length > 0) return cached[0];
    }
  }
  const vNode = toArray(txNode['c:v'])[0];
  if (vNode) return typeof vNode === 'string' ? vNode : extractText(vNode);
  return '';
}

/** 解析标记符号 */
function parseMarker(ser: any): { marker: MarkerSymbol; size: number } | null {
  const markerNode = toArray(ser['c:marker'])[0];
  if (!markerNode) return null;
  const symbol = markerNode['@_symbol'] || markerNode['c:symbol']?.[0]?.['@_val'];
  const size = getNum(markerNode['c:size']?.[0]?.['@_val'] || markerNode['@_size']);
  let sym: MarkerSymbol = 'circle';
  if (symbol === 'none') sym = 'none';
  else if (symbol === 'square') sym = 'square';
  else if (symbol === 'diamond') sym = 'diamond';
  else if (symbol === 'triangle') sym = 'triangle';
  return { marker: sym, size: size || 5 };
}

/**
 * 解析单个数据系列
 */
function parseSeries(
  chartType: ChartType,
  serNode: any,
  workbook: ExcelJS.Workbook,
  seriesIndex: number
): ChartSeriesModel | null {
  const name = parseSeriesName(serNode, workbook);
  const color = parseColorFromSpPr(serNode['c:spPr']);
  const smooth = serNode['c:smooth']?.[0]?.['@_val'] === '1';

  const markerInfo = parseMarker(serNode);

  const series: ChartSeriesModel = {
    name: name || `Series ${seriesIndex + 1}`,
    type: chartType,
    color: color || DEFAULT_SERIES_COLORS[seriesIndex % DEFAULT_SERIES_COLORS.length],
    smooth,
    marker: markerInfo?.marker,
    markerSize: markerInfo?.size,
  };

  // 根据图表类型解析数据
  switch (chartType) {
    case 'bar':
    case 'line':
    case 'area':
    case 'radar': {
      const valNode = toArray(serNode['c:val'])[0];
      series.data = readNumRef(valNode, workbook);
      break;
    }

    case 'pie':
    case 'doughnut': {
      const valNode = toArray(serNode['c:val'])[0];
      series.data = readNumRef(valNode, workbook);
      break;
    }

    case 'scatter': {
      const xValNode = toArray(serNode['c:xVal'])[0];
      const yValNode = toArray(serNode['c:yVal'])[0];
      const xValues = readNumRef(xValNode, workbook);
      const yValues = readNumRef(yValNode, workbook);
      const count = Math.max(xValues.length, yValues.length);
      series.points = [];
      for (let i = 0; i < count; i++) {
        series.points.push({
          x: xValues[i] ?? i,
          y: yValues[i] ?? 0,
        });
      }
      break;
    }

    case 'bubble': {
      const xValNode = toArray(serNode['c:xVal'])[0];
      const yValNode = toArray(serNode['c:yVal'])[0];
      const sizeNode = toArray(serNode['c:bubbleSize'])[0];
      const xValues = readNumRef(xValNode, workbook);
      const yValues = readNumRef(yValNode, workbook);
      const sizes = readNumRef(sizeNode, workbook);
      const count = Math.max(xValues.length, yValues.length);
      series.points = [];
      for (let i = 0; i < count; i++) {
        series.points.push({
          x: xValues[i] ?? i,
          y: yValues[i] ?? 0,
          size: sizes[i] ?? 10,
        });
      }
      break;
    }

    case 'stock': {
      const highNode = toArray(serNode['c:high'])[0];
      const lowNode = toArray(serNode['c:low'])[0];
      const openNode = toArray(serNode['c:open'])[0];
      const closeNode = toArray(serNode['c:close'])[0];
      const high = readNumRef(highNode, workbook);
      const low = readNumRef(lowNode, workbook);
      const open = readNumRef(openNode, workbook);
      const close = readNumRef(closeNode, workbook);
      const ohlc: StockOhlcData = { open, close, high, low };
      // volume (可选)
      const volNode = toArray(serNode['c:volume'])[0];
      if (volNode) ohlc.volume = readNumRef(volNode, workbook);
      series.ohlc = ohlc;
      series.data = close; // 默认用收盘价作为主数据
      break;
    }

    case 'surface': {
      const valNode = toArray(serNode['c:val'])[0];
      series.data = readNumRef(valNode, workbook);
      break;
    }
  }

  // 线条样式
  const ln = toArray(serNode['c:spPr']?.[0]?.['a:ln'] || serNode['c:spPr']?.['a:ln'])[0];
  if (ln) {
    const dash = ln['@_dash'] || ln['a:prstDash']?.[0]?.['@_val'];
    if (dash === 'dash') series.lineStyle = 'dash';
    else if (dash === 'dot') series.lineStyle = 'dot';
    else if (dash === 'dashDot') series.lineStyle = 'dashDot';
    else series.lineStyle = 'solid';
    const w = getNum(ln['@_w']);
    if (w) series.lineWidth = Math.round(w / 12700); // EMU to px
  }

  // 有效数据检查
  const hasData =
    (series.data && series.data.length > 0) ||
    (series.points && series.points.length > 0) ||
    (series.ohlc && series.ohlc.close.length > 0);

  if (!hasData && !name) return null;

  return series;
}

// ===== 分类标签解析 =====

/** 从第一个系列解析分类标签 */
function parseCategories(plotArea: any, chartKey: string, workbook: ExcelJS.Workbook): string[] {
  const chartEl = plotArea?.[chartKey];
  const chartElArr = toArray(chartEl);
  const firstSer = toArray(chartElArr[0]?.['c:ser'])[0];
  if (!firstSer) return [];

  const catNode = toArray(firstSer['c:cat'])[0];
  if (!catNode) return [];

  // 字符串引用
  const strRef = toArray(catNode['c:strRef'])[0];
  if (strRef) {
    const f = toArray(strRef['c:f'])[0];
    const refText = typeof f === 'string' ? f : extractText(f);
    const parsed = parseCellRangeRef(refText);
    if (parsed) {
      const values = readCellValues(workbook, parsed.sheetName, parsed.startCol, parsed.startRow, parsed.endCol, parsed.endRow);
      const cats = values.flat().map(v => String(v || ''));
      if (cats.length > 0) return cats;
    }
    const strCache = toArray(strRef['c:strCache'])[0];
    if (strCache) {
      const cached = parseStrPtValues(strCache);
      if (cached.length > 0) return cached;
    }
  }

  // 数值引用（分类为数字时）
  const numRef = toArray(catNode['c:numRef'])[0];
  if (numRef) {
    const f = toArray(numRef['c:f'])[0];
    const refText = typeof f === 'string' ? f : extractText(f);
    const parsed = parseCellRangeRef(refText);
    if (parsed) {
      const values = readCellValues(workbook, parsed.sheetName, parsed.startCol, parsed.startRow, parsed.endCol, parsed.endRow);
      const cats = values.flat().map(v => String(v ?? ''));
      if (cats.length > 0) return cats;
    }
    const numCache = toArray(numRef['c:numCache'])[0];
    if (numCache) {
      const cached = parseNumPtValues(numCache);
      if (cached.length > 0) return cached.map(v => String(v));
    }
  }

  // 直接字符串
  const strLit = toArray(catNode['c:strLit'])[0];
  if (strLit) return parseStrPtValues(strLit);

  return [];
}

// ===== 图表类型元素扫描 =====

interface DetectedChartElement {
  key: string;
  info: ChartElementInfo;
  node: any;
}

/** 扫描 plotArea 中的所有图表类型元素 */
function detectChartElements(plotArea: any): DetectedChartElement[] {
  const detected: DetectedChartElement[] = [];
  for (const info of CHART_ELEMENT_MAP) {
    const node = plotArea?.[info.ooxmlKey];
    if (node) {
      const nodeArr = toArray(node);
      for (const n of nodeArr) {
        if (n) detected.push({ key: info.ooxmlKey, info, node: n });
      }
    }
  }
  return detected;
}

// ===== 坐标轴解析 =====

/** 解析坐标轴配置 */
function parseAxes(plotArea: any, chartElements: DetectedChartElement[]): {
  xAxis?: ChartAxisModel;
  yAxis?: ChartAxisModel;
  yAxisSecondary?: ChartAxisModel;
  hasSecondaryAxis: boolean;
} {
  // 收集所有 axId 引用
  const allAxisIds: string[] = [];
  for (const el of chartElements) {
    const axIds = toArray(el.node['c:axId']);
    for (const ax of axIds) {
      const id = ax?.['@_val'];
      if (id) allAxisIds.push(String(id));
    }
  }

  // 查找坐标轴定义
  const catAxes = toArray(plotArea?.['c:catAx']);
  const valAxes = toArray(plotArea?.['c:valAx']);
  const dateAxes = toArray(plotArea?.['c:dateAx']);

  const hasSecondaryAxis = valAxes.length > 1;
  let yAxis: ChartAxisModel | undefined;
  let yAxisSecondary: ChartAxisModel | undefined;
  let xAxis: ChartAxisModel | undefined;

  // 类目轴
  if (catAxes.length > 0 || dateAxes.length > 0) {
    const catAx = catAxes[0] || dateAxes[0];
    xAxis = createDefaultAxis('category', 'bottom');
    const titleNode = toArray(catAx?.['c:title'])[0];
    if (titleNode) {
      const tx = toArray(titleNode['c:tx'])[0];
      const rich = toArray(tx?.['c:rich'])[0];
      if (rich) {
        const runs = toArray(toArray(rich['a:p'])[0]?.['a:r']);
        const t = runs.map((r: any) => extractText(r['a:t'])).join('');
        if (t) xAxis.title = t;
      }
    }
  }

  // 数值轴
  for (let i = 0; i < valAxes.length; i++) {
    const va = valAxes[i];
    const axis = createDefaultAxis('value', 'left');
    const titleNode = toArray(va?.['c:title'])[0];
    if (titleNode) {
      const tx = toArray(titleNode['c:tx'])[0];
      const rich = toArray(tx?.['c:rich'])[0];
      if (rich) {
        const runs = toArray(toArray(rich['a:p'])[0]?.['a:r']);
        const t = runs.map((r: any) => extractText(r['a:t'])).join('');
        if (t) axis.title = t;
      }
    }
    // min/max
    const min = toArray(va?.['c:scaling']?.[0]?.['c:min'])[0];
    const max = toArray(va?.['c:scaling']?.[0]?.['c:max'])[0];
    if (min) axis.min = getNum(min['@_val'] || min);
    if (max) axis.max = getNum(max['@_val'] || max);

    if (i === 0) yAxis = axis;
    else yAxisSecondary = axis;
  }

  return { xAxis, yAxis, yAxisSecondary, hasSecondaryAxis };
}

// ===== 图例解析 =====

/** 解析图例配置 */
function parseLegend(chartXmlObj: any): ChartLegendModel | undefined {
  const legendNode = chartXmlObj?.['c:chartSpace']?.['c:chart']?.['c:legend'];
  const legendArr = toArray(legendNode);
  if (!legendArr.length) return undefined;
  const ln = legendArr[0];
  if (!ln) return undefined;

  const pos = ln['c:legendPos']?.[0]?.['@_val'] || 'b';
  const positionMap: Record<string, 'top' | 'bottom' | 'left' | 'right'> = {
    b: 'bottom', t: 'top', l: 'left', r: 'right',
  };
  const overlay = ln['@_overlay'] === '1';

  return {
    visible: true,
    position: positionMap[pos] || 'bottom',
  };
}

// ===== 柱状图方向 & 分组 =====

/** 解析柱状图方向 */
function parseBarDirection(chartNode: any): BarDirection {
  const dir = chartNode?.['@_barDir'];
  return dir === 'bar' ? 'bar' : 'col';
}

/** 解析分组方式 */
function parseGrouping(chartNode: any): Grouping {
  const g = chartNode?.['@_grouping'] || chartNode?.['c:grouping']?.[0]?.['@_val'];
  if (g === 'stacked') return 'stacked';
  if (g === 'percentStacked') return 'percentStacked';
  if (g === 'standard') return 'standard';
  return 'clustered';
}

/** 解析 gapWidth */
function parseGapWidth(chartNode: any): number | undefined {
  const gap = chartNode?.['c:gapWidth']?.[0]?.['@_val'];
  return gap !== undefined ? getNum(gap) : undefined;
}

// ===== 雷达图样式 =====

function parseRadarStyle(chartNode: any): RadarStyle {
  const style = chartNode?.['c:radarStyle']?.[0]?.['@_val'];
  if (style === 'marker') return 'marker';
  if (style === 'filled') return 'filled';
  return 'standard';
}

// ===== 股价图子类型 =====

function parseStockChartType(chartNode: any): StockChartType {
  // OOXML stockChart 没有显式子类型标志，通过系列数据推断
  // 如果有 open 数据则为 K线图
  const ser = toArray(chartNode?.['c:ser'])[0];
  if (ser?.['c:open']) return 'openHighLowClose';
  if (ser?.['c:volume']) return 'volumeHighLowClose';
  return 'highLowClose';
}

// ===== 主解析函数 =====

/**
 * 将单个图表 XML 解析为 ChartModel
 *
 * @param chartXmlObj - fast-xml-parser 解析后的图表 XML 对象
 * @param workbook - exceljs Workbook（用于读取单元格数据）
 * @param anchor - 图表锚点位置
 * @param sheetIndex - 所属工作表索引
 * @param chartId - 图表唯一标识
 * @returns ChartModel
 */
export function parseChartXmlToModel(
  chartXmlObj: any,
  workbook: ExcelJS.Workbook,
  anchor: ChartAnchor,
  sheetIndex: number,
  chartId: string,
  theme?: ChartTheme
): ChartModel | null {
  // 应用主题色（如果传入）
  if (theme) {
    THEME_COLOR_MAP = themeColorsToMap(theme);
  }
  const plotArea = getPlotArea(chartXmlObj);
  if (!plotArea) return null;

  // 1. 检测所有图表类型元素
  const detected = detectChartElements(plotArea);
  if (detected.length === 0) return null;

  // 2. 解析坐标轴
  const { xAxis, yAxis, yAxisSecondary, hasSecondaryAxis } = parseAxes(plotArea, detected);

  // 3. 解析标题
  const title = extractTitle(chartXmlObj, workbook);

  // 4. 解析图例
  const legend = parseLegend(chartXmlObj);

  // 5. 判断是单一图表还是组合图
  const isCombo = detected.length > 1;

  // 收集所有系列
  const allSeries: ChartSeriesModel[] = [];
  const plotGroups: PlotGroup[] = [];
  let categories: string[] = [];
  let primaryChartType: ChartType = detected[0].info.chartType;
  let is3D = false;
  let barDirection: BarDirection | undefined;
  let grouping: Grouping | undefined;
  let gapWidth: number | undefined;
  let radarStyle: RadarStyle | undefined;
  let stockChartType: StockChartType | undefined;

  for (let di = 0; di < detected.length; di++) {
    const { key, info, node } = detected[di];
    if (info.is3D) is3D = true;

    // 解析该图表类型下的系列
    const serNodes = toArray(node['c:ser']);
    const seriesIndices: number[] = [];

    // 解析分类（仅从第一个图表类型取）
    if (di === 0) {
      categories = parseCategories(plotArea, key, workbook);
    }

    // 图表类型特定属性
    if (info.chartType === 'bar') {
      if (barDirection === undefined) barDirection = parseBarDirection(node);
      if (grouping === undefined) grouping = parseGrouping(node);
      if (gapWidth === undefined) gapWidth = parseGapWidth(node);
    }
    if (info.chartType === 'radar') {
      if (radarStyle === undefined) radarStyle = parseRadarStyle(node);
    }
    if (info.chartType === 'stock') {
      if (stockChartType === undefined) stockChartType = parseStockChartType(node);
    }

    for (let si = 0; si < serNodes.length; si++) {
      const series = parseSeries(info.chartType, serNodes[si], workbook, allSeries.length);
      if (series) {
        // 组合图中：第二个图表类型元素使用次坐标轴
        if (isCombo && di > 0) {
          series.yAxisIndex = 1;
        }
        allSeries.push(series);
        seriesIndices.push(allSeries.length - 1);
      }
    }

    if (isCombo) {
      plotGroups.push({
        type: info.chartType,
        seriesIndices,
        yAxisIndex: di === 0 ? 0 : 1,
        grouping: info.chartType === 'bar' ? parseGrouping(node) : undefined,
      });
    }
  }

  // 6. 构建 PlotAreaModel
  const plotAreaModel: PlotAreaModel = {
    hasSecondaryAxis,
    gapWidth,
  };
  if (primaryChartType === 'radar') {
    plotAreaModel.radarShape = 'polygon';
  }

  // 7. 组装 ChartModel
  const model: ChartModel = {
    id: chartId,
    type: isCombo ? 'combo' : primaryChartType,
    title,
    series: allSeries,
    categories: categories.length > 0 ? categories : undefined,
    xAxis,
    yAxis,
    yAxisSecondary: hasSecondaryAxis ? yAxisSecondary : undefined,
    legend,
    plotArea: plotAreaModel,
    plotGroups: isCombo ? plotGroups : undefined,
    barDirection,
    grouping,
    is3D,
    radarStyle,
    stockChartType,
    anchor,
    sheetIndex,
  };

  return model;
}
