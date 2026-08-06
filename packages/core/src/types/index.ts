import type { ChartType } from '../chart/chart-model';
export type { ChartType } from '../chart/chart-model';

// ===== 数据源类型 =====

/** 支持的数据源类型 */
export type DataSource = string | ArrayBuffer | Blob | File;

/** 网络请求配置 */
export interface LoadOptions {
  /** 请求方法 */
  method?: 'GET' | 'POST';
  /** 自定义请求头 */
  headers?: Record<string, string>;
  /** 是否携带凭证 */
  withCredentials?: boolean;
  /** 请求体 (POST 时使用) */
  body?: any;
}

// ===== 单元格类型 =====

export type CellType =
  | 'string'
  | 'number'
  | 'date'
  | 'boolean'
  | 'formula'
  | 'richText'
  | 'hyperlink'
  | 'empty';

/** 字体样式 */
export interface CellFontStyle {
  name?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  underline?: boolean;
  strike?: boolean;
}

/** 边框样式: [边框样式, 颜色] */
export type BorderItem = [string, string];

/** 单元格样式 */
export interface CellStyle {
  font?: CellFontStyle;
  bgcolor?: string; // 背景色 (#RRGGBB)
  color?: string; // 文字颜色
  align?: 'left' | 'center' | 'right';
  valign?: 'top' | 'middle' | 'bottom';
  textwrap?: boolean;
  border?: {
    top?: BorderItem;
    bottom?: BorderItem;
    left?: BorderItem;
    right?: BorderItem;
  };
  numFmt?: string;
}

/** 合并单元格信息 */
export interface MergeInfo {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

/** 单元格批注 */
export interface CellComment {
  text: string;
  author?: string;
}

/** 条件格式规则类型 */
export type CfRuleType = 'cellIs' | 'expression' | 'colorScale' | 'dataBar' | 'iconSet' | 'top10' | 'aboveAverage' | 'duplicateValues' | 'uniqueValues' | 'containsText' | 'notContainsText' | 'beginsWith' | 'endsWith' | 'containsBlanks' | 'notContainsBlanks' | 'containsErrors' | 'notContainsErrors' | 'timePeriod' | 'formula';

/** 色阶条件格式 */
export interface CfColorScale {
  type: 'colorScale';
  cfvo: Array<{ type: 'min' | 'max' | 'num' | 'percent' | 'percentile' | 'formula'; value?: number | string }>;
  colors: string[];
}

/** 数据条条件格式 */
export interface CfDataBar {
  type: 'dataBar';
  cfvo: Array<{ type: 'min' | 'max' | 'num' | 'percent' | 'percentile' | 'formula'; value?: number | string }>;
  color: string;
  showValue?: boolean;
}

/** 图标集条件格式 */
export interface CfIconSet {
  type: 'iconSet';
  iconSet?: string;
  cfvo: Array<{ type: 'percent' | 'num' | 'percentile' | 'formula'; value?: number | string }>;
  showValue?: boolean;
}

/** 条件格式规则 */
export interface CfRule {
  type: CfRuleType;
  priority: number;
  formula?: string[];
  operator?: string;
  text?: string;
  fill?: string;
  font?: Partial<CellStyle['font']>;
  colorScale?: CfColorScale;
  dataBar?: CfDataBar;
  iconSet?: CfIconSet;
}

/** 条件格式分段 */
export interface ConditionalFormatting {
  ref: string;
  rules: CfRule[];
}

/** 富文本片段 */
export interface RichTextRun {
  text: string;
  font?: CellFontStyle;
}

export interface ParsedCell {
  value: any;
  text: string;
  type: CellType;
  style: CellStyle;
  merge?: MergeInfo;
  comment?: CellComment;
  richTextRuns?: RichTextRun[];
}

// ===== 工作表类型 =====

/** 冻结窗格 */
export interface FreezePane {
  xSplit: number;
  ySplit: number;
}

/** 解析后的工作表 */
export interface ParsedSheet {
  name: string;
  id: string;
  rows: ParsedCell[][];
  merges: MergeInfo[];
  colWidths: number[];
  rowHeights: number[];
  freezePane?: FreezePane;
  conditionalFormatting?: ConditionalFormatting[];
  hiddenRows?: Set<number>;
  hiddenCols?: Set<number>;
}

// ===== 图表类型 =====

/** 图表锚点定位（基于单元格坐标） */
export interface ChartAnchor {
  fromCol: number;
  fromColOff: number; // 列偏移 (EMU)
  fromRow: number;
  fromRowOff: number; // 行偏移 (EMU)
  toCol: number;
  toColOff: number;
  toRow: number;
  toRowOff: number;
}

/** 图表数据系列 */
export interface ChartSeries {
  name: string;
  data: number[];
  type?: ChartType;
  color?: string; // 系列颜色 (hex)
}

/**
 * 解析后的图表（旧版接口，已由 ChartModel 替代）
 * @deprecated 请使用 `import { ChartModel } from '@excel-preview/core'` 代替
 */
export interface ParsedChart {
  id: string;
  type: ChartType;
  title?: string;
  anchor: ChartAnchor;
  series: ChartSeries[];
  categories?: string[];
  echartsOption: object;
  sheetIndex: number;
}

/** 解析后的图片 */
export interface ParsedImage {
  id: string;
  anchor: ChartAnchor;
  imageData: string; // base64 或 blob URL
  mimeType: string;
  name?: string;
  sheetIndex: number; // 所属 Sheet 索引
  /** 图片在 Excel 中的实际尺寸（EMU），用于保持原始宽高比 */
  extent?: { width: number; height: number };
}

// ===== 数据透视表类型 =====

/** 透视表缓存字段 */
export interface PivotCacheField {
  name: string;
  sharedItems: string[];
  isNumeric: boolean;
}

/** 透视表缓存记录 */
export interface PivotCacheRecord {
  values: (string | number | null)[];
}

/** 透视表数据字段 */
export interface PivotDataField {
  fieldIndex: number;
  name: string;
  summarizeFunction: 'sum' | 'count' | 'average' | 'max' | 'min';
}

/** 解析后的数据透视表 */
export interface ParsedPivotTable {
  name: string;
  sheetIndex: number;
  /** 裸缓存字段定义 */
  cacheFields: PivotCacheField[];
  /** 裸缓存记录（行数据） */
  cacheRecords: PivotCacheRecord[];
  /** 行字段在 cacheFields 中的索引 */
  rowFieldIndices: number[];
  /** 列字段在 cacheFields 中的索引 */
  colFieldIndices: number[];
  /** 数据字段 */
  dataFields: PivotDataField[];
}

// ===== 工作簿类型 =====

/** 解析后的完整工作簿 */
export interface ParsedWorkbook {
  sheets: ParsedSheet[];
  pivotTables?: ParsedPivotTable[];
}

// ===== 渲染选项 =====

/** 图表渲染共享配置 */
export interface ChartRenderOptions {
  /** ECharts 实例；不传时由 ExcelViewer 按需加载 */
  echarts?: any;
  chartBackend?: 'echarts' | 'canvas' | 'auto';
  echartsRenderer?: 'svg' | 'canvas';
}

/** 预览器配置选项 */
export interface ViewerOptions extends ChartRenderOptions {
  /** 最小列数 */
  minColLength?: number;
  /** 最小行数 */
  minRowLength?: number;
  /** 是否显示底部 Sheet 切换栏 */
  showToolbar?: boolean;
}

// ===== ExcelViewer 高阶选项 =====

/** Excel 数据源类型 */
export type ExcelSource = string | File | Blob | ArrayBuffer;

/** ExcelViewer 构造选项 */
export interface ExcelViewerOptions extends ChartRenderOptions {
  /** 挂载目标：CSS 选择器或 HTMLElement */
  target?: HTMLElement | string;
  /** 初始数据源 */
  src?: ExcelSource;
  /** 容器宽度 */
  width?: string;
  /** 容器高度 */
  height?: string;
  /** 是否显示底部 Sheet 切换栏 */
  showToolbar?: boolean;
  /** 是否解析数据透视表缓存（默认 false，解析大文件会增加开销） */
  parsePivotTables?: boolean;
  /** 渲染完成回调 */
  onRendered?: () => void;
  /** 错误回调 */
  onError?: (error: Error) => void;
  /** Sheet 切换回调 */
  onSheetChange?: (sheetName: string, sheetIndex: number) => void;
}

// ===== 事件类型 =====

/** 渲染完成回调参数 */
export interface RenderedEvent {
  sheetCount: number;
  chartCount: number;
}

/** Sheet 切换事件参数 */
export interface SwitchSheetEvent {
  index: number;
  name: string;
}

/** 单元格选中事件参数 */
export interface CellSelectedEvent {
  cell: any;
  rowIndex: number;
  columnIndex: number;
}

/** 区域选中事件参数 */
export interface CellsSelectedEvent {
  startRowIndex: number;
  startColumnIndex: number;
  endRowIndex: number;
  endColumnIndex: number;
}
