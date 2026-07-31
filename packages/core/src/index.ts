/**
 * @excel-preview/core
 *
 * Excel 预览与图表渲染核心引擎（框架无关）
 */

// 类型导出
export type {
  DataSource,
  LoadOptions,
  CellType,
  CellFontStyle,
  BorderItem,
  CellStyle,
  MergeInfo,
  ParsedCell,
  ParsedSheet,
  ChartType,
  ChartAnchor,
  ChartSeries,
  ParsedChart,
  ParsedImage,
  ParsedWorkbook,
  ViewerOptions,
  ExcelSource,
  ExcelViewerOptions,
  RenderedEvent,
  SwitchSheetEvent,
  CellSelectedEvent,
  CellsSelectedEvent,
  ParsedPivotTable,
} from './types';

// Loader
export { loadData, isUrlSource, isBinarySource } from './loader';

// Parser
export { parseExcel, loadRawWorkbook } from './parser/excel-parser';
export { parseCharts } from './parser/chart-parser';
export { parsePivotTables } from './parser/pivot-parser';

// Chart 模块（OOXML Parser + ChartModel + Layout Engine + ECharts Adapter + Canvas Renderer + Google Charts Adapter）
export type {
  ChartType as ChartModelType,
  BarDirection,
  Grouping,
  MarkerSymbol,
  LineStyle,
  RadarStyle,
  StockChartType,
  ChartDataPoint,
  ChartSeriesModel,
  StockOhlcData,
  AxisType,
  ChartAxisModel,
  ChartLegendModel,
  ChartTitleModel,
  PlotAreaModel,
  PlotGroup,
  ChartModel,
  ChartContainerSize,
  LayoutRect,
  ChartLayout,
  ThemeColorScheme,
  ThemeFontScheme,
  ChartTheme,
  GoogleChartConfig,
  GoogleRendererConfig,
} from './chart';
export {
  createDefaultAxis,
  createDefaultLegend,
  computeLayout,
  layoutToGrid,
  parseChartXmlToModel,
  convertToEChartsOption,
  convertToGoogleChart,
  parseThemeXml,
  parseThemeFromZip,
  themeColorsToMap,
  DEFAULT_THEME,
  CanvasChartRenderer,
  GoogleChartsRenderer,
} from './chart';

// Renderer
export { TableRenderer } from './renderer/table-renderer';
export { ChartRenderer } from './renderer/chart-renderer';
export type { PositionFn, ChartBackend } from './renderer/chart-renderer';
export { ImageRenderer } from './renderer/image-renderer';

// High-level Viewer
export { ExcelViewer } from './excel-viewer';
