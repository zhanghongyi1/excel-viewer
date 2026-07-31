/**
 * Chart 模块 — 完整图表架构
 *
 * 五层架构:
 *   1. OOXML Parser (theme-parser + ooxml-chart-parser) — 解析 XML
 *   2. ChartModel — 领域模型（与渲染库解耦）
 *   3. Layout Engine — 计算标题/图例/坐标轴/绘图区布局
 *   4. ECharts Adapter — ChartModel → ECharts option（高质量渲染）
 *   5. Canvas Renderer — 自研 Canvas 渲染引擎（零外部依赖）
 *
 * 数据流:
 *   theme1.xml ──► ThemeParser ──► ChartTheme ─────────────────┐
 *   chart*.xml  ──► OoxmlChartParser ──► ChartModel ────────────┤
 *                                                                ▼
 *   LayoutEngine.computeLayout(model, size) ──► ChartLayout ──► ┬── EChartsConverter → ECharts option
 *                                                                └── CanvasChartRenderer → <canvas>
 */

// 领域模型
export type {
  ChartType,
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
} from './chart-model';

export { createDefaultAxis, createDefaultLegend } from './chart-model';

// 布局引擎
export type { ChartContainerSize, LayoutRect, ChartLayout } from './layout-engine';
export { computeLayout, layoutToGrid } from './layout-engine';

// 主题解析器
export type { ThemeColorScheme, ThemeFontScheme, ChartTheme } from './theme-parser';
export { parseThemeXml, parseThemeFromZip, themeColorsToMap, DEFAULT_THEME } from './theme-parser';

// OOXML 图表解析器
export { parseChartXmlToModel } from './ooxml-chart-parser';

// ECharts 转换器
export { convertToEChartsOption } from './echarts-converter';

// Canvas 渲染引擎
export { CanvasChartRenderer } from './canvas-chart-renderer';
export type { PixelRect, CanvasRendererConfig } from './canvas-chart-renderer';
