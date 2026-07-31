/**
 * ChartModel — 图表领域模型
 *
 * 描述 Excel 图表的完整语义信息，与渲染层（ECharts）解耦。
 *
 * 数据流:
 *   OOXML XML ──► OoxmlChartParser ──► ChartModel
 *                                          │
 *                          LayoutEngine ───┤ (输入: model + 容器尺寸)
 *                                          │
 *                                  EChartsConverter ──► ECharts option
 *
 * 支持的图表类型:
 *   bar / line / area / pie / doughnut / scatter / bubble
 *   radar / stock / surface / combo（组合图）
 */

import type { ChartAnchor } from '../types';

// ===== 图表类型 =====

/** 图表类型枚举 */
export type ChartType =
  | 'bar' // 柱状图 / 条形图（方向由 barDirection 决定）
  | 'line' // 折线图
  | 'area' // 面积图
  | 'pie' // 饼图
  | 'doughnut' // 环形图
  | 'scatter' // 散点图
  | 'bubble' // 气泡图
  | 'radar' // 雷达图
  | 'stock' // 股价图（K线 / OHLC）
  | 'surface' // 曲面图（降级为等高线/热力图渲染）
  | 'combo'; // 组合图（多类型混合 + 双坐标轴）

/** 柱状图方向: col=垂直柱状, bar=水平条形 */
export type BarDirection = 'col' | 'bar';

/** 系列分组方式 */
export type Grouping =
  | 'clustered' // 簇状
  | 'stacked' // 堆积
  | 'percentStacked' // 百分比堆积
  | 'standard'; // 标准

/** 标记形状 */
export type MarkerSymbol = 'circle' | 'square' | 'diamond' | 'triangle' | 'none';

/** 线型 */
export type LineStyle = 'solid' | 'dash' | 'dot' | 'dashDot';

/** 雷达图样式 */
export type RadarStyle = 'standard' | 'marker' | 'filled';

/** 股价图子类型 */
export type StockChartType =
  | 'highLowClose' // 最高-最低-收盘
  | 'openHighLowClose' // 开-高-低-收（K线）
  | 'volumeHighLowClose'; // 成交量-高-低-收

// ===== 数据结构 =====

/** 单个数据点（散点/气泡图使用多维数据） */
export interface ChartDataPoint {
  /** X 值（散点图） */
  x?: number;
  /** Y 值 */
  y: number;
  /** 气泡大小 */
  size?: number;
  /** 显示标签 */
  label?: string;
}

/** 图表系列 */
export interface ChartSeriesModel {
  /** 系列名称 */
  name: string;
  /** 系列类型（组合图中不同系列可为不同类型） */
  type: ChartType;
  /** 一维数值数据（柱/线/面/饼/雷达 等常规图） */
  data?: number[];
  /** 多维数据点（散点/气泡图） */
  points?: ChartDataPoint[];
  /** 系列颜色 (hex, 如 #5B9BD5) */
  color?: string;
  /** 线条样式 */
  lineStyle?: LineStyle;
  /** 线宽 */
  lineWidth?: number;
  /** 是否平滑曲线 */
  smooth?: boolean;
  /** 标记形状 */
  marker?: MarkerSymbol;
  /** 标记大小 */
  markerSize?: number;
  /** 柱状图宽度（百分比或像素） */
  barWidth?: number | string;
  /** 所属 Y 轴索引（0=主轴, 1=次轴，用于组合图） */
  yAxisIndex?: 0 | 1;
  /** 面积填充透明度 (0-1) */
  areaOpacity?: number;
  /** Excel 数据标签配置 */
  dataLabels?: {
    showValue?: boolean;
    showCategoryName?: boolean;
    showSeriesName?: boolean;
    showPercent?: boolean;
    position?: 'top' | 'bottom' | 'left' | 'right' | 'inside' | 'insideTop' | 'insideBottom' | 'center';
  };
  /** 是否堆积 */
  stack?: string;
  /** 股价图 OHLC 数据 */
  ohlc?: StockOhlcData;
}

/** 股价图 OHLC 数据 */
export interface StockOhlcData {
  open: number[];
  close: number[];
  high: number[];
  low: number[];
  /** 成交量（可选） */
  volume?: number[];
}

// ===== 坐标轴 =====

/** 坐标轴类型 */
export type AxisType = 'category' | 'value' | 'time' | 'log';

/** 坐标轴模型 */
export interface ChartAxisModel {
  type: AxisType;
  title?: string;
  min?: number;
  max?: number;
  /** 是否反向 */
  inverse?: boolean;
  /** 标签旋转角度 */
  labelRotate?: number;
  /** 是否显示 */
  visible: boolean;
  position?: 'bottom' | 'top' | 'left' | 'right';
  /** 网格线样式 */
  splitLine?: 'solid' | 'dashed' | 'none';
  /** 标签字号 */
  labelFontSize?: number;
}

// ===== 图例 & 标题 =====

/** 图例配置 */
export interface ChartLegendModel {
  visible: boolean;
  position: 'top' | 'bottom' | 'left' | 'right';
  orientation?: 'horizontal' | 'vertical';
}

/** 标题配置 */
export interface ChartTitleModel {
  text: string;
  /** 是否覆盖在绘图区上方（而非占用独立空间） */
  overlay?: boolean;
}

// ===== 绘图区 & 组合图 =====

/** 绘图区配置 */
export interface PlotAreaModel {
  /** 柱状图簇间距百分比 (0-500) */
  gapWidth?: number;
  /** 是否有次坐标轴 */
  hasSecondaryAxis?: boolean;
  /** 雷达图形状 */
  radarShape?: 'polygon' | 'circle';
}

/** 组合图分组：同一组系列共享坐标轴和图表类型 */
export interface PlotGroup {
  type: ChartType;
  /** 该组包含的系列索引列表 */
  seriesIndices: number[];
  yAxisIndex: 0 | 1;
  grouping?: Grouping;
}

// ===== 完整图表模型 =====

/** 图表领域模型（完整语义描述） */
export interface ChartModel {
  /** 唯一标识 */
  id: string;
  /** 图表类型 */
  type: ChartType;
  /** 标题 */
  title?: ChartTitleModel;
  /** 数据系列 */
  series: ChartSeriesModel[];
  /** 分类标签（X 轴） */
  categories?: string[];
  /** X 轴配置 */
  xAxis?: ChartAxisModel;
  /** Y 轴（主）配置 */
  yAxis?: ChartAxisModel;
  /** Y 轴（次）配置（组合图） */
  yAxisSecondary?: ChartAxisModel;
  /** 图例配置 */
  legend?: ChartLegendModel;
  /** 绘图区配置 */
  plotArea?: PlotAreaModel;
  /** 组合图分组定义 */
  plotGroups?: PlotGroup[];

  // 通用样式属性
  /** 柱状图方向 */
  barDirection?: BarDirection;
  /** 分组方式 */
  grouping?: Grouping;
  /** 全局平滑 */
  smooth?: boolean;
  /** 全局显示标记点 */
  showMarkers?: boolean;
  /** 是否 3D 图表 */
  is3D?: boolean;
  /** 雷达图样式 */
  radarStyle?: RadarStyle;
  /** 股价图子类型 */
  stockChartType?: StockChartType;

  // 定位信息
  /** 在 Excel 中的锚点位置 */
  anchor: ChartAnchor;
  /** 所属工作表索引 */
  sheetIndex: number;
}

// ===== 辅助工厂函数 =====

/** 创建默认坐标轴 */
export function createDefaultAxis(type: AxisType, position: ChartAxisModel['position']): ChartAxisModel {
  return {
    type,
    visible: true,
    position,
    splitLine: type === 'value' ? 'dashed' : 'none',
    labelFontSize: 9,
  };
}

/** 创建默认图例 */
export function createDefaultLegend(): ChartLegendModel {
  return { visible: true, position: 'bottom' };
}
