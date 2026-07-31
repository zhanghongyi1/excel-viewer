/**
 * Google Charts Converter — ChartModel → Google Visualization API
 *
 * 将 ChartModel 转换为 Google Charts 所需的:
 *   1. google.visualization.DataTable — 数据表
 *   2. Google Charts options — 图表配置
 *
 * 类型映射:
 *   ChartModel.type   →   Google Charts class
 *   bar (col)         →   ColumnChart
 *   bar (bar)         →   BarChart
 *   line              →   LineChart
 *   area              →   AreaChart
 *   pie               →   PieChart
 *   doughnut          →   PieChart (pieHole: 0.5)
 *   scatter           →   ScatterChart
 *   bubble            →   BubbleChart
 *   stock             →   CandlestickChart
 *   radar             →   降级为 LineChart (Google Charts 无原生雷达图)
 *   combo             →   ComboChart
 *   surface           →   降级为 AreaChart
 */

import type { ChartModel, ChartSeriesModel, ChartType } from './chart-model';

// ===== 类型定义 =====

/** Google Charts 渲染所需的数据包 */
export interface GoogleChartConfig {
  /** Google Charts 构造函数名称 (如 'ColumnChart', 'LineChart') */
  chartType: string;
  /** 数据表 (二维数组格式，便于 google.visualization.arrayToDataTable) */
  dataTable: (string | number | Date | null)[][];
  /** Google Charts 选项 */
  options: Record<string, any>;
}

// ===== 图表类型映射 =====

/** ChartModel 类型 → Google Charts 构造函数名称 */
function getGoogleChartType(model: ChartModel): string {
  // 水平条形图 → BarChart，垂直柱状图 → ColumnChart
  if (model.type === 'bar') {
    return model.barDirection === 'bar' ? 'BarChart' : 'ColumnChart';
  }
  if (model.type === 'combo') return 'ComboChart';

  const map: Record<string, string> = {
    line: 'LineChart',
    area: 'AreaChart',
    pie: 'PieChart',
    doughnut: 'PieChart',
    scatter: 'ScatterChart',
    bubble: 'BubbleChart',
    stock: 'CandlestickChart',
    radar: 'LineChart', // 降级
    surface: 'AreaChart', // 降级
  };
  return map[model.type] || 'ColumnChart';
}

// ===== 数据表构建 =====

/**
 * 构建常规图表（柱/线/面积/饼图）的 DataTable
 *
 * 格式:
 *   [Category, Series1, Series2, ...]
 *   ['Q1',     10,       20 ],
 *   ['Q2',     20,       30 ],
 */
function buildCartesianDataTable(model: ChartModel): (string | number | null)[][] {
  const categories = model.categories || [];
  const series = model.series;
  if (series.length === 0) return [['Category', 'Value']];

  // 表头
  const header: (string | number)[] = ['Category'];
  for (const s of series) {
    header.push(s.name || `Series ${series.indexOf(s) + 1}`);
  }

  const table: (string | number | null)[][] = [header];

  // 数据行
  const dataCount = categories.length || getMaxDataCount(series);
  for (let i = 0; i < dataCount; i++) {
    const row: (string | number | null)[] = [categories[i] || `Item ${i + 1}`];
    for (const s of series) {
      row.push(s.data?.[i] ?? null);
    }
    table.push(row);
  }

  return table;
}

/**
 * 构建散点图的 DataTable
 *
 * 格式:
 *   [X, Y, Series]
 *   [1, 10, 'S1']
 *   [2, 20, 'S1']
 */
function buildScatterDataTable(model: ChartModel): (string | number | null)[][] {
  const table: (string | number | null)[][] = [['X', 'Y', 'Series']];
  for (const s of model.series) {
    if (!s.points) continue;
    for (const p of s.points) {
      table.push([p.x ?? 0, p.y, s.name || 'Series']);
    }
  }
  return table;
}

/**
 * 构建气泡图的 DataTable
 *
 * 格式:
 *   [X, Y, Series, Size]
 *   [1, 10, 'S1', 5]
 */
function buildBubbleDataTable(model: ChartModel): (string | number | null)[][] {
  const table: (string | number | null)[][] = [['X', 'Y', 'Series', 'Size']];
  for (const s of model.series) {
    if (!s.points) continue;
    for (const p of s.points) {
      table.push([p.x ?? 0, p.y, s.name || 'Series', p.size ?? 10]);
    }
  }
  return table;
}

/**
 * 构建股价图 (K线) 的 DataTable
 *
 * 格式 (CandlestickChart):
 *   [Date, Low, Open, Close, High]
 */
function buildStockDataTable(model: ChartModel): (string | number | null)[][] {
  const header: (string | number)[] = ['Date', 'Low', 'Open', 'Close', 'High'];

  const series = model.series[0];
  if (!series || !series.ohlc) return [header];

  const categories = model.categories || [];
  const ohlc = series.ohlc;
  const count = ohlc.close.length;

  const rows: (string | number | null)[][] = [header];
  for (let i = 0; i < count; i++) {
    rows.push([
      categories[i] || `Day ${i + 1}`,
      ohlc.low[i] ?? 0,
      ohlc.open[i] ?? ohlc.close[i],
      ohlc.close[i] ?? 0,
      ohlc.high[i] ?? 0,
    ]);
  }
  return rows;
}

/**
 * 构建组合图的 DataTable（与常规图表相同，但系列类型在 options 中指定）
 */
function buildComboDataTable(model: ChartModel): (string | number | null)[][] {
  return buildCartesianDataTable(model);
}

// ===== 选项构建 =====

/**
 * 构建常规 Cartesian 图表选项
 */
function buildCartesianOptions(model: ChartModel): Record<string, any> {
  const options: Record<string, any> = {
    backgroundColor: '#ffffff',
    legend: buildLegendOptions(model),
    chartArea: { width: '85%', height: '75%' },
  };

  // 标题
  if (model.title?.text) {
    options.title = model.title.text;
    options.titleTextStyle = { fontSize: 14, color: '#333' };
  }

  // 水平条形图 → 交换轴
  if (model.barDirection === 'bar') {
    options.hAxis = { format: 'short' };
    options.vAxis = { format: 'short' };
  } else {
    options.hAxis = { format: 'short' };
    options.vAxis = { format: 'short' };
  }

  // 堆积
  if (model.grouping === 'stacked') {
    options.isStacked = true;
  } else if (model.grouping === 'percentStacked') {
    options.isStacked = 'percent';
  }

  // 系列颜色
  if (model.series.length > 0) {
    const colors = model.series.map((s, i) => s.color || getDefaultColor(i));
    options.colors = colors;
  }

  // 面积图透明度
  if (model.type === 'area') {
    options.areaOpacity = 0.3;
  }

  return options;
}

/**
 * 构建饼图选项
 */
function buildPieOptions(model: ChartModel): Record<string, any> {
  const options: Record<string, any> = {
    backgroundColor: '#ffffff',
    chartArea: { width: '80%', height: '80%' },
  };

  if (model.title?.text) {
    options.title = model.title.text;
    options.titleTextStyle = { fontSize: 14, color: '#333' };
  }

  if (model.legend?.visible) {
    options.legend = { position: mapLegendPosition(model.legend.position) };
  } else {
    options.legend = { position: 'none' };
  }

  if (model.type === 'doughnut') {
    options.pieHole = 0.5;
  }

  // 系列颜色
  if (model.series.length > 0) {
    const colors = model.series.map((s, i) => s.color || getDefaultColor(i));
    options.colors = colors;
  }

  return options;
}

/**
 * 构建散点图选项
 */
function buildScatterOptions(model: ChartModel): Record<string, any> {
  const options: Record<string, any> = {
    backgroundColor: '#ffffff',
    chartArea: { width: '85%', height: '75%' },
    hAxis: { format: 'short' },
    vAxis: { format: 'short' },
    pointSize: 5,
  };

  if (model.title?.text) {
    options.title = model.title.text;
    options.titleTextStyle = { fontSize: 14, color: '#333' };
  }

  if (model.legend?.visible) {
    options.legend = { position: mapLegendPosition(model.legend.position) };
  } else {
    options.legend = { position: 'none' };
  }

  return options;
}

/**
 * 构建气泡图选项
 */
function buildBubbleOptions(model: ChartModel): Record<string, any> {
  const options: Record<string, any> = {
    backgroundColor: '#ffffff',
    chartArea: { width: '85%', height: '75%' },
    hAxis: { format: 'short' },
    vAxis: { format: 'short' },
    bubble: { textStyle: { fontSize: 0, color: 'transparent' } },
  };

  if (model.title?.text) {
    options.title = model.title.text;
    options.titleTextStyle = { fontSize: 14, color: '#333' };
  }

  return options;
}

/**
 * 构建股价图选项
 */
function buildStockOptions(model: ChartModel): Record<string, any> {
  const options: Record<string, any> = {
    backgroundColor: '#ffffff',
    chartArea: { width: '85%', height: '75%' },
    candlestick: {
      fallingColor: { fill: '#ec0000', stroke: '#ec0000' },
      risingColor: { fill: '#00da3c', stroke: '#00da3c' },
    },
    hAxis: { format: 'short' },
    vAxis: { format: 'short' },
  };

  if (model.title?.text) {
    options.title = model.title.text;
    options.titleTextStyle = { fontSize: 14, color: '#333' };
  }

  if (model.legend?.visible) {
    options.legend = { position: mapLegendPosition(model.legend.position) };
  } else {
    options.legend = { position: 'none' };
  }

  return options;
}

/**
 * 构建组合图选项
 */
function buildComboOptions(model: ChartModel): Record<string, any> {
  const options: Record<string, any> = buildCartesianOptions(model);

  // 为每个系列指定类型
  const seriesOpts: Record<number, any> = {};
  model.series.forEach((s, i) => {
    const typeMap: Record<string, string> = {
      bar: model.barDirection === 'bar' ? 'bars' : 'columns',
      line: 'line',
      area: 'area',
    };
    seriesOpts[i] = {
      type: typeMap[s.type] || 'line',
      color: s.color || getDefaultColor(i),
    };
    if (s.yAxisIndex === 1) {
      seriesOpts[i].targetAxisIndex = 1;
    }
  });

  options.series = seriesOpts;

  // 双坐标轴
  if (model.yAxisSecondary) {
    options.vAxes = {
      0: { format: 'short' },
      1: { format: 'short' },
    };
  }

  return options;
}

// ===== 辅助函数 =====

/** 默认颜色序列 */
const DEFAULT_COLORS = [
  '#5B9BD5', '#ED7D31', '#A5A5A5', '#FFC000',
  '#4472C4', '#70AD47', '#264478', '#9B59B6',
];

function getDefaultColor(index: number): string {
  return DEFAULT_COLORS[index % DEFAULT_COLORS.length];
}

/** 获取系列中最大数据点数 */
function getMaxDataCount(series: ChartSeriesModel[]): number {
  return Math.max(...series.map(s => s.data?.length || s.points?.length || 0), 0);
}

/** ChartModel 图例位置 → Google Charts 位置 */
function mapLegendPosition(pos: string | undefined): string {
  const map: Record<string, string> = {
    top: 'top',
    bottom: 'bottom',
    left: 'left',
    right: 'right',
  };
  return map[pos || 'bottom'] || 'bottom';
}

/** 构建图例选项 */
function buildLegendOptions(model: ChartModel): Record<string, any> {
  if (!model.legend?.visible) {
    return { position: 'none' };
  }
  // 单系列非组合图通常不显示图例
  if (model.series.length <= 1 && model.type !== 'combo') {
    return { position: 'none' };
  }
  return { position: mapLegendPosition(model.legend.position) };
}

// ===== 主转换函数 =====

/**
 * 将 ChartModel 转换为 Google Charts 配置
 *
 * @param model 图表领域模型
 * @returns GoogleChartConfig (chartType + dataTable + options)
 */
export function convertToGoogleChart(model: ChartModel): GoogleChartConfig {
  const chartType = getGoogleChartType(model);

  // 构建 DataTable
  let dataTable: (string | number | null)[][];

  switch (model.type) {
    case 'scatter':
      dataTable = buildScatterDataTable(model);
      break;
    case 'bubble':
      dataTable = buildBubbleDataTable(model);
      break;
    case 'stock':
      dataTable = buildStockDataTable(model);
      break;
    case 'combo':
      dataTable = buildComboDataTable(model);
      break;
    case 'pie':
    case 'doughnut':
      dataTable = buildCartesianDataTable(model);
      break;
    default:
      dataTable = buildCartesianDataTable(model);
  }

  // 构建 Options
  let options: Record<string, any>;

  switch (model.type) {
    case 'pie':
    case 'doughnut':
      options = buildPieOptions(model);
      break;
    case 'scatter':
      options = buildScatterOptions(model);
      break;
    case 'bubble':
      options = buildBubbleOptions(model);
      break;
    case 'stock':
      options = buildStockOptions(model);
      break;
    case 'combo':
      options = buildComboOptions(model);
      break;
    default:
      options = buildCartesianOptions(model);
  }

  return { chartType, dataTable, options };
}
