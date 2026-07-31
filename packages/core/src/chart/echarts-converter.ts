/**
 * ECharts Converter — 图表模型到 ECharts 配置的转换器
 *
 * 将 ChartModel（领域模型）+ ChartLayout（布局）转换为 ECharts option。
 * 这是渲染层的适配器，使领域模型与具体渲染库解耦。
 *
 * 职责:
 *   1. 根据图表类型映射 ECharts series 类型
 *   2. 应用布局引擎计算的 grid / title / legend 位置
 *   3. 处理坐标轴、标记、颜色等样式
 *   4. 支持组合图（多系列类型 + 双坐标轴）
 *   5. 支持股价图（candlestick）、雷达图、气泡图等
 */

import type { ChartModel, ChartSeriesModel, ChartType } from './chart-model';
import type { ChartLayout } from './layout-engine';
import { layoutToGrid } from './layout-engine';

// ===== 类型映射 =====

/** Excel 图表类型 → ECharts series type */
const SERIES_TYPE_MAP: Record<ChartType, string> = {
  bar: 'bar',
  line: 'line',
  area: 'line',
  pie: 'pie',
  doughnut: 'pie',
  scatter: 'scatter',
  bubble: 'scatter',
  radar: 'radar',
  stock: 'candlestick',
  surface: 'heatmap',
  combo: 'bar', // 组合图按系列各自类型处理
};

// ===== 主转换函数 =====

/**
 * 将 ChartModel + Layout 转换为 ECharts option
 *
 * @param model 图表模型
 * @param layout 布局计算结果
 * @returns ECharts option 对象
 */
export function convertToEChartsOption(model: ChartModel, layout: ChartLayout): Record<string, any> {
  const option: Record<string, any> = {
    backgroundColor: '#ffffff',
    animation: false,
  };

  // 1. 标题
  applyTitle(option, model, layout);

  // 2. 提示框
  applyTooltip(option, model);

  // 3. 图例
  applyLegend(option, model, layout);

  // 4. 坐标轴 & grid（无坐标轴图表跳过）
  if (!layout.isAxisless) {
    applyGrid(option, layout);
    applyXAxis(option, model, layout);
    applyYAxis(option, model, layout);
  }

  // 5. 雷达图特殊处理
  if (model.type === 'radar') {
    applyRadar(option, model, layout);
  }

  // 6. 系列数据
  applySeries(option, model, layout);

  // 清理 undefined 值
  return cleanOption(option);
}

// ===== 标题 =====

function applyTitle(
  option: Record<string, any>,
  model: ChartModel,
  layout: ChartLayout
): void {
  if (!model.title || !model.title.text) return;
  option.title = {
    text: model.title.text,
    left: 'center',
    top: model.title.overlay ? layout.plotArea.top + 4 : 4,
    textStyle: {
      fontSize: 13,
      fontWeight: 'normal',
      color: '#333',
    },
    itemGap: 0,
  };
}

// ===== 提示框 =====

function applyTooltip(option: Record<string, any>, model: ChartModel): void {
  const isPie = model.type === 'pie' || model.type === 'doughnut';
  const isScatter = model.type === 'scatter' || model.type === 'bubble';
  option.tooltip = {
    trigger: isPie ? 'item' : isScatter ? 'item' : 'axis',
    axisPointer: { type: 'cross' },
    textStyle: { fontSize: 11 },
  };
}

// ===== 图例 =====

function applyLegend(
  option: Record<string, any>,
  model: ChartModel,
  _layout: ChartLayout
): void {
  const legend = model.legend;
  if (!legend || !legend.visible) return;
  if (model.series.length <= 1 && model.type !== 'combo') return;

  const posMap: Record<string, { left?: string | number; right?: string | number; top?: string | number; bottom?: number | string; orient?: string }> = {
    bottom: { bottom: 2, left: 'center' },
    top: { top: 24, left: 'center' },
    left: { left: 4, top: 'middle', orient: 'vertical' },
    right: { right: 4, top: 'middle', orient: 'vertical' },
  };

  option.legend = {
    data: model.series.map(s => s.name),
    type: 'scroll',
    textStyle: { fontSize: 10 },
    itemWidth: 12,
    itemHeight: 8,
    ...posMap[legend.position],
  };
}

// ===== Grid & 坐标轴 =====

function applyGrid(option: Record<string, any>, layout: ChartLayout): void {
  option.grid = layoutToGrid(layout);
}

function applyXAxis(
  option: Record<string, any>,
  model: ChartModel,
  layout: ChartLayout
): void {
  const categories = model.categories || [];
  const isHorizontalBar = model.barDirection === 'bar';

  if (isHorizontalBar) {
    // 水平条形图: X 轴为数值轴, Y 轴为类目轴
    option.xAxis = buildValueAxis(model.xAxis, layout, 'x');
    return;
  }

  option.xAxis = {
    type: 'category',
    data: categories.length > 0 ? categories : generateDefaultCategories(model),
    axisLabel: {
      fontSize: 9,
      interval: 0,
      rotate: layout.xAxisLabelRotate || 0,
      hideOverlap: true,
    },
    axisLine: { lineStyle: { color: '#ccc' } },
    splitLine: { show: false },
  };
}

function applyYAxis(
  option: Record<string, any>,
  model: ChartModel,
  layout: ChartLayout
): void {
  const isHorizontalBar = model.barDirection === 'bar';

  if (isHorizontalBar) {
    option.yAxis = {
      type: 'category',
      data: (model.categories || []).slice().reverse(),
      axisLabel: { fontSize: 9 },
      axisLine: { lineStyle: { color: '#ccc' } },
      splitLine: { show: false },
    };
  } else {
    option.yAxis = [buildValueAxis(model.yAxis, layout, 'y')];
  }

  // 次坐标轴（组合图）
  if (model.plotArea?.hasSecondaryAxis && model.yAxisSecondary) {
    if (Array.isArray(option.yAxis)) {
      option.yAxis.push(buildValueAxis(model.yAxisSecondary, layout, 'y', 1));
    }
  }
}

function buildValueAxis(
  axisModel: ChartModel['yAxis'],
  layout: ChartLayout,
  orient: 'x' | 'y',
  index?: number
): Record<string, any> {
  const axis: Record<string, any> = {
    type: 'value',
    axisLabel: { fontSize: 9 },
    splitLine: { lineStyle: { type: 'dashed', color: '#eee' } },
  };
  if (axisModel?.min !== undefined) axis.min = axisModel.min;
  if (axisModel?.max !== undefined) axis.max = axisModel.max;
  if (axisModel?.title) {
    axis.name = axisModel.title;
    axis.nameTextStyle = { fontSize: 10, color: '#666' };
  }
  if (index !== undefined) {
    axis.position = orient === 'y' ? 'right' : 'top';
  }
  return axis;
}

// ===== 雷达图坐标 =====

function applyRadar(
  option: Record<string, any>,
  model: ChartModel,
  _layout: ChartLayout
): void {
  const categories = model.categories || [];
  const indicators: { name: string; max: number | undefined }[] = categories.length > 0
    ? categories.map(name => ({ name, max: undefined }))
    : generateDefaultCategories(model).map(name => ({ name, max: undefined }));

  // 自动计算 max
  for (const s of model.series) {
    if (s.data) {
      for (let i = 0; i < s.data.length; i++) {
        if (indicators[i]) {
          const currentMax = indicators[i].max ?? 0;
          indicators[i].max = Math.max(currentMax, s.data[i] || 0);
        }
      }
    }
  }
  // 加 10% 余量
  for (const ind of indicators) {
    if (typeof ind.max === 'number' && ind.max > 0) {
      ind.max = Math.ceil(ind.max * 1.1);
    } else if (ind.max === undefined || ind.max === 0) {
      ind.max = 100;
    }
  }

  option.radar = {
    indicator: indicators,
    shape: model.plotArea?.radarShape || 'polygon',
    radius: '62%',
    center: ['50%', '52%'],
    axisName: { fontSize: 9, color: '#666' },
    splitArea: { areaStyle: { color: ['#fafafa', '#fff'] } },
    splitLine: { lineStyle: { color: '#ddd' } },
  };
}

// ===== 系列数据 =====

function applySeries(
  option: Record<string, any>,
  model: ChartModel,
  _layout: ChartLayout
): void {
  const series: Record<string, any>[] = [];

  for (let i = 0; i < model.series.length; i++) {
    const s = model.series[i];
    const echartsSeries = convertSeries(s, model, i);
    if (echartsSeries) series.push(echartsSeries);
  }

  if (series.length > 0) {
    option.series = series;
  }
}

/** 转换单个系列 */
function convertSeries(
  s: ChartSeriesModel,
  model: ChartModel,
  index: number
): Record<string, any> | null {
  const baseType = SERIES_TYPE_MAP[s.type] || 'bar';

  const series: Record<string, any> = {
    name: s.name,
    type: baseType,
  };

  // 坐标轴索引（组合图）
  if (s.yAxisIndex !== undefined) {
    series.yAxisIndex = s.yAxisIndex;
    series.xAxisIndex = s.yAxisIndex;
  }

  // 颜色
  if (s.color) {
    series.itemStyle = { color: s.color };
    if (baseType === 'line') {
      series.lineStyle = { color: s.color, width: s.lineWidth || 2 };
    }
  }

  // 线宽
  if (s.lineWidth) {
    if (!series.lineStyle) series.lineStyle = {};
    series.lineStyle.width = s.lineWidth;
  }

  // 平滑
  if (s.smooth) series.smooth = true;

  // 标记
  if (s.marker && s.marker !== 'none') {
    series.symbol = s.marker;
    series.symbolSize = s.markerSize || 5;
  } else if (s.marker === 'none') {
    series.symbol = 'none';
  }

  // 线型
  if (s.lineStyle) {
    const dashMap: Record<string, string> = {
      dash: 'dashed',
      dot: 'dotted',
      dashDot: 'dashed',
      solid: 'solid',
    };
    if (!series.lineStyle) series.lineStyle = {};
    series.lineStyle.type = dashMap[s.lineStyle] || 'solid';
  }

  // 按图表类型处理数据
  switch (s.type) {
    case 'bar': {
      series.data = s.data || [];
      series.barMaxWidth = 40;
      if (model.grouping === 'stacked' || model.grouping === 'percentStacked') {
        series.stack = model.grouping;
      }
      if (model.barDirection === 'bar') {
        // 水平条形图
      }
      break;
    }

    case 'line': {
      series.data = s.data || [];
      if (!model.categories || model.categories.length === 0) {
        series.showSymbol = true;
      }
      break;
    }

    case 'area': {
      series.data = s.data || [];
      series.areaStyle = { opacity: s.areaOpacity ?? 0.3 };
      break;
    }

    case 'pie':
    case 'doughnut': {
      const cats = model.categories || [];
      const pieData = cats.length > 0
        ? cats.map((name, i) => ({ name, value: s.data?.[i] || 0 }))
        : (s.data || []).map((v, i) => ({ name: `${s.name} ${i + 1}`, value: v }));
      series.data = pieData.filter((d: any) => d.value > 0);
      series.radius = s.type === 'doughnut' ? ['38%', '68%'] : '65%';
      series.center = ['50%', '52%'];
      series.label = { formatter: '{b}: {c} ({d}%)', fontSize: 10 };
      series.labelLine = { length: 8, length2: 8 };
      series.itemStyle = { borderColor: '#fff', borderWidth: 1 };
      break;
    }

    case 'scatter': {
      if (s.points && s.points.length > 0) {
        series.data = s.points.map(p => [p.x ?? 0, p.y]);
      } else if (s.data) {
        const cats = model.categories || [];
        series.data = s.data.map((v, i) => [cats[i] ? parseFloat(cats[i]) : i, v]);
      }
      series.symbolSize = 6;
      break;
    }

    case 'bubble': {
      if (s.points && s.points.length > 0) {
        series.data = s.points.map(p => [p.x ?? 0, p.y, p.size ?? 10]);
      }
      series.symbolSize = (data: any[]) => Math.sqrt(Math.abs(data[2] || 10)) * 2;
      break;
    }

    case 'radar': {
      series.data = [{ value: s.data || [], name: s.name }];
      if (model.radarStyle === 'filled') {
        series.areaStyle = { opacity: 0.3 };
      }
      break;
    }

    case 'stock': {
      if (s.ohlc) {
        const ohlc = s.ohlc;
        const count = ohlc.close.length;
        const kData: any[] = [];
        for (let i = 0; i < count; i++) {
          kData.push([
            ohlc.open[i] ?? ohlc.close[i],
            ohlc.close[i],
            ohlc.low[i],
            ohlc.high[i],
          ]);
        }
        series.data = kData;
        series.itemStyle = {
          color: '#ec0000',
          color0: '#00da3c',
          borderColor: '#8A0000',
          borderColor0: '#008F28',
        };
      }
      break;
    }

    case 'surface': {
      // 曲面图降级为热力图
      series.data = s.data || [];
      break;
    }
  }

  return series;
}

// ===== 辅助函数 =====

/** 生成默认分类标签 (1, 2, 3, ...) */
function generateDefaultCategories(model: ChartModel): string[] {
  const maxLen = Math.max(
    ...model.series.map(s => s.data?.length || s.points?.length || 0),
    0
  );
  return Array.from({ length: maxLen }, (_, i) => String(i + 1));
}

/** 清理 option 中的 undefined 值 */
function cleanOption(obj: any): any {
  return JSON.parse(JSON.stringify(obj, (_k, v) => (v === undefined ? undefined : v)));
}
