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
import { OFFICE_CHART_COLORS } from './palette';

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
  waterfall: 'bar',
  funnel: 'funnel',
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
    color: [...OFFICE_CHART_COLORS],
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
  const isFunnel = model.type === 'funnel';
  const isScatter = model.type === 'scatter' || model.type === 'bubble';
  option.tooltip = {
    trigger: isPie || isFunnel ? 'item' : isScatter ? 'item' : 'axis',
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

  if (model.type === 'scatter' || model.type === 'bubble') {
    option.xAxis = buildValueAxis(model.xAxis, layout, 'x');
    return;
  }

  if (isHorizontalBar) {
    // 水平条形图: X 轴为数值轴, Y 轴为类目轴
    option.xAxis = buildValueAxis(model.xAxis, layout, 'x');
    return;
  }

  option.xAxis = {
    type: 'category',
    data: categories.length > 0 ? categories : generateDefaultCategories(model),
    axisLabel: {
      fontSize: model.xAxis?.labelFontSize || 9,
      interval: 0,
      rotate: model.xAxis?.labelRotate ?? layout.xAxisLabelRotate ?? 0,
      hideOverlap: true,
    },
    axisLine: { lineStyle: { color: '#ccc' } },
    splitLine: { show: false },
    inverse: model.xAxis?.inverse || false,
    show: model.xAxis?.visible !== false,
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
      data: model.categories || [],
      axisLabel: { fontSize: 9 },
      axisLine: { lineStyle: { color: '#ccc' } },
      splitLine: { show: false },
      inverse: model.yAxis?.inverse ?? true,
    };
  } else {
    option.yAxis = [buildValueAxis(model.yAxis, layout, 'y')];
    if (model.grouping === 'percentStacked') {
      option.yAxis[0].min = 0;
      option.yAxis[0].max = 100;
      option.yAxis[0].axisLabel.formatter = '{value}%';
    }
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
    splitLine: { lineStyle: { type: 'solid', color: '#d9d9d9', width: 1 } },
  };
  if (axisModel?.min !== undefined) axis.min = axisModel.min;
  if (axisModel?.max !== undefined) axis.max = axisModel.max;
  if (axisModel?.inverse) axis.inverse = true;
  if (axisModel?.visible === false) axis.show = false;
  if (axisModel?.labelRotate !== undefined) axis.axisLabel.rotate = axisModel.labelRotate;
  if (axisModel?.labelFontSize !== undefined) axis.axisLabel.fontSize = axisModel.labelFontSize;
  if (axisModel?.splitLine === 'none') axis.splitLine.show = false;
  else if (axisModel?.splitLine) axis.splitLine.lineStyle.type = axisModel.splitLine;
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
    if (s.type === 'waterfall') {
      series.push(...buildWaterfallSeries(s, i));
      continue;
    }
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

  if (s.dataLabels && Object.values(s.dataLabels).some(Boolean)) {
    series.label = {
      show: true,
      position: s.dataLabels.position || 'top',
      color: '#595959',
      fontSize: 12,
      formatter: (params: any) => {
        const parts: string[] = [];
        if (s.dataLabels?.showSeriesName) parts.push(s.name);
        if (s.dataLabels?.showCategoryName) parts.push(String(params.name ?? ''));
        if (s.dataLabels?.showValue) {
          const raw = Array.isArray(params.value) ? params.value[params.value.length - 1] : params.value;
          parts.push(typeof raw === 'number' ? String(Number(raw.toPrecision(12))) : String(raw ?? ''));
        }
        if (s.dataLabels?.showPercent && params.percent !== undefined) parts.push(`${params.percent}%`);
        return parts.join(' ');
      },
    };
  }

  // 按图表类型处理数据
  switch (s.type) {
    case 'bar': {
      series.data = model.grouping === 'percentStacked'
        ? normalizePercentSeries(model, index)
        : (s.data || []);
      if (s.barWidth !== undefined) series.barWidth = s.barWidth;
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
      series.data = model.grouping === 'percentStacked'
        ? normalizePercentSeries(model, index)
        : (s.data || []);
      series.areaStyle = { opacity: s.areaOpacity ?? 0.3 };
      if (model.grouping === 'stacked' || model.grouping === 'percentStacked') {
        series.stack = 'area';
      }
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
      if (!s.dataLabels) {
        series.label = { show: false };
      } else {
        series.label = {
          ...series.label,
          show: true,
          formatter: s.dataLabels.showPercent ? '{b}: {d}%' : '{b}: {c}',
          fontSize: 10,
        };
        series.labelLine = { show: true, length: 8, length2: 8 };
      }
      series.itemStyle = { borderColor: '#fff', borderWidth: 1 };
      break;
    }

    case 'waterfall': {
      const values = s.data || [];
      const base: number[] = [];
      const increases: Array<number | string> = [];
      const decreases: Array<number | string> = [];
      let cumulative = 0;
      for (const value of values) {
        if (value >= 0) {
          base.push(cumulative);
          increases.push(value);
          decreases.push('-');
        } else {
          base.push(cumulative + value);
          increases.push('-');
          decreases.push(Math.abs(value));
        }
        cumulative += value;
      }
      return {
        name: s.name,
        type: 'bar',
        data: increases,
        stack: `waterfall-${index}`,
        itemStyle: { color: '#70AD47' },
        emphasis: { focus: 'series' },
        label: series.label,
        __waterfallBase: base,
        __waterfallDecrease: decreases,
      };
    }

    case 'funnel': {
      const cats = model.categories || [];
      series.data = (s.data || []).map((value, i) => ({
        name: cats[i] || `${s.name} ${i + 1}`,
        value,
      }));
      series.sort = 'descending';
      series.left = '10%';
      series.width = '80%';
      series.top = 36;
      series.bottom = 12;
      series.gap = 2;
      series.label = series.label || { show: true, position: 'inside', formatter: '{b}: {c}' };
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

function normalizePercentSeries(model: ChartModel, seriesIndex: number): number[] {
  const source = model.series[seriesIndex]?.data || [];
  return source.map((value, pointIndex) => {
    const total = model.series.reduce((sum, item) => sum + Math.abs(item.data?.[pointIndex] || 0), 0);
    return total === 0 ? 0 : value / total * 100;
  });
}

function buildWaterfallSeries(s: ChartSeriesModel, index: number): Record<string, any>[] {
  const values = s.data || [];
  const base: number[] = [];
  const increases: Array<number | string> = [];
  const decreases: Array<number | string> = [];
  let cumulative = 0;
  for (const value of values) {
    if (value >= 0) {
      base.push(cumulative);
      increases.push(value);
      decreases.push('-');
    } else {
      base.push(cumulative + value);
      increases.push('-');
      decreases.push(Math.abs(value));
    }
    cumulative += value;
  }
  const stack = `waterfall-${index}`;
  return [
    { type: 'bar', stack, data: base, itemStyle: { color: 'transparent' }, silent: true, tooltip: { show: false } },
    { name: `${s.name} +`, type: 'bar', stack, data: increases, itemStyle: { color: '#70AD47' }, label: { show: true, position: 'top' } },
    { name: `${s.name} -`, type: 'bar', stack, data: decreases, itemStyle: { color: '#C00000' }, label: { show: true, position: 'bottom' } },
  ];
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
  if (Array.isArray(obj)) return obj.map(cleanOption);
  if (obj && typeof obj === 'object') {
    const cleaned: Record<string, any> = {};
    for (const [key, value] of Object.entries(obj)) {
      if (value !== undefined) cleaned[key] = cleanOption(value);
    }
    return cleaned;
  }
  return obj;
}
