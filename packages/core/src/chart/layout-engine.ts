/**
 * Layout Engine — 图表布局引擎
 *
 * 根据图表模型 (ChartModel) 和容器像素尺寸，计算各组成部分的空间分配:
 *   - 标题 (title)
 *   - 图例 (legend)
 *   - 坐标轴 (xAxis / yAxis / yAxisSecondary)
 *   - 绘图区 (plotArea)
 *
 * 输出标准化的 ChartLayout，供 ECharts Converter 映射为
 *   grid / title.top / legend 位置 / axis 标签 等配置。
 *
 * 设计目标:
 *   1. 让图表紧凑铺满容器，减少空白
 *   2. 根据数据量自适应坐标轴标签旋转
 *   3. 组合图双坐标轴时正确分配左右空间
 *   4. 饼图/雷达图等无坐标轴图表充分利用空间
 */

import type { ChartModel, ChartSeriesModel } from './chart-model';

// ===== 布局类型 =====

/** 容器尺寸 */
export interface ChartContainerSize {
  width: number;
  height: number;
}

/** 矩形区域 */
export interface LayoutRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

/** 完整布局计算结果 */
export interface ChartLayout {
  /** 容器尺寸 */
  container: ChartContainerSize;
  /** 标题区域（null 表示无标题） */
  title: LayoutRect | null;
  /** 图例区域（null 表示无图例） */
  legend: LayoutRect | null;
  /** 绘图区（核心数据绘制区域） */
  plotArea: LayoutRect;
  /** X 轴高度（底部/顶部预留空间） */
  xAxisHeight: number;
  /** 主 Y 轴宽度（左侧预留空间） */
  yAxisWidth: number;
  /** 次 Y 轴宽度（右侧预留空间，0 表示无次轴） */
  yAxisSecondaryWidth: number;
  /** 是否需要旋转 X 轴标签 */
  xAxisLabelRotate: number;
  /** 是否为饼图/雷达图等无坐标轴图表 */
  isAxisless: boolean;
}

// ===== 常量 =====

const TITLE_HEIGHT = 22;
const LEGEND_HORIZONTAL_HEIGHT = 20;
const LEGEND_VERTICAL_WIDTH = 70;
const DEFAULT_AXIS_WIDTH = 40;
const DEFAULT_AXIS_HEIGHT = 22;
const MIN_PLOT_PADDING = 8;
const MAX_LABEL_LENGTH_BEFORE_ROTATE = 8;
const ROTATE_ANGLE = 30;
const MIN_CONTAINER_SIZE = 50;

// ===== 核心计算 =====

/**
 * 计算图表布局
 *
 * @param model 图表模型
 * @param size 容器像素尺寸
 * @returns ChartLayout 布局结果
 */
export function computeLayout(model: ChartModel, size: ChartContainerSize): ChartLayout {
  const width = Math.max(size.width, MIN_CONTAINER_SIZE);
  const height = Math.max(size.height, MIN_CONTAINER_SIZE);

  const isAxisless = isAxislessChart(model.type);

  // 1. 标题区域
  const title = computeTitleLayout(model, width);

  // 2. 图例区域
  const legend = computeLegendLayout(model, width, height);

  // 3. 计算坐标轴标签旋转
  const categories = model.categories || [];
  const dataCount = categories.length || getMaxSeriesDataCount(model.series);
  const xAxisLabelRotate = computeLabelRotate(categories, dataCount, width);

  // 4. 坐标轴占用空间
  let xAxisHeight = 0;
  let yAxisWidth = 0;
  let yAxisSecondaryWidth = 0;

  if (!isAxisless) {
    xAxisHeight = DEFAULT_AXIS_HEIGHT + (xAxisLabelRotate > 0 ? 12 : 0);
    yAxisWidth = computeYAxisWidth(model, width);
    if (model.plotArea?.hasSecondaryAxis || model.yAxisSecondary) {
      yAxisSecondaryWidth = DEFAULT_AXIS_WIDTH;
    }
  }

  // 5. 计算绘图区
  let top = MIN_PLOT_PADDING;
  let bottom = MIN_PLOT_PADDING;
  let left = MIN_PLOT_PADDING;
  let right = MIN_PLOT_PADDING;

  // 标题占用顶部空间（非 overlay 时）
  if (title && !model.title?.overlay) {
    top += title.height;
  }

  // 图例占用空间
  if (legend) {
    if (legend.top >= 0 && legend.height > 0) {
      if (model.legend?.position === 'top') {
        top += legend.height;
      } else {
        bottom += legend.height;
      }
    } else if (model.legend?.position === 'left') {
      left += legend.width;
    } else if (model.legend?.position === 'right') {
      right += legend.width;
    }
  }

  // 坐标轴占用空间
  if (!isAxisless) {
    left += yAxisWidth;
    right += yAxisSecondaryWidth;
    bottom += xAxisHeight;
  }

  // 确保不溢出
  const plotWidth = Math.max(width - left - right, 40);
  const plotHeight = Math.max(height - top - bottom, 40);

  const plotArea: LayoutRect = {
    left,
    top,
    width: plotWidth,
    height: plotHeight,
  };

  return {
    container: { width, height },
    title,
    legend,
    plotArea,
    xAxisHeight,
    yAxisWidth,
    yAxisSecondaryWidth,
    xAxisLabelRotate,
    isAxisless,
  };
}

// ===== 辅助函数 =====

/** 判断是否为无坐标轴图表（饼图/环形图/雷达图） */
function isAxislessChart(type: ChartModel['type']): boolean {
  return type === 'pie' || type === 'doughnut' || type === 'radar';
}

/** 获取系列中最大数据点数量 */
function getMaxSeriesDataCount(series: ChartSeriesModel[]): number {
  if (!series.length) return 0;
  return Math.max(
    ...series.map(s => (s.data ? s.data.length : s.points ? s.points.length : 0))
  );
}

/** 计算标题布局 */
function computeTitleLayout(model: ChartModel, containerWidth: number): LayoutRect | null {
  if (!model.title || !model.title.text) return null;

  const titleHeight = TITLE_HEIGHT;
  return {
    left: 0,
    top: 4,
    width: containerWidth,
    height: titleHeight,
  };
}

/** 计算图例布局 */
function computeLegendLayout(
  model: ChartModel,
  containerWidth: number,
  _containerHeight: number
): LayoutRect | null {
  const legend = model.legend;
  if (!legend || !legend.visible) return null;
  if (model.series.length <= 1 && model.type !== 'combo') {
    // 单系列非组合图通常不显示图例
    return null;
  }

  const position = legend.position;
  if (position === 'top' || position === 'bottom') {
    return {
      left: 0,
      top: position === 'top' ? 0 : -1,
      width: containerWidth,
      height: LEGEND_HORIZONTAL_HEIGHT,
    };
  }
  // left / right
  return {
    left: 0,
    top: 0,
    width: LEGEND_VERTICAL_WIDTH,
    height: _containerHeight,
  };
}

/**
 * 计算 X 轴标签是否需要旋转
 * 当分类标签较长或数量过多时自动旋转以避免重叠
 */
function computeLabelRotate(
  categories: string[],
  dataCount: number,
  containerWidth: number
): number {
  if (categories.length === 0) return 0;

  // 估算每个标签可用宽度
  const avgLabelWidth = containerWidth / Math.max(dataCount, 1);

  // 检查标签平均长度
  const avgLabelLength =
    categories.reduce((sum, c) => sum + String(c).length, 0) / categories.length;

  if (avgLabelLength > MAX_LABEL_LENGTH_BEFORE_ROTATE || avgLabelWidth < 40) {
    return ROTATE_ANGLE;
  }

  return 0;
}

/**
 * 估算 Y 轴所需宽度
 * 基于数据最大值的位数估算标签宽度
 */
function computeYAxisWidth(model: ChartModel, containerWidth: number): number {
  // 收集所有 Y 值
  let maxValue = 0;
  let minValue = Infinity;
  let maxDigits = 0;

  for (const s of model.series) {
    if (s.data) {
      for (const v of s.data) {
        if (typeof v === 'number' && !isNaN(v)) {
          maxValue = Math.max(maxValue, Math.abs(v));
          minValue = Math.min(minValue, v);
          const digits = formatNumberForAxis(v).length;
          maxDigits = Math.max(maxDigits, digits);
        }
      }
    } else if (s.points) {
      for (const p of s.points) {
        if (typeof p.y === 'number') {
          maxValue = Math.max(maxValue, Math.abs(p.y));
          const digits = formatNumberForAxis(p.y).length;
          maxDigits = Math.max(maxDigits, digits);
        }
      }
    }
  }

  // 如果有自定义 max/min，考虑其位数
  if (model.yAxis?.max !== undefined) {
    maxDigits = Math.max(maxDigits, formatNumberForAxis(model.yAxis.max).length);
  }
  if (model.yAxis?.min !== undefined) {
    maxDigits = Math.max(maxDigits, formatNumberForAxis(model.yAxis.min).length);
  }

  // 每位约 6px + 边距
  const estimatedWidth = Math.min(maxDigits * 6 + 12, containerWidth * 0.3);
  return Math.max(DEFAULT_AXIS_WIDTH, Math.round(estimatedWidth));
}

/** 格式化数值用于坐标轴标签估算 */
function formatNumberForAxis(v: number): string {
  if (v === 0) return '0';
  const abs = Math.abs(v);
  if (abs >= 1e9) return (v / 1e9).toFixed(1) + 'B';
  if (abs >= 1e6) return (v / 1e6).toFixed(1) + 'M';
  if (abs >= 1e4) return (v / 1e3).toFixed(1) + 'K';
  if (abs >= 100) return Math.round(v).toString();
  if (abs >= 1) return v.toFixed(1);
  return v.toFixed(2);
}

// ===== 布局辅助: 获取 ECharts grid 配置 =====

/**
 * 将布局结果转换为 ECharts grid 配置
 * 供 ECharts Converter 直接使用
 */
export function layoutToGrid(layout: ChartLayout): {
  left: number | string;
  right: number | string;
  top: number | string;
  bottom: number | string;
  containLabel: boolean;
} {
  if (layout.isAxisless) {
    return { left: '8%', right: '8%', top: '8%', bottom: '8%', containLabel: false };
  }

  return {
    left: layout.yAxisWidth,
    right: Math.max(layout.yAxisSecondaryWidth, 12),
    top: layout.plotArea.top,
    bottom: Math.max(layout.xAxisHeight + 4, 8),
    containLabel: true,
  };
}
