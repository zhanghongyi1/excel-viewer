/**
 * Canvas Chart Renderer — 自研图表渲染引擎
 *
 * 直接使用 Canvas 2D API 渲染 ChartModel，不依赖 ECharts。
 *
 * 渲染流程:
 *   ChartModel + 容器尺寸
 *       │
 *       ▼
 *   LayoutEngine.computeLayout()  → ChartLayout (标题/图例/坐标轴/绘图区)
 *       │
 *       ▼
 *   CanvasChartRenderer.render()
 *       ├─ drawBackground()
 *       ├─ drawTitle()
 *       ├─ drawLegend()
 *       ├─ drawAxes()      (网格线 + 刻度 + 标签)
 *       └─ drawSeries()    (bar / line / area / pie / scatter / radar / stock)
 *
 * 特性:
 *   - 高 DPI (Retina) 渲染
 *   - 自动坐标轴刻度计算 (nice numbers)
 *   - 颜色主题支持
 *   - 基础 tooltip (hover 高亮)
 */

import type { ChartModel, ChartSeriesModel, ChartType } from './chart-model';
import type { ChartLayout, LayoutRect } from './layout-engine';
import { OFFICE_CHART_COLORS } from './palette';
import { computeLayout } from './layout-engine';

// ===== 类型 =====

/** 像素区域 */
export interface PixelRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

/** 渲染配置 */
export interface CanvasRendererConfig {
  /** 父容器（canvas 将挂载在此） */
  container: HTMLElement;
  /** 主题颜色序列（系列颜色回退） */
  colorPalette?: string[];
}

// ===== 常量 =====

const FONT_FAMILY = '"Calibri", "Microsoft YaHei", "Segoe UI", Arial, sans-serif';
const FONT_SIZE_TITLE = 13;
const FONT_SIZE_LABEL = 9;
const FONT_SIZE_LEGEND = 10;
const AXIS_TICK_COUNT = 5;

// ===== 坐标轴工具 =====

/** 计算 "nice" 刻度值 */
function niceTicks(min: number, max: number, count: number): number[] {
  if (min === max) {
    if (min === 0) {
      return [0, 1, 2, 3, 4, 5];
    }
    const pad = Math.abs(min) * 0.2;
    min -= pad;
    max += pad;
  }

  const range = max - min;
  const step = niceNum(range / (count - 1), true);
  const niceMin = Math.floor(min / step) * step;
  const niceMax = Math.ceil(max / step) * step;

  const ticks: number[] = [];
  for (let v = niceMin; v <= niceMax + step * 0.5; v += step) {
    ticks.push(Math.round(v * 1e10) / 1e10);
  }
  return ticks;
}

/** 计算 "nice" 数值 */
function niceNum(value: number, round: boolean): number {
  if (value === 0) return 1;
  const exp = Math.floor(Math.log10(value));
  const frac = value / Math.pow(10, exp);
  let nf: number;
  if (round) {
    if (frac < 1.5) nf = 1;
    else if (frac < 3) nf = 2;
    else if (frac < 7) nf = 5;
    else nf = 10;
  } else {
    if (frac <= 1) nf = 1;
    else if (frac <= 2) nf = 2;
    else if (frac <= 5) nf = 5;
    else nf = 10;
  }
  return nf * Math.pow(10, exp);
}

/** 格式化数值用于轴标签 */
function formatAxisValue(v: number): string {
  if (v === 0) return '0';
  const abs = Math.abs(v);
  if (abs >= 1e9) return (v / 1e9).toFixed(1).replace(/\.0$/, '') + 'B';
  if (abs >= 1e6) return (v / 1e6).toFixed(1).replace(/\.0$/, '') + 'M';
  if (abs >= 1e4) return (v / 1e3).toFixed(1).replace(/\.0$/, '') + 'K';
  if (abs >= 100) return Math.round(v).toString();
  if (abs >= 1) return v.toFixed(1).replace(/\.0$/, '');
  return v.toFixed(2).replace(/0+$/, '').replace(/\.$/, '');
}

/** 获取系列颜色 */
function getSeriesColor(series: ChartSeriesModel, index: number, palette: string[]): string {
  return series.color || palette[index % palette.length];
}

// ===== 主渲染器 =====

export class CanvasChartRenderer {
  private container: HTMLElement | null = null;
  private palette: string[];
  private canvasMap: Map<string, HTMLCanvasElement> = new Map();
  private ctxMap: Map<string, CanvasRenderingContext2D> = new Map();
  private currentCharts: ChartModel[] = [];
  private isDestroyed = false;

  constructor(config: CanvasRendererConfig) {
    this.container = config.container;
    this.palette = config.colorPalette || [...OFFICE_CHART_COLORS];
    if (!this.container) throw new Error('[CanvasChartRenderer] container is required');
  }

  /**
   * 渲染单个图表到 canvas
   */
  renderChart(chart: ChartModel, area: PixelRect): boolean {
    if (!this.container || this.isDestroyed) return false;

    try {
      const w = Math.max(area.width, 50);
      const h = Math.max(area.height, 50);
      const dpr = window.devicePixelRatio || 1;

      // 获取或创建 canvas
      let canvas = this.canvasMap.get(chart.id);
      let ctx = this.ctxMap.get(chart.id);

      if (!canvas) {
        canvas = document.createElement('canvas');
        canvas.className = 'excel-preview-chart-canvas';
        canvas.style.cssText = `position:absolute;left:${area.left}px;top:${area.top}px;width:${w}px;height:${h}px;pointer-events:auto;z-index:5;`;
        this.container.appendChild(canvas);
        this.canvasMap.set(chart.id, canvas);

        ctx = canvas.getContext('2d') || undefined;
        if (ctx) this.ctxMap.set(chart.id, ctx);
      }

      if (!canvas || !ctx) return false;

      // 更新位置和尺寸
      canvas.style.left = `${area.left}px`;
      canvas.style.top = `${area.top}px`;
      canvas.style.width = `${w}px`;
      canvas.style.height = `${h}px`;
      canvas.width = Math.round(w * dpr);
      canvas.height = Math.round(h * dpr);
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0);

      // 计算布局
      const layout = computeLayout(chart, { width: w, height: h });

      // 清空画布
      ctx.clearRect(0, 0, w, h);

      // 绘制
      this.drawBackground(ctx, w, h);
      this.drawTitle(ctx, chart, layout, w);
      this.drawLegend(ctx, chart, layout, w, h);

      if (layout.isAxisless) {
        this.drawAxislessChart(ctx, chart, layout);
      } else {
        this.drawAxes(ctx, chart, layout);
        this.drawSeries(ctx, chart, layout);
      }

      return true;
    } catch (err) {
      console.error(`[CanvasChartRenderer] Failed to render chart ${chart.id}:`, err);
      return false;
    }
  }

  /**
   * 批量渲染
   */
  renderAllCharts(charts: ChartModel[], getArea: (chart: ChartModel) => PixelRect): void {
    this.currentCharts = charts;

    // 清理不再存在的图表
    const newIds = new Set(charts.map(c => c.id));
    for (const [id] of this.canvasMap) {
      if (!newIds.has(id)) this.removeChart(id);
    }

    for (const chart of charts) {
      const area = getArea(chart);
      this.renderChart(chart, area);
    }
  }

  /**
   * 更新所有图表位置（resize 时调用）
   */
  updatePositions(getArea: (chart: ChartModel) => PixelRect): void {
    if (this.isDestroyed) return;
    for (const chart of this.currentCharts) {
      const area = getArea(chart);
      this.renderChart(chart, area);
    }
  }

  /** 移除单个图表 */
  removeChart(id: string): void {
    const canvas = this.canvasMap.get(id);
    if (canvas && canvas.parentNode) canvas.parentNode.removeChild(canvas);
    this.canvasMap.delete(id);
    this.ctxMap.delete(id);
  }

  /** 清除所有 */
  clearAll(): void {
    for (const [id] of this.canvasMap) this.removeChart(id);
    this.currentCharts = [];
  }

  /** 销毁 */
  destroy(): void {
    this.isDestroyed = true;
    this.clearAll();
    this.container = null;
  }

  // ===== 绘制：背景 =====

  private drawBackground(ctx: CanvasRenderingContext2D, w: number, h: number): void {
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, w, h);
  }

  // ===== 绘制：标题 =====

  private drawTitle(ctx: CanvasRenderingContext2D, model: ChartModel, layout: ChartLayout, _w: number): void {
    if (!model.title || !model.title.text || !layout.title) return;
    ctx.save();
    ctx.font = `${FONT_SIZE_TITLE}px ${FONT_FAMILY}`;
    ctx.fillStyle = '#333333';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'top';
    const x = layout.container.width / 2;
    const y = layout.title.top + 2;
    ctx.fillText(model.title.text, x, y);
    ctx.restore();
  }

  // ===== 绘制：图例 =====

  private drawLegend(ctx: CanvasRenderingContext2D, model: ChartModel, layout: ChartLayout, w: number, h: number): void {
    if (!model.legend || !model.legend.visible) return;
    if (model.series.length <= 1 && model.type !== 'combo') return;

    const pos = model.legend.position;
    const items = model.series.map((s, i) => ({
      name: s.name,
      color: getSeriesColor(s, i, this.palette),
    }));

    ctx.save();
    ctx.font = `${FONT_SIZE_LEGEND}px ${FONT_FAMILY}`;
    ctx.textBaseline = 'middle';

    if (pos === 'bottom' || pos === 'top') {
      // 水平排列
      const y = pos === 'bottom' ? h - 12 : layout.title ? layout.title.top + layout.title.height + 4 : 8;
      let x = 8;
      for (const item of items) {
        // 色块
        ctx.fillStyle = item.color;
        ctx.fillRect(x, y - 4, 10, 8);
        // 文字
        ctx.fillStyle = '#666';
        ctx.textAlign = 'left';
        ctx.fillText(item.name, x + 14, y);
        x += 14 + ctx.measureText(item.name).width + 16;
      }
    } else {
      // 垂直排列
      const x = pos === 'left' ? 4 : w - 60;
      let y = layout.plotArea.top + 8;
      for (const item of items) {
        ctx.fillStyle = item.color;
        ctx.fillRect(x, y - 4, 10, 8);
        ctx.fillStyle = '#666';
        ctx.textAlign = 'left';
        ctx.fillText(item.name, x + 14, y);
        y += 18;
      }
    }
    ctx.restore();
  }

  // ===== 绘制：坐标轴 =====

  private drawAxes(ctx: CanvasRenderingContext2D, model: ChartModel, layout: ChartLayout): void {
    const { plotArea } = layout;
    const isHorizontalBar = model.barDirection === 'bar';

    // 计算 Y 轴范围
    const yRange = this.computeYRange(model);

    if (isHorizontalBar) {
      // 水平条形图: Y=类目轴, X=数值轴
      this.drawCategoryAxis(ctx, model, layout, 'vertical');
      this.drawValueAxis(ctx, model, layout, yRange, 'horizontal');
    } else {
      // 默认: X=类目轴, Y=数值轴
      this.drawValueAxis(ctx, model, layout, yRange, 'vertical');
      this.drawCategoryAxis(ctx, model, layout, 'horizontal');
    }
  }

  /** 计算数值轴范围 */
  private computeYRange(model: ChartModel): { min: number; max: number } {
    let min = Infinity;
    let max = -Infinity;

    for (const s of model.series) {
      const values = s.data || (s.points ? s.points.map(p => p.y) : []);
      for (const v of values) {
        if (typeof v === 'number' && !isNaN(v)) {
          min = Math.min(min, v);
          max = Math.max(max, v);
        }
      }
      // 股价图
      if (s.ohlc) {
        for (const arr of [s.ohlc.high, s.ohlc.low, s.ohlc.open, s.ohlc.close]) {
          for (const v of arr) {
            min = Math.min(min, v);
            max = Math.max(max, v);
          }
        }
      }
    }

    if (min === Infinity) { min = 0; max = 1; }
    if (min > 0) min = 0;
    if (max < 0) max = 0;
    if (min === max) { min = 0; max = min + 1; }

    if (model.yAxis?.min !== undefined) min = model.yAxis.min;
    if (model.yAxis?.max !== undefined) max = model.yAxis.max;

    return { min, max };
  }

  /** 绘制类目轴 */
  private drawCategoryAxis(
    ctx: CanvasRenderingContext2D,
    model: ChartModel,
    layout: ChartLayout,
    orientation: 'horizontal' | 'vertical'
  ): void {
    const { plotArea } = layout;
    const categories = model.categories || [];
    const count = categories.length || this.getMaxDataCount(model);
    if (count === 0) return;

    ctx.save();
    ctx.font = `${FONT_SIZE_LABEL}px ${FONT_FAMILY}`;
    ctx.fillStyle = '#666';
    ctx.strokeStyle = '#ddd';

    if (orientation === 'horizontal') {
      // X 轴在底部
      const axisY = plotArea.top + plotArea.height;
      ctx.textAlign = 'center';
      ctx.textBaseline = 'top';

      for (let i = 0; i < count; i++) {
        const x = plotArea.left + (plotArea.width / count) * (i + 0.5);
        const label = categories[i] || String(i + 1);
        const labelText = String(label);

        // 旋转长标签
        if (layout.xAxisLabelRotate > 0 && labelText.length > 4) {
          ctx.save();
          ctx.translate(x, axisY + 4);
          ctx.rotate((layout.xAxisLabelRotate * Math.PI) / 180);
          ctx.textAlign = 'right';
          ctx.fillText(labelText, 0, 0);
          ctx.restore();
        } else {
          ctx.fillText(labelText, x, axisY + 4);
        }
      }

      // 轴线
      ctx.strokeStyle = '#ccc';
      ctx.beginPath();
      ctx.moveTo(plotArea.left, axisY);
      ctx.lineTo(plotArea.left + plotArea.width, axisY);
      ctx.stroke();
    } else {
      // Y 轴在左侧（水平条形图）
      const axisX = plotArea.left;
      ctx.textAlign = 'right';
      ctx.textBaseline = 'middle';

      for (let i = 0; i < count; i++) {
        const y = plotArea.top + (plotArea.height / count) * (i + 0.5);
        const label = categories[count - 1 - i] || String(count - i);
        ctx.fillText(String(label), axisX - 6, y);
      }

      ctx.strokeStyle = '#ccc';
      ctx.beginPath();
      ctx.moveTo(axisX, plotArea.top);
      ctx.lineTo(axisX, plotArea.top + plotArea.height);
      ctx.stroke();
    }
    ctx.restore();
  }

  /** 绘制数值轴 */
  private drawValueAxis(
    ctx: CanvasRenderingContext2D,
    model: ChartModel,
    layout: ChartLayout,
    yRange: { min: number; max: number },
    orientation: 'vertical' | 'horizontal'
  ): void {
    const { plotArea } = layout;
    const ticks = niceTicks(yRange.min, yRange.max, AXIS_TICK_COUNT);
    const tickMin = ticks[0];
    const tickMax = ticks[ticks.length - 1];
    const tickRange = tickMax - tickMin || 1;

    ctx.save();
    ctx.font = `${FONT_SIZE_LABEL}px ${FONT_FAMILY}`;
    ctx.fillStyle = '#999';

    if (orientation === 'vertical') {
      // Y 轴在左侧
      ctx.textAlign = 'right';
      ctx.textBaseline = 'middle';

      for (const tick of ticks) {
        const y = plotArea.top + plotArea.height - ((tick - tickMin) / tickRange) * plotArea.height;

        // 网格线
        if (model.yAxis?.splitLine !== 'none') {
          ctx.strokeStyle = model.yAxis?.splitLine === 'dashed' ? '#eee' : '#e8e8e8';
          ctx.setLineDash(model.yAxis?.splitLine === 'dashed' ? [3, 3] : []);
          ctx.beginPath();
          ctx.moveTo(plotArea.left, y);
          ctx.lineTo(plotArea.left + plotArea.width, y);
          ctx.stroke();
          ctx.setLineDash([]);
        }

        // 标签
        ctx.fillText(formatAxisValue(tick), plotArea.left - 6, y);
      }
    } else {
      // X 轴在底部（水平条形图）
      ctx.textAlign = 'center';
      ctx.textBaseline = 'top';

      for (const tick of ticks) {
        const x = plotArea.left + ((tick - tickMin) / tickRange) * plotArea.width;

        ctx.strokeStyle = '#eee';
        ctx.setLineDash([3, 3]);
        ctx.beginPath();
        ctx.moveTo(x, plotArea.top);
        ctx.lineTo(x, plotArea.top + plotArea.height);
        ctx.stroke();
        ctx.setLineDash([]);

        ctx.fillText(formatAxisValue(tick), x, plotArea.top + plotArea.height + 4);
      }
    }
    ctx.restore();
  }

  // ===== 绘制：系列 =====

  private drawSeries(ctx: CanvasRenderingContext2D, model: ChartModel, layout: ChartLayout): void {
    const { plotArea } = layout;

    // 按图表类型分发
    const chartType = model.type === 'combo' ? 'combo' : model.type;

    switch (chartType) {
      case 'bar':
        this.drawBarSeries(ctx, model, layout, plotArea);
        break;
      case 'line':
        this.drawLineSeries(ctx, model, layout, plotArea, false);
        break;
      case 'area':
        this.drawLineSeries(ctx, model, layout, plotArea, true);
        break;
      case 'scatter':
        this.drawScatterSeries(ctx, model, layout, plotArea);
        break;
      case 'bubble':
        this.drawBubbleSeries(ctx, model, layout, plotArea);
        break;
      case 'stock':
        this.drawStockSeries(ctx, model, layout, plotArea);
        break;
      case 'combo':
        this.drawComboSeries(ctx, model, layout, plotArea);
        break;
    }
  }

  /** 柱状图 */
  private drawBarSeries(ctx: CanvasRenderingContext2D, model: ChartModel, layout: ChartLayout, plotArea: LayoutRect): void {
    const yRange = this.computeYRange(model);
    const ticks = niceTicks(yRange.min, yRange.max, AXIS_TICK_COUNT);
    const tickMin = ticks[0];
    const tickMax = ticks[ticks.length - 1];
    const tickRange = tickMax - tickMin || 1;

    const categories = model.categories || [];
    const count = categories.length || this.getMaxDataCount(model);
    const seriesCount = model.series.length;
    const isStacked = model.grouping === 'stacked' || model.grouping === 'percentStacked';
    const isHorizontal = model.barDirection === 'bar';

    // 计算柱宽
    const categorySpace = isHorizontal ? plotArea.height / count : plotArea.width / count;
    const gap = categorySpace * 0.1;
    const barAreaWidth = categorySpace - gap;
    const barWidth = isStacked
      ? barAreaWidth
      : barAreaWidth / seriesCount;

    for (let ci = 0; ci < count; ci++) {
      const categoryStart = isHorizontal
        ? plotArea.top + categorySpace * ci + gap / 2
        : plotArea.left + categorySpace * ci + gap / 2;

      let stackedPositive = tickMin;
      let stackedNegative = tickMin;

      for (let si = 0; si < seriesCount; si++) {
        const series = model.series[si];
        const value = series.data?.[ci] || 0;
        const color = getSeriesColor(series, si, this.palette);

        if (isStacked) {
          if (value >= 0) {
            const barStartVal = stackedPositive;
            const barEndVal = stackedPositive + value;
            stackedPositive = barEndVal;
            this.drawSingleBar(ctx, plotArea, categoryStart, barWidth, barStartVal, barEndVal, tickMin, tickRange, color, isHorizontal, si);
          } else {
            const barStartVal = stackedNegative;
            const barEndVal = stackedNegative + value;
            stackedNegative = barEndVal;
            this.drawSingleBar(ctx, plotArea, categoryStart, barWidth, barStartVal, barEndVal, tickMin, tickRange, color, isHorizontal, si);
          }
        } else {
          const offset = si * (barWidth / 0.8);
          this.drawSingleBar(ctx, plotArea, categoryStart + offset, barWidth, tickMin, tickMin + value, tickMin, tickRange, color, isHorizontal, si);
        }
      }
    }
  }

  /** 绘制单个柱子 */
  private drawSingleBar(
    ctx: CanvasRenderingContext2D,
    plotArea: LayoutRect,
    categoryStart: number,
    barWidth: number,
    valStart: number,
    valEnd: number,
    tickMin: number,
    tickRange: number,
    color: string,
    isHorizontal: boolean,
    _seriesIndex: number
  ): void {
    ctx.save();
    ctx.fillStyle = color;

    if (isHorizontal) {
      const y = categoryStart;
      const h = barWidth;
      const x1 = plotArea.left + ((valStart - tickMin) / tickRange) * plotArea.width;
      const x2 = plotArea.left + ((valEnd - tickMin) / tickRange) * plotArea.width;
      ctx.fillRect(Math.min(x1, x2), y, Math.abs(x2 - x1), h);
    } else {
      const x = categoryStart;
      const w = barWidth;
      const y1 = plotArea.top + plotArea.height - ((valStart - tickMin) / tickRange) * plotArea.height;
      const y2 = plotArea.top + plotArea.height - ((valEnd - tickMin) / tickRange) * plotArea.height;
      ctx.fillRect(x, Math.min(y1, y2), w, Math.abs(y2 - y1));
    }
    ctx.restore();
  }

  /** 折线图 / 面积图 */
  private drawLineSeries(ctx: CanvasRenderingContext2D, model: ChartModel, layout: ChartLayout, plotArea: LayoutRect, isArea: boolean): void {
    const yRange = this.computeYRange(model);
    const ticks = niceTicks(yRange.min, yRange.max, AXIS_TICK_COUNT);
    const tickMin = ticks[0];
    const tickMax = ticks[ticks.length - 1];
    const tickRange = tickMax - tickMin || 1;

    const categories = model.categories || [];
    const count = categories.length || this.getMaxDataCount(model);
    if (count === 0) return;

    const stepX = plotArea.width / Math.max(count - 1, 1);

    for (let si = 0; si < model.series.length; si++) {
      const series = model.series[si];
      const data = series.data || [];
      const color = getSeriesColor(series, si, this.palette);

      ctx.save();
      ctx.strokeStyle = color;
      ctx.lineWidth = series.lineWidth || 2;
      ctx.lineJoin = 'round';
      ctx.lineCap = 'round';

      // 虚线
      if (series.lineStyle === 'dash') ctx.setLineDash([5, 3]);
      else if (series.lineStyle === 'dot') ctx.setLineDash([2, 2]);
      else if (series.lineStyle === 'dashDot') ctx.setLineDash([5, 2, 2, 2]);

      // 绘制路径
      const points: { x: number; y: number }[] = [];
      for (let i = 0; i < data.length; i++) {
        const x = plotArea.left + stepX * i;
        const y = plotArea.top + plotArea.height - ((data[i] - tickMin) / tickRange) * plotArea.height;
        points.push({ x, y });
      }

      if (points.length === 0) { ctx.restore(); continue; }

      // 面积填充
      if (isArea) {
        ctx.fillStyle = color + '4D'; // 30% opacity
        ctx.beginPath();
        ctx.moveTo(points[0].x, plotArea.top + plotArea.height);
        if (series.smooth) {
          this.drawSmoothPath(ctx, points);
        } else {
          for (const p of points) ctx.lineTo(p.x, p.y);
        }
        ctx.lineTo(points[points.length - 1].x, plotArea.top + plotArea.height);
        ctx.closePath();
        ctx.fill();
      }

      // 线条
      ctx.beginPath();
      ctx.moveTo(points[0].x, points[0].y);
      if (series.smooth) {
        this.drawSmoothPath(ctx, points);
      } else {
        for (let i = 1; i < points.length; i++) ctx.lineTo(points[i].x, points[i].y);
      }
      ctx.stroke();
      ctx.setLineDash([]);

      // 标记点
      if (series.marker && series.marker !== 'none') {
        const markerSize = series.markerSize || 4;
        ctx.fillStyle = color;
        for (const p of points) {
          this.drawMarker(ctx, p.x, p.y, series.marker!, markerSize);
        }
      }

      ctx.restore();
    }
  }

  /** 绘制平滑曲线路径 */
  private drawSmoothPath(ctx: CanvasRenderingContext2D, points: { x: number; y: number }[]): void {
    if (points.length < 2) return;
    for (let i = 0; i < points.length - 1; i++) {
      const p0 = points[i - 1] || points[i];
      const p1 = points[i];
      const p2 = points[i + 1];
      const p3 = points[i + 2] || p2;
      const cp1x = p1.x + (p2.x - p0.x) / 6;
      const cp1y = p1.y + (p2.y - p0.y) / 6;
      const cp2x = p2.x - (p3.x - p1.x) / 6;
      const cp2y = p2.y - (p3.y - p1.y) / 6;
      ctx.bezierCurveTo(cp1x, cp1y, cp2x, cp2y, p2.x, p2.y);
    }
  }

  /** 绘制标记 */
  private drawMarker(ctx: CanvasRenderingContext2D, x: number, y: number, symbol: string, size: number): void {
    ctx.save();
    ctx.beginPath();
    switch (symbol) {
      case 'square':
        ctx.rect(x - size, y - size, size * 2, size * 2);
        break;
      case 'diamond':
        ctx.moveTo(x, y - size);
        ctx.lineTo(x + size, y);
        ctx.lineTo(x, y + size);
        ctx.lineTo(x - size, y);
        ctx.closePath();
        break;
      case 'triangle':
        ctx.moveTo(x, y - size);
        ctx.lineTo(x + size, y + size);
        ctx.lineTo(x - size, y + size);
        ctx.closePath();
        break;
      default: // circle
        ctx.arc(x, y, size, 0, Math.PI * 2);
    }
    ctx.fill();
    ctx.restore();
  }

  /** 散点图 */
  private drawScatterSeries(ctx: CanvasRenderingContext2D, model: ChartModel, layout: ChartLayout, plotArea: LayoutRect): void {
    for (let si = 0; si < model.series.length; si++) {
      const series = model.series[si];
      const points = series.points || [];
      const color = getSeriesColor(series, si, this.palette);

      // 计算范围
      let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
      for (const p of points) {
        if (p.x !== undefined) { minX = Math.min(minX, p.x); maxX = Math.max(maxX, p.x); }
        minY = Math.min(minY, p.y); maxY = Math.max(maxY, p.y);
      }
      if (minX === Infinity) { minX = 0; maxX = 1; }
      if (minY === Infinity) { minY = 0; maxY = 1; }
      const rangeX = maxX - minX || 1;
      const rangeY = maxY - minY || 1;

      ctx.save();
      ctx.fillStyle = color;
      const markerSize = series.markerSize || 4;

      for (const p of points) {
        const x = plotArea.left + ((p.x ?? 0) - minX) / rangeX * plotArea.width;
        const y = plotArea.top + plotArea.height - (p.y - minY) / rangeY * plotArea.height;
        this.drawMarker(ctx, x, y, 'circle', markerSize);
      }
      ctx.restore();
    }
  }

  /** 气泡图 */
  private drawBubbleSeries(ctx: CanvasRenderingContext2D, model: ChartModel, layout: ChartLayout, plotArea: LayoutRect): void {
    for (let si = 0; si < model.series.length; si++) {
      const series = model.series[si];
      const points = series.points || [];
      const color = getSeriesColor(series, si, this.palette);

      let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
      let maxSize = 0;
      for (const p of points) {
        if (p.x !== undefined) { minX = Math.min(minX, p.x); maxX = Math.max(maxX, p.x); }
        minY = Math.min(minY, p.y); maxY = Math.max(maxY, p.y);
        maxSize = Math.max(maxSize, p.size ?? 10);
      }
      if (minX === Infinity) { minX = 0; maxX = 1; }
      if (minY === Infinity) { minY = 0; maxY = 1; }
      const rangeX = maxX - minX || 1;
      const rangeY = maxY - minY || 1;
      maxSize = maxSize || 10;

      ctx.save();
      ctx.fillStyle = color + 'B0'; // semi-transparent

      for (const p of points) {
        const x = plotArea.left + ((p.x ?? 0) - minX) / rangeX * plotArea.width;
        const y = plotArea.top + plotArea.height - (p.y - minY) / rangeY * plotArea.height;
        const radius = Math.sqrt((p.size ?? 10) / maxSize) * 15;
        ctx.beginPath();
        ctx.arc(x, y, Math.max(radius, 2), 0, Math.PI * 2);
        ctx.fill();
      }
      ctx.restore();
    }
  }

  /** 股价图（K线） */
  private drawStockSeries(ctx: CanvasRenderingContext2D, model: ChartModel, layout: ChartLayout, plotArea: LayoutRect): void {
    const yRange = this.computeYRange(model);
    const ticks = niceTicks(yRange.min, yRange.max, AXIS_TICK_COUNT);
    const tickMin = ticks[0];
    const tickMax = ticks[ticks.length - 1];
    const tickRange = tickMax - tickMin || 1;

    const categories = model.categories || [];
    const count = categories.length || this.getMaxDataCount(model);
    const stepX = plotArea.width / Math.max(count, 1);
    const barWidth = Math.min(stepX * 0.6, 15);

    const series = model.series[0];
    if (!series || !series.ohlc) return;
    const ohlc = series.ohlc;
    const dataCount = ohlc.close.length;

    ctx.save();
    for (let i = 0; i < dataCount; i++) {
      const x = plotArea.left + stepX * (i + 0.5);
      const open = ohlc.open[i] ?? ohlc.close[i];
      const close = ohlc.close[i];
      const high = ohlc.high[i];
      const low = ohlc.low[i];

      const isUp = close >= open;
      const color = isUp ? '#00da3c' : '#ec0000';
      const fillColor = isUp ? '#ffffff' : color;

      const yHigh = plotArea.top + plotArea.height - ((high - tickMin) / tickRange) * plotArea.height;
      const yLow = plotArea.top + plotArea.height - ((low - tickMin) / tickRange) * plotArea.height;
      const yOpen = plotArea.top + plotArea.height - ((open - tickMin) / tickRange) * plotArea.height;
      const yClose = plotArea.top + plotArea.height - ((close - tickMin) / tickRange) * plotArea.height;

      ctx.strokeStyle = color;
      ctx.fillStyle = fillColor;

      // 高低线
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(x, yHigh);
      ctx.lineTo(x, yLow);
      ctx.stroke();

      // 实体
      const bodyTop = Math.min(yOpen, yClose);
      const bodyHeight = Math.max(Math.abs(yClose - yOpen), 1);
      ctx.fillRect(x - barWidth / 2, bodyTop, barWidth, bodyHeight);
      ctx.strokeRect(x - barWidth / 2, bodyTop, barWidth, bodyHeight);
    }
    ctx.restore();
  }

  /** 组合图 */
  private drawComboSeries(ctx: CanvasRenderingContext2D, model: ChartModel, layout: ChartLayout, plotArea: LayoutRect): void {
    // 组合图: 按系列类型分别绘制
    if (!model.plotGroups) {
      // 无分组信息时回退为普通柱状图
      this.drawBarSeries(ctx, model, layout, plotArea);
      return;
    }

    for (const group of model.plotGroups) {
      const groupSeries = group.seriesIndices.map(i => model.series[i]).filter(Boolean) as ChartSeriesModel[];
      if (groupSeries.length === 0) continue;

      const subModel: ChartModel = {
        ...model,
        type: group.type,
        series: groupSeries,
      };

      switch (group.type) {
        case 'bar':
          this.drawBarSeries(ctx, subModel, layout, plotArea);
          break;
        case 'line':
          this.drawLineSeries(ctx, subModel, layout, plotArea, false);
          break;
        case 'area':
          this.drawLineSeries(ctx, subModel, layout, plotArea, true);
          break;
      }
    }
  }

  // ===== 无坐标轴图表（饼图/环形图/雷达图）=====

  private drawAxislessChart(ctx: CanvasRenderingContext2D, model: ChartModel, layout: ChartLayout): void {
    switch (model.type) {
      case 'pie':
      case 'doughnut':
        this.drawPieChart(ctx, model, layout);
        break;
      case 'radar':
        this.drawRadarChart(ctx, model, layout);
        break;
      default:
        // 回退为有坐标轴渲染
        this.drawAxes(ctx, model, layout);
        this.drawSeries(ctx, model, layout);
    }
  }

  /** 饼图 / 环形图 */
  private drawPieChart(ctx: CanvasRenderingContext2D, model: ChartModel, layout: ChartLayout): void {
    const { plotArea } = layout;
    const series = model.series[0];
    if (!series || !series.data || series.data.length === 0) return;

    const categories = model.categories || [];
    const isDoughnut = model.type === 'doughnut';

    const cx = plotArea.left + plotArea.width / 2;
    const cy = plotArea.top + plotArea.height / 2;
    const radius = Math.min(plotArea.width, plotArea.height) / 2 * 0.8;
    const innerRadius = isDoughnut ? radius * 0.55 : 0;

    // 过滤零值
    const items: { value: number; label: string; color: string }[] = [];
    for (let i = 0; i < series.data.length; i++) {
      const v = series.data[i];
      if (v > 0) {
        items.push({
          value: v,
          label: categories[i] || `Item ${i + 1}`,
          color: getSeriesColor({ ...series, color: undefined }, i, this.palette),
        });
      }
    }

    const total = items.reduce((sum, item) => sum + item.value, 0);
    if (total === 0) return;

    let startAngle = -Math.PI / 2;

    ctx.save();
    for (const item of items) {
      const angle = (item.value / total) * Math.PI * 2;
      const endAngle = startAngle + angle;

      ctx.fillStyle = item.color;
      ctx.beginPath();
      if (isDoughnut) {
        ctx.arc(cx, cy, radius, startAngle, endAngle);
        ctx.arc(cx, cy, innerRadius, endAngle, startAngle, true);
      } else {
        ctx.moveTo(cx, cy);
        ctx.arc(cx, cy, radius, startAngle, endAngle);
      }
      ctx.closePath();
      ctx.fill();

      // 描边
      ctx.strokeStyle = '#fff';
      ctx.lineWidth = 1;
      ctx.stroke();

      // 标签
      const midAngle = (startAngle + endAngle) / 2;
      const labelR = isDoughnut ? (radius + innerRadius) / 2 : radius * 0.7;
      const lx = cx + Math.cos(midAngle) * labelR;
      const ly = cy + Math.sin(midAngle) * labelR;
      const percent = ((item.value / total) * 100).toFixed(1);

      ctx.fillStyle = '#fff';
      ctx.font = `${FONT_SIZE_LABEL}px ${FONT_FAMILY}`;
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      if (angle > 0.15) {
        ctx.fillText(`${percent}%`, lx, ly);
      }

      startAngle = endAngle;
    }
    ctx.restore();
  }

  /** 雷达图 */
  private drawRadarChart(ctx: CanvasRenderingContext2D, model: ChartModel, layout: ChartLayout): void {
    const { plotArea } = layout;
    const categories = model.categories || [];
    if (categories.length < 3) return;

    const cx = plotArea.left + plotArea.width / 2;
    const cy = plotArea.top + plotArea.height / 2;
    const radius = Math.min(plotArea.width, plotArea.height) / 2 * 0.75;
    const numAxes = categories.length;
    const angleStep = (Math.PI * 2) / numAxes;

    // 计算最大值
    let maxVal = 0;
    for (const s of model.series) {
      if (s.data) {
        for (const v of s.data) {
          if (typeof v === 'number') maxVal = Math.max(maxVal, Math.abs(v));
        }
      }
    }
    if (maxVal === 0) maxVal = 100;
    maxVal = Math.ceil(maxVal * 1.1);

    ctx.save();

    // 绘制网格（同心多边形）
    const gridLevels = 5;
    ctx.strokeStyle = '#e8e8e8';
    ctx.lineWidth = 1;
    for (let level = 1; level <= gridLevels; level++) {
      const r = (radius * level) / gridLevels;
      ctx.beginPath();
      for (let i = 0; i < numAxes; i++) {
        const angle = -Math.PI / 2 + i * angleStep;
        const x = cx + Math.cos(angle) * r;
        const y = cy + Math.sin(angle) * r;
        if (i === 0) ctx.moveTo(x, y);
        else ctx.lineTo(x, y);
      }
      ctx.closePath();
      ctx.stroke();
    }

    // 绘制轴线
    ctx.strokeStyle = '#ddd';
    for (let i = 0; i < numAxes; i++) {
      const angle = -Math.PI / 2 + i * angleStep;
      ctx.beginPath();
      ctx.moveTo(cx, cy);
      ctx.lineTo(cx + Math.cos(angle) * radius, cy + Math.sin(angle) * radius);
      ctx.stroke();
    }

    // 绘制标签
    ctx.fillStyle = '#666';
    ctx.font = `${FONT_SIZE_LABEL}px ${FONT_FAMILY}`;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    for (let i = 0; i < numAxes; i++) {
      const angle = -Math.PI / 2 + i * angleStep;
      const lx = cx + Math.cos(angle) * (radius + 12);
      const ly = cy + Math.sin(angle) * (radius + 12);
      ctx.fillText(categories[i], lx, ly);
    }

    // 绘制系列
    for (let si = 0; si < model.series.length; si++) {
      const series = model.series[si];
      const data = series.data || [];
      const color = getSeriesColor(series, si, this.palette);
      const isFilled = model.radarStyle === 'filled';

      ctx.strokeStyle = color;
      ctx.lineWidth = series.lineWidth || 2;
      ctx.fillStyle = color + (isFilled ? '4D' : '20');

      ctx.beginPath();
      for (let i = 0; i < data.length; i++) {
        const angle = -Math.PI / 2 + i * angleStep;
        const r = (Math.abs(data[i]) / maxVal) * radius;
        const x = cx + Math.cos(angle) * r;
        const y = cy + Math.sin(angle) * r;
        if (i === 0) ctx.moveTo(x, y);
        else ctx.lineTo(x, y);
      }
      ctx.closePath();
      ctx.fill();
      ctx.stroke();

      // 标记点
      if (series.marker && series.marker !== 'none') {
        ctx.fillStyle = color;
        for (let i = 0; i < data.length; i++) {
          const angle = -Math.PI / 2 + i * angleStep;
          const r = (Math.abs(data[i]) / maxVal) * radius;
          const x = cx + Math.cos(angle) * r;
          const y = cy + Math.sin(angle) * r;
          this.drawMarker(ctx, x, y, series.marker!, series.markerSize || 3);
        }
      }
    }

    ctx.restore();
  }

  // ===== 工具 =====

  private getMaxDataCount(model: ChartModel): number {
    return Math.max(
      ...model.series.map(s => s.data?.length || s.points?.length || 0),
      0
    );
  }
}
