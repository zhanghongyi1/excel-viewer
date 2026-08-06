/**
 * 图表渲染器（统一入口）
 *
 * 支持两种渲染后端:
 *   1. ECharts Adapter      — 使用 ECharts 库高质量渲染（需要传入 echartsLib）
 *   2. Canvas Renderer      — 自研 Canvas 渲染引擎（零外部依赖）
 *
 * 自动选择策略:
 *   - 如果传入了 echartsLib → 使用 ECharts
 *   - 否则 → 使用 Canvas Renderer
 *
 * 两种模式都基于 ChartModel，
 * 用户可通过 renderer 选项强制指定。
 */

import type { ChartAnchor } from '../types';
import type { ChartModel } from '../chart/chart-model';
import { computeLayout } from '../chart/layout-engine';
import { convertToEChartsOption } from '../chart/echarts-converter';
import { CanvasChartRenderer } from '../chart/canvas-chart-renderer';
import { EMU_PER_PIXEL } from '../utils/ooxml';

/** 像素区域 */
interface Rect {
  left: number;
  top: number;
  width: number;
  height: number;
}

/** 单元格位置查询函数类型 */
export type PositionFn = (col: number, row: number) => Rect;

/** 渲染后端类型（默认 echarts） */
export type ChartBackend = 'echarts' | 'canvas' | 'auto';

interface ChartRendererConfig {
  /** 表格容器（图表浮层将叠加在此容器上方） */
  container: HTMLElement;
  /** ECharts 库引用（可选，不传则使用 Canvas 渲染） */
  echartsLib?: any;
  /** 渲染后端（默认 auto） */
  backend?: ChartBackend;
  /** ECharts 渲染器: 'svg' | 'canvas'（默认 svg） */
  renderer?: 'svg' | 'canvas';
  /** 主题颜色序列 */
  colorPalette?: string[];
}

export class ChartRenderer {
  private container: HTMLElement | null = null;
  private echartsLib: any = null;
  private backend: ChartBackend = 'auto';
  private renderer: 'svg' | 'canvas' = 'svg';

  // ECharts 模式
  private chartContainers: Map<string, HTMLElement> = new Map();
  private echartsInstances: Map<string, any> = new Map();

  // Canvas 模式
  private canvasRenderer: CanvasChartRenderer | null = null;

  // 实际使用的后端（auto 解析后的结果）
  private activeBackend: 'echarts' | 'canvas' = 'echarts';

  private currentCharts: ChartModel[] = [];
  private positionFn: PositionFn | null = null;
  private isDestroyed = false;

  // resize 监听
  private resizeObserver: ResizeObserver | null = null;
  private resizeTimer: ReturnType<typeof setTimeout> | null = null;

  constructor() {}

  /**
   * 初始化渲染器
   */
  init(config: ChartRendererConfig): void {
    this.container = config.container;
    this.echartsLib = config.echartsLib;
    this.backend = config.backend || 'auto';
    this.renderer = config.renderer || 'svg';

    if (!this.container) {
      throw new Error('[ChartRenderer] Container element is required');
    }

    // 解析 auto → 选择实际后端
    // 优先级: ECharts > Canvas
    if (this.backend === 'auto') {
      if (this.echartsLib) {
        this.activeBackend = 'echarts';
      } else {
        this.activeBackend = 'canvas';
      }
    } else {
      this.activeBackend = this.backend;
    }

    // 初始化对应后端
    if (this.activeBackend === 'echarts') {
      if (!this.echartsLib) {
        console.warn('[ChartRenderer] ECharts backend selected but echartsLib not provided, falling back to canvas.');
        this.activeBackend = 'canvas';
      }
    }

    if (this.activeBackend === 'canvas') {
      this.canvasRenderer = new CanvasChartRenderer({
        container: this.container,
        colorPalette: config.colorPalette,
      });
    }

    // 设置容器为相对定位
    const computedStyle = window.getComputedStyle(this.container);
    if (computedStyle.position === 'static') {
      this.container.style.position = 'relative';
    }

    // 监听容器尺寸变化
    this.setupResizeObserver();
  }

  /** 获取当前使用的渲染后端 */
  getActiveBackend(): 'echarts' | 'canvas' {
    return this.activeBackend;
  }

  private setupResizeObserver(): void {
    if (!this.container || typeof ResizeObserver === 'undefined') return;
    this.resizeObserver = new ResizeObserver(() => this.debouncedUpdatePositions());
    this.resizeObserver.observe(this.container);
  }

  private debouncedUpdatePositions(): void {
    if (this.resizeTimer) clearTimeout(this.resizeTimer);
    this.resizeTimer = setTimeout(() => this.updatePositions(), 200);
  }

  setPositionFn(fn: PositionFn): void {
    this.positionFn = fn;
  }

  /**
   * 渲染单个图表
   */
  renderChart(chart: ChartModel): boolean {
    if (!this.container || !this.positionFn || this.isDestroyed) return false;

    try {
      const area = this.calculateArea(chart.anchor);
      const minSize = 50;
      if (area.width < minSize || area.height < minSize) {
        return false;
      }

      if (this.activeBackend === 'canvas' && this.canvasRenderer) {
        return this.renderChartCanvas(chart, area);
      } else {
        return this.renderChartECharts(chart, area);
      }
    } catch (error) {
      console.error(`[ChartRenderer] Failed to render chart ${chart.id}:`, error);
      return false;
    }
  }

  // ===== ECharts 渲染 =====

  private renderChartECharts(chart: ChartModel, area: Rect): boolean {
    if (!this.container || !this.echartsLib) return false;

    // 创建或复用容器
    let containerEl = this.chartContainers.get(chart.id);

    if (containerEl) {
      containerEl.style.left = `${area.left}px`;
      containerEl.style.top = `${area.top}px`;
      containerEl.style.width = `${area.width}px`;
      containerEl.style.height = `${area.height}px`;
      containerEl.style.display = 'block';
    } else {
      containerEl = document.createElement('div');
      containerEl.className = 'excel-preview-chart';
      containerEl.dataset.chartId = chart.id;
      containerEl.style.cssText = `
        position: absolute;
        left: ${area.left}px;
        top: ${area.top}px;
        width: ${area.width}px;
        height: ${area.height}px;
        pointer-events: auto;
        z-index: 5;
        overflow: hidden;
        background: #fff;
        border: 1px solid #e0e0e0;
        box-shadow: 0 1px 4px rgba(0,0,0,0.08);
      `;
      this.container.appendChild(containerEl);
      this.chartContainers.set(chart.id, containerEl);

      const instance = this.echartsLib.init(containerEl, null, { renderer: this.renderer });
      this.echartsInstances.set(chart.id, instance);
    }

    const instance = this.echartsInstances.get(chart.id);
    if (instance) {
      const layout = computeLayout(chart, { width: area.width, height: area.height });
      const echartsOption = convertToEChartsOption(chart, layout);
      instance.setOption(echartsOption, true);
      instance.resize();
    }

    return true;
  }

  // ===== Canvas 渲染 =====

  private renderChartCanvas(chart: ChartModel, area: Rect): boolean {
    if (!this.canvasRenderer) return false;

    // Canvas renderer 内部管理 canvas 元素的创建和定位
    this.canvasRenderer.renderChart(chart, {
      left: area.left,
      top: area.top,
      width: area.width,
      height: area.height,
    });

    return true;
  }

  // ===== 批量渲染 =====

  renderAllCharts(charts: ChartModel[], positionFn?: PositionFn): void {
    if (positionFn) this.positionFn = positionFn;
    this.currentCharts = charts;

    if (this.activeBackend === 'canvas' && this.canvasRenderer) {
      // Canvas 模式
      this.canvasRenderer.renderAllCharts(charts, (chart) => this.calculateArea(chart.anchor));
    } else {
      // ECharts 模式 — 先清理不存在的
      const newIds = new Set(charts.map(c => c.id));
      for (const [id] of this.chartContainers) {
        if (!newIds.has(id)) this.removeChart(id);
      }
      for (const chart of charts) {
        this.renderChart(chart);
      }
    }
  }

  // ===== 位置计算 =====

  calculateArea(anchor: ChartAnchor): Rect {
    if (this.positionFn) {
      const startPos = this.positionFn(anchor.fromCol, anchor.fromRow);
      const endPos = this.positionFn(anchor.toCol, anchor.toRow);
      const left = startPos.left + anchor.fromColOff / EMU_PER_PIXEL;
      const top = startPos.top + anchor.fromRowOff / EMU_PER_PIXEL;
      const right = endPos.left + anchor.toColOff / EMU_PER_PIXEL;
      const bottom = endPos.top + anchor.toRowOff / EMU_PER_PIXEL;
      return {
        left: Math.round(left),
        top: Math.round(top),
        width: Math.round(right - left),
        height: Math.round(bottom - top),
      };
    }

    const DEFAULT_COL_PX = 80;
    const DEFAULT_ROW_PX = 24;
    return {
      left: anchor.fromCol * DEFAULT_COL_PX,
      top: anchor.fromRow * DEFAULT_ROW_PX,
      width: (anchor.toCol - anchor.fromCol + 1) * DEFAULT_COL_PX,
      height: (anchor.toRow - anchor.fromRow + 1) * DEFAULT_ROW_PX,
    };
  }

  /**
   * 更新所有图表位置（resize 时调用）
   */
  updatePositions(): void {
    if (this.isDestroyed || !this.currentCharts.length) return;

    if (this.activeBackend === 'canvas' && this.canvasRenderer) {
      this.canvasRenderer.updatePositions((chart) => this.calculateArea(chart.anchor));
      return;
    }

    // ECharts 模式
    for (const chart of this.currentCharts) {
      const area = this.calculateArea(chart.anchor);
      const containerEl = this.chartContainers.get(chart.id);
      if (containerEl) {
        const w = Math.max(area.width, 100);
        const h = Math.max(area.height, 100);
        containerEl.style.left = `${area.left}px`;
        containerEl.style.top = `${area.top}px`;
        containerEl.style.width = `${w}px`;
        containerEl.style.height = `${h}px`;

        const instance = this.echartsInstances.get(chart.id);
        if (instance) {
          const layout = computeLayout(chart, { width: w, height: h });
          const echartsOption = convertToEChartsOption(chart, layout);
          instance.setOption(echartsOption, true);
          instance.resize();
        }
      }
    }
  }

  // ===== 清理 =====

  removeChart(chartId: string): void {
    // ECharts 模式
    const instance = this.echartsInstances.get(chartId);
    if (instance) {
      try { instance.dispose(); } catch {}
      this.echartsInstances.delete(chartId);
    }
    const containerEl = this.chartContainers.get(chartId);
    if (containerEl && containerEl.parentNode) {
      containerEl.parentNode.removeChild(containerEl);
    }
    this.chartContainers.delete(chartId);
  }

  clearAll(): void {
    // ECharts 模式
    for (const [id] of this.chartContainers) {
      this.removeChart(id);
    }
    // Canvas 模式
    this.canvasRenderer?.clearAll();
    this.currentCharts = [];
  }

  getChartCount(): number {
    if (this.activeBackend === 'canvas' && this.canvasRenderer) {
      return this.currentCharts.length;
    }
    return this.echartsInstances.size;
  }

  destroy(): void {
    this.isDestroyed = true;
    this.clearAll();

    this.canvasRenderer?.destroy();
    this.canvasRenderer = null;

    if (this.resizeObserver) {
      this.resizeObserver.disconnect();
      this.resizeObserver = null;
    }
    if (this.resizeTimer) {
      clearTimeout(this.resizeTimer);
      this.resizeTimer = null;
    }

    this.container = null;
    this.positionFn = null;
    this.echartsLib = null;
  }
}
