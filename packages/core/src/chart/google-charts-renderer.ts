/**
 * Google Charts Renderer — Google Visualization API 渲染器
 *
 * 基于 google-charts npm 包（ES6 封装），管理 Google Charts 实例生命周期:
 *   - GoogleCharts.load() 异步加载 Google Charts 库
 *   - 创建图表实例并绘制
 *   - resize 时重绘
 *   - 销毁释放资源
 *
 * 加载流程:
 *   ensureLoaded() → google-charts 包自动加载 loader.js + corechart
 *   renderChart() → convertToGoogleChart() → new google.visualization[ChartType]() → draw()
 */

import { GoogleCharts } from 'google-charts';
import type { ChartModel } from './chart-model';
import { convertToGoogleChart } from './google-charts-converter';

// ===== 类型 =====

export interface PixelRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface GoogleRendererConfig {
  container: HTMLElement;
  /** 已加载的 google 对象（可选，不传则自动加载） */
  google?: any;
}

// ===== 渲染器 =====

export class GoogleChartsRenderer {
  private container: HTMLElement | null = null;
  private google: any = null;
  private googleReady: boolean = false;
  private loadPromise: Promise<void> | null = null;

  private chartContainers: Map<string, HTMLElement> = new Map();
  private chartInstances: Map<string, any> = new Map();
  private currentCharts: ChartModel[] = [];
  private isDestroyed = false;

  constructor(config: GoogleRendererConfig) {
    this.container = config.container;
    if (!this.container) throw new Error('[GoogleChartsRenderer] container is required');

    // 使用传入的 google 对象，或从 window 获取
    this.google = config.google || (typeof window !== 'undefined' ? (window as any).google : null);
  }

  /**
   * 确保 Google Charts 已加载
   *
   * 通过 google-charts npm 包加载，自动处理 loader.js + corechart 包
   */
  async ensureLoaded(): Promise<void> {
    if (this.googleReady && this.google) return;

    // 如果已有 google 对象且包含 visualization API
    if (this.google && this.google.visualization?.DataTable) {
      this.googleReady = true;
      return;
    }

    // 尝试从 window 获取（用户可能已在外部加载）
    if (typeof window !== 'undefined') {
      const g = (window as any).google;
      if (g?.visualization?.DataTable) {
        this.google = g;
        this.googleReady = true;
        return;
      }
    }

    // 通过 google-charts npm 包加载
    await new Promise<void>((resolve, reject) => {
      GoogleCharts.load(() => {
        this.google = GoogleCharts.api;
        if (!this.google?.visualization?.DataTable) {
          reject(new Error('[GoogleChartsRenderer] google-charts loaded but visualization API missing'));
          return;
        }
        this.googleReady = true;
        resolve();
      });
    });
  }

  /**
   * 渲染单个图表
   */
  async renderChart(chart: ChartModel, area: PixelRect): Promise<boolean> {
    if (!this.container || this.isDestroyed) return false;

    try {
      // 确保 Google Charts 已加载
      if (!this.googleReady) {
        await this.ensureLoaded();
      }

      if (!this.google || !this.google.visualization) {
        console.error('[GoogleChartsRenderer] Google Charts not loaded');
        return false;
      }

      const w = Math.max(area.width, 50);
      const h = Math.max(area.height, 50);

      // 创建或复用容器 div
      let containerEl = this.chartContainers.get(chart.id);
      if (!containerEl) {
        containerEl = document.createElement('div');
        containerEl.className = 'excel-preview-chart-google';
        containerEl.dataset.chartId = chart.id;
        this.container.appendChild(containerEl);
        this.chartContainers.set(chart.id, containerEl);
      }

      // 更新位置和尺寸
      containerEl.style.cssText = `
        position: absolute;
        left: ${area.left}px;
        top: ${area.top}px;
        width: ${w}px;
        height: ${h}px;
        pointer-events: auto;
        z-index: 5;
        overflow: hidden;
        background: #fff;
        border: 1px solid #e0e0e0;
        box-shadow: 0 1px 4px rgba(0,0,0,0.08);
      `;

      // 转换为 Google Charts 配置
      const config = convertToGoogleChart(chart);

      // 构建 DataTable
      const dataTable = this.google.visualization.arrayToDataTable(config.dataTable);

      // 合并尺寸到选项
      const options = {
        ...config.options,
        width: w,
        height: h,
      };

      // 创建或复用图表实例
      let instance = this.chartInstances.get(chart.id);
      const ChartConstructor = this.google.visualization[config.chartType];

      if (!ChartConstructor) {
        console.error(`[GoogleChartsRenderer] Unknown chart type: ${config.chartType}`);
        return false;
      }

      if (!instance) {
        instance = new ChartConstructor(containerEl);
        this.chartInstances.set(chart.id, instance);
      }

      // 绘制
      instance.draw(dataTable, options);

      return true;
    } catch (err) {
      console.error(`[GoogleChartsRenderer] Failed to render chart ${chart.id}:`, err);
      return false;
    }
  }

  /**
   * 批量渲染
   */
  async renderAllCharts(charts: ChartModel[], getArea: (chart: ChartModel) => PixelRect): Promise<void> {
    this.currentCharts = charts;

    // 清理不再存在的图表
    const newIds = new Set(charts.map(c => c.id));
    for (const [id] of this.chartContainers) {
      if (!newIds.has(id)) this.removeChart(id);
    }

    // 确保 Google Charts 已加载
    if (!this.googleReady) {
      await this.ensureLoaded();
    }

    // 渲染每个图表
    for (const chart of charts) {
      const area = getArea(chart);
      await this.renderChart(chart, area);
    }
  }

  /**
   * 更新所有图表位置（resize 时调用）
   */
  async updatePositions(getArea: (chart: ChartModel) => PixelRect): Promise<void> {
    if (this.isDestroyed) return;

    for (const chart of this.currentCharts) {
      const area = getArea(chart);
      // 更新容器尺寸
      const containerEl = this.chartContainers.get(chart.id);
      if (containerEl) {
        const w = Math.max(area.width, 50);
        const h = Math.max(area.height, 50);
        containerEl.style.left = `${area.left}px`;
        containerEl.style.top = `${area.top}px`;
        containerEl.style.width = `${w}px`;
        containerEl.style.height = `${h}px`;
      }

      // 重新绘制
      await this.renderChart(chart, area);
    }
  }

  /** 移除单个图表 */
  removeChart(id: string): void {
    const instance = this.chartInstances.get(id);
    if (instance) {
      // Google Charts 没有 dispose 方法，清除容器即可
      this.chartInstances.delete(id);
    }
    const containerEl = this.chartContainers.get(id);
    if (containerEl && containerEl.parentNode) {
      containerEl.parentNode.removeChild(containerEl);
    }
    this.chartContainers.delete(id);
  }

  /** 清除所有 */
  clearAll(): void {
    for (const [id] of this.chartContainers) {
      this.removeChart(id);
    }
    this.currentCharts = [];
  }

  /** 销毁 */
  destroy(): void {
    this.isDestroyed = true;
    this.clearAll();
    this.container = null;
    this.google = null;
    this.googleReady = false;
  }
}
