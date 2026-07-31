/**
 * 图片渲染器
 *
 * 在表格上方的浮层中渲染图片:
 * 1. 根据图片锚点计算像素坐标
 * 2. 创建绝对定位的 img 元素
 * 3. 支持 resize 自适应
 */

import type { ParsedImage, ChartAnchor } from '../types';
import type { PositionFn } from './chart-renderer';

/** 像素区域 */
interface Rect {
  left: number;
  top: number;
  width: number;
  height: number;
}

interface ImageRendererConfig {
  /** 表格容器（图片浮层将叠加在此容器上方） */
  container: HTMLElement;
}

export class ImageRenderer {
  private container: HTMLElement | null = null;
  private imageContainers: Map<string, HTMLElement> = new Map();
  private currentImages: ParsedImage[] = [];
  private positionFn: PositionFn | null = null;
  private isDestroyed = false;

  // resize 监听
  private resizeObserver: ResizeObserver | null = null;
  private resizeTimer: ReturnType<typeof setTimeout> | null = null;
  private scrollHandler: (() => void) | null = null;

  constructor() {}

  /**
   * 初始化渲染器
   */
  init(config: ImageRendererConfig): void {
    this.container = config.container;

    if (!this.container) {
      throw new Error('[ImageRenderer] Container element is required');
    }

    // 设置容器为相对定位（作为浮层的定位参考）
    const computedStyle = window.getComputedStyle(this.container);
    if (computedStyle.position === 'static') {
      this.container.style.position = 'relative';
    }

    // 监听容器尺寸变化
    this.setupResizeObserver();
    // 监听容器滚动，动态创建可见区域的图片
    this.setupScrollListener();
  }

  /**
   * 设置 ResizeObserver 监听容器变化
   */
  private setupResizeObserver(): void {
    if (!this.container || typeof ResizeObserver === 'undefined') return;

    this.resizeObserver = new ResizeObserver(() => {
      this.debouncedUpdatePositions();
    });

    this.resizeObserver.observe(this.container);
  }

  /**
   * 监听容器滚动，动态渲染可见图片
   */
  private setupScrollListener(): void {
    if (!this.container) return;
    this.scrollHandler = () => {
      this.renderVisibleImages();
    };
    this.container.addEventListener('scroll', this.scrollHandler, { passive: true });
  }

  /**
   * 渲染当前可见区域内的图片
   */
  private renderVisibleImages(): void {
    if (this.isDestroyed || !this.currentImages.length || !this.positionFn) return;
    for (const image of this.currentImages) {
      const containerEl = this.imageContainers.get(image.id);
      if (!containerEl) {
        this.renderImage(image);
      }
    }
  }

  /**
   * 防抖更新位置
   */
  private debouncedUpdatePositions(): void {
    if (this.resizeTimer) {
      clearTimeout(this.resizeTimer);
    }
    this.resizeTimer = setTimeout(() => {
      this.updatePositions();
    }, 200);
  }

  /**
   * 设置单元格位置查询函数
   */
  setPositionFn(fn: PositionFn): void {
    this.positionFn = fn;
  }

  /**
   * 根据锚点计算像素区域
   */
  calculateArea(anchor: ChartAnchor): Rect {
    if (this.positionFn) {
      const startPos = this.positionFn(anchor.fromCol, anchor.fromRow);
      const endPos = this.positionFn(anchor.toCol, anchor.toRow);

      const left = startPos.left + anchor.fromColOff / 9525;
      const top = startPos.top + anchor.fromRowOff / 9525;
      const right = endPos.left + anchor.toColOff / 9525;
      const bottom = endPos.top + anchor.toRowOff / 9525;

      return {
        left: Math.round(left),
        top: Math.round(top),
        width: Math.round(right - left),
        height: Math.round(bottom - top),
      };
    }

    // 回退：使用默认估算值
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
   * 渲染单个图片
   */
  renderImage(image: ParsedImage): boolean {
    if (!this.container || !this.positionFn || this.isDestroyed) {
      return false;
    }

    try {
      // 计算像素区域
      const area = this.calculateArea(image.anchor);

      // 创建或复用容器
      let containerEl = this.imageContainers.get(image.id);

      if (containerEl) {
        // 更新位置和大小
        containerEl.style.left = `${area.left}px`;
        containerEl.style.top = `${area.top}px`;
        containerEl.style.width = `${area.width}px`;
        containerEl.style.height = `${area.height}px`;
        containerEl.style.display = 'block';
      } else {
        // 创建新容器
        containerEl = document.createElement('div');
        containerEl.className = 'excel-preview-image';
        containerEl.dataset.imageId = image.id;
        containerEl.style.cssText = `
          position: absolute;
          left: ${area.left}px;
          top: ${area.top}px;
          width: ${area.width}px;
          height: ${area.height}px;
          pointer-events: auto;
          z-index: 0;
          overflow: hidden;
          display: flex;
          align-items: center;
          justify-content: center;
        `;

        // 加载中占位
        const loadingEl = document.createElement('div');
        loadingEl.className = 'excel-image-loading';
        loadingEl.style.cssText = `
          position: absolute;
          inset: 0;
          display: flex;
          align-items: center;
          justify-content: center;
          background: #f5f5f5;
          color: #999;
          font-size: 12px;
          z-index: 1;
        `;
        loadingEl.textContent = 'Loading...';
        containerEl.appendChild(loadingEl);

        // 创建 img 元素
        const img = document.createElement('img');
        img.alt = image.name || 'Excel Image';
        img.style.cssText = `
          max-width: 100%;
          max-height: 100%;
          object-fit: contain;
          display: none;
        `;
        img.decoding = 'async';

        // 加载完成
        img.onload = () => {
          loadingEl.style.display = 'none';
          img.style.display = 'block';
        };
        // 加载失败
        img.onerror = () => {
          loadingEl.textContent = 'Image load failed';
          loadingEl.style.color = '#d93025';
        };

        img.src = image.imageData;
        containerEl.appendChild(img);
        this.container.appendChild(containerEl);
        this.imageContainers.set(image.id, containerEl);
      }

      return true;
    } catch (error) {
      console.error(`[ImageRenderer] Failed to render image ${image.id}:`, error);
      return false;
    }
  }

  /**
   * 批量渲染所有图片
   */
  renderAllImages(images: ParsedImage[], positionFn?: PositionFn): void {
    if (positionFn) {
      this.positionFn = positionFn;
    }

    this.currentImages = images;

    // 先清理不再存在的图片
    const newIds = new Set(images.map((i) => i.id));
    for (const [id] of this.imageContainers) {
      if (!newIds.has(id)) {
        this.removeImage(id);
      }
    }

    // 渲染每个图片
    for (const image of images) {
      this.renderImage(image);
    }
  }

  /**
   * 更新所有图片位置（scroll/resize 时调用）
   */
  updatePositions(): void {
    if (this.isDestroyed || !this.currentImages.length) return;

    for (const image of this.currentImages) {
      const area = this.calculateArea(image.anchor);
      const containerEl = this.imageContainers.get(image.id);

      if (containerEl) {
        containerEl.style.left = `${area.left}px`;
        containerEl.style.top = `${area.top}px`;
        containerEl.style.width = `${Math.max(area.width, 20)}px`;
        containerEl.style.height = `${Math.max(area.height, 20)}px`;
      }
    }
  }

  /**
   * 移除单个图片
   */
  removeImage(imageId: string): void {
    const containerEl = this.imageContainers.get(imageId);

    if (containerEl && containerEl.parentNode) {
      containerEl.parentNode.removeChild(containerEl);
    }
    this.imageContainers.delete(imageId);
  }

  /**
   * 清除所有图片
   */
  clearAll(): void {
    for (const [id] of this.imageContainers) {
      this.removeImage(id);
    }
    this.currentImages = [];
  }

  /**
   * 获取当前图片数量
   */
  getImageCount(): number {
    return this.imageContainers.size;
  }

  /**
   * 销毁实例，释放资源
   */
  destroy(): void {
    this.isDestroyed = true;

    // 清除所有图片
    this.clearAll();

    // 断开 ResizeObserver
    if (this.resizeObserver) {
      this.resizeObserver.disconnect();
      this.resizeObserver = null;
    }

    if (this.resizeTimer) {
      clearTimeout(this.resizeTimer);
      this.resizeTimer = null;
    }

    // 移除滚动监听
    if (this.scrollHandler && this.container) {
      this.container.removeEventListener('scroll', this.scrollHandler);
      this.scrollHandler = null;
    }

    this.container = null;
    this.positionFn = null;
  }
}
