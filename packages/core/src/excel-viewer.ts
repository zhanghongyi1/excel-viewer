import type { ExcelSource, ExcelViewerOptions, ParsedWorkbook, ParsedSheet } from './types';
import type { ChartModel } from './chart/chart-model';
import ExcelJS from 'exceljs';
import { loadData } from './loader';
import { parseExcel } from './parser/excel-parser';
import { parseCharts } from './parser/chart-parser';
import { parsePivotTables } from './parser/pivot-parser';
import { TableRenderer } from './renderer/table-renderer';
import { ChartRenderer } from './renderer/chart-renderer';
import { ImageRenderer } from './renderer/image-renderer';

export class ExcelViewer {
  private options: Required<Omit<ExcelViewerOptions, 'target' | 'src' | 'onRendered' | 'onError' | 'onSheetChange' | 'echarts' | 'chartBackend' | 'echartsRenderer'>> & Pick<ExcelViewerOptions, 'target' | 'src' | 'onRendered' | 'onError' | 'onSheetChange' | 'echarts' | 'chartBackend' | 'echartsRenderer'>;
  private rootElement: HTMLElement | null = null;
  private wrapperEl: HTMLDivElement | null = null;
  private loadingEl: HTMLDivElement | null = null;
  private errorEl: HTMLDivElement | null = null;

  private tableRenderer: TableRenderer | null = null;
  private chartRenderer: ChartRenderer | null = null;
  private imageRenderer: ImageRenderer | null = null;
  private workbook: ParsedWorkbook | null = null;
  private allCharts: ChartModel[] = [];
  private allImages: any[] = [];
  private fileData: ArrayBuffer | null = null;
  private isDestroyed = false;
  private renderVersion = 0;

  constructor(options: ExcelViewerOptions = {}) {
    this.options = {
      width: '100%',
      height: '100%',
      showToolbar: true,
      extraColCount: 5,
      extraRowCount: 20,
      initialZoom: 100,
      parsePivotTables: false,
      // ECharts 是默认主渲染后端，仅在明确指定 canvas 时切换。
      chartBackend: 'echarts',
      // 默认使用 SVG 渲染
      echartsRenderer: 'svg',
      ...options,
    };

    if (this.options.target) {
      this.mount(this.options.target);
    }
    if (this.options.src) {
      this.render(this.options.src);
    }
  }

  mount(target: HTMLElement | string): void {
    if (typeof target === 'string') {
      const el = document.querySelector<HTMLElement>(target);
      if (!el) throw new Error(`[ExcelViewer] Target not found: "${target}"`);
      this.rootElement = el;
    } else {
      this.rootElement = target;
    }
    this.buildLayout();
    this.renderBlankSheet();
  }

  async render(src?: ExcelSource): Promise<void> {
    const source = src || this.options.src;
    if (!source) return;
    if (!this.rootElement) {
      throw new Error('[ExcelViewer] Call mount() first or pass target in constructor.');
    }

    const renderVersion = ++this.renderVersion;
    this.showLoading();
    this.hideError();

    // 先销毁已有图表和图片实例
    this.chartRenderer?.clearAll();
    this.imageRenderer?.clearAll();

    try {
      const fileData = await loadData(source as any);
      if (this.isDestroyed || renderVersion !== this.renderVersion) return;
      this.fileData = fileData;
      this.workbook = await parseExcel(fileData);
      if (this.isDestroyed || renderVersion !== this.renderVersion) return;

      if (!this.workbook || this.workbook.sheets.length === 0) {
        throw new Error('[ExcelViewer] No readable worksheets found.');
      }

      const rawWb = new ExcelJS.Workbook();
      if (this.isDestroyed || renderVersion !== this.renderVersion) return;
      await rawWb.xlsx.load(fileData);
      const { charts, images } = await parseCharts(fileData, rawWb);
      if (this.isDestroyed || renderVersion !== this.renderVersion) return;
      this.allCharts = charts;
      this.allImages = images;

      if (this.allCharts.length > 0 && this.options.chartBackend !== 'canvas' && !this.options.echarts) {
        this.options.echarts = await import('echarts');
        if (this.isDestroyed || renderVersion !== this.renderVersion) return;
      }

      if (this.options.parsePivotTables) {
        try {
          const pivotTables = await parsePivotTables(fileData);
          if (this.isDestroyed || renderVersion !== this.renderVersion) return;
          if (pivotTables.length > 0) {
            this.workbook.pivotTables = pivotTables;
          }
        } catch {
          // 数据透视表不影响基础预览
        }
      }

      const tableRenderer = this.ensureTableRenderer();

      // 计算图表和图片需要的最大行/列数（图表区域之后多加载 5 行缓冲）
      const ROW_BUFFER = 5;
      const COL_BUFFER = 3;
      let maxChartRow = 0;
      let maxChartCol = 0;
      for (const chart of this.allCharts) {
        maxChartRow = Math.max(maxChartRow, chart.anchor.toRow + 1 + ROW_BUFFER);
        maxChartCol = Math.max(maxChartCol, chart.anchor.toCol + 1 + COL_BUFFER);
      }
      for (const img of this.allImages) {
        maxChartRow = Math.max(maxChartRow, img.anchor.toRow + 1 + ROW_BUFFER);
        maxChartCol = Math.max(maxChartCol, img.anchor.toCol + 1 + COL_BUFFER);
      }

      // 设置最小行/列数以容纳图表
      // 每次加载都覆盖上次文件的扩展尺寸，避免旧文件留下大量空白行列。
      tableRenderer.setMinRowCount(maxChartRow);
      tableRenderer.setMinColCount(maxChartCol);

      // loadData 会重建表格（scrollEl.innerHTML = ''），清除所有旧 DOM
      tableRenderer.loadData(this.workbook);

      // 清除图表和图片渲染器的旧引用
      this.chartRenderer?.clearAll();
      this.imageRenderer?.clearAll();

      if (this.allCharts.length > 0) {
        if (!this.chartRenderer) {
          this.chartRenderer = new ChartRenderer();
          this.chartRenderer.init({
            container: tableRenderer.getScrollContainer()!,
            echarts: this.options.echarts,
            chartBackend: this.options.chartBackend || 'auto',
            echartsRenderer: this.options.echartsRenderer,
          });
        }
        this.chartRenderer.setPositionFn((col, row) => tableRenderer.getCellPosition(col, row));
        this.renderCurrentSheetCharts();
      }

      if (this.allImages.length > 0) {
        if (!this.imageRenderer) {
          this.imageRenderer = new ImageRenderer();
          this.imageRenderer.init({
            container: tableRenderer.getScrollContainer()!,
          });
        }
        this.imageRenderer.setPositionFn((col, row) => tableRenderer.getCellPosition(col, row));
        this.renderCurrentSheetImages();
      }

      if (this.isDestroyed || renderVersion !== this.renderVersion) return;
      this.hideLoading();
      this.options.onRendered?.();
    } catch (err: any) {
      if (this.isDestroyed || renderVersion !== this.renderVersion) return;
      this.hideLoading();
      this.showError(err.message || 'Render failed');
      this.options.onError?.(err);
    }
  }

  setSheet(indexOrName: number | string): void {
    if (!this.workbook) return;
    let idx = -1;
    if (typeof indexOrName === 'number') {
      if (indexOrName >= 0 && indexOrName < this.workbook.sheets.length) idx = indexOrName;
    } else {
      idx = this.workbook.sheets.findIndex(s => s.name === indexOrName);
    }
    if (idx >= 0) this.tableRenderer?.switchSheet(idx);
  }

  getWorkbook(): ParsedWorkbook | null { return this.workbook; }

  destroy(): void {
    this.isDestroyed = true;
    this.renderVersion++;
    this.chartRenderer?.destroy();
    this.imageRenderer?.destroy();
    this.chartRenderer = null;
    this.imageRenderer = null;
    this.clearRenderers();
    if (this.wrapperEl?.parentNode) this.wrapperEl.parentNode.removeChild(this.wrapperEl);
    this.rootElement = null;
    this.wrapperEl = null;
    this.workbook = null;
    this.fileData = null;
  }

  private buildLayout(): void {
    if (!this.rootElement) return;
    this.rootElement.innerHTML = '';
    this.rootElement.style.position = 'relative';

    this.wrapperEl = document.createElement('div');
    this.wrapperEl.style.cssText = `position:relative;display:flex;flex-direction:column;width:${this.options.width};height:${this.options.height};overflow:hidden;background:#fff;border:1px solid #e0e0e0;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;`;

    this.loadingEl = document.createElement('div');
    this.loadingEl.style.cssText = 'position:absolute;inset:0;background:rgba(255,255,255,0.85);display:none;align-items:center;justify-content:center;font-size:14px;z-index:100;';
    this.loadingEl.innerHTML = '<div style="display:flex;align-items:center;gap:8px;color:#1890ff"><span class="excel-spinner"></span><span>Loading Excel...</span></div>';

    this.errorEl = document.createElement('div');
    this.errorEl.style.cssText = 'position:absolute;inset:0;background:#fff;display:none;flex-direction:column;align-items:center;justify-content:center;padding:24px;color:#d93025;font-size:14px;text-align:center;z-index:90;';

    this.wrapperEl.appendChild(this.loadingEl);
    this.wrapperEl.appendChild(this.errorEl);
    this.rootElement.appendChild(this.wrapperEl);

    if (!document.getElementById('excel-viewer-style')) {
      const s = document.createElement('style');
      s.id = 'excel-viewer-style';
      s.textContent = '.excel-spinner{display:inline-block;width:18px;height:18px;border:2px solid #1890ff;border-top-color:transparent;border-radius:50%;animation:excel-spin .8s linear infinite}@keyframes excel-spin{to{transform:rotate(360deg)}}';
      document.head.appendChild(s);
    }
  }

  private renderCurrentSheetCharts(): void {
    if (!this.chartRenderer || !this.tableRenderer || !this.workbook) return;
    const idx = this.tableRenderer.getCurrentSheetIndex();
    const sheet = this.workbook.sheets[idx];
    if (!sheet) return;
    // 只渲染当前 Sheet 的图表
    const sheetCharts = this.allCharts.filter(c => c.sheetIndex === idx);
    const posFn = (col: number, row: number) => this.tableRenderer!.getCellPosition(col, row);
    this.chartRenderer.renderAllCharts(sheetCharts, posFn);
  }

  private renderCurrentSheetImages(): void {
    if (!this.imageRenderer || !this.tableRenderer || !this.workbook) return;
    const idx = this.tableRenderer.getCurrentSheetIndex();
    const sheet = this.workbook.sheets[idx];
    if (!sheet) return;
    // 只渲染当前 Sheet 的图片
    const sheetImages = this.allImages.filter(img => img.sheetIndex === idx);
    const posFn = (col: number, row: number) => this.tableRenderer!.getCellPosition(col, row);
    this.imageRenderer.renderAllImages(sheetImages, posFn);
  }

  private clearRenderers(): void {
    this.chartRenderer?.clearAll();
    this.imageRenderer?.clearAll();
    this.tableRenderer?.destroy();
    this.tableRenderer = null;
    this.chartRenderer = null;
    this.imageRenderer = null;
    if (this.wrapperEl) {
      const overlays = this.wrapperEl.querySelectorAll('.excel-preview-chart, .excel-preview-chart-box, .excel-preview-image');
      overlays.forEach(el => el.remove());
    }
  }

  private ensureTableRenderer(): TableRenderer {
    if (this.tableRenderer) return this.tableRenderer;

    const tableRenderer = new TableRenderer();
    tableRenderer.init({
      container: this.wrapperEl!,
      options: {
        echarts: this.options.echarts,
        showToolbar: this.options.showToolbar,
        extraColCount: this.options.extraColCount,
        extraRowCount: this.options.extraRowCount,
        initialZoom: this.options.initialZoom,
      },
    });
    tableRenderer.onSwitchSheet((idx) => {
      this.chartRenderer?.clearAll();
      this.imageRenderer?.clearAll();
      this.renderCurrentSheetCharts();
      this.renderCurrentSheetImages();
      const sheet = this.workbook?.sheets[idx];
      if (sheet) this.options.onSheetChange?.(sheet.name, idx);
    });
    tableRenderer.onDimensionsChange(() => {
      this.chartRenderer?.updatePositions();
      this.imageRenderer?.updatePositions();
    });
    this.tableRenderer = tableRenderer;
    return tableRenderer;
  }

  /** 在用户选择文件前提供与 Excel 一致的空白工作区。 */
  private renderBlankSheet(): void {
    const tableRenderer = this.ensureTableRenderer();
    tableRenderer.setMinRowCount(50);
    tableRenderer.setMinColCount(26);
    tableRenderer.loadData({
      sheets: [{
        name: 'Sheet1',
        id: 'blank-sheet',
        rows: [],
        merges: [],
        colWidths: [],
        rowHeights: [],
      }],
    });
  }

  private showLoading(): void { if (this.loadingEl) this.loadingEl.style.display = 'flex'; }
  private hideLoading(): void { if (this.loadingEl) this.loadingEl.style.display = 'none'; }
  private showError(msg: string): void {
    if (this.errorEl) {
      this.errorEl.replaceChildren();
      const title = document.createElement('div');
      title.style.cssText = 'font-weight:bold;margin-bottom:8px';
      title.textContent = 'Render Error';
      const detail = document.createElement('div');
      detail.textContent = msg;
      this.errorEl.append(title, detail);
      this.errorEl.style.display = 'flex';
    }
  }
  private hideError(): void { if (this.errorEl) this.errorEl.style.display = 'none'; }
}
