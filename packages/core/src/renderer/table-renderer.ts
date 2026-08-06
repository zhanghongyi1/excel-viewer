import type { ParsedWorkbook, ParsedSheet, ParsedCell, ParsedPivotTable, PivotCacheRecord, ViewerOptions, CellStyle, CfRule, CfColorScale, CfDataBar } from '../types';
import { colLetterToNumber } from '../utils/ooxml';

const DEFAULT_COL_WIDTH = 80;
const DEFAULT_ROW_HEIGHT = 20; // Excel 默认行高更紧凑
const HEADER_COL_WIDTH = 40; // 行号列更窄
const HEADER_ROW_HEIGHT = 20;
const EXCEL_FONT = 'Calibri, "Microsoft YaHei", Arial, sans-serif';

interface TableRendererConfig {
  container: HTMLElement;
  options?: ViewerOptions;
}

function getColWidth(sheet: ParsedSheet, c: number, overrides?: Map<number, number>): number {
  if (overrides?.has(c)) return overrides.get(c)!;
  return Math.max(Math.round(sheet.colWidths[c] || DEFAULT_COL_WIDTH), 30);
}

function getRowHeight(sheet: ParsedSheet, r: number, overrides?: Map<number, number>): number {
  if (overrides?.has(r)) return overrides.get(r)!;
  return Math.max(Math.round(sheet.rowHeights[r] || DEFAULT_ROW_HEIGHT), 18);
}

function colNumToLetter(n: number): string {
  let s = '';
  while (n >= 0) {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

export class TableRenderer {
  private container: HTMLElement | null = null;
  private rootEl: HTMLElement | null = null;
  private scrollEl: HTMLElement | null = null;
  private sheetBar: HTMLElement | null = null;
  private workbookData: ParsedWorkbook | null = null;
  private currentSheetIndex = 0;
  private isDestroyed = false;
  private selectedRow = -1;
  private selectedCol = -1;
  private selectedMode: 'cell' | 'row' | 'col' | 'all' = 'cell';

  private onSwitchSheetCallback?: (index: number) => void;
  private onCellSelectedCallback?: (cell: any, rowIndex: number, colIndex: number) => void;

  private options: ViewerOptions = {};

  // 最小行数和列数（用于容纳图表）
  private minRowCount = 0;
  private minColCount = 0;

  // 冻结窗格
  private freezeRowCount = 1;
  private freezeColCount = 1;

  // 列宽/行高拖拽调整
  private colWidthOverrides: Map<number, number> = new Map();
  private rowHeightOverrides: Map<number, number> = new Map();
  private resizeState: {
    type: 'col' | 'row';
    index: number;
    startPos: number;
    startSize: number;
  } | null = null;

  // 数据透视表 UI 将在后续主版本清理，当前保留公开方法兼容性。
  private pivotTableMode = false;
  private pivotTableEl: HTMLDivElement | null = null;

  private buildStickyPositions(sheet: ParsedSheet): { top: number[]; left: number[] } {
    this.freezeRowCount = Math.max(1, sheet.freezePane?.ySplit ?? 1);
    this.freezeColCount = Math.max(1, sheet.freezePane?.xSplit ?? 1);

    const top: number[] = [];
    let cumHeight = 0;
    for (let r = 0; r < this.freezeRowCount; r++) {
      top[r] = cumHeight;
      if (r === 0) {
        cumHeight += HEADER_ROW_HEIGHT;
      } else {
        cumHeight += getRowHeight(sheet, r - 1, this.rowHeightOverrides);
      }
    }

    const left: number[] = [];
    let cumWidth = 0;
    for (let c = 0; c < this.freezeColCount; c++) {
      left[c] = cumWidth;
      if (c === 0) {
        cumWidth += HEADER_COL_WIDTH;
      } else {
        cumWidth += getColWidth(sheet, c - 1, this.colWidthOverrides);
      }
    }

    return { top, left };
  }

  private stickyStyle(r: number, c: number, freezeTop: number[], freezeLeft: number[]): string {
    const top = r < freezeTop.length ? freezeTop[r] : undefined;
    const left = c < freezeLeft.length ? freezeLeft[c] : undefined;

    // z-index 层级需高于图表浮层(z-index:5)和图片浮层(z-index:0)，
    // 确保冻结表头始终覆盖在图表/图片之上
    if (top !== undefined && left !== undefined) {
      return `position:sticky;top:${top}px;left:${left}px;z-index:13;`;
    }
    if (top !== undefined) {
      return `position:sticky;top:${top}px;z-index:12;`;
    }
    if (left !== undefined) {
      return `position:sticky;left:${left}px;z-index:11;`;
    }
    return '';
  }

  init(config: TableRendererConfig): void {
    this.container = config.container;
    this.options = config.options || {};
    this.minColCount = this.options.minColLength || 0;
    this.minRowCount = this.options.minRowLength || 0;
    if (!this.container) throw new Error('[TableRenderer] Container element is required');
    this.createDOM();
  }

  private createDOM(): void {
    if (!this.container) return;
    this.rootEl = document.createElement('div');
    this.rootEl.className = 'excel-preview-table';
    this.rootEl.style.cssText = 'display:flex;flex-direction:column;height:100%;overflow:hidden;font-family:' + EXCEL_FONT + ';font-size:14px;';

    this.scrollEl = document.createElement('div');
    this.scrollEl.className = 'excel-preview-scroll';
    this.scrollEl.style.cssText = 'flex:1;overflow:auto;position:relative;outline:none;';

    this.sheetBar = document.createElement('div');
    this.sheetBar.className = 'excel-preview-sheet-bar';
    this.sheetBar.style.cssText = 'display:flex;background:#f0f0f0;border-top:1px solid #c0c0c0;overflow-x:auto;';

    this.rootEl.appendChild(this.scrollEl);
    if (this.options.showToolbar !== false) this.rootEl.appendChild(this.sheetBar);
    this.container.appendChild(this.rootEl);

    // 注入选中高亮样式
    if (!document.getElementById('excel-table-style')) {
      const s = document.createElement('style');
      s.id = 'excel-table-style';
      s.textContent = `
        .excel-selected-row { background:#d3e3fd !important; }
        .excel-selected-col { background:#d3e3fd !important; }
        .excel-selected-cell { background:#c7d9f5 !important; outline:2px solid #1a73e8; outline-offset:-1px; }
        .excel-has-comment { position:relative; }
        .excel-has-comment::after {
          content:''; position:absolute; top:0; right:0;
          width:0; height:0;
          border-style:solid;
          border-width:0 6px 6px 0;
          border-color:transparent #e84c3d transparent transparent;
        }
      `;
      document.head.appendChild(s);
    }
  }

  loadData(workbook: ParsedWorkbook): void {
    this.workbookData = workbook;
    this.currentSheetIndex = 0;
    this.selectedRow = -1;
    this.selectedCol = -1;
    this.renderSheet(0);
    this.renderSheetBar();
    this.onSwitchSheetCallback?.(0);
  }

  private renderSheet(index: number): void {
    const scrollEl = this.scrollEl;
    if (!scrollEl || !this.workbookData) return;
    const sheet = this.workbookData.sheets[index];
    if (!sheet) return;

    scrollEl.innerHTML = '';
    const table = document.createElement('table');
    table.style.cssText = 'border-collapse:collapse;table-layout:fixed;min-width:100%;font-size:13px;';
    table.cellSpacing = '0';
    scrollEl.appendChild(table);

    let maxRow = Math.max(sheet.rows.length, this.minRowCount);
    let maxCol = 0;
    for (const row of sheet.rows) {
      if (row && row.length > maxCol) maxCol = row.length;
    }
    maxCol = Math.max(maxCol, this.minColCount, 1);
    maxRow = Math.max(maxRow, 1);

    // Column headers
    const { top: freezeTop, left: freezeLeft } = this.buildStickyPositions(sheet);

    const headerTr = document.createElement('tr');
    const cornerTh = document.createElement('th');
    cornerTh.style.cssText = `
      ${this.stickyStyle(0, 0, freezeTop, freezeLeft)}
      width:${HEADER_COL_WIDTH}px;min-width:${HEADER_COL_WIDTH}px;max-width:${HEADER_COL_WIDTH}px;
      height:${HEADER_ROW_HEIGHT}px;
      background:#f3f3f3;border-bottom:1px solid #d0d0d0;border-right:1px solid #d0d0d0;
      padding:0;
    `;
    cornerTh.addEventListener('click', () => this.selectAllCells());
    headerTr.appendChild(cornerTh);

    for (let c = 0; c < maxCol; c++) {
      if (sheet.hiddenCols?.has(c)) continue;
      const th = document.createElement('th');
      th.dataset.col = String(c);
      const colWidth = getColWidth(sheet, c, this.colWidthOverrides);
      th.textContent = colNumToLetter(c);
      th.style.cssText = `
        ${this.stickyStyle(0, c + 1, freezeTop, freezeLeft)}
        width:${colWidth}px;min-width:${colWidth}px;max-width:${colWidth}px;
        height:${HEADER_ROW_HEIGHT}px;
        background:#f3f3f3;border-bottom:1px solid #d0d0d0;border-right:1px solid #e0e0e0;
        padding:0 4px;font-weight:400;font-size:11px;color:#333;
        text-align:center;user-select:none;overflow:hidden;white-space:nowrap;
      `;

      const colResize = document.createElement('div');
      colResize.className = 'excel-col-resize';
      colResize.style.cssText = 'position:absolute;top:0;right:-2px;width:4px;height:100%;cursor:col-resize;z-index:10;';
      colResize.dataset.colIndex = String(c);
      colResize.addEventListener('mousedown', (e) => this.startColResize(e, c));
      th.appendChild(colResize);
      th.addEventListener('click', () => this.selectColumn(c));

      headerTr.appendChild(th);
    }
    table.appendChild(headerTr);

    // Data rows
    for (let r = 0; r < maxRow; r++) {
      if (sheet.hiddenRows?.has(r)) continue;
      const tr = document.createElement('tr');

      const rowH = getRowHeight(sheet, r, this.rowHeightOverrides);
      tr.style.height = rowH + 'px';

      const rowTh = document.createElement('th');
      rowTh.dataset.row = String(r);
      rowTh.textContent = String(r + 1);
      rowTh.style.cssText = `
        ${this.stickyStyle(r + 1, 0, freezeTop, freezeLeft)}
        width:${HEADER_COL_WIDTH}px;min-width:${HEADER_COL_WIDTH}px;max-width:${HEADER_COL_WIDTH}px;
        height:${rowH}px;
        background:#f3f3f3;border-right:1px solid #d0d0d0;border-bottom:1px solid #e0e0e0;
        padding:0 4px;font-weight:400;font-size:11px;color:#333;
        text-align:center;user-select:none;overflow:hidden;white-space:nowrap;
      `;

      const rowResize = document.createElement('div');
      rowResize.className = 'excel-row-resize';
      rowResize.style.cssText = 'position:absolute;left:0;bottom:-2px;height:4px;width:100%;cursor:row-resize;z-index:10;';
      rowResize.addEventListener('mousedown', (e) => this.startRowResize(e, r));
      // 使用 absolute 子元素需要父元素有定位上下文，但不能覆盖 sticky
      // 通过包装层解决：在 th 内部加一个 relative 容器
      const resizeWrap = document.createElement('div');
      resizeWrap.style.cssText = 'position:absolute;left:0;right:0;top:0;bottom:0;pointer-events:none;';
      resizeWrap.appendChild(rowResize);
      rowResize.style.pointerEvents = 'auto';
      rowTh.appendChild(resizeWrap);
      rowTh.addEventListener('click', () => this.selectRow(r));
      tr.appendChild(rowTh);

      for (let c = 0; c < maxCol; c++) {
        if (sheet.hiddenCols?.has(c)) continue;
        const cell = sheet.rows[r]?.[c];

        // 检查是否在合并区域内但不是起始单元格（跳过）
        const isInMergeButNotStart = sheet.merges.some(
          m =>
            r >= m.startRow &&
            r <= m.endRow &&
            c >= m.startCol &&
            c <= m.endCol &&
            !(m.startRow === r && m.startCol === c)
        );

        if (isInMergeButNotStart) {
          continue; // 跳过被合并的单元格
        }

        const colWidth = getColWidth(sheet, c, this.colWidthOverrides);
        const td = document.createElement('td');
        td.dataset.row = String(r);
        td.dataset.col = String(c);

        const sticky = this.stickyStyle(r + 1, c + 1, freezeTop, freezeLeft);
        const posStyle = sticky ? '' : 'position:relative;';
        td.style.cssText = `
          ${sticky}${posStyle}
          width:${colWidth}px;min-width:${colWidth}px;max-width:${colWidth}px;
          height:${rowH}px;
          border-right:1px solid #e0e0e0;border-bottom:1px solid #e0e0e0;
          padding:0 3px 0 2px;font-size:11px;line-height:1.2;
          overflow:hidden;text-overflow:ellipsis;white-space:nowrap;
          cursor:cell;
        `;

        if (cell) {
          if (cell.richTextRuns) {
            this.applyCellStyle(td, cell.style, cell.type);
            td.innerHTML = '';
            for (const run of cell.richTextRuns) {
              const span = document.createElement('span');
              span.textContent = run.text;
              if (run.font) {
                if (run.font.bold) span.style.fontWeight = 'bold';
                if (run.font.italic) span.style.fontStyle = 'italic';
                if (run.font.size && run.font.size > 0) span.style.fontSize = run.font.size + 'px';
                if (run.font.color) span.style.color = run.font.color;
                if (run.font.name) span.style.fontFamily = run.font.name;
                if (run.font.underline) span.style.textDecoration = 'underline';
                if (run.font.strike) span.style.textDecoration = 'line-through';
              }
              td.appendChild(span);
            }
          } else {
            td.textContent = cell.text;
            this.applyCellStyle(td, cell.style, cell.type);
          }

          if (cell.merge) {
            td.rowSpan = cell.merge.endRow - cell.merge.startRow + 1;
            td.colSpan = cell.merge.endCol - cell.merge.startCol + 1;
          }

          if (cell.comment) {
            td.classList.add('excel-has-comment');
            td.title = cell.comment.text;
          }
        }

        const rowIdx = r;
        const colIdx = c;
        td.addEventListener('click', (e) => {
          this.selectCell(td, rowIdx, colIdx, cell);
        });

        tr.appendChild(td);
      }
      table.appendChild(tr);
    }

    this.applyConditionalFormatting(table, sheet);
  }

  private applyConditionalFormatting(table: HTMLTableElement, sheet: ParsedSheet): void {
    const cfList = sheet.conditionalFormatting;
    if (!cfList) return;

    for (const cf of cfList) {
      try {
        const { startRow, startCol, endRow, endCol } = this.parseRange(cf.ref, sheet);
        const numericValues: number[] = [];
        for (let r = startRow; r <= endRow; r++) {
          for (let c = startCol; c <= endCol; c++) {
            const value = this.getParsedCell(sheet, r, c)?.value;
            if (typeof value === 'number' && Number.isFinite(value)) numericValues.push(value);
          }
        }
        numericValues.sort((a, b) => a - b);

        for (let r = startRow; r <= endRow; r++) {
          for (let c = startCol; c <= endCol; c++) {
            const cellEl = table.querySelector<HTMLElement>(`td[data-row="${r}"][data-col="${c}"]`);
            if (!cellEl) continue;

            const cell = this.getParsedCell(sheet, r, c);
            if (!cell) continue;

            for (const rule of cf.rules) {
              this.applyCfRule(cellEl, cell, rule, numericValues);
            }
          }
        }
      } catch {
        // ignore invalid ref
      }
    }
  }

  private parseRange(ref: string, sheet: ParsedSheet): { startRow: number; startCol: number; endRow: number; endCol: number } {
    const match = ref.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    if (!match) {
      return { startRow: 0, startCol: 0, endRow: sheet.rows.length - 1, endCol: 10 };
    }
    const startCol = colLetterToNumber(match[1]);
    const startRow = parseInt(match[2], 10) - 1;
    const endCol = colLetterToNumber(match[3]);
    const endRow = parseInt(match[4], 10) - 1;
    return { startRow, startCol, endRow, endCol };
  }

  private getParsedCell(sheet: ParsedSheet, row: number, col: number): ParsedCell | undefined {
    const r = sheet.rows[row];
    if (!r) return undefined;
    return r[col];
  }

  private applyCfRule(cellEl: HTMLElement, cell: ParsedCell, rule: CfRule, rangeValues: number[]): void {
    if (rule.type === 'colorScale' && rule.colorScale) {
      this.applyColorScale(cellEl, cell, rule.colorScale, rangeValues);
      return;
    }

    if (rule.type === 'dataBar' && rule.dataBar) {
      this.applyDataBar(cellEl, cell, rule.dataBar, rangeValues);
      return;
    }

    if (rule.type === 'iconSet' && rule.iconSet) {
      // Icon set rendering skipped for now (requires icon font/SVG)
      return;
    }

    // For cellIs and expression rules, check condition and apply fill/font
    const matched = rule.type === 'expression'
      ? this.evaluateExpression(cell, rule)
      : this.evaluateCellIs(cell, rule);

    if (matched) {
      if (rule.fill) {
        cellEl.style.backgroundColor = rule.fill;
      }
      if (rule.font) {
        if (rule.font.color) cellEl.style.color = rule.font.color;
        if (rule.font.bold) cellEl.style.fontWeight = 'bold';
        if (rule.font.italic) cellEl.style.fontStyle = 'italic';
        if (rule.font.strike) cellEl.style.textDecoration = 'line-through';
        if (rule.font.underline) cellEl.style.textDecoration = 'underline';
      }
    }
  }

  private evaluateCellIs(cell: ParsedCell, rule: CfRule): boolean {
    const val = cell.value;
    if (typeof val !== 'number') return false;
    const target = parseFloat(rule.formula?.[0] ?? '0');
    if (isNaN(target)) return false;

    switch (rule.operator) {
      case 'greaterThan': return val > target;
      case 'lessThan': return val < target;
      case 'between': {
        const target2 = parseFloat(rule.formula?.[1] ?? '0');
        return !isNaN(target2) && val >= target && val <= target2;
      }
      case 'notBetween': {
        const target2 = parseFloat(rule.formula?.[1] ?? '0');
        return !isNaN(target2) && (val < target || val > target2);
      }
      case 'equal': return val === target;
      case 'notEqual': return val !== target;
      case 'greaterThanOrEqual': return val >= target;
      case 'lessThanOrEqual': return val <= target;
      default: return false;
    }
  }

  private evaluateExpression(cell: ParsedCell, rule: CfRule): boolean {
    // Deliberately support only a safe comparison subset. Workbook formulas are untrusted input.
    const formula = (rule.formula?.[0] ?? '').trim().replace(/^=/, '');
    const match = formula.match(/^(?:\$?[A-Z]+\$?\d+)\s*(>=|<=|<>|=|>|<)\s*(?:"([^"]*)"|(-?\d+(?:\.\d+)?))$/i);
    if (!match) return false;
    const expected: string | number = match[2] !== undefined ? match[2] : Number(match[3]);
    const actual = cell.value;
    switch (match[1]) {
      case '>': return actual > expected;
      case '<': return actual < expected;
      case '>=': return actual >= expected;
      case '<=': return actual <= expected;
      case '=': return actual === expected;
      case '<>': return actual !== expected;
      default: return false;
    }
  }

  private applyColorScale(cellEl: HTMLElement, cell: ParsedCell, cs: CfColorScale, rangeValues: number[]): void {
    const val = cell.value;
    if (typeof val !== 'number') return;

    if (rangeValues.length === 0) return;
    const thresholds = cs.cfvo.map(v => {
      if (v.type === 'min') return rangeValues[0];
      if (v.type === 'max') return rangeValues[rangeValues.length - 1];
      if (v.type === 'num') return Number(v.value ?? 0);
      if (v.type === 'percent') return rangeValues[0] + (rangeValues[rangeValues.length - 1] - rangeValues[0]) * Number(v.value ?? 50) / 100;
      if (v.type === 'percentile') return rangeValues[Math.round((rangeValues.length - 1) * Number(v.value ?? 50) / 100)];
      return Number(v.value ?? 0);
    });

    const colors = cs.colors;

    // Find where value falls between thresholds
    if (thresholds.length >= 2 && colors.length >= 2) {
      const min = thresholds[0];
      const max = thresholds[thresholds.length - 1];
      const ratio = max === min ? 0 : Math.max(0, Math.min(1, (val - min) / (max - min)));

      if (thresholds.length === 2 || colors.length === 2) {
        cellEl.style.backgroundColor = this.lerpColor(colors[0], colors[1], ratio);
      } else if (colors.length >= 3) {
        // 3-color scale
        if (ratio <= 0.5) {
          cellEl.style.backgroundColor = this.lerpColor(colors[0], colors[1], ratio * 2);
        } else {
          cellEl.style.backgroundColor = this.lerpColor(colors[1], colors[2], (ratio - 0.5) * 2);
        }
      }
    }
  }

  private applyDataBar(cellEl: HTMLElement, cell: ParsedCell, db: CfDataBar, rangeValues: number[]): void {
    const val = cell.value;
    if (typeof val !== 'number') return;

    if (rangeValues.length === 0) return;
    const resolve = (index: number, fallback: number): number => {
      const threshold = db.cfvo[index];
      if (!threshold) return fallback;
      if (threshold.type === 'min') return rangeValues[0];
      if (threshold.type === 'max') return rangeValues[rangeValues.length - 1];
      if (threshold.type === 'percent') return rangeValues[0] + (rangeValues[rangeValues.length - 1] - rangeValues[0]) * Number(threshold.value ?? 0) / 100;
      if (threshold.type === 'percentile') return rangeValues[Math.round((rangeValues.length - 1) * Number(threshold.value ?? 0) / 100)];
      return Number(threshold.value ?? fallback);
    };
    const min = resolve(0, rangeValues[0]);
    const max = resolve(1, rangeValues[rangeValues.length - 1]);
    const ratio = max === min ? 0 : Math.max(0, Math.min(1, (val - min) / (max - min)));

    cellEl.style.background = `linear-gradient(to right, ${db.color}36 0%, ${db.color}36 ${ratio * 100}%, transparent ${ratio * 100}%)`;
  }

  private lerpColor(c1: string, c2: string, t: number): string {
    const r1 = parseInt(c1.slice(1,3), 16), g1 = parseInt(c1.slice(3,5), 16), b1 = parseInt(c1.slice(5,7), 16);
    const r2 = parseInt(c2.slice(1,3), 16), g2 = parseInt(c2.slice(3,5), 16), b2 = parseInt(c2.slice(5,7), 16);
    const r = Math.round(r1 + (r2 - r1) * t);
    const g = Math.round(g1 + (g2 - g1) * t);
    const b = Math.round(b1 + (b2 - b1) * t);
    return `#${r.toString(16).padStart(2,'0')}${g.toString(16).padStart(2,'0')}${b.toString(16).padStart(2,'0')}`;
  }

  private selectCell(td: HTMLElement, row: number, col: number, cell: any): void {
    this.selectedRow = row;
    this.selectedCol = col;
    this.selectedMode = 'cell';
    this.highlightSelection();
    this.onCellSelectedCallback?.(cell, row, col);
  }

  private selectRow(row: number): void {
    this.selectedRow = row;
    this.selectedCol = -1;
    this.selectedMode = 'row';
    this.highlightSelection();
    this.onCellSelectedCallback?.(null, row, -1);
  }

  private selectColumn(col: number): void {
    this.selectedRow = -1;
    this.selectedCol = col;
    this.selectedMode = 'col';
    this.highlightSelection();
    this.onCellSelectedCallback?.(null, -1, col);
  }

  private selectAllCells(): void {
    this.selectedRow = -1;
    this.selectedCol = -1;
    this.selectedMode = 'all';
    this.highlightSelection();
    this.onCellSelectedCallback?.(null, -1, -1);
  }

  private highlightSelection(): void {
    if (!this.scrollEl) return;
    const table = this.scrollEl.querySelector('table');
    if (!table) return;

    // clear previous highlights
    table.querySelectorAll('.excel-selected-row, .excel-selected-col, .excel-selected-cell').forEach(el => {
      el.classList.remove('excel-selected-row', 'excel-selected-col', 'excel-selected-cell');
    });

    if (this.selectedMode === 'cell' && this.selectedRow >= 0 && this.selectedCol >= 0) {
      table.querySelector<HTMLElement>(`td[data-row="${this.selectedRow}"][data-col="${this.selectedCol}"]`)?.classList.add('excel-selected-cell');
      table.querySelector<HTMLElement>(`th[data-row="${this.selectedRow}"]`)?.classList.add('excel-selected-row');
      table.querySelector<HTMLElement>(`th[data-col="${this.selectedCol}"]`)?.classList.add('excel-selected-col');
    } else if (this.selectedMode === 'row' && this.selectedRow >= 0) {
      table.querySelectorAll<HTMLElement>(`th[data-row="${this.selectedRow}"], td[data-row="${this.selectedRow}"]`).forEach(cell => cell.classList.add('excel-selected-row'));
    } else if (this.selectedMode === 'col' && this.selectedCol >= 0) {
      table.querySelectorAll<HTMLElement>(`th[data-col="${this.selectedCol}"], td[data-col="${this.selectedCol}"]`).forEach(cell => cell.classList.add('excel-selected-col'));
    } else if (this.selectedMode === 'all') {
      table.querySelectorAll('td').forEach(el => el.classList.add('excel-selected-cell'));
    }
  }

  private applyCellStyle(td: HTMLElement, style: CellStyle, cellType?: string): void {
    // 默认对齐：数字右对齐，文本左对齐
    if (!style.align) {
      if (cellType === 'number' || cellType === 'date') {
        td.style.textAlign = 'right';
      } else {
        td.style.textAlign = 'left';
      }
    } else {
      td.style.textAlign = style.align;
    }

    if (style.font) {
      if (style.font.bold) td.style.fontWeight = 'bold';
      if (style.font.italic) td.style.fontStyle = 'italic';
      if (style.font.size && style.font.size > 0) td.style.fontSize = style.font.size + 'px';
      if (style.font.color) td.style.color = style.font.color;
      if (style.font.name) td.style.fontFamily = style.font.name + ', ' + EXCEL_FONT;
      if (style.font.underline) td.style.textDecoration = 'underline';
      if (style.font.strike) td.style.textDecoration = td.style.textDecoration ? td.style.textDecoration + ' line-through' : 'line-through';
    }
    if (style.bgcolor) td.style.backgroundColor = style.bgcolor;
    if (style.valign) td.style.verticalAlign = style.valign;
    if (style.textwrap) {
      td.style.whiteSpace = 'normal';
      td.style.wordBreak = 'break-word';
    }
    if (style.color) td.style.color = style.color;

    // 应用边框样式 - 将 Excel 边框样式映射到 CSS
    if (style.border) {
      const borderStyleMap: Record<string, string> = {
        'thin': '1px solid',
        'medium': '2px solid',
        'thick': '3px solid',
        'double': 'double',
        'dashed': 'dashed',
        'dotted': 'dotted',
        'hair': '1px solid',
        'mediumDashed': '2px dashed',
        'dashDot': 'dashed',
        'mediumDashDot': '2px dashed',
        'dashDotDot': 'dotted',
        'mediumDashDotDot': '2px dotted',
        'slantDashDot': 'dashed',
      };

      const applyBorder = (pos: 'top' | 'bottom' | 'left' | 'right') => {
        const border = style.border?.[pos];
        if (border) {
          const [excelStyle, color] = border;
          const cssStyle = borderStyleMap[excelStyle] || '1px solid';
          (td.style as any)[`border${pos.charAt(0).toUpperCase() + pos.slice(1)}`] = `${cssStyle} ${color}`;
        }
      };

      applyBorder('top');
      applyBorder('bottom');
      applyBorder('left');
      applyBorder('right');
    }
  }

  private renderSheetBar(): void {
    const sheetBar = this.sheetBar;
    if (!sheetBar || !this.workbookData) return;
    sheetBar.innerHTML = '';

    const tabsWrap = document.createElement('div');
    tabsWrap.style.cssText = 'display:flex;flex:1;overflow:hidden;';
    sheetBar.appendChild(tabsWrap);

    const tabsScroll = document.createElement('div');
    tabsScroll.style.cssText = 'display:flex;overflow-x:auto;flex:1;';
    tabsWrap.appendChild(tabsScroll);

    this.workbookData.sheets.forEach((sheet, i) => {
      const tab = document.createElement('div');
      tab.textContent = sheet.name;
      const isActive = i === this.currentSheetIndex;
      tab.style.cssText = `
        padding:4px 16px 3px;cursor:pointer;font-size:13px;user-select:none;
        white-space:nowrap;min-width:60px;text-align:center;
        border-right:1px solid #c0c0c0;
        ${isActive
          ? 'background:#fff;border-bottom:2px solid #217346;font-weight:500;color:#217346;margin-bottom:-1px;'
          : 'background:#e8e8e8;color:#555;border-bottom:1px solid #c0c0c0;'}
      `;
      tab.addEventListener('click', () => this.switchSheet(i));
      tabsScroll.appendChild(tab);
    });

    const emptyTab = document.createElement('div');
    emptyTab.style.cssText = 'flex:1;border-bottom:1px solid #c0c0c0;background:#e8e8e8;';
    tabsScroll.appendChild(emptyTab);
  }

  switchSheet(index: number): void {
    if (!this.workbookData || index < 0 || index >= this.workbookData.sheets.length) return;
    this.currentSheetIndex = index;
    this.selectedRow = -1;
    this.selectedCol = -1;
    if (this.scrollEl) this.scrollEl.scrollTop = 0;
    this.renderSheet(index);
    this.renderSheetBar();
    this.onSwitchSheetCallback?.(index);
  }

  /**
   * 获取滚动容器元素（图表浮层应挂载在此容器内，随表格内容一起滚动）
   */
  getScrollContainer(): HTMLElement | null {
    return this.scrollEl;
  }

  getCurrentSheetIndex(): number {
    return this.currentSheetIndex;
  }

  getCurrentSheet(): ParsedSheet | null {
    if (!this.workbookData) return null;
    return this.workbookData.sheets[this.currentSheetIndex] || null;
  }

  getSheets(): ParsedSheet[] {
    return this.workbookData?.sheets || [];
  }

  /**
   * 设置最小行数（用于容纳图表）
   */
  setMinRowCount(count: number): void {
    this.minRowCount = count;
  }

  /**
   * 设置最小列数（用于容纳图表）
   */
  setMinColCount(count: number): void {
    this.minColCount = count;
  }

  getCellPosition(col: number, row: number): { left: number; top: number; width: number; height: number } {
    const sheet = this.getCurrentSheet();
    if (!sheet) {
      return { left: HEADER_COL_WIDTH + col * DEFAULT_COL_WIDTH, top: HEADER_ROW_HEIGHT + row * DEFAULT_ROW_HEIGHT, width: DEFAULT_COL_WIDTH, height: DEFAULT_ROW_HEIGHT };
    }
    let left = HEADER_COL_WIDTH;
    for (let i = 0; i < col; i++) {
      left += i < sheet.colWidths.length ? getColWidth(sheet, i, this.colWidthOverrides) : DEFAULT_COL_WIDTH;
    }
    let top = HEADER_ROW_HEIGHT;
    for (let i = 0; i < row; i++) {
      top += i < sheet.rowHeights.length ? getRowHeight(sheet, i, this.rowHeightOverrides) : DEFAULT_ROW_HEIGHT;
    }
    const width = col < sheet.colWidths.length ? getColWidth(sheet, col, this.colWidthOverrides) : DEFAULT_COL_WIDTH;
    const height = row < sheet.rowHeights.length ? getRowHeight(sheet, row, this.rowHeightOverrides) : DEFAULT_ROW_HEIGHT;
    return { left, top, width, height };
  }

  onSwitchSheet(callback: (index: number) => void): void {
    this.onSwitchSheetCallback = callback;
  }

  onCellSelected(callback: (cell: any, rowIndex: number, colIndex: number) => void): void {
    this.onCellSelectedCallback = callback;
  }

  // ===== 列宽/行高拖拽调整 =====

  private startColResize(e: MouseEvent, colIndex: number): void {
    e.preventDefault();
    e.stopPropagation();
    const sheet = this.getCurrentSheet();
    if (!sheet) return;
    this.resizeState = {
      type: 'col',
      index: colIndex,
      startPos: e.clientX,
      startSize: getColWidth(sheet, colIndex, this.colWidthOverrides),
    };
    document.addEventListener('mousemove', this.onResizeMove);
    document.addEventListener('mouseup', this.onResizeEnd);
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';
  }

  private startRowResize(e: MouseEvent, rowIndex: number): void {
    e.preventDefault();
    e.stopPropagation();
    const sheet = this.getCurrentSheet();
    if (!sheet) return;
    this.resizeState = {
      type: 'row',
      index: rowIndex,
      startPos: e.clientY,
      startSize: getRowHeight(sheet, rowIndex, this.rowHeightOverrides),
    };
    document.addEventListener('mousemove', this.onResizeMove);
    document.addEventListener('mouseup', this.onResizeEnd);
    document.body.style.cursor = 'row-resize';
    document.body.style.userSelect = 'none';
  }

  private onResizeMove = (e: MouseEvent): void => {
    const state = this.resizeState;
    if (!state) return;
    const sheet = this.getCurrentSheet();
    if (!sheet) return;

    if (state.type === 'col') {
      const diff = e.clientX - state.startPos;
      const newWidth = Math.max(30, state.startSize + diff);
      this.colWidthOverrides.set(state.index, newWidth);
      this.applyColWidth(state.index, newWidth);
    } else {
      const diff = e.clientY - state.startPos;
      const newHeight = Math.max(18, state.startSize + diff);
      this.rowHeightOverrides.set(state.index, newHeight);
      this.applyRowHeight(state.index, newHeight);
    }
  };

  private onResizeEnd = (): void => {
    this.resizeState = null;
    document.removeEventListener('mousemove', this.onResizeMove);
    document.removeEventListener('mouseup', this.onResizeEnd);
    document.body.style.cursor = '';
    document.body.style.userSelect = '';
  };

  private applyColWidth(colIndex: number, width: number): void {
    if (!this.scrollEl) return;
    const table = this.scrollEl.querySelector('table');
    if (!table) return;
    const cells = table.querySelectorAll<HTMLElement>(`th, td`);
    const colIdx = colIndex + 1;
    for (const cell of cells) {
      const cellIndex = (cell.parentElement?.querySelectorAll('th, td') as NodeListOf<HTMLElement>) || [];
      let idx = 0;
      for (let i = 0; i < cellIndex.length; i++) {
        if (cellIndex[i] === cell) { idx = i; break; }
      }
      if (idx === colIdx) {
        cell.style.width = `${width}px`;
        cell.style.minWidth = `${width}px`;
        cell.style.maxWidth = `${width}px`;
      }
    }
  }

  private applyRowHeight(rowIndex: number, height: number): void {
    if (!this.scrollEl) return;
    const table = this.scrollEl.querySelector('table');
    if (!table) return;
    const rows = table.querySelectorAll('tr');
    const row = rows[rowIndex + 1];
    if (row) {
      row.style.height = `${height}px`;
      const cells = row.querySelectorAll<HTMLElement>('th, td');
      for (const cell of cells) {
        cell.style.height = `${height}px`;
      }
    }
  }

  // ===== 数据透视表 =====

  isPivotTableMode(): boolean {
    return this.pivotTableMode;
  }

  togglePivotTable(): void {
    this.pivotTableMode = !this.pivotTableMode;
    if (this.pivotTableMode) {
      this.renderPivotTable();
    } else {
      this.hidePivotTable();
    }
  }

  private getCurrentPivotTable(): ParsedPivotTable | undefined {
    if (!this.workbookData?.pivotTables) return undefined;
    return this.workbookData.pivotTables.find(pt => pt.sheetIndex === this.currentSheetIndex);
  }

  private renderPivotTable(): void {
    if (!this.scrollEl) return;
    const pivotTable = this.getCurrentPivotTable();
    if (!pivotTable) {
      this.pivotTableMode = false;
      return;
    }

    // Hide normal table
    const table = this.scrollEl.querySelector('table');
    if (table) table.style.display = 'none';

    // Create or show pivot table overlay
    if (!this.pivotTableEl) {
      this.pivotTableEl = document.createElement('div');
      this.pivotTableEl.className = 'excel-pivot-table';
      this.pivotTableEl.style.cssText = 'overflow:auto;height:100%;padding:8px;';
      this.scrollEl.appendChild(this.pivotTableEl);
    }

    this.pivotTableEl.style.display = 'block';

    // Group records by row fields
    const rowIndices = pivotTable.rowFieldIndices;
    const colIndices = pivotTable.colFieldIndices;
    const dataFields = pivotTable.dataFields;
    const records = pivotTable.cacheRecords;
    const fields = pivotTable.cacheFields;

    // Build grouped data
    const groups: Map<string, {
      keys: string[];
      records: PivotCacheRecord[];
      subtotals: Map<number, number>;
    }> = new Map();

    for (const rec of records) {
      const keyParts = rowIndices.map(i => String(rec.values[i] ?? ''));
      const key = keyParts.join('||');
      if (!groups.has(key)) {
        groups.set(key, { keys: keyParts, records: [], subtotals: new Map() });
      }
      groups.get(key)!.records.push(rec);

      // Calculate subtotals
      for (const df of dataFields) {
        const val = parseFloat(String(rec.values[df.fieldIndex] ?? '0'));
        const current = groups.get(key)!.subtotals.get(df.fieldIndex) || 0;
        groups.get(key)!.subtotals.set(df.fieldIndex, current + val);
      }
    }

    // Build HTML
    const html: string[] = [];
    html.push(`<div style="font-weight:bold;margin-bottom:8px;font-size:13px;color:#333;">${pivotTable.name}</div>`);
    html.push('<table style="border-collapse:collapse;font-size:12px;width:100%;">');

    // Header row
    html.push('<tr>');
    for (const ri of rowIndices) {
      html.push(`<th style="border:1px solid #d0d0d0;background:#f3f3f3;padding:4px 8px;text-align:left;font-weight:500;">${fields[ri]?.name || `Row ${ri}`}</th>`);
    }
    for (const df of dataFields) {
      html.push(`<th style="border:1px solid #d0d0d0;background:#f3f3f3;padding:4px 8px;text-align:right;font-weight:500;">${df.name}</th>`);
    }
    html.push('</tr>');

    // Data rows
    for (const [key, group] of groups.entries()) {
      html.push('<tr>');
      for (const k of group.keys) {
        html.push(`<td style="border:1px solid #e0e0e0;padding:3px 8px;">${k}</td>`);
      }
      for (const df of dataFields) {
        const total = group.subtotals.get(df.fieldIndex) || 0;
        html.push(`<td style="border:1px solid #e0e0e0;padding:3px 8px;text-align:right;">${total}</td>`);
      }
      html.push('</tr>');
    }

    // Grand total
    if (groups.size > 1) {
      html.push('<tr>');
      html.push(`<td style="border:1px solid #e0e0e0;padding:3px 8px;font-weight:500;background:#f9f9f9;" colspan="${rowIndices.length}">总计</td>`);
      const grandTotals = new Map<number, number>();
      for (const group of groups.values()) {
        for (const [fi, val] of group.subtotals) {
          grandTotals.set(fi, (grandTotals.get(fi) || 0) + val);
        }
      }
      for (const df of dataFields) {
        const total = grandTotals.get(df.fieldIndex) || 0;
        html.push(`<td style="border:1px solid #e0e0e0;padding:3px 8px;text-align:right;font-weight:500;background:#f9f9f9;">${total}</td>`);
      }
      html.push('</tr>');
    }

    html.push('</table>');

    // Record count
    html.push(`<div style="margin-top:6px;font-size:11px;color:#888;">共 ${records.length} 条记录</div>`);

    this.pivotTableEl.innerHTML = html.join('');

    // Notify callback
    this.onCellSelectedCallback?.(null, -1, -1);
  }

  private hidePivotTable(): void {
    if (this.scrollEl) {
      const table = this.scrollEl.querySelector('table');
      if (table) table.style.display = '';
    }
    if (this.pivotTableEl) {
      this.pivotTableEl.style.display = 'none';
    }
  }

  destroy(): void {
    this.isDestroyed = true;
    this.onResizeEnd();
    this.colWidthOverrides.clear();
    this.rowHeightOverrides.clear();
    if (this.rootEl && this.rootEl.parentNode) {
      this.rootEl.parentNode.removeChild(this.rootEl);
    }
    this.container = null;
    this.rootEl = null;
    this.scrollEl = null;
    this.sheetBar = null;
    this.workbookData = null;
    this.onSwitchSheetCallback = undefined;
    this.onCellSelectedCallback = undefined;
  }
}
