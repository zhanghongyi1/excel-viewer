import type { ChartAnchor } from '../types';
import { EMU_PER_PIXEL } from '../utils/ooxml';

export interface PixelRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

export type PositionFn = (col: number, row: number) => PixelRect;

/** 将 Excel 单元格锚点转换为浮层像素区域。 */
export function calculateAnchorRect(anchor: ChartAnchor, positionFn: PositionFn): PixelRect {
  const start = positionFn(anchor.fromCol, anchor.fromRow);
  const end = positionFn(anchor.toCol, anchor.toRow);
  const left = start.left + anchor.fromColOff / EMU_PER_PIXEL;
  const top = start.top + anchor.fromRowOff / EMU_PER_PIXEL;
  const right = end.left + anchor.toColOff / EMU_PER_PIXEL;
  const bottom = end.top + anchor.toRowOff / EMU_PER_PIXEL;
  return {
    left: Math.round(left),
    top: Math.round(top),
    width: Math.round(right - left),
    height: Math.round(bottom - top),
  };
}
