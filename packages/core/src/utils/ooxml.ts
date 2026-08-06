/** OpenXML 长度单位 EMU 转 CSS 像素（96 DPI）。 */
export const EMU_PER_PIXEL = 9525;

/** 将 fast-xml-parser 的单节点/数组统一为数组。 */
export function toArray<T = any>(value: T | T[] | null | undefined): T[] {
  if (value === undefined || value === null) return [];
  return Array.isArray(value) ? value : [value];
}

/** 安全读取 OpenXML 数值或文本节点。 */
export function getOoxmlNumber(value: any): number {
  if (value === undefined || value === null) return 0;
  if (typeof value === 'number') return value;
  const raw = typeof value === 'object' ? value['#text'] : value;
  const parsed = Number(raw);
  return Number.isFinite(parsed) ? parsed : 0;
}

/** Excel 列名转 0-based 索引：A=0, Z=25, AA=26。 */
export function colLetterToNumber(letters: string): number {
  let result = 0;
  for (const char of letters.toUpperCase()) {
    result = result * 26 + char.charCodeAt(0) - 64;
  }
  return result - 1;
}
