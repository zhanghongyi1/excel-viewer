/**
 * 数据加载器
 * 支持多种数据源统一转换为 ArrayBuffer:
 * - string (URL) → fetch 请求 → ArrayBuffer
 * - ArrayBuffer → 直接返回
 * - Blob / File → .arrayBuffer() → ArrayBuffer
 */

import type { DataSource, LoadOptions } from '../types';

/**
 * 将多种数据源统一转换为 ArrayBuffer
 * @param src 数据源（URL字符串、ArrayBuffer、Blob 或 File）
 * @param options 网络请求配置（仅当 src 为 URL 时生效）
 * @returns Promise<ArrayBuffer>
 */
export async function loadData(
  src: DataSource,
  options: LoadOptions = {}
): Promise<ArrayBuffer> {
  if (typeof src === 'string') {
    return loadFromUrl(src, options);
  }

  if (src instanceof ArrayBuffer) {
    return src;
  }

  if (src instanceof Blob) {
    return src.arrayBuffer();
  }

  throw new Error(`Unsupported data source type: ${typeof src}`);
}

/**
 * 从 URL 加载数据
 */
async function loadFromUrl(url: string, options: LoadOptions): Promise<ArrayBuffer> {
  const response = await fetch(url, {
    method: options.method || 'GET',
    headers: options.headers,
    credentials: options.withCredentials ? 'include' : 'omit',
    body: options.body,
  });

  if (!response.ok) {
    const errorMsg = `Failed to load Excel file: HTTP ${response.status} ${response.statusText}`;
    if (response.status === 0) {
      throw new Error(`${errorMsg}. Possible CORS error or network issue.`);
    }
    if (response.status === 404) {
      throw new Error(`${errorMsg}. File not found at: ${url}`);
    }
    if (response.status === 403) {
      throw new Error(`${errorMsg}. Access denied. Check CORS configuration.`);
    }
    throw new Error(errorMsg);
  }

  // 检查 content-type 是否为 Excel 文件
  const contentType = response.headers.get('content-type') || '';
  const isExcelFile =
    contentType.includes('spreadsheet') ||
    contentType.includes('excel') ||
    contentType.includes('octet-stream') ||
    url.endsWith('.xlsx') ||
    url.endsWith('.xls');

  if (!isExcelFile) {
    console.warn(
      `[excel-preview] Warning: Response content-type is "${contentType}", expected an Excel file.`
    );
  }

  return response.arrayBuffer();
}

/**
 * 判断数据源是否为 URL 字符串
 */
export function isUrlSource(src: DataSource): src is string {
  return typeof src === 'string';
}

/**
 * 判断数据源是否为二进制数据
 */
export function isBinarySource(src: DataSource): src is ArrayBuffer | Blob | File {
  return src instanceof ArrayBuffer || src instanceof Blob;
}
