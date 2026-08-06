/**
 * OpenXML 图表 & 图片解析入口
 *
 * 数据流:
 *   xlsx (zip) ──► JSZip 解压
 *      ├─► xl/drawings/*.xml → 锚点位置 + 图片
 *      └─► xl/charts/*.xml   → OoxmlChartParser → ChartModel
 *
 * 本模块负责:
 *   1. 解压 xlsx，建立 Sheet ↔ drawing 映射
 *   2. 遍历 drawing 锚点，解析图表锚点 & 图片
 *   3. 对每个图表 XML 调用 parseChartXmlToModel 生成 ChartModel
 *   4. 解析图片 (base64)
 */

import JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';
import ExcelJS from 'exceljs';
import type { ParsedImage, ChartAnchor } from '../types';
import type { ChartModel } from '../chart/chart-model';
import { parseChartXmlToModel } from '../chart/ooxml-chart-parser';
import { parseThemeFromZip } from '../chart/theme-parser';
import type { ChartTheme } from '../chart/theme-parser';
import { getOoxmlNumber as getNum, toArray } from '../utils/ooxml';

// ===== XML 解析配置 =====

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  isArray: (name) => {
    const arrayNodes = [
      'c:ser', 'c:cat', 'c:val', 'c:dLbls',
      'xdr:twoCellAnchor', 'xdr:oneCellAnchor', 'xdr:absoluteAnchor', 'Relationship',
      // 所有图表类型元素（确保单次出现时也为数组）
      'c:barChart', 'c:bar3DChart', 'c:lineChart', 'c:line3DChart',
      'c:pieChart', 'c:pie3DChart', 'c:areaChart', 'c:area3DChart',
      'c:scatterChart', 'c:doughnutChart', 'c:bubbleChart',
      'c:radarChart', 'c:stockChart', 'c:surfaceChart', 'c:surface3DChart',
      // 引用/缓存节点
      'c:title', 'c:tx', 'c:rich', 'c:strRef', 'c:numRef',
      'c:numLit', 'c:numCache', 'c:strCache', 'c:strLit',
      'c:yVal', 'c:xVal', 'c:high', 'c:low', 'c:open', 'c:close',
      'c:bubbleSize', 'c:volume', 'c:cat', 'c:val',
      // 坐标轴
      'c:catAx', 'c:valAx', 'c:dateAx', 'c:serAx',
      'c:axId',
      // drawing 内部
      'a:p', 'a:r', 'a:t',
    ];
    return arrayNodes.includes(name);
  },
});

/** 解析相对路径 */
function resolveRelativePath(baseDir: string, relativePath: string): string {
  const parts = baseDir.split('/');
  const relParts = relativePath.split('/');
  for (const part of relParts) {
    if (part === '..') parts.pop();
    else if (part !== '.') parts.push(part);
  }
  return parts.join('/');
}

/**
 * 解析图表/图片锚点位置
 *
 * 支持三种锚点类型:
 *   - twoCellAnchor: from + to（最常见）
 *   - oneCellAnchor: from + ext（绝对尺寸）
 *   - absoluteAnchor: pos + ext（绝对坐标，EMU）
 */
function parseAnchor(anchorData: any): ChartAnchor | null {
  if (!anchorData) return null;

  // twoCellAnchor: from + to
  const from = anchorData['xdr:from'];
  if (from) {
    const to = anchorData['xdr:to'];
    if (to) {
      return {
        fromCol: getNum(from['xdr:col']),
        fromColOff: getNum(from['xdr:colOff']),
        fromRow: getNum(from['xdr:row']),
        fromRowOff: getNum(from['xdr:rowOff']),
        toCol: getNum(to['xdr:col']),
        toColOff: getNum(to['xdr:colOff']),
        toRow: getNum(to['xdr:row']),
        toRowOff: getNum(to['xdr:rowOff']),
      };
    }

    // oneCellAnchor: from + ext（使用实际尺寸）
    const ext = anchorData['xdr:ext'];
    if (ext) {
      const fromCol = getNum(from['xdr:col']);
      const fromRow = getNum(from['xdr:row']);
      const fromColOff = getNum(from['xdr:colOff']);
      const fromRowOff = getNum(from['xdr:rowOff']);
      // ext 的 cx/cy 是 EMU (914400 EMU = 1 inch)
      // 转换为列/行偏移量: 1 列 ≈ 635000 EMU (≈7px), 1 行 ≈ 190500 EMU (≈5px)
      const extCx = getNum(ext['@_cx']);
      const extCy = getNum(ext['@_cy']);
      // 估算跨越的列/行数 (使用 EMU 转换)
      const colSpan = Math.max(1, Math.round(extCx / 635000));
      const rowSpan = Math.max(1, Math.round(extCy / 190500));
      return {
        fromCol,
        fromColOff,
        fromRow,
        fromRowOff,
        toCol: fromCol + colSpan,
        toColOff: 0,
        toRow: fromRow + rowSpan,
        toRowOff: 0,
      };
    }
  }

  // absoluteAnchor: pos + ext（绝对坐标）
  const pos = anchorData['xdr:pos'];
  if (pos) {
    const x = getNum(pos['@_x']);
    const y = getNum(pos['@_y']);
    const ext = anchorData['xdr:ext'];
    const cx = ext ? getNum(ext['@_cx']) : 0;
    const cy = ext ? getNum(ext['@_cy']) : 0;
    // EMU 转列/行: 1 列 ≈ 635000 EMU, 1 行 ≈ 190500 EMU
    const fromCol = Math.floor(x / 635000);
    const fromRow = Math.floor(y / 190500);
    const fromColOff = x % 635000;
    const fromRowOff = y % 190500;
    const toCol = Math.floor((x + cx) / 635000);
    const toRow = Math.floor((y + cy) / 190500);
    const toColOff = (x + cx) % 635000;
    const toRowOff = (y + cy) % 190500;
    return {
      fromCol, fromColOff, fromRow, fromRowOff,
      toCol, toColOff, toRow, toRowOff,
    };
  }

  return null;
}

// ===== 图片解析 =====

/**
 * 从 drawing 中解析所有图片
 */
async function parseImagesFromDrawing(
  zip: JSZip,
  drawingFile: string,
  drawingXml: any,
  sheetIndex: number = 0
): Promise<ParsedImage[]> {
  const images: ParsedImage[] = [];
  const wsDr = drawingXml['xdr:wsDr'];
  if (!wsDr) return images;

  const twoCellAnchors = wsDr['xdr:twoCellAnchor'] || [];
  const oneCellAnchors = wsDr['xdr:oneCellAnchor'] || [];
  const absoluteAnchors = wsDr['xdr:absoluteAnchor'] || [];
  const anchors = [
    ...(Array.isArray(twoCellAnchors) ? twoCellAnchors : [twoCellAnchors]),
    ...(Array.isArray(oneCellAnchors) ? oneCellAnchors : [oneCellAnchors]),
    ...(Array.isArray(absoluteAnchors) ? absoluteAnchors : [absoluteAnchors]),
  ].filter(Boolean);

  // 获取关系文件
  const idx = drawingFile.lastIndexOf('/');
  const drawingRelsFile = `${drawingFile.substring(0, idx)}/_rels/${drawingFile.substring(idx + 1)}.rels`;
  const relsContent = await zip.file(drawingRelsFile)?.async('text');
  if (!relsContent) return images;

  const relsXml = xmlParser.parse(relsContent);
  const relationships = relsXml['Relationships']?.['Relationship'] || [];

  for (let i = 0; i < anchors.length; i++) {
    const anchor = anchors[i];
    const pics = Array.isArray(anchor['xdr:pic'])
      ? anchor['xdr:pic']
      : [anchor['xdr:pic']].filter(Boolean);

    for (const pic of pics) {
      if (!pic) continue;

      const imageAnchor = parseAnchor(anchor);
      if (!imageAnchor) continue;

      const nvPicPr = pic['xdr:nvPicPr'];
      const cNvPr = nvPicPr?.['xdr:cNvPr'];
      const picName = cNvPr?.['@_name'] || `Image_${i}`;

      const blipFill = pic['xdr:blipFill'];
      const blip = blipFill?.['a:blip'];
      const imageRef = blip?.['@_r:embed'] || blip?.['@_embed'];
      if (!imageRef) continue;

      const rel = Array.isArray(relationships)
        ? relationships.find((r: any) => r?.['@_Id'] === imageRef)
        : relationships['@_Id'] === imageRef ? relationships : null;
      if (!rel) continue;

      const drawingDir = drawingFile.substring(0, drawingFile.lastIndexOf('/'));
      const imageTarget = resolveRelativePath(drawingDir, rel['@_Target']);
      const imageFile = zip.file(imageTarget);
      if (!imageFile) continue;

      const imageData = await imageFile.async('base64');
      const ext = imageTarget.split('.').pop()?.toLowerCase() || 'png';
      const mimeType = ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : `image/${ext}`;

      const spPr = pic['xdr:spPr'];
      const xfrm = spPr?.['a:xfrm'];
      const aExt = xfrm?.['a:ext'];
      const extent = aExt ? {
        width: getNum(aExt['@_cx']),
        height: getNum(aExt['@_cy']),
      } : undefined;

      images.push({
        id: `img_${images.length}_${i}`,
        anchor: imageAnchor,
        imageData: `data:${mimeType};base64,${imageData}`,
        mimeType,
        name: picName,
        sheetIndex,
        extent,
      });
    }
  }

  return images;
}

// ===== 主入口函数 =====

export interface ParseChartsResult {
  charts: ChartModel[];
  images: ParsedImage[];
}

/**
 * 解析 Excel 文件中的所有图表和图片
 *
 * @param buffer - Excel 文件的 ArrayBuffer 数据
 * @param workbook - 已解析的 exceljs Workbook 对象（可选）
 * @returns Promise<ParseChartsResult> 解析后的图表模型和图片列表
 */
export async function parseCharts(
  buffer: ArrayBuffer,
  workbook?: ExcelJS.Workbook
): Promise<ParseChartsResult> {
  const charts: ChartModel[] = [];
  const images: ParsedImage[] = [];

  try {
    // 1. 解压 xlsx
    const zip = await JSZip.loadAsync(buffer);

    // 1.5 解析主题色
    const theme = await parseThemeFromZip(buffer);

    // 2. 创建/复用 workbook
    let excelWorkbook: ExcelJS.Workbook;
    if (workbook) {
      excelWorkbook = workbook;
    } else {
      excelWorkbook = new ExcelJS.Workbook();
      await excelWorkbook.xlsx.load(buffer);
    }

    // 3. 建立 Sheet ↔ drawing 映射
    const sheetDrawingMap: Map<number, string> = new Map();
    for (let i = 0; i < excelWorkbook.worksheets.length; i++) {
      const sheet = excelWorkbook.worksheets[i];
      const sheetId = i;
      const relsPath = `xl/worksheets/_rels/sheet${sheet.id}.xml.rels`;
      const relsContent = await zip.file(relsPath)?.async('text');
      if (relsContent) {
        const relsXml = xmlParser.parse(relsContent);
        const relationships = relsXml['Relationships']?.['Relationship'] || [];
        for (const rel of relationships) {
          if (rel?.['@_Type']?.includes('/drawing')) {
            const target = rel['@_Target'];
            const drawingPath = target.startsWith('../')
              ? `xl/${target.substring(3)}`
              : `xl/worksheets/${target}`;
            sheetDrawingMap.set(sheetId, drawingPath);
            break;
          }
        }
      }
    }

    // 4. 查找所有 drawing 文件
    const drawingFiles: string[] = [];
    zip.forEach((relativePath) => {
      if (relativePath.startsWith('xl/drawings/') && relativePath.endsWith('.xml')) {
        drawingFiles.push(relativePath);
      }
    });

    if (drawingFiles.length === 0) {
      return { charts, images };
    }

    // 5. 解析每个 drawing 文件
    for (const drawingFile of drawingFiles) {
      // 找到对应的 Sheet 索引
      let sheetIndex = 0;
      for (const [idx, path] of sheetDrawingMap.entries()) {
        if (path === drawingFile) {
          sheetIndex = idx;
          break;
        }
      }

      try {
        const drawingContent = await zip.file(drawingFile)?.async('text');
        if (!drawingContent) continue;

        const drawingXml = xmlParser.parse(drawingContent);
        const wsDr = drawingXml['xdr:wsDr'];
        if (!wsDr) continue;

        const twoCellAnchors = wsDr['xdr:twoCellAnchor'] || [];
        const oneCellAnchors = wsDr['xdr:oneCellAnchor'] || [];
        const absoluteAnchors = wsDr['xdr:absoluteAnchor'] || [];
        const anchors = [
          ...(Array.isArray(twoCellAnchors) ? twoCellAnchors : [twoCellAnchors]),
          ...(Array.isArray(oneCellAnchors) ? oneCellAnchors : [oneCellAnchors]),
          ...(Array.isArray(absoluteAnchors) ? absoluteAnchors : [absoluteAnchors]),
        ].filter(Boolean);

        // 先解析图片
        const drawingImages = await parseImagesFromDrawing(zip, drawingFile, drawingXml, sheetIndex);
        images.push(...drawingImages);

        // 解析图表锚点
        for (let i = 0; i < anchors.length; i++) {
          const anchor = anchors[i];
          const chartAnchor = parseAnchor(anchor);
          if (!chartAnchor) continue;

          // 查找 graphicFrame → chart 引用
          const graphicFrames = toArray(anchor['xdr:graphicFrame']);
          let chartRef = '';
          for (const gf of graphicFrames) {
            const graphics = toArray(gf['a:graphic']);
            for (const g of graphics) {
              const gdList = toArray(g['a:graphicData']);
              for (const gd of gdList) {
                if (gd?.['c:chart']) {
                  const c = gd['c:chart'];
                  chartRef = c['@_r:id'] || c['@_id'] || '';
                  break;
                }
              }
              if (chartRef) break;
            }
            if (chartRef) break;
          }
          if (!chartRef) continue;

          // 通过关系文件查找 chart 文件路径
          const idx = drawingFile.lastIndexOf('/');
          const drawingRelsFile = `${drawingFile.substring(0, idx)}/_rels/${drawingFile.substring(idx + 1)}.rels`;
          const relsContent = await zip.file(drawingRelsFile)?.async('text');
          if (!relsContent) continue;

          const relsXml = xmlParser.parse(relsContent);
          const relationships = toArray(relsXml['Relationships']?.['Relationship']);
          const relationship = relationships.find((r: any) => r?.['@_Id'] === chartRef);
          if (!relationship) continue;

          // 解析 chart 文件路径
          const drawingDir = drawingFile.substring(0, drawingFile.lastIndexOf('/'));
          const chartTarget = resolveRelativePath(drawingDir, relationship['@_Target']);
          const chartContent = await zip.file(chartTarget)?.async('text');
          if (!chartContent) continue;

          // 6. 解析图表 XML → ChartModel
          const chartXmlObj = xmlParser.parse(chartContent);
          const chartId = `chart_${charts.length}_${i}`;

          // 6.5 解析图表关系文件 (colors*.xml, style*.xml)
          const chartDir = chartTarget.substring(0, chartTarget.lastIndexOf('/'));
          const chartFileName = chartTarget.substring(chartTarget.lastIndexOf('/') + 1);
          const chartRelsPath = `${chartDir}/_rels/${chartFileName}.rels`;
          const chartRelsContent = await zip.file(chartRelsPath)?.async('text');
          let chartColorStyle: any = undefined;
          let chartStyle: any = undefined;

          if (chartRelsContent) {
            const chartRelsXml = xmlParser.parse(chartRelsContent);
            const chartRels = toArray(chartRelsXml['Relationships']?.['Relationship']);
            for (const rel of chartRels) {
              const relType = rel?.['@_Type'] || '';
              const relTarget = rel?.['@_Target'] || '';
              if (relType.includes('chartColorStyle')) {
                const colorPath = resolveRelativePath(chartDir, relTarget);
                const colorContent = await zip.file(colorPath)?.async('text');
                if (colorContent) {
                  chartColorStyle = xmlParser.parse(colorContent);
                }
              } else if (relType.includes('chartStyle')) {
                const stylePath = resolveRelativePath(chartDir, relTarget);
                const styleContent = await zip.file(stylePath)?.async('text');
                if (styleContent) {
                  chartStyle = xmlParser.parse(styleContent);
                }
              }
            }
          }

          // 将 colors/style XML 挂载到 chart XML 对象上供 parser 使用
          if (chartColorStyle) {
            (chartXmlObj as any).__colorStyle = chartColorStyle;
          }
          if (chartStyle) {
            (chartXmlObj as any).__chartStyle = chartStyle;
          }

          const model = parseChartXmlToModel(chartXmlObj, excelWorkbook, chartAnchor, sheetIndex, chartId, theme);

          if (model) {
            charts.push(model);
          }
        }
      } catch (err) {
        console.warn(`[excel-preview] Failed to parse drawing file ${drawingFile}:`, err);
      }
    }
  } catch (error) {
    console.error('[excel-preview] Failed to parse charts:', error);
  }

  return { charts, images };
}
