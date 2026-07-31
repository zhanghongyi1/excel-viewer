/**
 * Theme Parser — OOXML 主题色解析器
 *
 * 解析 xl/theme/theme1.xml，提取:
 *   - 颜色方案 (clrScheme): accent1-6, dk1/2, lt1/2, hlink, folHlink
 *   - 字体方案 (fontScheme): majorFont, minorFont
 *
 * 解析后的 ThemeColorScheme 传入 OOXML Chart Parser，
 * 使图表系列颜色与 Excel 主题完全一致。
 *
 * XML 结构:
 *   <a:theme>
 *     <a:themeElements>
 *       <a:clrScheme>
 *         <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
 *         <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
 *         <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
 *         ...
 *       </a:clrScheme>
 *       <a:fontScheme>
 *         <a:majorFont><a:latin typeface="Cambria"/></a:majorFont>
 *         <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
 *       </a:fontScheme>
 *     </a:themeElements>
 *   </a:theme>
 */

import JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';

// ===== 类型定义 =====

/** 主题颜色方案 */
export interface ThemeColorScheme {
  /** 深色1（文字色） */
  dk1: string;
  /** 浅色1（背景色） */
  lt1: string;
  /** 深色2 */
  dk2: string;
  /** 浅色2 */
  lt2: string;
  /** 强调色1-6 */
  accent1: string;
  accent2: string;
  accent3: string;
  accent4: string;
  accent5: string;
  accent6: string;
  /** 超链接颜色 */
  hlink: string;
  /** 已访问超链接颜色 */
  folHlink: string;
}

/** 主题字体方案 */
export interface ThemeFontScheme {
  majorFont: string;
  minorFont: string;
}

/** 完整主题 */
export interface ChartTheme {
  colors: ThemeColorScheme;
  fonts: ThemeFontScheme;
}

// ===== 默认主题（Office 2007-2010）=====

export const DEFAULT_THEME: ChartTheme = {
  colors: {
    dk1: '#000000',
    lt1: '#FFFFFF',
    dk2: '#1F497D',
    lt2: '#EEECE1',
    accent1: '#4F81BD',
    accent2: '#C0504D',
    accent3: '#9BBB59',
    accent4: '#8064A2',
    accent5: '#4BACC6',
    accent6: '#F79646',
    hlink: '#0000FF',
    folHlink: '#800080',
  },
  fonts: {
    majorFont: 'Cambria',
    minorFont: 'Calibri',
  },
};

// ===== XML 解析 =====

const themeXmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
});

/** 从颜色节点提取 hex 颜色 */
function extractThemeColor(colorNode: any): string {
  if (!colorNode) return '#000000';

  // srgbClr: 直接 RGB 值
  const srgbClr = colorNode['a:srgbClr'];
  if (srgbClr) {
    const val = srgbClr['@_val'] || srgbClr;
    return `#${val}`;
  }

  // sysClr: 系统颜色，lastClr 为回退 RGB 值
  const sysClr = colorNode['a:sysClr'];
  if (sysClr) {
    const lastClr = sysClr['@_lastClr'] || sysClr['lastClr'];
    if (lastClr) return `#${lastClr}`;
    // 根据 val 映射常见系统颜色
    const sysVal = sysClr['@_val'] || sysClr;
    const sysColorMap: Record<string, string> = {
      windowText: '#000000',
      window: '#FFFFFF',
      ButtonFace: '#C0C0C0',
      ButtonShadow: '#808080',
      ButtonText: '#000000',
      CaptionText: '#000000',
      GrayText: '#808080',
      Highlight: '#0000FF',
      HighlightText: '#FFFFFF',
      InfoText: '#000000',
      InfoBk: '#FFFFE1',
    };
    return sysColorMap[sysVal] || '#000000';
  }

  return '#000000';
}

/**
 * 从 theme XML 对象解析主题
 */
export function parseThemeXml(themeXmlObj: any): ChartTheme {
  try {
    const themeElements = themeXmlObj?.['a:theme']?.['a:themeElements'];
    if (!themeElements) return DEFAULT_THEME;

    const clrScheme = themeElements['a:clrScheme'];
    const fontScheme = themeElements['a:fontScheme'];

    // 解析颜色方案
    let colors: ThemeColorScheme = { ...DEFAULT_THEME.colors };
    if (clrScheme) {
      colors = {
        dk1: extractThemeColor(clrScheme['a:dk1']),
        lt1: extractThemeColor(clrScheme['a:lt1']),
        dk2: extractThemeColor(clrScheme['a:dk2']),
        lt2: extractThemeColor(clrScheme['a:lt2']),
        accent1: extractThemeColor(clrScheme['a:accent1']),
        accent2: extractThemeColor(clrScheme['a:accent2']),
        accent3: extractThemeColor(clrScheme['a:accent3']),
        accent4: extractThemeColor(clrScheme['a:accent4']),
        accent5: extractThemeColor(clrScheme['a:accent5']),
        accent6: extractThemeColor(clrScheme['a:accent6']),
        hlink: extractThemeColor(clrScheme['a:hlink']),
        folHlink: extractThemeColor(clrScheme['a:folHlink']),
      };
    }

    // 解析字体方案
    let fonts: ThemeFontScheme = { ...DEFAULT_THEME.fonts };
    if (fontScheme) {
      const majorFont = fontScheme['a:majorFont']?.['a:latin']?.['@_typeface'];
      const minorFont = fontScheme['a:minorFont']?.['a:latin']?.['@_typeface'];
      if (majorFont) fonts.majorFont = majorFont;
      if (minorFont) fonts.minorFont = minorFont;
    }

    return { colors, fonts };
  } catch {
    return DEFAULT_THEME;
  }
}

/**
 * 从 xlsx 的 ArrayBuffer 解析主题
 *
 * @param buffer xlsx 文件 ArrayBuffer
 * @returns ChartTheme（解析失败时返回默认主题）
 */
export async function parseThemeFromZip(buffer: ArrayBuffer): Promise<ChartTheme> {
  try {
    const zip = await JSZip.loadAsync(buffer);
    const themeContent = await zip.file('xl/theme/theme1.xml')?.async('text');
    if (!themeContent) return DEFAULT_THEME;

    const themeXmlObj = themeXmlParser.parse(themeContent);
    return parseThemeXml(themeXmlObj);
  } catch {
    return DEFAULT_THEME;
  }
}

/**
 * 将主题颜色方案转换为 OOXML schemeClr 名称 → hex 的映射表
 * 供 OoxmlChartParser 使用
 */
export function themeColorsToMap(theme: ChartTheme): Record<string, string> {
  return {
    accent1: theme.colors.accent1,
    accent2: theme.colors.accent2,
    accent3: theme.colors.accent3,
    accent4: theme.colors.accent4,
    accent5: theme.colors.accent5,
    accent6: theme.colors.accent6,
    dk1: theme.colors.dk1,
    dk2: theme.colors.dk2,
    lt1: theme.colors.lt1,
    lt2: theme.colors.lt2,
    tx1: theme.colors.dk1, // tx1 映射到 dk1
    tx2: theme.colors.dk2, // tx2 映射到 dk2
    bg1: theme.colors.lt1, // bg1 映射到 lt1
    bg2: theme.colors.lt2, // bg2 映射到 lt2
  };
}
