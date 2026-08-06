/**
 * Chart Parser Snapshot Tests
 *
 * 对每种图表类型的解析结果进行快照回归验证。
 * 基准文件保存在 __snapshots__/ 目录下。
 *
 * 运行: pnpm test
 * 更新基准: pnpm test -- -u
 */

import { describe, it, expect } from 'vitest';
import { XMLParser } from 'fast-xml-parser';
import ExcelJS from 'exceljs';
import { parseChartXmlToModel } from '../chart/ooxml-chart-parser';
import { parseThemeXml, DEFAULT_THEME } from '../chart/theme-parser';
import { computeLayout } from '../chart/layout-engine';
import { convertToEChartsOption } from '../chart/echarts-converter';
import { parseExcel } from '../parser/excel-parser';
import type { ChartModel } from '../chart/chart-model';
import type { ChartAnchor } from '../types';

// ===== 测试工具 =====

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  isArray: (name) => [
    'c:ser', 'c:cat', 'c:val', 'c:dLbls',
    'c:barChart', 'c:lineChart', 'c:areaChart', 'c:pieChart',
    'c:scatterChart', 'c:doughnutChart', 'c:bubbleChart',
    'c:radarChart', 'c:stockChart', 'c:surfaceChart',
    'c:title', 'c:tx', 'c:rich', 'c:strRef', 'c:numRef',
    'c:numLit', 'c:numCache', 'c:strCache', 'c:strLit',
    'c:yVal', 'c:xVal', 'c:high', 'c:low', 'c:open', 'c:close',
    'c:bubbleSize', 'c:volume', 'c:cat', 'c:val',
    'c:catAx', 'c:valAx', 'c:dateAx', 'c:serAx', 'c:axId',
    'a:p', 'a:r', 'a:t',
  ].includes(name),
});

const DEFAULT_ANCHOR: ChartAnchor = {
  fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
  toCol: 10, toColOff: 0, toRow: 20, toRowOff: 0,
};

function createMockWorkbook(): ExcelJS.Workbook {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  // 添加测试数据
  ws.getCell('A1').value = 'Category';
  ws.getCell('B1').value = 'Series1';
  ws.getCell('C1').value = 'Series2';
  for (let i = 2; i <= 6; i++) {
    ws.getCell(`A${i}`).value = `Item ${i - 1}`;
    ws.getCell(`B${i}`).value = Math.round(Math.random() * 100);
    ws.getCell(`C${i}`).value = Math.round(Math.random() * 100);
  }
  return wb;
}

function parseChartXml(xml: string, chartId: string, workbook?: ExcelJS.Workbook): ChartModel | null {
  const chartXmlObj = xmlParser.parse(xml);
  const wb = workbook || createMockWorkbook();
  return parseChartXmlToModel(chartXmlObj, wb, DEFAULT_ANCHOR, 0, chartId, DEFAULT_THEME);
}

// 用于快照的序列化（移除不稳定字段）
function sanitizeForSnapshot(model: ChartModel | null): object {
  if (!model) return { error: 'Model is null' };
  return {
    id: model.id,
    type: model.type,
    title: model.title,
    seriesCount: model.series.length,
    series: model.series.map(s => ({
      name: s.name,
      type: s.type,
      dataLength: s.data?.length || 0,
      pointsLength: s.points?.length || 0,
      hasOhlc: !!s.ohlc,
      color: s.color,
      lineStyle: s.lineStyle,
      lineWidth: s.lineWidth,
      smooth: s.smooth,
      marker: s.marker,
    })),
    categories: model.categories,
    barDirection: model.barDirection,
    grouping: model.grouping,
    is3D: model.is3D,
    hasLegend: !!model.legend,
    legendPosition: model.legend?.position,
    hasSecondaryAxis: !!model.yAxisSecondary,
    plotGroups: model.plotGroups,
  };
}

// ===== 测试用例 =====

describe('Chart Parser — Theme Parser', () => {
  it('should parse default theme correctly', () => {
    const themeXml = `<?xml version="1.0"?>
      <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:themeElements>
          <a:clrScheme name="Test">
            <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
            <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
            <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
            <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
            <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
            <a:accent2><a:srgbClr val="C0504D"/></a:accent2>
            <a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
            <a:accent4><a:srgbClr val="8064A2"/></a:accent4>
            <a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
            <a:accent6><a:srgbClr val="F79646"/></a:accent6>
            <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
            <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
          </a:clrScheme>
          <a:fontScheme name="Test">
            <a:majorFont><a:latin typeface="Cambria"/></a:majorFont>
            <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
          </a:fontScheme>
        </a:themeElements>
      </a:theme>`;

    const theme = parseThemeXml(xmlParser.parse(themeXml));
    expect(theme.colors.accent1).toBe('#4F81BD');
    expect(theme.colors.accent2).toBe('#C0504D');
    expect(theme.colors.dk1).toBe('#000000');
    expect(theme.colors.lt1).toBe('#FFFFFF');
    expect(theme.fonts.majorFont).toBe('Cambria');
    expect(theme.fonts.minorFont).toBe('Calibri');
  });
});

describe('Excel Parser — formula calculation', () => {
  it('calculates formulas and cross-sheet references for read-only preview', async () => {
    const workbook = new ExcelJS.Workbook();
    const source = workbook.addWorksheet('Source');
    const summary = workbook.addWorksheet('Summary');
    source.getCell('A1').value = 2;
    source.getCell('A2').value = 3;
    source.getCell('A3').value = { formula: 'SUM(A1:A2)' };
    summary.getCell('A1').value = { formula: "Source!A3*2" };

    const buffer = await workbook.xlsx.writeBuffer();
    const parsed = await parseExcel(buffer);

    expect(parsed.sheets[0].rows[2][0]).toMatchObject({ value: 5, text: '5', type: 'formula' });
    expect(parsed.sheets[1].rows[0][0]).toMatchObject({ value: 10, text: '10', type: 'formula' });
  });

  it('shows Excel-compatible error for unsupported formulas', async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Sheet1');
    const cell = sheet.getCell('A1');
    cell.value = { formula: 'UNSUPPORTED_FUNCTION(1)' };
    cell.numFmt = '0.00';

    const buffer = await workbook.xlsx.writeBuffer();
    const parsed = await parseExcel(buffer);

    expect(parsed.sheets[0].rows[0][0]).toMatchObject({
      value: '#NAME?',
      text: '#NAME?',
      type: 'formula',
    });
  });

  it('cleans up floating-point artifacts in cached numeric values like Excel', async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Sheet1');
    sheet.getCell('A1').value = 0.1 + 0.2;
    sheet.getCell('A2').value = 1 / 3;
    sheet.getCell('A3').value = 37.057000000000002;

    const buffer = await workbook.xlsx.writeBuffer();
    const parsed = await parseExcel(buffer);

    expect(parsed.sheets[0].rows[0][0]).toMatchObject({ text: '0.3', type: 'number' });
    expect(parsed.sheets[0].rows[1][0]).toMatchObject({ text: '0.333333333333333', type: 'number' });
    expect(parsed.sheets[0].rows[2][0]).toMatchObject({ text: '37.057', type: 'number' });
  });
});

describe('Chart Parser — Bar Chart', () => {
  const barChartXml = `<?xml version="1.0"?>
    <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <c:chart>
        <c:plotArea>
          <c:barChart>
            <c:barDir val="col"/>
            <c:grouping val="clustered"/>
            <c:ser>
              <c:idx val="0"/>
              <c:order val="0"/>
              <c:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></c:spPr>
              <c:val>
                <c:numRef>
                  <c:f>Sheet1!$B$2:$B$6</c:f>
                  <c:numCache>
                    <c:formatCode>General</c:formatCode>
                    <c:ptCount val="5"/>
                    <c:pt idx="0"><c:v>10</c:v></c:pt>
                    <c:pt idx="1"><c:v>20</c:v></c:pt>
                    <c:pt idx="2"><c:v>30</c:v></c:pt>
                    <c:pt idx="3"><c:v>40</c:v></c:pt>
                    <c:pt idx="4"><c:v>50</c:v></c:pt>
                  </c:numCache>
                </c:numRef>
              </c:val>
            </c:ser>
            <c:gapWidth val="150"/>
            <c:axId val="1"/>
            <c:axId val="2"/>
          </c:barChart>
          <c:catAx>
            <c:axId val="1"/>
            <c:axPos val="b"/>
            <c:crossAx val="2"/>
          </c:catAx>
          <c:valAx>
            <c:axId val="2"/>
            <c:axPos val="l"/>
            <c:crossAx val="1"/>
          </c:valAx>
        </c:plotArea>
      </c:chart>
    </c:chartSpace>`;

  it('should parse bar chart correctly', () => {
    const model = parseChartXml(barChartXml, 'bar_test');
    expect(model).not.toBeNull();
    expect(model!.type).toBe('bar');
    expect(model!.barDirection).toBe('col');
    expect(model!.grouping).toBe('clustered');
    expect(model!.series).toHaveLength(1);
    // 解析器优先读取 workbook 实时数据，numCache 作为回退
    expect(model!.series[0].data).toHaveLength(5);
    // 确保数据都是数字
    expect(model!.series[0].data!.every(v => typeof v === 'number')).toBe(true);
  });

  it('should match snapshot', () => {
    const model = parseChartXml(barChartXml, 'bar_snapshot');
    expect(sanitizeForSnapshot(model)).toMatchSnapshot();
  });
});

describe('Chart Parser — Line Chart', () => {
  const lineChartXml = `<?xml version="1.0"?>
    <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <c:chart>
        <c:title>
          <c:tx><c:rich><a:p><a:r><a:t>Line Chart Title</a:t></a:r></a:p></c:rich></c:tx>
        </c:title>
        <c:plotArea>
          <c:lineChart>
            <c:grouping val="standard"/>
            <c:ser>
              <c:idx val="0"/>
              <c:order val="0"/>
              <c:spPr>
                <a:ln w="28575">
                  <a:solidFill><a:schemeClr val="accent6"/></a:solidFill>
                </a:ln>
              </c:spPr>
              <c:marker><c:symbol val="none"/></c:marker>
              <c:val>
                <c:numRef>
                  <c:f>Sheet1!$B$2:$B$6</c:f>
                  <c:numCache>
                    <c:ptCount val="5"/>
                    <c:pt idx="0"><c:v>15</c:v></c:pt>
                    <c:pt idx="1"><c:v>25</c:v></c:pt>
                    <c:pt idx="2"><c:v>35</c:v></c:pt>
                    <c:pt idx="3"><c:v>45</c:v></c:pt>
                    <c:pt idx="4"><c:v>55</c:v></c:pt>
                  </c:numCache>
                </c:numRef>
              </c:val>
              <c:smooth val="0"/>
            </c:ser>
            <c:axId val="1"/>
            <c:axId val="2"/>
          </c:lineChart>
          <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
          <c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
        </c:plotArea>
      </c:chart>
    </c:chartSpace>`;

  it('should parse line chart correctly', () => {
    const model = parseChartXml(lineChartXml, 'line_test');
    expect(model).not.toBeNull();
    expect(model!.type).toBe('line');
    // 标题解析 (富文本格式)
    expect(model!.title).toBeDefined();
    expect(model!.series).toHaveLength(1);
    expect(model!.series[0].data).toHaveLength(5);
    expect(model!.series[0].data!.every(v => typeof v === 'number')).toBe(true);
  });

  it('should match snapshot', () => {
    const model = parseChartXml(lineChartXml, 'line_snapshot');
    expect(sanitizeForSnapshot(model)).toMatchSnapshot();
  });
});

describe('Chart Parser — Pie Chart', () => {
  const pieChartXml = `<?xml version="1.0"?>
    <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <c:chart>
        <c:plotArea>
          <c:pieChart>
            <c:varyColors val="1"/>
            <c:ser>
              <c:idx val="0"/>
              <c:order val="0"/>
              <c:val>
                <c:numRef>
                  <c:numCache>
                    <c:ptCount val="4"/>
                    <c:pt idx="0"><c:v>30</c:v></c:pt>
                    <c:pt idx="1"><c:v>20</c:v></c:pt>
                    <c:pt idx="2"><c:v>15</c:v></c:pt>
                    <c:pt idx="3"><c:v>35</c:v></c:pt>
                  </c:numCache>
                </c:numRef>
              </c:val>
            </c:ser>
          </c:pieChart>
        </c:plotArea>
      </c:chart>
    </c:chartSpace>`;

  it('should parse pie chart correctly', () => {
    const model = parseChartXml(pieChartXml, 'pie_test');
    expect(model).not.toBeNull();
    expect(model!.type).toBe('pie');
    expect(model!.series[0].data).toEqual([30, 20, 15, 35]);
  });

  it('should match snapshot', () => {
    const model = parseChartXml(pieChartXml, 'pie_snapshot');
    expect(sanitizeForSnapshot(model)).toMatchSnapshot();
  });
});

describe('Chart Parser — Scatter Chart', () => {
  const scatterXml = `<?xml version="1.0"?>
    <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <c:chart>
        <c:plotArea>
          <c:scatterChart>
            <c:scatterStyle val="lineMarker"/>
            <c:ser>
              <c:idx val="0"/>
              <c:order val="0"/>
              <c:xVal>
                <c:numRef>
                  <c:numCache>
                    <c:ptCount val="3"/>
                    <c:pt idx="0"><c:v>1</c:v></c:pt>
                    <c:pt idx="1"><c:v>2</c:v></c:pt>
                    <c:pt idx="2"><c:v>3</c:v></c:pt>
                  </c:numCache>
                </c:numRef>
              </c:xVal>
              <c:yVal>
                <c:numRef>
                  <c:numCache>
                    <c:ptCount val="3"/>
                    <c:pt idx="0"><c:v>10</c:v></c:pt>
                    <c:pt idx="1"><c:v>20</c:v></c:pt>
                    <c:pt idx="2"><c:v>30</c:v></c:pt>
                  </c:numCache>
                </c:numRef>
              </c:yVal>
            </c:ser>
            <c:axId val="1"/>
            <c:axId val="2"/>
          </c:scatterChart>
          <c:valAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:valAx>
          <c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
        </c:plotArea>
      </c:chart>
    </c:chartSpace>`;

  it('should parse scatter chart correctly', () => {
    const model = parseChartXml(scatterXml, 'scatter_test');
    expect(model).not.toBeNull();
    expect(model!.type).toBe('scatter');
    expect(model!.series[0].points).toHaveLength(3);
    expect(model!.series[0].points![0]).toEqual({ x: 1, y: 10 });
  });

  it('should match snapshot', () => {
    const model = parseChartXml(scatterXml, 'scatter_snapshot');
    expect(sanitizeForSnapshot(model)).toMatchSnapshot();
  });
});

describe('Chart Parser — Area Chart', () => {
  const areaXml = `<?xml version="1.0"?>
    <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <c:chart>
        <c:plotArea>
          <c:areaChart>
            <c:grouping val="standard"/>
            <c:ser>
              <c:idx val="0"/>
              <c:order val="0"/>
              <c:val>
                <c:numRef>
                  <c:numCache>
                    <c:ptCount val="4"/>
                    <c:pt idx="0"><c:v>5</c:v></c:pt>
                    <c:pt idx="1"><c:v>15</c:v></c:pt>
                    <c:pt idx="2"><c:v>25</c:v></c:pt>
                    <c:pt idx="3"><c:v>35</c:v></c:pt>
                  </c:numCache>
                </c:numRef>
              </c:val>
            </c:ser>
            <c:axId val="1"/>
            <c:axId val="2"/>
          </c:areaChart>
          <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
          <c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
        </c:plotArea>
      </c:chart>
    </c:chartSpace>`;

  it('should parse area chart correctly', () => {
    const model = parseChartXml(areaXml, 'area_test');
    expect(model).not.toBeNull();
    expect(model!.type).toBe('area');
    expect(model!.series[0].data).toEqual([5, 15, 25, 35]);
  });

  it('should match snapshot', () => {
    const model = parseChartXml(areaXml, 'area_snapshot');
    expect(sanitizeForSnapshot(model)).toMatchSnapshot();
  });
});

describe('Chart Parser — Combo Chart', () => {
  const comboXml = `<?xml version="1.0"?>
    <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <c:chart>
        <c:plotArea>
          <c:barChart>
            <c:barDir val="col"/>
            <c:grouping val="clustered"/>
            <c:ser>
              <c:idx val="0"/>
              <c:order val="0"/>
              <c:val>
                <c:numRef>
                  <c:numCache>
                    <c:ptCount val="3"/>
                    <c:pt idx="0"><c:v>10</c:v></c:pt>
                    <c:pt idx="1"><c:v>20</c:v></c:pt>
                    <c:pt idx="2"><c:v>30</c:v></c:pt>
                  </c:numCache>
                </c:numRef>
              </c:val>
            </c:ser>
            <c:axId val="1"/>
            <c:axId val="2"/>
          </c:barChart>
          <c:lineChart>
            <c:grouping val="standard"/>
            <c:ser>
              <c:idx val="1"/>
              <c:order val="1"/>
              <c:val>
                <c:numRef>
                  <c:numCache>
                    <c:ptCount val="3"/>
                    <c:pt idx="0"><c:v>5</c:v></c:pt>
                    <c:pt idx="1"><c:v>15</c:v></c:pt>
                    <c:pt idx="2"><c:v>25</c:v></c:pt>
                  </c:numCache>
                </c:numRef>
              </c:val>
            </c:ser>
            <c:axId val="1"/>
            <c:axId val="3"/>
          </c:lineChart>
          <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
          <c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
          <c:valAx><c:axId val="3"/><c:axPos val="r"/><c:crossAx val="1"/></c:valAx>
        </c:plotArea>
      </c:chart>
    </c:chartSpace>`;

  it('should parse combo chart correctly', () => {
    const model = parseChartXml(comboXml, 'combo_test');
    expect(model).not.toBeNull();
    expect(model!.type).toBe('combo');
    expect(model!.series).toHaveLength(2);
    expect(model!.series[0].type).toBe('bar');
    expect(model!.series[1].type).toBe('line');
    expect(model!.series[1].yAxisIndex).toBe(1);
    expect(model!.plotGroups).toBeDefined();
    expect(model!.plotGroups).toHaveLength(2);
  });

  it('should match snapshot', () => {
    const model = parseChartXml(comboXml, 'combo_snapshot');
    expect(sanitizeForSnapshot(model)).toMatchSnapshot();
  });
});

describe('Chart Parser — Stock Chart', () => {
  const stockXml = `<?xml version="1.0"?>
    <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <c:chart>
        <c:plotArea>
          <c:stockChart>
            <c:ser>
              <c:idx val="0"/>
              <c:order val="0"/>
              <c:high>
                <c:numRef><c:numCache>
                  <c:ptCount val="3"/>
                  <c:pt idx="0"><c:v>105</c:v></c:pt>
                  <c:pt idx="1"><c:v>115</c:v></c:pt>
                  <c:pt idx="2"><c:v>125</c:v></c:pt>
                </c:numCache></c:numRef>
              </c:high>
              <c:low>
                <c:numRef><c:numCache>
                  <c:ptCount val="3"/>
                  <c:pt idx="0"><c:v>95</c:v></c:pt>
                  <c:pt idx="1"><c:v>105</c:v></c:pt>
                  <c:pt idx="2"><c:v>110</c:v></c:pt>
                </c:numCache></c:numRef>
              </c:low>
              <c:open>
                <c:numRef><c:numCache>
                  <c:ptCount val="3"/>
                  <c:pt idx="0"><c:v>100</c:v></c:pt>
                  <c:pt idx="1"><c:v>110</c:v></c:pt>
                  <c:pt idx="2"><c:v>115</c:v></c:pt>
                </c:numCache></c:numRef>
              </c:open>
              <c:close>
                <c:numRef><c:numCache>
                  <c:ptCount val="3"/>
                  <c:pt idx="0"><c:v>102</c:v></c:pt>
                  <c:pt idx="1"><c:v>112</c:v></c:pt>
                  <c:pt idx="2"><c:v>120</c:v></c:pt>
                </c:numCache></c:numRef>
              </c:close>
            </c:ser>
            <c:axId val="1"/>
            <c:axId val="2"/>
          </c:stockChart>
          <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
          <c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
        </c:plotArea>
      </c:chart>
    </c:chartSpace>`;

  it('should parse stock chart correctly', () => {
    const model = parseChartXml(stockXml, 'stock_test');
    expect(model).not.toBeNull();
    expect(model!.type).toBe('stock');
    expect(model!.series[0].ohlc).toBeDefined();
    expect(model!.series[0].ohlc!.high).toEqual([105, 115, 125]);
    expect(model!.series[0].ohlc!.low).toEqual([95, 105, 110]);
    expect(model!.series[0].ohlc!.open).toEqual([100, 110, 115]);
    expect(model!.series[0].ohlc!.close).toEqual([102, 112, 120]);
  });

  it('should match snapshot', () => {
    const model = parseChartXml(stockXml, 'stock_snapshot');
    expect(sanitizeForSnapshot(model)).toMatchSnapshot();
  });
});

describe('Chart Model — Renderer Decoupling', () => {
  it('ChartModel should not contain any ECharts-specific fields', () => {
    const barXml = `<?xml version="1.0"?>
      <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <c:chart><c:plotArea>
          <c:barChart>
            <c:barDir val="col"/><c:grouping val="clustered"/>
            <c:ser><c:idx val="0"/><c:order val="0"/>
              <c:val><c:numRef><c:numCache>
                <c:ptCount val="2"/>
                <c:pt idx="0"><c:v>10</c:v></c:pt>
                <c:pt idx="1"><c:v>20</c:v></c:pt>
              </c:numCache></c:numRef></c:val>
            </c:ser>
            <c:axId val="1"/><c:axId val="2"/>
          </c:barChart>
          <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
          <c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
        </c:plotArea></c:chart>
      </c:chartSpace>`;

    const model = parseChartXml(barXml, 'decoupling_test');
    expect(model).not.toBeNull();

    // ChartModel 不应包含任何 ECharts 字段
    const modelStr = JSON.stringify(model);
    expect(modelStr).not.toContain('echarts');
    expect(modelStr).not.toContain('series.type');
    expect(modelStr).not.toContain('grid');
    expect(modelStr).not.toContain('xAxis.type');
    expect(modelStr).not.toContain('itemStyle');
  });
});

describe('Layout Engine', () => {
  it('should compute layout for bar chart', () => {
    const model: ChartModel = {
      id: 'layout_test',
      type: 'bar',
      series: [{ name: 'S1', type: 'bar', data: [10, 20, 30] }],
      categories: ['A', 'B', 'C'],
      xAxis: { type: 'category', visible: true, position: 'bottom' },
      yAxis: { type: 'value', visible: true, position: 'left', splitLine: 'dashed' },
      anchor: DEFAULT_ANCHOR,
      sheetIndex: 0,
    };

    const layout = computeLayout(model, { width: 400, height: 300 });
    expect(layout.container.width).toBe(400);
    expect(layout.container.height).toBe(300);
    expect(layout.plotArea.width).toBeGreaterThan(0);
    expect(layout.plotArea.height).toBeGreaterThan(0);
    expect(layout.isAxisless).toBe(false);
  });

  it('should compute layout for pie chart (axisless)', () => {
    const model: ChartModel = {
      id: 'layout_pie',
      type: 'pie',
      series: [{ name: 'S1', type: 'pie', data: [30, 70] }],
      anchor: DEFAULT_ANCHOR,
      sheetIndex: 0,
    };

    const layout = computeLayout(model, { width: 300, height: 300 });
    expect(layout.isAxisless).toBe(true);
  });
});

describe('ECharts Converter', () => {
  it('should convert bar chart to ECharts option', () => {
    const model: ChartModel = {
      id: 'convert_test',
      type: 'bar',
      series: [{ name: 'Sales', type: 'bar', data: [10, 20, 30] }],
      categories: ['Q1', 'Q2', 'Q3'],
      xAxis: { type: 'category', visible: true, position: 'bottom' },
      yAxis: { type: 'value', visible: true, position: 'left', splitLine: 'dashed' },
      anchor: DEFAULT_ANCHOR,
      sheetIndex: 0,
    };

    const layout = computeLayout(model, { width: 500, height: 400 });
    const option = convertToEChartsOption(model, layout);

    expect(option).toBeDefined();
    expect(option.series).toBeDefined();
    expect(option.series[0].type).toBe('bar');
    expect(option.series[0].data).toEqual([10, 20, 30]);
    expect(option.xAxis).toBeDefined();
    expect(option.yAxis).toBeDefined();
  });
});

describe('Formula Parser — Cell Range References', () => {
  it('should parse simple cell range reference', () => {
    // 测试内联函数 parseCellRangeRef 通过解析包含引用的 XML
    const xmlWithRef = `<?xml version="1.0"?>
      <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart><c:plotArea>
          <c:barChart>
            <c:barDir val="col"/><c:grouping val="clustered"/>
            <c:ser><c:idx val="0"/><c:order val="0"/>
              <c:val><c:numRef>
                <c:f>Sheet1!$B$2:$B$6</c:f>
                <c:numCache>
                  <c:ptCount val="3"/>
                  <c:pt idx="0"><c:v>10</c:v></c:pt>
                  <c:pt idx="1"><c:v>20</c:v></c:pt>
                  <c:pt idx="2"><c:v>30</c:v></c:pt>
                </c:numCache>
              </c:numRef></c:val>
            </c:ser>
            <c:axId val="1"/><c:axId val="2"/>
          </c:barChart>
          <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
          <c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
        </c:plotArea></c:chart>
      </c:chartSpace>`;

    // 使用 mock workbook 测试引用解析
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell('B2').value = 100;
    ws.getCell('B3').value = 200;
    ws.getCell('B4').value = 300;
    ws.getCell('B5').value = 400;
    ws.getCell('B6').value = 500;

    const model = parseChartXml(xmlWithRef, 'ref_test', wb);
    expect(model).not.toBeNull();
    // 应该优先使用 workbook 实时数据
    expect(model!.series[0].data).toEqual([100, 200, 300, 400, 500]);
  });

  it('should fall back to numCache when workbook data is unavailable', () => {
    const xmlWithRef = `<?xml version="1.0"?>
      <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart><c:plotArea>
          <c:barChart>
            <c:barDir val="col"/><c:grouping val="clustered"/>
            <c:ser><c:idx val="0"/><c:order val="0"/>
              <c:val><c:numRef>
                <c:f>Sheet1!$B$2:$B$6</c:f>
                <c:numCache>
                  <c:ptCount val="3"/>
                  <c:pt idx="0"><c:v>10</c:v></c:pt>
                  <c:pt idx="1"><c:v>20</c:v></c:pt>
                  <c:pt idx="2"><c:v>30</c:v></c:pt>
                </c:numCache>
              </c:numRef></c:val>
            </c:ser>
            <c:axId val="1"/><c:axId val="2"/>
          </c:barChart>
          <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
          <c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
        </c:plotArea></c:chart>
      </c:chartSpace>`;

    // 空 workbook（无 Sheet1）→ 应该回退到 numCache
    const wb = new ExcelJS.Workbook();

    const model = parseChartXml(xmlWithRef, 'ref_fallback', wb);
    expect(model).not.toBeNull();
    expect(model!.series[0].data).toEqual([10, 20, 30]);
  });

  it('should parse quoted sheet name reference', () => {
    // 'Page 1'!$C$8:$C$28 格式的引用
    const xmlWithQuotedRef = `<?xml version="1.0"?>
      <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart><c:plotArea>
          <c:barChart>
            <c:barDir val="col"/><c:grouping val="clustered"/>
            <c:ser><c:idx val="0"/><c:order val="0"/>
              <c:val><c:numRef>
                <c:f>'Page 1'!$C$8:$C$10</c:f>
                <c:numCache>
                  <c:ptCount val="3"/>
                  <c:pt idx="0"><c:v>5</c:v></c:pt>
                  <c:pt idx="1"><c:v>15</c:v></c:pt>
                  <c:pt idx="2"><c:v>25</c:v></c:pt>
                </c:numCache>
              </c:numRef></c:val>
            </c:ser>
            <c:axId val="1"/><c:axId val="2"/>
          </c:barChart>
          <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
          <c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
        </c:plotArea></c:chart>
      </c:chartSpace>`;

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Page 1');
    ws.getCell('C8').value = 50;
    ws.getCell('C9').value = 60;
    ws.getCell('C10').value = 70;

    const model = parseChartXml(xmlWithQuotedRef, 'quoted_ref', wb);
    expect(model).not.toBeNull();
    expect(model!.series[0].data).toEqual([50, 60, 70]);
  });
});
