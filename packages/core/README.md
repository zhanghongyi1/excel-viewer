# @excel-preview/core

浏览器端 Excel `.xlsx` 预览组件：渲染工作表、样式、合并单元格、图片和 OpenXML 图表。核心包不依赖 Vue、React 等框架，可直接用于原生 JavaScript 或任意前端框架。

## 功能总览

`@excel-preview/core` 的定位是**只读工作簿预览**：以 `.xlsx` 的 OpenXML 结构为输入，尽可能还原报表中的数据、版式、计算结果和图表，而不是将其转成图片或只输出单元格文本。

| 能力 | 组件提供的功能 | 使用价值 |
| --- | --- | --- |
| 多源加载 | URL、`File`、`Blob`、`ArrayBuffer` | 文件选择、接口下载、对象存储和内存数据使用同一套 API |
| 工作表浏览 | 仅解析可见 Sheet、底部 Sheet 标签切换、滚动与 50%–200% 缩放 | 可直接阅读多工作表报表，无需下载到本地 Excel |
| 单元格与版式 | 文本、数值、日期、布尔值、公式、富文本、超链接、批注、数字格式、合并单元格、隐藏行列、冻结窗格 | 表头层级、金额/日期/百分比格式与备注等业务上下文可保留 |
| 样式与强调 | 字体、填充、边框、对齐、换行、Office 主题色；数值比较、受限表达式、色阶与数据条条件格式 | 报表中的重点、异常和趋势不会退化为纯文本 |
| 动态计算 | 基于 HyperFormula 计算同表/跨表公式、聚合、条件与常用日期函数；提供缓存结果回退 | 常见汇总报表在浏览器中可按当前引用关系展示 |
| 图表与图片 | 解析 OpenXML 图表与嵌入图片，按单元格锚点定位；支持 ECharts 和 Canvas 后端 | 图表保持在原工作表中的业务位置，而不是成为静态附件 |
| 结构化能力 | 提供工作簿、图表、主题、布局和数据透视表缓存的解析 API | 接入方可在预览之外继续做摘要、审计或二次展示 |

### 工作簿处理流程

![Excel 工作簿到浏览器预览的处理链路](https://raw.githubusercontent.com/zhanghongyi1/excel-viewer/main/docs/assets/excel-to-echarts.svg)

```text
.xlsx / URL / File / Blob / ArrayBuffer
                │
                ▼
       OpenXML + ExcelJS 解析
                │
        ┌───────┼────────┐
        ▼       ▼        ▼
   表格与样式  公式计算  图表/图片锚点
        │       │        │
        └───────┴────────┘
                ▼
     浏览器 DOM 表格 + ECharts / Canvas 图表
```

## 安装

```bash
pnpm add @excel-preview/core
# 或 npm install @excel-preview/core
```

### Vite 集成：依赖预构建配置

在使用 `@excel-preview/core` 的 Vite 项目中，建议将以下配置加入**使用方项目**的 `vite.config.ts`。这不是组件运行时 API，也不需要加到 Excel 文件处理逻辑中。

原因是：Vite 预构建本包时，可能将 ExcelJS 的一个默认导入改写为动态命名空间导入，从而导致 `Workbook` 在开发服务器中不可用；同时，HyperFormula 的嵌套依赖需要经过 Vite 的 CommonJS 互操作转换。

```ts
import { defineConfig } from 'vite';

export default defineConfig({
  optimizeDeps: {
    // 避免预构建 core 时改写 ExcelJS 的默认导入。
    exclude: ['@excel-preview/core'],
    // 为 core 的依赖启用 Vite 预构建 / CommonJS 互操作。
    include: [
      '@excel-preview/core > exceljs',
      '@excel-preview/core > fast-xml-parser',
      '@excel-preview/core > hyperformula',
      '@excel-preview/core > jszip',
    ],
  },
});
```

若项目已有 `optimizeDeps`，请将上述 `exclude` / `include` 项合并进现有数组，不要覆盖已有依赖。若已为 ExcelJS 配置浏览器端入口别名，也可保留，例如：

```ts
resolve: {
  alias: {
    exceljs: 'exceljs/dist/exceljs.min.js',
  },
},
```

该配置只影响 Vite 的本地依赖预构建；不影响组件的运行时 API，也不需要修改工作簿文件。配置生效后，请重启 Vite 开发服务器；若 Vite 仍复用旧的依赖预构建缓存，可删除使用方项目的 `node_modules/.vite` 缓存后重启。

## 快速开始

HTML 中准备一个具有明确高度的挂载节点：

```html
<div id="excel-viewer" style="height: 700px"></div>
```

```ts
import { ExcelViewer } from '@excel-preview/core';

const viewer = new ExcelViewer({
  target: '#excel-viewer',
  src: '/reports/quarterly-report.xlsx',
  onError: (error) => console.error('Excel 加载失败：', error),
});
```

未传入 `src` 时，组件会显示一个空白 `Sheet1`。之后可通过 `render()` 加载文件：

```ts
const input = document.querySelector<HTMLInputElement>('#file-input')!;

input.addEventListener('change', async () => {
  const file = input.files?.[0];
  if (file) await viewer.render(file);
});
```

使用完毕请销毁实例：

```ts
viewer.destroy();
```

## 数据源

`render()` 和 `src` 支持以下输入：

```ts
await viewer.render('/files/report.xlsx');     // URL
await viewer.render(file);                     // File
await viewer.render(blob);                     // Blob
await viewer.render(arrayBuffer);              // ArrayBuffer
```

从 URL 加载时，目标服务器必须允许浏览器跨域访问（CORS）。当前仅支持 `.xlsx` / OpenXML 工作簿，不支持旧版 `.xls` 二进制格式。

### URL 请求与错误处理

高阶 `ExcelViewer` 使用 `src` 或 `render()` 读取数据源。需要自定义 URL 请求方法、请求头或凭证时，可先使用底层 `loadData()` 获取 `ArrayBuffer`，再传给 `render()`：

```ts
import { ExcelViewer, loadData } from '@excel-preview/core';

const buffer = await loadData('/api/reports/monthly.xlsx', {
  headers: { Authorization: 'Bearer <token>' },
  withCredentials: true,
});

const viewer = new ExcelViewer({ target: '#excel-viewer' });
await viewer.render(buffer);
```

URL 返回 `403`、`404`、非成功状态或浏览器 CORS 拦截时，加载会失败并触发 `onError`。建议接入方在页面上提供加载中和错误提示，并确保资源服务允许浏览器读取文件。

## 表格、样式与交互

### 已还原的工作表信息

| Excel 信息 | 浏览器预览行为 |
| --- | --- |
| 文本、数值、日期、布尔值和数字格式 | 按单元格值与格式显示；公式得到的日期序列号会结合数字格式转换为日期 |
| 字体、填充、边框、对齐、换行 | 映射为网页单元格样式，并解析 Office 主题色引用 |
| 合并单元格 | 保留跨行、跨列的合并区域，适用于多级表头和汇总标题 |
| 隐藏行列与隐藏 Sheet | 隐藏行列不渲染；隐藏与 veryHidden 工作表不进入预览 |
| 冻结窗格 | 对应表头和首列保持固定，便于滚动阅读大表 |
| 富文本、超链接与批注 | 保留富文本片段和链接文本；批注以单元格提示方式展示 |
| 条件格式 | 支持数值比较、受限的单元格比较表达式、色阶与数据条 |
| Sheet 标签栏与缩放 | 默认显示工作表标签和 50%–200% 缩放控制；可用 `showToolbar: false` 隐藏 |

图片与图表以 Excel 的起止单元格及偏移量为锚点渲染。切换工作表、调整缩放、行高或列宽变化时，组件会重新计算其位置；它们会随工作表内容滚动。

## 图表渲染

默认会在工作簿包含图表时加载 ECharts 并以 SVG 渲染。若项目已自行引入 ECharts，建议注入实例，避免重复加载：

![Excel 图表到 ECharts 的结构化映射](https://raw.githubusercontent.com/zhanghongyi1/excel-viewer/main/docs/assets/echarts-mapping.svg)

```ts
import * as echarts from 'echarts';
import { ExcelViewer } from '@excel-preview/core';

const viewer = new ExcelViewer({
  target: '#excel-viewer',
  echarts,
  chartBackend: 'echarts',
  echartsRenderer: 'svg', // 也可为 'canvas'
});
```

也可以使用内置 Canvas 后端（不使用 ECharts）：

```ts
const viewer = new ExcelViewer({
  target: '#excel-viewer',
  chartBackend: 'canvas',
});
```

`chartBackend` 可选值为 `'echarts'`、`'canvas'` 和 `'auto'`。`auto` 在已注入 ECharts 时使用 ECharts，否则回退到 Canvas。

### 图表类型

解析器会读取 `xl/charts/chart*.xml` 中的 OOXML 图表定义，并根据图表元素生成图表模型。当前可识别的类型如下：

| 类型            | 说明                                           | ECharts 后端    | Canvas 后端          |
| --------------- | ---------------------------------------------- | --------------- | -------------------- |
| 柱状图 / 条形图 | 簇状、堆积、百分比堆积；支持垂直柱状与水平条形 | 原生映射        | 支持                 |
| 折线图          | 普通折线、平滑线、数据点标记                   | 原生映射        | 支持                 |
| 面积图          | 普通、堆积及透明区域填充                       | 折线 + 面积填充 | 支持                 |
| 饼图 / 环形图   | 分类数据及数据标签                             | 饼图映射        | 支持                 |
| 散点图 / 气泡图 | X/Y 数据点；气泡图读取气泡大小                 | 散点映射        | 支持基础散点         |
| 雷达图          | 多维分类数据；支持标准、标记、填充样式         | 原生映射        | 支持                 |
| 股价图          | OHLC / K 线数据，包含可选成交量                | K 线映射        | 支持基础 OHLC        |
| 组合图          | 柱、线、面积等多类型系列混合；支持主次 Y 轴    | 支持            | 不支持完整组合图语义 |
| 曲面图          | 读取为曲面数据                                 | 降级为热力图    | 不支持               |
| 瀑布图          | 读取新式或 classic chartSpace 定义             | 降级为柱状图    | 不支持               |
| 漏斗图          | 读取新式或 classic chartSpace 定义             | 漏斗图映射      | 不支持               |

通用图表属性包括标题、图例位置、分类轴和数值轴、轴范围、标签旋转、系列颜色、数据标签、堆积方式、图表锚点及主题色。OOXML 中的 3D 柱状图、折线图、面积图、饼图和曲面图会保留 `is3D` 标记，但当前仍以二维方式渲染；这不等同于完整的三维透视效果。

图表数据优先从工作簿单元格引用读取，缺少可用引用时回退到图表 XML 中的 `numCache` / `strCache` 缓存值。因此，工作簿中的公式结果应在加载时可计算，或文件本身包含可用的缓存结果。

### 公式表达式

单元格公式会以 HyperFormula 语法计算，支持范围、同表和跨工作表引用，以及常见运算符：

```excel
=A1+B1
=A1*10%
=SUM(B2:B10)
=IF(C2>0, C2, 0)
=Source!A3*2
=SUM(Source!B2:B10)
```

常见可用函数包括 `SUM`、`AVERAGE`、`MIN`、`MAX`、`COUNT`、`COUNTA`、`IF`、`AND`、`OR`、`NOT`、`ROUND`、`ABS`、`DATE`、`YEAR`、`MONTH` 和 `DAY`。这不是完整函数清单，实际支持范围以当前依赖版本的 HyperFormula 为准；不支持的函数会产生 Excel 兼容错误（例如 `#NAME?`）。

公式支持的主要表达能力包括：

- 算术运算：`+`、`-`、`*`、`/`、`^`，以及括号和百分比
- 比较与逻辑：`=`、`<>`、`>`、`>=`、`<`、`<=`，配合 `IF`、`AND`、`OR`、`NOT`
- 单元格和区域引用：`A1`、`$A$1`、`A1:B10`、`Sheet2!A1`；含空格的工作表名称使用 Excel 的单引号写法
- 文本和日期值：文本参与函数计算时遵循 HyperFormula 的类型转换规则；日期结果会依据单元格数字格式还原为日期显示

以下内容不应按“完整 Excel 公式引擎”理解：宏和 VBA、外部工作簿引用、数据表公式、部分新版动态数组或 Excel 专有函数，以及 HyperFormula 尚未实现的函数。计算只发生在预览加载阶段，不会修改或保存原工作簿。

### 公式计算与回退策略

![浏览器端公式计算与缓存回退流程](https://raw.githubusercontent.com/zhanghongyi1/excel-viewer/main/docs/assets/dynamic-calculation.svg)

公式预览遵循“**计算结果优先，文件缓存兜底**”的顺序：

1. 加载时，组件为工作簿的工作表建立 HyperFormula 计算上下文，并写入单元格内容和公式。
2. 对可计算的公式，使用计算引擎结果并按原单元格数字格式显示。
3. 公式无法计算时，优先使用 `.xlsx` 文件中保存的可用缓存结果。
4. 缓存也不可用时，保留公式文本或显示 Excel 兼容错误值，例如 `#NAME?`、`#DIV/0!`。

因此，本包可用于阅读“数据 + 公式 + 图表”型业务报表，但不应被用作完整 Excel 编辑器或公式兼容性验证工具。

## 图片与数据透视表缓存

### 图片

工作簿中的嵌入图片会被解析为 `ParsedImage`，并按所属 Sheet 和单元格锚点叠加到预览区域。图片保持原始宽高比，并与图表一样随滚动、缩放和尺寸变化重新定位。

### 数据透视表缓存

`parsePivotTables` 默认关闭，因为解析大型工作簿的数据透视表缓存会增加开销。开启后，组件会将解析结果挂在 `getWorkbook()?.pivotTables` 上；其中包含缓存字段、缓存记录、行/列字段与数据字段定义。

```ts
const viewer = new ExcelViewer({
  target: '#excel-viewer',
  parsePivotTables: true,
});

await viewer.render(file);
const pivotTables = viewer.getWorkbook()?.pivotTables ?? [];
```

这项能力用于读取与利用透视表缓存数据，**不提供** Excel 数据透视表的筛选、拖拽或刷新交互。

## API

### `new ExcelViewer(options)`

| 选项                   | 类型                                      | 说明                                    |
| ---------------------- | ----------------------------------------- | --------------------------------------- |
| `target`             | `HTMLElement \| string`                  | 挂载节点或 CSS 选择器                   |
| `src`                | `string \| File \| Blob \| ArrayBuffer`    | 可选的初始工作簿                        |
| `width` / `height` | `string`                                | 容器尺寸，默认均为`100%`              |
| `showToolbar`        | `boolean`                               | 是否显示底部 Sheet 标签栏，默认`true` |
| `initialZoom`        | `number`                                | 初始缩放比例，范围 50–200，默认`100` |
| `extraColCount`      | `number`                                | 已用区域右侧额外渲染空白列数，默认`5` |
| `extraRowCount`      | `number`                                | 已用区域下方额外渲染空白行数，默认`20` |
| `echarts`            | `any`                                   | 可选的 ECharts 实例                     |
| `chartBackend`       | `'echarts' \| 'canvas' \| 'auto'`         | 图表渲染后端，默认`'echarts'`         |
| `echartsRenderer`    | `'svg' \| 'canvas'`                      | ECharts 渲染器，默认`'svg'`           |
| `parsePivotTables`   | `boolean`                               | 是否解析数据透视表缓存，默认`false`   |
| `onRendered`         | `() => void`                            | 渲染成功回调                            |
| `onError`            | `(error: Error) => void`                | 渲染失败回调                            |
| `onSheetChange`      | `(name: string, index: number) => void` | 工作表切换回调                          |

### 实例方法

| 方法                      | 说明                                     |
| ------------------------- | ---------------------------------------- |
| `mount(target)`         | 挂载到一个节点                           |
| `render(source?)`       | 加载并渲染工作簿                         |
| `setSheet(indexOrName)` | 按索引或名称切换工作表                   |
| `getWorkbook()`         | 获取解析后的工作簿；未加载时返回`null` |
| `destroy()`             | 释放图表、图片和 DOM 资源                |

### 配置与生命周期示例

```ts
const viewer = new ExcelViewer({
  target: '#excel-viewer',
  width: '100%',
  height: '720px',
  showToolbar: true,
  initialZoom: 100,
  parsePivotTables: false,
  chartBackend: 'echarts',
  echartsRenderer: 'svg',
  onRendered: () => console.log('工作簿渲染完成'),
  onSheetChange: (name, index) => console.log(`切换到 ${index}: ${name}`),
  onError: (error) => console.error('Excel 预览失败：', error),
});

await viewer.render(file);
viewer.setSheet('汇总');

// 页面卸载、弹窗关闭或替换预览器时调用。
viewer.destroy();
```

`render()` 可被重复调用加载新文件。组件会清理旧的图表、图片与 DOM 引用，并通过渲染版本控制避免先发起的异步加载覆盖后发起的文件。

## 低层解析与渲染 API

大多数场景只需要 `ExcelViewer`。如果接入方需要把解析结果用于自定义界面、审计或二次分析，可按需使用以下导出：

| 导出 | 作用 |
| --- | --- |
| `loadData()` / `isUrlSource()` / `isBinarySource()` | 将 URL、文件或二进制输入规范化，并判断输入类型 |
| `parseExcel()` / `loadRawWorkbook()` | 解析工作簿为 `ParsedWorkbook`，或取得 ExcelJS 原始工作簿用于高级操作 |
| `parseCharts()` / `parseChartXmlToModel()` | 解析图表关系、图表 XML 与图表模型 |
| `parsePivotTables()` | 独立读取数据透视表缓存 |
| `convertToEChartsOption()` / `computeLayout()` | 将 `ChartModel` 计算为 ECharts 配置与布局 |
| `TableRenderer` / `ChartRenderer` / `ImageRenderer` | 需要自行组织渲染流程时使用；普通场景不必直接实例化 |

低层 API 面向具备 OpenXML、ExcelJS 或自定义渲染需求的接入方；常规业务页面优先使用 `ExcelViewer`，以便自动处理 Sheet 切换、锚点定位与资源释放。

## 已支持内容

- 单元格文本、数值、日期、布尔值、超链接、富文本、公式结果及数字格式
- 使用 HyperFormula 在加载时计算公式，支持同表引用、跨工作表引用、算术和逻辑运算、代表性聚合函数及日期结果还原；详细范围见“公式表达式”
- 对无法计算的公式优先使用 Excel 文件内的缓存结果；仍不可用时保留公式文本或显示 Excel 兼容错误值（如 `#NAME?`、`#DIV/0!`）
- 字体、填充、边框、对齐、换行、数字格式、批注提示
- 合并单元格、隐藏行列、冻结窗格、工作表切换
- 条件格式：数值比较、受限的单元格比较表达式、色阶与数据条
- OpenXML 图表：柱状、折线、面积、饼、环形、散点、气泡、雷达、股价、组合、曲面、瀑布和漏斗
- 图表标题、图例、系列样式、坐标轴、数据标签、主题色、锚点定位及嵌入图片
- ECharts 与 Canvas 两种图表渲染后端；Canvas 适合常见二维图表，完整类型映射请参见上表

## 注意事项

- 本包是只读预览器，不提供单元格编辑或保存能力。
- 复杂或 HyperFormula 未实现的 Excel 函数、宏、数据验证、切片器、图标集条件格式及完整的数据透视表交互不在当前支持范围内。公式计算仅用于只读预览，不会回写工作簿。
- 公式计算不保证覆盖 Excel 与 HyperFormula 的全部兼容差异；当计算引擎不可用或公式无法求值时，会按前述缓存结果回退策略处理。
- 3D 图表目前不提供真正的三维渲染；曲面、瀑布、漏斗图在不同后端可能以降级形式显示。
- 超大工作簿会创建较多 DOM 节点；建议在接入前控制文件大小与工作表规模。
- 图表和图片会跟随工作表滚动；调整行高或列宽后会自动重新定位。

## 本地开发

```bash
pnpm install
pnpm build
pnpm --filter @excel-preview/core test
pnpm dev
```

`pnpm dev` 启动 playground，默认地址为 `http://localhost:3000`。

## 发布

```bash
pnpm --filter @excel-preview/core build
pnpm --filter @excel-preview/core test
pnpm --filter @excel-preview/core pack --dry-run
```

确认产物包含 `dist/` 与本 README 后，发布请走 GitHub Actions CI（支持 provenance），完整流程见仓库根目录 [PUBLISH.md](../../PUBLISH.md)。
