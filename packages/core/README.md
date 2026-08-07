# @excel-preview/core

浏览器端 Excel `.xlsx` 预览组件：渲染工作表、样式、合并单元格、图片和 OpenXML 图表。核心包不依赖 Vue、React 等框架，可直接用于原生 JavaScript 或任意前端框架。

## 安装

```bash
pnpm add @excel-preview/core
# 或 npm install @excel-preview/core
```

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

## 图表渲染

默认会在工作簿包含图表时加载 ECharts 并以 SVG 渲染。若项目已自行引入 ECharts，建议注入实例，避免重复加载：

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

## API

### `new ExcelViewer(options)`

| 选项                   | 类型                                      | 说明                                    |
| ---------------------- | ----------------------------------------- | --------------------------------------- |
| `target`             | `HTMLElement \| string`                  | 挂载节点或 CSS 选择器                   |
| `src`                | `string \| File \| Blob \| ArrayBuffer`    | 可选的初始工作簿                        |
| `width` / `height` | `string`                                | 容器尺寸，默认均为`100%`              |
| `showToolbar`        | `boolean`                               | 是否显示底部 Sheet 标签栏，默认`true` |
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
