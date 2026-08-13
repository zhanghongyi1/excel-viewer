# @excel-preview/core

浏览器端 `.xlsx` 只读预览组件。解析 OpenXML 工作簿并渲染表格、公式、图片和图表；无 Vue/React 依赖。

![工作簿处理链路](https://raw.githubusercontent.com/zhanghongyi1/excel-viewer/main/docs/assets/excel-to-echarts.svg)

## 安装

```bash
pnpm add @excel-preview/core
# 或 npm install @excel-preview/core
```

## Vite

在**使用本包的 Vite 项目**中加入以下配置。它避免 Vite 预构建 core 时改写 ExcelJS 默认导入，并处理 HyperFormula 的 CommonJS 互操作。

```ts
// vite.config.ts
import { defineConfig } from 'vite';

export default defineConfig({
  optimizeDeps: {
    exclude: ['@excel-preview/core'],
    include: [
      '@excel-preview/core > exceljs',
      '@excel-preview/core > fast-xml-parser',
      '@excel-preview/core > hyperformula',
      '@excel-preview/core > jszip',
    ],
  },
});
```

如已配置 `optimizeDeps`，合并数组即可。若 ExcelJS 被解析到 Node 入口，再加入：

```ts
resolve: {
  alias: { exceljs: 'exceljs/dist/exceljs.min.js' },
},
```

修改后重启 Vite；仍命中旧缓存时删除使用方项目的 `node_modules/.vite`。

## 最小示例

挂载目标需要明确高度：

```html
<div id="excel-viewer" style="height: 700px"></div>
```

```ts
import { ExcelViewer } from '@excel-preview/core';

const viewer = new ExcelViewer({
  target: '#excel-viewer',
  src: '/reports/quarterly.xlsx',
  onError: console.error,
});

// 后续加载：URL / File / Blob / ArrayBuffer
await viewer.render(file);
viewer.setSheet('汇总');

// 路由卸载、弹窗关闭时调用
viewer.destroy();
```

需要认证请求时，先调用 `loadData()` 再传入 `render()`：

```ts
import { ExcelViewer, loadData } from '@excel-preview/core';

const buffer = await loadData('/api/report.xlsx', {
  headers: { Authorization: 'Bearer <token>' },
  withCredentials: true,
});
await viewer.render(buffer);
```

URL 源需要服务端正确配置 CORS；错误会传给 `onError`。

## ExcelViewer API

### 选项

| 选项 | 类型 / 默认值 | 说明 |
| --- | --- | --- |
| `target` | `HTMLElement \| string` | 挂载节点或选择器 |
| `src` | `string \| File \| Blob \| ArrayBuffer` | 初始数据源 |
| `width` / `height` | `'100%'` | 容器尺寸 |
| `showToolbar` | `true` | Sheet 标签栏和缩放控件 |
| `initialZoom` | `100`，范围 50–200 | 初始缩放百分比 |
| `extraColCount` / `extraRowCount` | `5` / `20` | 已用区域后的额外空白列/行 |
| `chartBackend` | `'echarts'` | `'echarts'`、`'canvas'` 或 `'auto'` |
| `echartsRenderer` | `'svg'` | `'svg'` 或 `'canvas'` |
| `echarts` | `any` | 宿主已加载的 ECharts 实例 |
| `parsePivotTables` | `false` | 是否解析透视表缓存；大文件会增加开销 |
| `onRendered` | `() => void` | 渲染完成回调 |
| `onError` | `(error) => void` | 加载或渲染错误回调 |
| `onSheetChange` | `(name, index) => void` | Sheet 切换回调 |

### 方法

| 方法 | 说明 |
| --- | --- |
| `mount(target)` | 挂载到节点 |
| `render(source?)` | 加载并渲染；支持重复调用加载新文件 |
| `setSheet(indexOrName)` | 切换可见工作表 |
| `getWorkbook()` | 获取 `ParsedWorkbook`，未加载时为 `null` |
| `destroy()` | 销毁图表、图片与 DOM 引用 |

## 功能范围

| 类别 | 支持内容 |
| --- | --- |
| 单元格与样式 | 文本、数值、日期、布尔值、公式、富文本、超链接、批注、数字格式、字体、填充、边框、对齐、换行、Office 主题色 |
| 表格结构 | 合并单元格、隐藏行列、冻结窗格、可见 Sheet 切换、50%–200% 缩放 |
| 条件格式 | 数值比较、受限单元格表达式、色阶、数据条 |
| 图片 | 嵌入图片；按 Sheet 和单元格锚点定位，随滚动/缩放/尺寸变化更新 |
| 透视表 | 可选读取缓存字段、记录、行列字段、数据字段；不提供筛选/拖拽/刷新 |
| 图表 | 按 OpenXML 图表模型和锚点渲染；标题、图例、轴、标签、系列颜色、主题色、堆叠和主次 Y 轴 |

隐藏和 veryHidden Sheet 不会进入预览。

## 图表

![Excel 图表到 ECharts 映射](https://raw.githubusercontent.com/zhanghongyi1/excel-viewer/main/docs/assets/echarts-mapping.svg)

```ts
import * as echarts from 'echarts';

const viewer = new ExcelViewer({
  target: '#excel-viewer',
  echarts,
  chartBackend: 'echarts',
  echartsRenderer: 'svg',
});
```

| 类型 | ECharts | Canvas |
| --- | --- | --- |
| 柱/条、折线、面积、饼/环、散点/气泡、雷达、股价 | 支持 | 支持基础版本 |
| 组合图（柱/线/面积、主次 Y 轴） | 支持 | 不支持完整组合语义 |
| 曲面、瀑布、漏斗 | 热力图/柱状图/漏斗降级 | 不支持 |

图表数据优先读取工作簿单元格引用；引用不可用时回退图表 XML 的 `numCache` / `strCache`。3D 图表保留识别标记，但以二维形式渲染。

## 公式

![动态计算与缓存回退](https://raw.githubusercontent.com/zhanghongyi1/excel-viewer/main/docs/assets/dynamic-calculation.svg)

使用 HyperFormula 在加载期计算，不会修改原文件。

```excel
=A1+B1
=SUM(B2:B10)
=IF(C2>0, C2, 0)
=SUM(Source!B2:B10)
```

支持算术、比较、百分比、同表/跨表/区域引用，以及 `SUM`、`AVERAGE`、`MIN`、`MAX`、`COUNT`、`COUNTA`、`IF`、`AND`、`OR`、`NOT`、`ROUND`、`ABS`、`DATE`、`YEAR`、`MONTH`、`DAY` 等常用函数。

计算顺序：**HyperFormula 结果 → Excel 文件缓存结果 → 公式文本或 `#NAME?` / `#DIV/0!` 等错误值**。日期结果会按单元格数字格式显示。

## 低层 API

除 `ExcelViewer` 外，还可按需导入：

| 导出 | 用途 |
| --- | --- |
| `loadData`、`isUrlSource`、`isBinarySource` | 数据源加载与判断 |
| `parseExcel`、`loadRawWorkbook` | 解析 `ParsedWorkbook` 或获取 ExcelJS Workbook |
| `parseCharts`、`parseChartXmlToModel` | 图表关系/XML/模型解析 |
| `parsePivotTables` | 透视表缓存解析 |
| `computeLayout`、`convertToEChartsOption` | 图表布局和 ECharts option 转换 |
| `TableRenderer`、`ChartRenderer`、`ImageRenderer` | 自定义渲染流程 |

## 限制

- 仅支持 `.xlsx` / OpenXML；不支持旧版 `.xls`。
- 只读，不提供编辑与保存回写。
- 不执行 VBA、宏、外部工作簿引用和数据表公式。
- 不覆盖全部 Excel / HyperFormula 函数；不支持的公式按上述策略回退。
- 不支持数据验证、切片器、图标集条件格式及完整数据透视表交互。
- 不提供像素级 3D 图表还原；复杂图表可能按后端能力降级。
- 超大工作簿会创建较多 DOM 节点，请控制文件大小和工作表规模。

## 开发与发布

```bash
pnpm install
pnpm --filter @excel-preview/core build
pnpm --filter @excel-preview/core test
pnpm --filter @excel-preview/core pack --dry-run
```

发布走 GitHub Actions，见仓库根目录 [PUBLISH.md](../../PUBLISH.md)。
