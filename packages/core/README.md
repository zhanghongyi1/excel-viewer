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

## API

### `new ExcelViewer(options)`

| 选项 | 类型 | 说明 |
| --- | --- | --- |
| `target` | `HTMLElement \| string` | 挂载节点或 CSS 选择器 |
| `src` | `string \| File \| Blob \| ArrayBuffer` | 可选的初始工作簿 |
| `width` / `height` | `string` | 容器尺寸，默认均为 `100%` |
| `showToolbar` | `boolean` | 是否显示底部 Sheet 标签栏，默认 `true` |
| `echarts` | `any` | 可选的 ECharts 实例 |
| `chartBackend` | `'echarts' \| 'canvas' \| 'auto'` | 图表渲染后端，默认 `'echarts'` |
| `echartsRenderer` | `'svg' \| 'canvas'` | ECharts 渲染器，默认 `'svg'` |
| `parsePivotTables` | `boolean` | 是否解析数据透视表缓存，默认 `false` |
| `onRendered` | `() => void` | 渲染成功回调 |
| `onError` | `(error: Error) => void` | 渲染失败回调 |
| `onSheetChange` | `(name: string, index: number) => void` | 工作表切换回调 |

### 实例方法

| 方法 | 说明 |
| --- | --- |
| `mount(target)` | 挂载到一个节点 |
| `render(source?)` | 加载并渲染工作簿 |
| `setSheet(indexOrName)` | 按索引或名称切换工作表 |
| `getWorkbook()` | 获取解析后的工作簿；未加载时返回 `null` |
| `destroy()` | 释放图表、图片和 DOM 资源 |

## 已支持内容

- 单元格文本、数值、日期、公式结果及基础公式计算
- 字体、填充、边框、对齐、换行、数字格式、批注提示
- 合并单元格、隐藏行列、冻结窗格、工作表切换
- 条件格式：数值比较、部分表达式、色阶与数据条
- OpenXML 图表、锚点定位及嵌入图片
- ECharts 与 Canvas 两种图表渲染后端

## 注意事项

- 本包是只读预览器，不提供单元格编辑或保存能力。
- 复杂 Excel 公式、宏、数据验证、切片器、图标集条件格式及完整的数据透视表交互不在当前支持范围内。
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

## 发布前检查

```bash
pnpm --filter @excel-preview/core build
pnpm --filter @excel-preview/core test
pnpm --filter @excel-preview/core pack --dry-run
```

确认产物包含 `dist/` 与本 README 后，再在 `packages/core` 目录执行 `npm publish` 或 `pnpm publish`。
