# Excel Preview & Chart Render

一个面向浏览器的 Excel `.xlsx` 只读预览方案。项目以 OpenXML 工作簿为输入，在页面中还原表格、样式、公式计算、嵌入图片与图表；核心包为纯 JavaScript，不绑定 Vue、React 或其他框架。

> 核心包：[`@excel-preview/core`](./packages/core/README.md) · 本地演示：[`playground`](./playground)

![Excel 工作簿到浏览器预览的处理链路](./docs/assets/excel-to-echarts.svg)

## 适用场景

- 在业务系统中直接阅读经营、财务、能耗、生产等 Excel 报表，无需下载并打开桌面 Excel。
- 保留“表格 + 公式 + 图表”混合型报表的上下文，而不是只展示二维文本或静态截图。
- 需要把解析后的工作簿、图表或数据透视表缓存交给上层应用做审计、摘要或二次展示。

## 核心能力

| 能力 | 说明 |
| --- | --- |
| 表格与版式还原 | 单元格文本、数值、日期、布尔值、公式、富文本、超链接、批注、数字格式、字体、填充、边框、对齐、换行、合并单元格、隐藏行列、冻结窗格与 Sheet 切换。 |
| 条件格式 | 支持数值比较、受限的单元格比较表达式、色阶与数据条，保留报表对异常和趋势的视觉强调。 |
| 动态公式计算 | 使用 HyperFormula 在加载期计算同表/跨表引用、常见聚合、条件和日期函数；计算失败时回退到文件缓存结果或 Excel 兼容错误值。 |
| OpenXML 图表 | 解析图表类型、数据系列、坐标轴、标题、图例、标签、主题色和单元格锚点；支持 ECharts 优先渲染与 Canvas 后备。 |
| 嵌入图片 | 按所属 Sheet 和单元格锚点叠加图片，随滚动、缩放及尺寸变化重新定位。 |
| 数据透视表缓存 | 可选解析透视表缓存字段、记录、行列字段和数据字段，供二次使用；不提供透视表交互。 |
| 框架无关集成 | 支持 URL、`File`、`Blob`、`ArrayBuffer`；通过 `ExcelViewer` 统一控制加载、切换、回调与销毁。 |

![Excel 图表到 ECharts 的结构化映射](./docs/assets/echarts-mapping.svg)

## 快速开始

### 安装

```bash
pnpm add @excel-preview/core
# 或 npm install @excel-preview/core
```

### 使用

```html
<div id="excel-viewer" style="height: 700px"></div>
```

```ts
import { ExcelViewer } from '@excel-preview/core';

const viewer = new ExcelViewer({
  target: '#excel-viewer',
  src: '/reports/monthly-report.xlsx',
  onRendered: () => console.log('渲染完成'),
  onError: (error) => console.error('加载失败：', error),
});

// File、Blob、URL 与 ArrayBuffer 均可作为数据源。
await viewer.render(file);

// 页面卸载或关闭预览时释放资源。
viewer.destroy();
```

### ECharts 图表渲染

默认图表后端为 ECharts，默认使用 SVG 渲染器。若宿主项目已安装 ECharts，可显式注入实例，避免重复加载：

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

仅需常见二维图表、且不希望使用 ECharts 时，可选择内置 Canvas 后端：

```ts
const viewer = new ExcelViewer({
  target: '#excel-viewer',
  chartBackend: 'canvas',
});
```

支持柱/条、折线、面积、饼/环、散点/气泡、雷达、股价、组合、曲面、瀑布和漏斗等图表。部分高级或 3D 图表会依据后端能力降级为二维表达；完整映射见 [core 图表说明](./packages/core/README.md#图表渲染)。

## 动态计算与兼容回退

![浏览器端公式计算与缓存回退流程](./docs/assets/dynamic-calculation.svg)

公式计算遵循以下优先级：

1. 使用 HyperFormula 对支持的公式进行加载期计算。
2. 无法实时计算时，使用 `.xlsx` 内保存的可用缓存结果。
3. 缓存不存在或不可用时，显示公式文本或 Excel 兼容错误值，例如 `#NAME?`、`#DIV/0!`。

典型支持范围：同表/跨表引用、区域引用、`SUM`、`AVERAGE`、`MIN`、`MAX`、`COUNT`、`COUNTA`、`IF`、`AND`、`OR`、`NOT`、`ROUND`、`ABS` 与常用日期函数。公式只用于**只读预览**，不会修改或保存原工作簿。

## Vite 项目接入提示

在 Vite 开发服务器中使用本包时，应在**使用方项目**的 `vite.config.ts` 中配置依赖预构建，避免 ExcelJS 默认导入改写导致 `Workbook` 不可用，同时为 HyperFormula 的嵌套依赖启用 CommonJS 互操作：

```ts
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

若项目已有 `optimizeDeps`，请合并数组而非覆盖。若已配置 ExcelJS 浏览器入口别名，也可保留：

```ts
resolve: {
  alias: { exceljs: 'exceljs/dist/exceljs.min.js' },
},
```

完整说明和缓存清理提示见 [core 的 Vite 集成文档](./packages/core/README.md#vite-集成依赖预构建配置)。

## 本地开发

环境要求：Node.js >= 18、pnpm >= 8。

```bash
pnpm install
pnpm build
pnpm --filter @excel-preview/core test
pnpm dev
```

`pnpm dev` 会启动 playground，默认访问地址为 `http://localhost:3000`。演示页支持加载 URL 或本地 Excel 文件。

## 项目结构

```text
.
├── packages/core/       # @excel-preview/core：解析、计算与渲染核心
│   └── src/
│       ├── loader/      # URL、File、Blob、ArrayBuffer 加载
│       ├── parser/      # 工作表、图表与透视表缓存解析
│       ├── chart/       # 图表模型、布局、ECharts/Canvas 适配
│       ├── renderer/    # 表格、图表、图片渲染
│       └── excel-viewer.ts
├── playground/          # Vite 演示应用
├── docs/                # 产品与差异化说明、架构插图
├── PLAN.md               # 开发计划与架构概要
└── PUBLISH.md            # npm 发布流程
```

## 使用边界

本项目是浏览器端**只读预览器**，不提供单元格编辑或保存回写。以下内容不属于当前承诺范围：

- 旧版 `.xls` 二进制文件；
- VBA、宏、外部工作簿引用与数据表公式；
- 部分新版动态数组、Excel 专有函数和 HyperFormula 未实现的函数；
- 数据验证、切片器、图标集条件格式，以及完整的数据透视表筛选/拖拽/刷新交互；
- 像素级还原的 3D 图表与所有复杂图表效果；
- 超大工作簿的无限性能保障（预览会创建相应 DOM 节点）。

建议为关键报表模板建立公式、格式与图表的回归用例，并在上传环节控制文件大小与工作表规模。

## 文档与发布

- [`@excel-preview/core` 完整 API 与功能说明](./packages/core/README.md)
- [npm 发布流程](./PUBLISH.md)

发布前建议执行：

```bash
pnpm --filter @excel-preview/core build
pnpm --filter @excel-preview/core test
pnpm --filter @excel-preview/core pack --dry-run
```
