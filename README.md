# Excel Preview & Chart Render

浏览器端 `.xlsx` 只读预览。核心包 [`@excel-preview/core`](./packages/core/README.md) 为纯 JavaScript，可用于原生 JS、Vue、React 等任意前端。

![工作簿处理链路](./docs/assets/excel-to-echarts.svg)

## 功能

- 表格：样式、数字格式、合并单元格、隐藏行列、冻结窗格、富文本、超链接、批注、条件格式与 Sheet 切换。
- 计算：同表/跨表公式、常用聚合/条件/日期函数；无法计算时使用 Excel 缓存结果。
- 图表：OpenXML 图表解析，ECharts 优先、Canvas 后备，按单元格锚点定位。
- 图片与透视表：嵌入图片预览；可选解析数据透视表缓存。
- 输入：URL、`File`、`Blob`、`ArrayBuffer`。

## 快速开始

```bash
pnpm add @excel-preview/core
```

```ts
import { ExcelViewer } from '@excel-preview/core';

const viewer = new ExcelViewer({
  target: '#excel-viewer',
  src: '/reports/monthly.xlsx',
  onError: console.error,
});

await viewer.render(file); // URL / File / Blob / ArrayBuffer
viewer.destroy();
```

完整的 Vite 配置、API、图表类型、公式范围与限制见 [`packages/core/README.md`](./packages/core/README.md)。

## 开发

```bash
pnpm install
pnpm build
pnpm --filter @excel-preview/core test
pnpm dev
```

`pnpm dev` 启动 playground，默认地址为 `http://localhost:3000`。

## 目录

```text
packages/core/  # 核心包：解析、计算、表格/图表/图片渲染
playground/     # Vite 演示页
docs/assets/    # README 插图
PUBLISH.md      # npm 发布流程
```

## 限制

仅支持 `.xlsx` 且为只读预览；不支持 `.xls`、编辑/保存、VBA/宏、外部工作簿引用、完整透视表交互和像素级 3D 图表还原。

## 相关文档

- [`@excel-preview/core` 开发者手册](./packages/core/README.md)
- [发布流程](./PUBLISH.md)
