# Excel Preview & Chart Render — 开发计划

> 纯 JS 核心包，无框架绑定，任何 JS 环境可用
> 包名: **`@excel-preview/core`**

---

## 一、架构

```
excel-vue-react/
├── packages/core/          ← 唯一包：纯 JS，框架无关
│   ├── src/
│   │   ├── types/          类型定义
│   │   ├── loader/         多源数据加载
│   │   ├── parser/         表格+图表解析
│   │   ├── renderer/       表格渲染 + ECharts 浮层
│   │   ├── excel-viewer.ts 高阶 API 封装
│   │   └── index.ts        统一导出
│   └── package.json
├── playground/             Vite 演示应用（纯 TS）
└── pnpm-workspace.yaml
```

## 二、实现状态

| 模块 | 状态 |
|------|------|
| 类型定义 | ✅ 已完成 |
| Loader 数据加载 | ✅ 已完成 |
| Excel Parser 表格解析 | ✅ 已完成 |
| Chart Parser 图表解析 | ✅ 已完成 |
| Table Renderer 表格渲染 | ✅ 已完成 |
| Chart Renderer 图表渲染 | ✅ 已完成 |
| ExcelViewer 高阶封装 | ✅ 已完成 |
| Playground 演示 | ✅ 已完成 |

## 三、使用方式

```ts
import { ExcelViewer } from '@excel-preview/core';

const viewer = new ExcelViewer({
  target: '#container',
  src: 'https://example.com/file.xlsx',
});

// 或分步调用
viewer.render(file);
viewer.setSheet(1);
```

## 四、构建验证

- `pnpm build` — 构建 core 包
- `pnpm dev` — 启动 playground 演示
