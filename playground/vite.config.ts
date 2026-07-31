import { defineConfig } from 'vite';

export default defineConfig({
  server: {
    port: 3000,
    open: true,
  },
  resolve: {
    alias: {
      // exceljs 默认入口 (excel.js) 使用 Node.js 内置模块 (crypto, stream 等)
      // alias 到其预构建的浏览器包，内部已 polyfill 所有 Node.js 依赖
      exceljs: 'exceljs/dist/exceljs.min.js',
    },
  },
  optimizeDeps: {
    include: ['exceljs'],
  },
});
