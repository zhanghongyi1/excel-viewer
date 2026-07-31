import { ExcelViewer } from '@excel-preview/core';

// 默认使用 Google Charts 套件渲染图表
// loader.js 由框架内部自动加载，无需在 HTML 中手动引入
const viewer = new ExcelViewer({
  target: '#viewer-container',
  width: '100%',
  height: '100%',
  showToolbar: true,
  // chartBackend 默认 'google'，无需手动指定
  // 如需切换后端:
  //   chartBackend: 'canvas'  — 自研 Canvas 渲染引擎
  //   chartBackend: 'echarts' — ECharts（需传入 echarts 实例）
  onRendered: () => showStatus('渲染完成', 'success'),
  onError: (err) => showStatus(err.message, 'error'),
  onSheetChange: (name, idx) => console.log(`切换到: ${name} (${idx})`),
});

const urlInput = document.getElementById('url-input') as HTMLInputElement;
const statusEl = document.getElementById('status') as HTMLSpanElement;

document.getElementById('btn-load-url')!.addEventListener('click', () => {
  const url = urlInput.value.trim();
  if (url) {
    showStatus('加载中...', 'loading');
    viewer.render(url);
  }
});

document.getElementById('file-input')!.addEventListener('change', (e: Event) => {
  const file = (e.target as HTMLInputElement).files?.[0];
  if (file) {
    showStatus('加载中...', 'loading');
    viewer.render(file);
  }
});

function showStatus(msg: string, type: 'loading' | 'error' | 'success') {
  statusEl.textContent = msg;
  statusEl.className = `status ${type}`;
  statusEl.style.display = 'inline-block';
}
