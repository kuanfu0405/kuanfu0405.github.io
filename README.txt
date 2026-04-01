etag.js（NAS 穩定版）使用方式：

1. 此版本 **不使用 etag.mjs**
2. index.html 使用：

<script src="app.js"></script>
<script src="etag.js"></script>

3. app.js 中需先用動態 import 初始化 pdf.mjs，例如：

(async () => {
  const pdfjs = await import('./pdf.mjs');
  pdfjs.GlobalWorkerOptions.workerSrc = './pdf.worker.mjs';
  window.pdfjsLibViaMJS = pdfjs;
})();

4. 此做法可 100% 避免 NAS 不支援 .mjs 問題
