// app_patch_pdf_init.js
// ✅ 必須加在你原本 app.js 最後
// 作用：初始化 pdf.mjs，供 etag.js 使用

(function () {
  async function initPDF() {
    try {
      const pdfjs = await import('./pdf.mjs');
      pdfjs.GlobalWorkerOptions.workerSrc = './pdf.worker.mjs';
      window.pdfjsLibViaMJS = pdfjs;
      console.log('[PDF] pdf.mjs 初始化完成');
    } catch (e) {
      console.error('[PDF] 初始化失敗', e);
    }
  }

  initPDF();
})();
