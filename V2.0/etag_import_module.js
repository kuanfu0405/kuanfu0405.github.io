// eTag PDF 匯入功能（Module 版，強化解析 v1.2）
// ✅ 已修正：PDF 文字被拆碎導致 0 筆金額的問題
// 使用方式：搭配 pdf.mjs / pdf.worker.mjs

import * as pdfjsLib from './pdf.mjs';
pdfjsLib.GlobalWorkerOptions.workerSrc = './pdf.worker.mjs';

window.addEventListener('DOMContentLoaded', () => {
  const btn = document.getElementById('importEtagBtn');
  const input = document.getElementById('etagFile');

  if (!btn || !input) {
    alert('❌ 找不到 eTag 按鈕或檔案欄位');
    return;
  }

  btn.addEventListener('click', () => input.click());

  input.addEventListener('change', async () => {
    const file = input.files[0];
    if (!file) return;

    try {
      const buf = await file.arrayBuffer();
      const pdf = await pdfjsLib.getDocument({ data: buf }).promise;

      const etagByDate = {};
      const dateRegex = /\d{4}\/\d{2}\/\d{2}/g;
      const feeRegex = /(\d+)\s*元/;

      for (let p = 1; p <= pdf.numPages; p++) {
        const page = await pdf.getPage(p);
        const tc = await page.getTextContent();
        const text = tc.items.map(i => i.str).join(' ');

        const dates = text.match(dateRegex) || [];

        for (const d of dates) {
          const idx = text.indexOf(d);
          if (idx === -1) continue;

          const slice = text.slice(idx, idx + 120);
          const feeMatch = slice.match(feeRegex);

          if (feeMatch) {
            const fee = parseInt(feeMatch[1], 10);
            etagByDate[d] = (etagByDate[d] || 0) + fee;
          }
        }
      }

      let count = 0;
      const rows = document.querySelectorAll('#resultTable tbody tr');
      rows.forEach(r => {
        const d = r.cells[0]?.innerText.trim();
        const c = r.cells[5];
        if (etagByDate[d] && c.innerText.trim() === '') {
          c.innerText = etagByDate[d];
          count++;
        }
      });

      alert(`✅ eTag 對帳完成，共填入 ${count} 筆停車費`);

    } catch (err) {
      console.error(err);
      alert('❌ 解析 eTag PDF 發生錯誤');
    }
  });
});
