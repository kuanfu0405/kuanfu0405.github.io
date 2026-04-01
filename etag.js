// etag.js — V2.2-B 通行費「欄位定位」解析版（NAS Stable）
// 核心修正：
// ✅ 不再假設「通行費」字樣會出現在每一筆前
// ✅ 先記錄「通行費」欄位的 X 座標
// ✅ 再抓取同一欄位上的數字作為通行費

(function () {
  function onReady(fn) {
    if (document.readyState !== 'loading') fn();
    else document.addEventListener('DOMContentLoaded', fn);
  }

  onReady(function () {
    console.log('[etag] ready');

    var btn = document.getElementById('etagBtn');
    var input = document.getElementById('etagFile');

    if (!btn || !input) {
      alert('❌ 找不到 eTagBtn 或 etagFile');
      return;
    }

    btn.addEventListener('click', function () {
      input.value = '';
      input.click();
    });

    input.addEventListener('change', function () {
      var file = input.files[0];
      if (!file) return;

      if (!window.pdfjsLibViaMJS) {
        alert('❌ pdf.mjs 尚未初始化');
        return;
      }

      var reader = new FileReader();
      reader.onload = function (e) {
        var buffer = e.target.result;

        window.pdfjsLibViaMJS.getDocument({ data: buffer }).promise.then(function (pdf) {
          var etagMap = {};
          var chain = Promise.resolve();

          for (var p = 1; p <= pdf.numPages; p++) {
            (function (pageNo) {
              chain = chain.then(function () {
                return pdf.getPage(pageNo).then(function (page) {
                  return page.getTextContent().then(function (tc) {
                    var currentDate = '';
                    var tollX = null;

                    tc.items.forEach(function (it) {
                      var text = (it.str || '').trim();
                      var x = it.transform && it.transform[4];

                      // 日期 YYYY/MM/DD
                      if (text.length === 10 && text.indexOf('/') === 4) {
                        currentDate = text;
                        if (!etagMap[currentDate]) etagMap[currentDate] = 0;
                        return;
                      }

                      // 找到「通行費」欄位位置（只需一次）
                      if (text.indexOf('通行費') !== -1) {
                        tollX = x;
                        console.log('[etag] 通行費欄位 X =', tollX);
                        return;
                      }

                      // 若已知通行費欄位，且是數字且在同一欄位
                      if (currentDate && tollX !== null && x !== null) {
                        if (Math.abs(x - tollX) < 5) {
                          var numText = text.replace(/[^0-9]/g, '');
                          if (numText) {
                            var val = parseInt(numText, 10);
                            etagMap[currentDate] += val;
                          }
                        }
                      }
                    });
                  });
                });
              });
            })(p);
          }

          chain.then(function () {
            console.log('[etag] parse done', etagMap);
            if (window.fillEtagByDate) {
              window.fillEtagByDate(etagMap);
              alert('✅ eTag PDF 通行費已寫入表格');
            }
          });
        });
      };

      reader.readAsArrayBuffer(file);
    });
  });
})();
