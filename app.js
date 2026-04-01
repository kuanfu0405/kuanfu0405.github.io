// app.js — V2.2.2 Stable
// 修正：匯入 eTag PDF 後，所有列（含無 eTag）皆可正確計算總和

var ORIGIN_ADDRESS = "台中市西屯區台灣大道四段771號";
var ORIGIN_DISPLAY = "台中市";

// ===== route-cell 對齊樣式 =====
(function injectRouteStyle() {
  var style = document.createElement('style');
  style.textContent = `
    #resultTable td.route-cell {
      display: flex;
      align-items: center;
      justify-content: center;
      vertical-align: middle;
    }
  `;
  document.head.appendChild(style);
})();

// ===== Excel 匯入 =====
document.addEventListener('DOMContentLoaded', function () {
  var excelInput = document.getElementById('excelFile');
  if (!excelInput) return;

  excelInput.addEventListener('change', function (e) {
    var file = e.target.files[0];
    if (!file) return;

    var reader = new FileReader();
    reader.onload = function (evt) {
      var wb = XLSX.read(evt.target.result, { type: 'binary' });
      var sheet = wb.Sheets[wb.SheetNames[0]];
      var raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

      var headerRow = -1;
      for (var i = 0; i < raw.length; i++) {
        for (var j = 0; j < raw[i].length; j++) {
          if (String(raw[i][j]).indexOf('申請') !== -1) { headerRow = i; break; }
        }
        if (headerRow !== -1) break;
      }
      if (headerRow === -1) { alert('❌ 找不到 Excel 表頭'); return; }

      var rows = XLSX.utils.sheet_to_json(sheet, { range: headerRow, defval: '' });
      var tbody = document.querySelector('#resultTable tbody');
      if (!tbody) return;
      tbody.innerHTML = '';

      function pick(row, keyword) {
        var keys = Object.keys(row);
        for (var k = 0; k < keys.length; k++) {
          if (String(keys[k]).indexOf(keyword) !== -1) return row[keys[k]];
        }
        return '';
      }

      for (var r = 0; r < rows.length; r++) {
        var row = rows[r];
        var formId = pick(row, '編號') || pick(row, '申請');
        var timeRange = pick(row, '出差時間') || pick(row, '出差日期');
        var location = pick(row, '地點') || pick(row, '縣市');
        var reason = pick(row, '原因') || pick(row, '事項');
        var transport = pick(row, '交通工具');

        var date = '';
        if (typeof timeRange === 'string' && timeRange.length >= 10) {
          date = timeRange.substring(0, 10).split('-').join('/');
        }
        if (!date && !location && !reason && !formId) continue;

        var tr = document.createElement('tr');
        function td(v){ var d=document.createElement('td'); d.textContent=v||''; return d; }

        tr.appendChild(td(date));
        tr.appendChild(td(ORIGIN_DISPLAY));
        tr.appendChild(td(location));
        tr.appendChild(td(reason));
        tr.appendChild(td('--'));
        tr.appendChild(td(transport));

        var parking = document.createElement('td'); parking.contentEditable = 'true'; tr.appendChild(parking);
        var etag = document.createElement('td'); tr.appendChild(etag);
        var hsr = document.createElement('td'); hsr.contentEditable = 'true'; tr.appendChild(hsr);
        var total = document.createElement('td'); tr.appendChild(total);
        tr.appendChild(td(formId));

        var mapTd = document.createElement('td'); mapTd.className='route-cell';
        var a = document.createElement('a'); a.target='_blank'; a.textContent='開啟路線';
        var mapUrl = 'https://www.google.com/maps/dir/?api=1';
        mapUrl += '&origin=' + encodeURIComponent(ORIGIN_ADDRESS);
        if (reason) mapUrl += '&waypoints=' + encodeURIComponent(reason);
        mapUrl += '&destination=' + encodeURIComponent(ORIGIN_ADDRESS);
        a.href = mapUrl; mapTd.appendChild(a); tr.appendChild(mapTd);

        tbody.appendChild(tr);

        function calc(){
          var p = parseFloat(parking.textContent) || 0;
          var e = parseFloat(etag.textContent) || 0;
          var h = parseFloat(hsr.textContent) || 0;
          total.textContent = (p + e + h) || '';
        }

        calc();
        parking.addEventListener('input', calc);
        hsr.addEventListener('input', calc);
      }
    };
    reader.readAsBinaryString(file);
  });
});

// ===== eTag 回填（重點修正） =====
window.fillEtagByDate = function (etagMap) {
  var rows = document.querySelectorAll('#resultTable tbody tr');
  for (var i = 0; i < rows.length; i++) {
    var tds = rows[i].getElementsByTagName('td');
    if (!tds || tds.length < 12) continue;

    var IDX_DATE=0, IDX_PARKING=6, IDX_ETAG=7, IDX_HSR=8, IDX_TOTAL=9;
    var date = tds[IDX_DATE].textContent.trim();

    var p = parseFloat(tds[IDX_PARKING].textContent) || 0;
    var h = parseFloat(tds[IDX_HSR].textContent) || 0;
    var e = etagMap[date] ? parseFloat(etagMap[date]) || 0 : 0;

    if (etagMap[date]) {
      tds[IDX_ETAG].textContent = e;
    }

    tds[IDX_TOTAL].textContent = (p + e + h) || '';
  }
};

console.log('[app] V2.2.2 修正完成：無 eTag 列亦會計算總和');
