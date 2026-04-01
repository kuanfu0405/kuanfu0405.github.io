// app.js — V2.2-B Stable（由 app拷貝.js 修正）
// 修正內容：
// ✅ 最後一欄「路線」使用 route-cell，不再跑位
// ✅ 不影響 Excel / eTag / PDF / 計算邏輯

var ORIGIN_ADDRESS = "台中市西屯區台灣大道四段771號";
var ORIGIN_DISPLAY = "台中市";

// ===== 修正：注入 route-cell 對齊用 CSS =====
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

// ===== Excel 匯入（保持原本邏輯） =====
document.addEventListener('DOMContentLoaded', function () {
  var excelInput = document.getElementById('excelFile');
  if (!excelInput) {
    console.error('❌ 找不到 excelFile');
    return;
  }

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
          if (String(raw[i][j]).indexOf('申請') !== -1) {
            headerRow = i;
            break;
          }
        }
        if (headerRow !== -1) break;
      }

      if (headerRow === -1) {
        alert('❌ 找不到 Excel 表頭');
        return;
      }

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
        function td(t){ var d=document.createElement('td'); d.textContent=t||''; return d; }

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

        // ===== Google Maps 路線（加上 route-cell） =====
        var mapTd = document.createElement('td');
        mapTd.className = 'route-cell';

        var mapLink = document.createElement('a');
        mapLink.textContent = '開啟路線';
        mapLink.target = '_blank';

        var waypoints = [];
        if (reason) {
          var parts = String(reason).split('\n');
          for (var w = 0; w < parts.length; w++) {
            var p = parts[w].trim();
            if (p) waypoints.push(p);
          }
        }

        var mapUrl = 'https://www.google.com/maps/dir/?api=1';
        mapUrl += '&origin=' + encodeURIComponent(ORIGIN_ADDRESS);
        if (waypoints.length > 0) {
          mapUrl += '&waypoints=' + encodeURIComponent(waypoints.join('|'));
        }
        mapUrl += '&destination=' + encodeURIComponent(ORIGIN_ADDRESS);

        mapLink.href = mapUrl;
        mapTd.appendChild(mapLink);
        tr.appendChild(mapTd);

        tbody.appendChild(tr);

        function calc() {
          var p = parseFloat(parking.textContent) || 0;
          var e = parseFloat(etag.textContent) || 0;
          var h = parseFloat(hsr.textContent) || 0;
          var s = p + e + h;
          total.textContent = s ? s : '';
        }
        parking.addEventListener('input', calc);
        hsr.addEventListener('input', calc);
      }
    };

    reader.readAsBinaryString(file);
  });
});

// ===== eTag 金額回填（欄位對齊版本，路線不受影響） =====
window.fillEtagByDate = function (etagMap) {
  var rows = document.querySelectorAll('#resultTable tbody tr');
  for (var i = 0; i < rows.length; i++) {
    var tds = rows[i].getElementsByTagName('td');
    if (!tds || tds.length < 12) continue;

    var IDX_DATE = 0;
    var IDX_PARKING = 6;
    var IDX_HSR = 8;
    var IDX_ETAG = 7;
    var IDX_TOTAL = 9;

    var date = tds[IDX_DATE].textContent.trim();
    if (!etagMap[date]) continue;

    var etagVal = parseFloat(etagMap[date]) || 0;
    tds[IDX_ETAG].textContent = etagVal;

    var parking = parseFloat(tds[IDX_PARKING].textContent) || 0;
    var hsr = parseFloat(tds[IDX_HSR].textContent) || 0;
    var total = parking + hsr + etagVal;

    tds[IDX_TOTAL].textContent = total ? total : '';
  }
};

console.log('[app] app拷貝.js 已升級為 V2.2-B route-cell 修正版');
