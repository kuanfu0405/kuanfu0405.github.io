const ORIGIN_ADDRESS = "台中市西屯區台灣大道四段771號"; // 路線/計算用
const ORIGIN_DISPLAY = "台中市"; // 畫面顯示用

window.addEventListener("DOMContentLoaded", () => {
  const input = document.getElementById("excelFile");
  if (!input) return;

  input.addEventListener("change", function () {
    const reader = new FileReader();

    reader.onload = (evt) => {
      // 讀取 Excel
      const wb = XLSX.read(evt.target.result, { type: "binary" });
      const sheet = wb.Sheets[wb.SheetNames[0]];

      // 先整張讀為陣列（因為 104HR 表頭不在第一列）
      const raw = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: ""
      });

      // 找出真正的表頭列（只要包含「申請」即可，最保險）
      const headerRowIndex = raw.findIndex(row =>
        row.some(cell => String(cell).includes("申請"))
      );

      if (headerRowIndex === -1) {
        alert("找不到 104HR 表頭列（申請單），請確認匯出檔是否正確");
        return;
      }

      // 從表頭列開始轉成物件
      const rows = XLSX.utils.sheet_to_json(sheet, {
        range: headerRowIndex,
        defval: ""
      });

      const tbody = document.querySelector("#resultTable tbody");
      tbody.innerHTML = "";

      rows.forEach(r => {
        // 欄位名稱容錯處理
        const keys = Object.keys(r).map(k => k.trim());
        const val = kw => {
          const key = keys.find(k => k.includes(kw));
          return key ? r[key] : "";
        };

        // 104 外出單號（不要再用這個當 return 條件）
        const formId = val("編號") || val("申請") || "";

        const timeRange = val("出差時間") || val("出差日期");
        const location = val("地點") || val("縣市") || "";
        const reason = val("原因") || val("事項") || "";

        // ✅ 穩定日期解析（只取前 10 個字）
        let date = "";

// 出差日期：直接取出差時間欄位的前 10 個字元
if (timeRange && typeof timeRange === "string") {
  date = timeRange.substring(0, 10).replace(/-/g, "/");
}
        // 真正完全空的列才跳過
        if (!date && !location && !reason && !formId) return;

        // 多點拜訪（出差原因每行一點）
        const waypoints = reason
          ? String(reason)
              .split(/\r?\n/)
              .map(x => x.trim())
              .filter(Boolean)
          : (location ? [location] : []);

        // Google Maps：去 → 多點 → 回台中
        const mapUrl =
          "https://www.google.com/maps/dir/?api=1" +
          `&origin=${encodeURIComponent(ORIGIN_ADDRESS)}` +
          (waypoints.length ? `&waypoints=${encodeURIComponent(waypoints.join("|"))}` : "") +
          `&destination=${encodeURIComponent(ORIGIN_ADDRESS)}`;

        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td>${date}</td>
          <td>${ORIGIN_DISPLAY}</td>
          <td>${location}</td>
          <td>${String(reason || "").replace(/\n/g, "<br>")}</td>
          <td contenteditable>--</td>
          <td contenteditable></td>
          <td>104 Max 外出單號：${formId}</td>
          <td><a href="${mapUrl}" target="_blank">開啟路線</a></td>
        `;

        tbody.appendChild(tr);
      });

      // 如果真的一列都沒畫出來，直接提醒
      if (!tbody.children.length) {
        alert("Excel 已讀取，但沒有任何可用的外出資料列");
      }
    };

    reader.readAsBinaryString(input.files[0]);
  });
});
``