// 外出經費申請輔助系統 V2.0
// 基於 V1.1 Excel 穩定解析邏輯延伸
// 新增：交通工具 + eTag + 高鐵 + 總和

const ORIGIN_ADDRESS = "台中市西屯區台灣大道四段771號";
const ORIGIN_DISPLAY = "台中市";

window.addEventListener("DOMContentLoaded", () => {
  const input = document.getElementById("excelFile");
  if (!input) return;

  input.addEventListener("change", function () {
    const reader = new FileReader();

    reader.onload = (evt) => {
      // 讀取 Excel
      const wb = XLSX.read(evt.target.result, { type: "binary" });
      const sheet = wb.Sheets[wb.SheetNames[0]];

      // 整張表先轉成陣列（因為 104HR 表頭不在第一列）
      const raw = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: ""
      });

      // 找出表頭列（包含「申請」）
      const headerRowIndex = raw.findIndex(row =>
        row.some(cell => String(cell).includes("申請"))
      );

      if (headerRowIndex === -1) {
        alert("找不到 104HR 表頭列（申請單），請確認檔案是否正確");
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
        // 欄位容錯處理
        const keys = Object.keys(r).map(k => k.trim());
        const val = kw => {
          const key = keys.find(k => k.includes(kw));
          return key ? r[key] : "";
        };

        const formId = val("編號") || val("申請") || "";
        const timeRange = val("出差時間") || val("出差日期") || "";
        const location = val("地點") || val("縣市") || "";
        const reason = val("原因") || val("事項") || "";

        // 解析日期（只取前 10 碼）
        let date = "";
        if (timeRange && typeof timeRange === "string") {
          date = timeRange.substring(0, 10).replace(/-/g, "/");
        }

        // 真正完全空列才跳過
        if (!date && !location && !reason && !formId) return;

        // 多點拜訪（處理事項）
        const waypoints = reason
          ? String(reason)
              .split(/\r?\n/)
              .map(x => x.trim())
              .filter(Boolean)
          : (location ? [location] : []);

        // Google Maps 路線
        const mapUrl =
          "https://www.google.com/maps/dir/?api=1" +
          `&origin=${encodeURIComponent(ORIGIN_ADDRESS)}` +
          (waypoints.length
            ? `&waypoints=${encodeURIComponent(waypoints.join("|"))}`
            : "") +
          `&destination=${encodeURIComponent(ORIGIN_ADDRESS)}`;

        // 建立資料列
        const tr = document.createElement("tr");

        // ===== 建立欄位 =====
        const tdDate = `<td>${date}</td>`;
        const tdFrom = `<td>${ORIGIN_DISPLAY}</td>`;
        const tdTo = `<td>${location}</td>`;
        const tdReason = `<td>${String(reason || "").replace(/\n/g, "<br>")}</td>`;
        const tdMileage = `<td>--</td>`;

        // 交通工具（select）
        const tdTransport = `
          <td>
            <select class="transport">
              <option value="">--</option>
              <option value="car">自行開車</option>
              <option value="hsr">高鐵</option>
              <option value="train">台鐵</option>
              <option value="bus">公車/捷運</option>
              <option value="other">其他</option>
            </select>
          </td>
        `;

        // 停車費 / eTag / 高鐵 / 總和
        const tdParking = `<td contenteditable class="parking"></td>`;
        const tdEtag = `<td class="etag"></td>`;
        const tdHSR = `<td contenteditable class="hsr"></td>`;
        const tdTotal = `<td class="total"></td>`;

        const tdFormId = `<td>104 Max 外出單號：${formId}</td>`;
        const tdMap = `<td><a href="${mapUrl}" target="_blank">開啟路線</a></td>`;

        tr.innerHTML =
          tdDate +
          tdFrom +
          tdTo +
          tdReason +
          tdMileage +
          tdTransport +
          tdParking +
          tdEtag +
          tdHSR +
          tdTotal +
          tdFormId +
          tdMap;

        tbody.appendChild(tr);

        // ===== 總和計算邏輯 =====
        const parkingCell = tr.querySelector(".parking");
        const etagCell = tr.querySelector(".etag");
        const hsrCell = tr.querySelector(".hsr");
        const totalCell = tr.querySelector(".total");

        const updateTotal = () => {
          const parking = parseFloat(parkingCell.innerText) || 0;
          const etag = parseFloat(etagCell.innerText) || 0;
          const hsr = parseFloat(hsrCell.innerText) || 0;
          const sum = parking + etag + hsr;
          totalCell.innerText = sum > 0 ? sum : "";
        };

        parkingCell.addEventListener("input", updateTotal);
        hsrCell.addEventListener("input", updateTotal);

        // eTag 後續由 module 填值時會再觸發 updateTotal()
      });

      if (!tbody.children.length) {
        alert("Excel 已讀取，但沒有任何可用的外出資料列");
      }
    };

    reader.readAsBinaryString(input.files[0]);
  });
});