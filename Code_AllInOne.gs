/**
 * =========================================================================================
 * BẢN TẤT CẢ TRONG 1 (ALL-IN-ONE): GOOGLE APPS SCRIPT FORM NHẬP LIỆU KHÁCH HÀNG KÈM NGÀY GIỜ
 * =========================================================================================
 * Ưu điểm: Chỉ cần DÁN 1 FILE DUY NHẤT này vào Code.gs, KHÔNG CẦN tạo thêm file Index.html.
 * Tránh hoàn toàn lỗi "Không tìm thấy tệp HTML có tên Index".
 */

// Cấu hình tên Sheet: Để trống "" sẽ tự động lấy trang tính đang mở
const SHEET_NAME = ""; 

/**
 * Tự động tạo Menu khi mở Google Sheet
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📋 Quản Lý Nhập Liệu")
    .addItem("▶ Mở Form Nhập Liệu (Bên phải - Sidebar)", "showSidebar")
    .addItem("▶ Mở Form Nhập Liệu (Cửa sổ giữa - Dialog)", "showModalDialog")
    .addSeparator()
    .addItem("ℹ Hướng Dẫn Sử Dụng", "showHelp")
    .addToUi();
}

/**
 * Hiển thị Sidebar bên phải
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutput(getFormHtml())
    .setTitle("Nhập Thông Tin Khách Hàng")
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Hiển thị Hộp thoại Popup ở giữa
 */
function showModalDialog() {
  const html = HtmlService.createHtmlOutput(getFormHtml())
    .setWidth(460)
    .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, "Nhập Thông Tin Khách Hàng");
}

/**
 * Web App chạy trên điện thoại/trình duyệt độc lập
 */
function doGet() {
  return HtmlService.createHtmlOutput(getFormHtml())
    .setTitle("Form Nhập Liệu Khách Hàng")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function showHelp() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    "HƯỚNG DẪN SỬ DỤNG",
    "1. Bấm 'Mở Form Nhập Liệu' để mở form.\n" +
    "2. Nhập đầy đủ thông tin Khách hàng, SĐT, Sản phẩm và Nhân viên.\n" +
    "3. Bấm 'Lưu Thông Tin' (hoặc nhấn Ctrl + Enter), hệ thống sẽ tự động tăng STT và lưu ngày giờ chính xác.\n" +
    "4. Số điện thoại được giữ nguyên số 0 ở đầu.",
    ui.ButtonSet.OK
  );
}

/**
 * Xử lý lưu thông tin vào Sheet
 */
function saveCustomerData(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, message: "Hệ thống đang bận, vui lòng thử lại sau giây lát!" };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet;
    
    if (SHEET_NAME && SHEET_NAME.trim() !== "") {
      sheet = ss.getSheetByName(SHEET_NAME.trim());
    } else {
      sheet = ss.getActiveSheet();
    }

    if (!sheet) {
      return { success: false, message: "Không tìm thấy trang tính phù hợp!" };
    }

    // 1. Kiểm tra và thêm tiêu đề cột F: NGÀY GIỜ LƯU nếu chưa có
    const headerF = sheet.getRange(1, 6).getValue();
    if (!headerF || headerF.toString().trim() === "") {
      const headerCell = sheet.getRange(1, 6);
      headerCell.setValue("NGÀY GIỜ LƯU");
      
      const sampleCell = sheet.getRange(1, 5);
      headerCell.setFontWeight("bold")
                .setBackground(sampleCell.getBackground() || "#fce5cd")
                .setFontColor(sampleCell.getFontColor() || "#783f04")
                .setFontFamily(sampleCell.getFontFamily() || "Arial")
                .setHorizontalAlignment("center")
                .setVerticalAlignment("middle");
      sheet.setColumnWidth(6, 175);
    }

    // 2. Tính STT tự động
    const lastRow = sheet.getLastRow();
    const stt = lastRow >= 1 ? lastRow : 1;
    const targetRow = lastRow + 1;

    // 3. Lấy ngày giờ Việt Nam
    const now = new Date();
    const formattedDate = Utilities.formatDate(now, "Asia/Ho_Chi_Minh", "dd/MM/yyyy HH:mm:ss");

    // 4. Chuẩn hóa dữ liệu
    const khachHang = (data.khachHang || "").trim();
    const sdtRaw = (data.sdt || "").trim();
    const sdt = sdtRaw.startsWith("'") ? sdtRaw : `'${sdtRaw}`;
    const sanPham = (data.sanPham || "").trim();
    const nhanVien = (data.nhanVien || "").trim();

    // 5. Ghi dữ liệu dòng mới
    const rowValues = [
      stt,
      khachHang,
      sdt,
      sanPham,
      nhanVien,
      formattedDate
    ];

    sheet.appendRow(rowValues);

    sheet.getRange(targetRow, 1).setHorizontalAlignment("center");
    sheet.getRange(targetRow, 3).setHorizontalAlignment("center");
    sheet.getRange(targetRow, 6).setHorizontalAlignment("center");

    return {
      success: true,
      message: `Đã lưu thành công khách hàng #${stt}!`,
      data: {
        stt: stt,
        khachHang: khachHang,
        sdt: sdtRaw,
        sanPham: sanPham,
        nhanVien: nhanVien,
        thoiGian: formattedDate
      }
    };

  } catch (error) {
    return {
      success: false,
      message: "Lỗi khi lưu vào trang tính: " + error.toString()
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Trả về chuỗi HTML giao diện Form (Được nhúng trực tiếp, không cần tệp Index.html)
 */
function getFormHtml() {
  return `<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Nhập Thông Tin Khách Hàng</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
  <style>
    :root {
      --primary: #c2881e;
      --primary-dark: #9d6d13;
      --primary-border: #f2dcab;
      --bg: #f8fafc;
      --card-bg: #ffffff;
      --text: #1e293b;
      --text-muted: #64748b;
      --border: #e2e8f0;
      --focus-ring: rgba(194, 136, 30, 0.25);
      --success-bg: #ecfdf5;
      --success-border: #a7f3d0;
      --success-text: #065f46;
      --error-bg: #fef2f2;
      --error-border: #fecaca;
      --error-text: #991b1b;
      --radius: 12px;
      --transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
    }
    * { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif; }
    body { background-color: var(--bg); color: var(--text); padding: 14px; display: flex; justify-content: center; align-items: flex-start; min-height: 100vh; }
    .form-container { width: 100%; max-width: 440px; background: var(--card-bg); border-radius: var(--radius); box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.05); border: 1px solid var(--border); overflow: hidden; }
    .form-header { background: linear-gradient(135deg, #fdf6e7 0%, #fae6be 100%); padding: 18px 16px 14px; border-bottom: 2px solid var(--primary-border); text-align: center; }
    .form-header h2 { font-size: 1.15rem; font-weight: 700; color: var(--primary-dark); text-transform: uppercase; display: flex; align-items: center; justify-content: center; gap: 6px; }
    .live-clock-badge { display: inline-flex; align-items: center; gap: 6px; margin-top: 8px; padding: 3px 12px; background: rgba(255, 255, 255, 0.9); border-radius: 20px; font-size: 0.78rem; color: var(--primary-dark); font-weight: 600; border: 1px solid rgba(194, 136, 30, 0.3); }
    .live-dot { width: 8px; height: 8px; background-color: #10b981; border-radius: 50%; display: inline-block; animation: pulse 1.8s infinite; }
    @keyframes pulse { 0% { transform: scale(0.95); } 70% { transform: scale(1); box-shadow: 0 0 0 5px rgba(16, 185, 129, 0); } 100% { transform: scale(0.95); } }
    .form-body { padding: 18px; }
    .form-group { margin-bottom: 14px; }
    .form-label { display: block; font-size: 0.84rem; font-weight: 600; color: #334155; margin-bottom: 5px; }
    .form-label span.req { color: #ef4444; margin-left: 2px; }
    .input-wrapper { position: relative; display: flex; align-items: center; }
    .input-icon { position: absolute; left: 12px; color: #94a3b8; font-size: 0.95rem; pointer-events: none; }
    .form-control { width: 100%; padding: 9px 12px 9px 36px; font-size: 0.9rem; color: var(--text); background-color: #ffffff; border: 1.5px solid var(--border); border-radius: 8px; outline: none; transition: var(--transition); }
    .form-control:focus { border-color: var(--primary); box-shadow: 0 0 0 3px var(--focus-ring); }
    textarea.form-control { resize: vertical; min-height: 65px; padding-top: 8px; }
    .button-group { display: flex; gap: 10px; margin-top: 18px; }
    .btn { display: inline-flex; align-items: center; justify-content: center; gap: 6px; padding: 10px 16px; font-size: 0.92rem; font-weight: 600; border-radius: 8px; border: none; cursor: pointer; transition: var(--transition); }
    .btn-submit { flex: 2; background: linear-gradient(135deg, #d99b26 0%, #b87a14 100%); color: #ffffff; box-shadow: 0 4px 10px rgba(184, 122, 20, 0.25); }
    .btn-submit:hover:not(:disabled) { background: linear-gradient(135deg, #c78b1d 0%, #a3690d 100%); transform: translateY(-1px); }
    .btn-submit:disabled { opacity: 0.65; cursor: not-allowed; }
    .btn-reset { flex: 1; background-color: #f1f5f9; color: #475569; border: 1px solid var(--border); }
    .btn-reset:hover { background-color: #e2e8f0; }
    .alert { padding: 10px 12px; border-radius: 8px; font-size: 0.83rem; display: none; align-items: center; gap: 8px; margin-bottom: 14px; }
    .alert.success { background-color: var(--success-bg); border: 1px solid var(--success-border); color: var(--success-text); display: flex; }
    .alert.error { background-color: var(--error-bg); border: 1px solid var(--error-border); color: var(--error-text); display: flex; }
    .last-saved-card { margin-top: 14px; padding: 10px 12px; background: #fafaf9; border: 1px dashed #d6d3d1; border-radius: 8px; font-size: 0.78rem; display: none; }
    .last-saved-card.active { display: block; }
    .last-saved-title { font-weight: 700; color: #78350f; margin-bottom: 4px; }
    .spinner { width: 14px; height: 14px; border: 2px solid rgba(255, 255, 255, 0.4); border-top-color: #ffffff; border-radius: 50%; animation: spin 0.8s linear infinite; display: none; }
    @keyframes spin { to { transform: rotate(360deg); } }
    .hint { text-align: center; margin-top: 10px; font-size: 0.72rem; color: var(--text-muted); }
  </style>
</head>
<body>
  <div class="form-container">
    <div class="form-header">
      <h2><span>📋</span> Nhập Dữ Liệu Khách Hàng</h2>
      <div class="live-clock-badge">
        <span class="live-dot"></span>
        <span id="liveClock">--:--:--</span>
      </div>
    </div>
    <div class="form-body">
      <div id="alertBox" class="alert">
        <span id="alertIcon"></span>
        <span id="alertMessage"></span>
      </div>
      <form id="customerForm" onsubmit="handleSubmit(event)">
        <div class="form-group">
          <label class="form-label" for="khachHang">Khách Hàng <span class="req">*</span></label>
          <div class="input-wrapper">
            <span class="input-icon">👤</span>
            <input type="text" id="khachHang" class="form-control" placeholder="VD: Nguyễn Văn A" required autocomplete="off" />
          </div>
        </div>
        <div class="form-group">
          <label class="form-label" for="sdt">Số Điện Thoại <span class="req">*</span></label>
          <div class="input-wrapper">
            <span class="input-icon">📞</span>
            <input type="tel" id="sdt" class="form-control" placeholder="VD: 0987654321" required pattern="[0-9]{9,11}" title="Vui lòng nhập 9-11 chữ số" autocomplete="off" />
          </div>
        </div>
        <div class="form-group">
          <label class="form-label" for="sanPham">Sản Phẩm Khách Xem</label>
          <div class="input-wrapper">
            <span class="input-icon" style="top: 10px;">🏷️</span>
            <textarea id="sanPham" class="form-control" rows="2" placeholder="VD: Áo sơ mi nam size L..."></textarea>
          </div>
        </div>
        <div class="form-group">
          <label class="form-label" for="nhanVien">Nhân Viên Phụ Trách</label>
          <div class="input-wrapper">
            <span class="input-icon">👔</span>
            <input type="text" id="nhanVien" class="form-control" placeholder="VD: Lan Anh" autocomplete="off" />
          </div>
        </div>
        <div class="button-group">
          <button type="button" class="btn btn-reset" onclick="resetForm()">🔄 Làm Mới</button>
          <button type="submit" id="btnSubmit" class="btn btn-submit">
            <span class="spinner" id="btnSpinner"></span>
            <span id="btnText">💾 Lưu Thông Tin</span>
          </button>
        </div>
        <div class="hint">Mẹo: Nhấn <b>Enter</b> hoặc <b>Ctrl + Enter</b> để lưu nhanh</div>
      </form>
      <div id="lastSavedCard" class="last-saved-card">
        <div class="last-saved-title">✅ Vừa lưu thành công:</div>
        <div id="lastSavedContent"></div>
      </div>
    </div>
  </div>
  <script>
    function updateClock() {
      const now = new Date();
      const options = { timeZone: 'Asia/Ho_Chi_Minh', day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false };
      document.getElementById('liveClock').textContent = new Intl.DateTimeFormat('vi-VN', options).format(now);
    }
    setInterval(updateClock, 1000); updateClock();

    window.addEventListener('DOMContentLoaded', () => { document.getElementById('khachHang').focus(); });

    document.addEventListener('keydown', (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
        const form = document.getElementById('customerForm');
        if (form.checkValidity()) handleSubmit(e);
        else form.reportValidity();
      }
    });

    function showAlert(type, message) {
      const box = document.getElementById('alertBox');
      box.className = 'alert ' + type;
      document.getElementById('alertIcon').textContent = type === 'success' ? '✅' : '⚠️';
      document.getElementById('alertMessage').textContent = message;
      if (type === 'success') setTimeout(() => { box.style.display = 'none'; }, 4000);
    }

    function resetForm() {
      document.getElementById('customerForm').reset();
      document.getElementById('alertBox').style.display = 'none';
      document.getElementById('khachHang').focus();
    }

    function handleSubmit(e) {
      if (e) e.preventDefault();
      const khachHang = document.getElementById('khachHang').value.trim();
      const sdt = document.getElementById('sdt').value.trim();
      const sanPham = document.getElementById('sanPham').value.trim();
      const nhanVien = document.getElementById('nhanVien').value.trim();

      if (!khachHang || !sdt) {
        showAlert('error', 'Vui lòng điền đủ Tên Khách Hàng và Số Điện Thoại!');
        return;
      }

      const btn = document.getElementById('btnSubmit');
      const spinner = document.getElementById('btnSpinner');
      const btnText = document.getElementById('btnText');
      btn.disabled = true; spinner.style.display = 'inline-block'; btnText.textContent = 'Đang lưu...';

      const payload = { khachHang, sdt, sanPham, nhanVien };

      if (typeof google !== 'undefined' && google.script && google.script.run) {
        google.script.run
          .withSuccessHandler(function(response) {
            btn.disabled = false; spinner.style.display = 'none'; btnText.textContent = '💾 Lưu Thông Tin';
            if (response && response.success) {
              showAlert('success', response.message);
              document.getElementById('lastSavedContent').innerHTML = 
                '<b>#' + response.data.stt + ':</b> ' + escapeHtml(response.data.khachHang) + ' (' + escapeHtml(response.data.sdt) + ')<br>' +
                '<b>SP:</b> ' + escapeHtml(response.data.sanPham || '---') + ' | <b>NV:</b> ' + escapeHtml(response.data.nhanVien || '---') + '<br>' +
                '<i>🕒 ' + response.data.thoiGian + '</i>';
              document.getElementById('lastSavedCard').classList.add('active');
              const nv = document.getElementById('nhanVien').value;
              document.getElementById('customerForm').reset();
              document.getElementById('nhanVien').value = nv;
              document.getElementById('khachHang').focus();
            } else {
              showAlert('error', response ? response.message : 'Lỗi khi lưu!');
            }
          })
          .withFailureHandler(function(err) {
            btn.disabled = false; spinner.style.display = 'none'; btnText.textContent = '💾 Lưu Thông Tin';
            showAlert('error', 'Lỗi kết nối: ' + err.toString());
          })
          .saveCustomerData(payload);
      }
    }

    function escapeHtml(str) {
      if (!str) return '';
      return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
    }
  </script>
</body>
</html>`;
}
