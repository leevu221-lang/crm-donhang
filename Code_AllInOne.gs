/**
 * =========================================================================================
 * BẢN TẤT CẢ TRONG 1 (ALL-IN-ONE): GOOGLE APPS SCRIPT FORM NHẬP LIỆU KHÁCH HÀNG KÈM NGÀY GIỜ
 * HỖ TRỢ: CHẠY TRỰC TIẾP TRONG GOOGLE SHEETS + KẾT NỐI VỚI GITHUB PAGES
 * =========================================================================================
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
 * Hiển thị Sidebar bên phải trang tính
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutput(getFormHtml())
    .setTitle("Nhập Thông Tin Khách Hàng")
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Hiển thị Hộp thoại Popup ở giữa trang tính
 */
function showModalDialog() {
  const html = HtmlService.createHtmlOutput(getFormHtml())
    .setWidth(460)
    .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, "Nhập Thông Tin Khách Hàng");
}

/**
 * Xử lý yêu cầu GET từ Web ngoài (GitHub Pages) hoặc khi mở trực tiếp link Web App
 */
function doGet(e) {
  // 1. Nhận dữ liệu gửi qua URL Query Parameters từ GitHub Pages
  if (e && e.parameter && (e.parameter.khachHang || e.parameter.sdt)) {
    const result = saveCustomerData(e.parameter);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 2. Ping kiểm tra kết nối
  if (e && e.parameter && e.parameter.action === "ping") {
    return ContentService.createTextOutput(JSON.stringify({ success: true, message: "Kết nối thành công!" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 3. Mở giao diện Web App trực tiếp
  return HtmlService.createHtmlOutput(getFormHtml())
    .setTitle("Form Nhập Liệu Khách Hàng")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Xử lý yêu cầu POST từ GitHub Pages (nhận JSON payload)
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: "Hệ thống đang bận, vui lòng thử lại sau vài giây!"
    })).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    let data = {};
    if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (parseErr) {
        data = e.parameter || {};
      }
    } else if (e.parameter) {
      data = e.parameter;
    }

    const result = saveCustomerData(data);

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: "Lỗi server Apps Script: " + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function showHelp() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    "HƯỚNG DẪN SỬ DỤNG",
    "1. Bấm 'Mở Form Nhập Liệu' để mở form.\n" +
    "2. Nhập đầy đủ thông tin Khách hàng, SĐT, Sản phẩm và Nhân viên.\n" +
    "3. Bấm 'Lưu Thông Tin' (hoặc nhấn Ctrl + Enter), hệ thống tự động tăng STT và lưu ngày giờ chính xác.\n" +
    "4. Số điện thoại được giữ nguyên số 0 ở đầu.",
    ui.ButtonSet.OK
  );
}

/**
 * Hàm xử lý cốt lõi: Ghi dòng mới vào Sheet kèm ngày giờ
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
    let sheet = (SHEET_NAME && SHEET_NAME.trim() !== "") ? ss.getSheetByName(SHEET_NAME.trim()) : ss.getActiveSheet();

    if (!sheet) return { success: false, message: "Không tìm thấy trang tính phù hợp!" };

    // 1. Tự động kiểm tra và thêm tiêu đề cột F: NGÀY GIỜ LƯU nếu chưa có
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

    // 3. Lấy thời gian hiện tại chuẩn giờ Việt Nam (GMT+7)
    const now = new Date();
    const formattedDate = Utilities.formatDate(now, "Asia/Ho_Chi_Minh", "dd/MM/yyyy HH:mm:ss");

    // 4. Chuẩn hóa dữ liệu
    const khachHang = (data.khachHang || "").trim();
    const sdtRaw = (data.sdt || "").trim();
    const sdt = sdtRaw.startsWith("'") ? sdtRaw : `'${sdtRaw}`;
    const sanPham = (data.sanPham || "").trim();
    const nhanVien = (data.nhanVien || "").trim();

    // 5. Ghi dòng mới vào Sheet: [STT, KHÁCH HÀNG, SĐT, SẢN PHẨM KHÁCH XEM, NHÂN VIÊN, NGÀY GIỜ LƯU]
    const rowValues = [stt, khachHang, sdt, sanPham, nhanVien, formattedDate];
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
    return { success: false, message: "Lỗi khi lưu vào trang tính: " + error.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Giao diện Form nhúng (đồng bộ hoàn toàn với index.html)
 */
function getFormHtml() {
  return '<!DOCTYPE html>\n<html lang="vi">\n<head>\n  <meta charset="UTF-8">\n  <meta name="viewport" content="width=device-width, initial-scale=1.0">\n  <title>Nhập Thông Tin Khách Hàng - CRM Đơn Hàng</title>\n  <link rel="preconnect" href="https://fonts.googleapis.com">\n  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>\n  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">\n  <style>\n    :root {\n      --primary: #c2881e;\n      --primary-dark: #9d6d13;\n      --primary-light: #fff9ed;\n      --primary-border: #f2dcab;\n      --bg: #f1f5f9;\n      --card-bg: #ffffff;\n      --text: #1e293b;\n      --text-muted: #64748b;\n      --border: #e2e8f0;\n      --focus-ring: rgba(194, 136, 30, 0.25);\n      --success-bg: #ecfdf5;\n      --success-border: #a7f3d0;\n      --success-text: #065f46;\n      --error-bg: #fef2f2;\n      --error-border: #fecaca;\n      --error-text: #991b1b;\n      --radius: 16px;\n      --transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);\n    }\n\n    * {\n      box-sizing: border-box;\n      margin: 0;\n      padding: 0;\n      font-family: \'Plus Jakarta Sans\', -apple-system, BlinkMacSystemFont, \'Segoe UI\', Roboto, sans-serif;\n    }\n\n    body {\n      background: linear-gradient(135deg, #f8fafc 0%, #eef2f6 100%);\n      color: var(--text);\n      padding: 24px 16px;\n      display: flex;\n      flex-direction: column;\n      justify-content: center;\n      align-items: center;\n      min-height: 100vh;\n    }\n\n    .form-container {\n      width: 100%;\n      max-width: 460px;\n      background: var(--card-bg);\n      border-radius: var(--radius);\n      box-shadow: 0 20px 35px -10px rgba(0, 0, 0, 0.08), 0 10px 15px -8px rgba(0, 0, 0, 0.04);\n      border: 1px solid rgba(226, 232, 240, 0.8);\n      overflow: hidden;\n      position: relative;\n    }\n\n    /* Top bar buttons (Config / Settings) */\n    .top-actions {\n      position: absolute;\n      top: 14px;\n      right: 14px;\n      z-index: 10;\n      display: flex;\n      gap: 6px;\n    }\n\n    .btn-icon {\n      background: rgba(255, 255, 255, 0.85);\n      border: 1px solid var(--primary-border);\n      border-radius: 50%;\n      width: 32px;\n      height: 32px;\n      display: flex;\n      align-items: center;\n      justify-content: center;\n      cursor: pointer;\n      font-size: 0.85rem;\n      transition: var(--transition);\n      color: var(--primary-dark);\n      box-shadow: 0 2px 4px rgba(0,0,0,0.04);\n    }\n\n    .btn-icon:hover {\n      background: #ffffff;\n      transform: scale(1.08);\n    }\n\n    /* Header */\n    .form-header {\n      background: linear-gradient(135deg, #fdf7ec 0%, #fae8c4 100%);\n      padding: 24px 20px 18px;\n      border-bottom: 2px solid var(--primary-border);\n      text-align: center;\n    }\n\n    .form-header h2 {\n      font-size: 1.25rem;\n      font-weight: 700;\n      color: #783f04;\n      text-transform: uppercase;\n      letter-spacing: 0.5px;\n      display: flex;\n      align-items: center;\n      justify-content: center;\n      gap: 8px;\n    }\n\n    .live-clock-badge {\n      display: inline-flex;\n      align-items: center;\n      gap: 6px;\n      margin-top: 10px;\n      padding: 4px 14px;\n      background: rgba(255, 255, 255, 0.95);\n      border-radius: 20px;\n      font-size: 0.8rem;\n      color: #92400e;\n      font-weight: 600;\n      border: 1px solid rgba(194, 136, 30, 0.35);\n      box-shadow: 0 2px 5px rgba(0, 0, 0, 0.03);\n    }\n\n    .live-dot {\n      width: 8px;\n      height: 8px;\n      background-color: #10b981;\n      border-radius: 50%;\n      display: inline-block;\n      animation: pulse 1.8s infinite;\n    }\n\n    @keyframes pulse {\n      0% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(16, 185, 129, 0.7); }\n      70% { transform: scale(1); box-shadow: 0 0 0 6px rgba(16, 185, 129, 0); }\n      100% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(16, 185, 129, 0); }\n    }\n\n    /* Form Body */\n    .form-body {\n      padding: 22px;\n    }\n\n    .form-group {\n      margin-bottom: 16px;\n    }\n\n    .form-label {\n      display: block;\n      font-size: 0.86rem;\n      font-weight: 600;\n      color: #334155;\n      margin-bottom: 6px;\n    }\n\n    .form-label span.req {\n      color: #ef4444;\n      margin-left: 2px;\n    }\n\n    .input-wrapper {\n      position: relative;\n      display: flex;\n      align-items: center;\n    }\n\n    .input-icon {\n      position: absolute;\n      left: 14px;\n      color: #94a3b8;\n      font-size: 1rem;\n      pointer-events: none;\n    }\n\n    .form-control {\n      width: 100%;\n      padding: 11px 14px 11px 42px;\n      font-size: 0.94rem;\n      color: var(--text);\n      background-color: #ffffff;\n      border: 1.5px solid var(--border);\n      border-radius: 10px;\n      outline: none;\n      transition: var(--transition);\n    }\n\n    .form-control:focus {\n      border-color: var(--primary);\n      box-shadow: 0 0 0 3.5px var(--focus-ring);\n    }\n\n    textarea.form-control {\n      resize: vertical;\n      min-height: 70px;\n      padding-top: 10px;\n    }\n\n    /* Buttons */\n    .button-group {\n      display: flex;\n      gap: 12px;\n      margin-top: 20px;\n    }\n\n    .btn {\n      display: inline-flex;\n      align-items: center;\n      justify-content: center;\n      gap: 8px;\n      padding: 12px 18px;\n      font-size: 0.95rem;\n      font-weight: 600;\n      border-radius: 10px;\n      border: none;\n      cursor: pointer;\n      transition: var(--transition);\n    }\n\n    .btn-submit {\n      flex: 2;\n      background: linear-gradient(135deg, #d99b26 0%, #b87a14 100%);\n      color: #ffffff;\n      box-shadow: 0 4px 14px rgba(184, 122, 20, 0.3);\n    }\n\n    .btn-submit:hover:not(:disabled) {\n      background: linear-gradient(135deg, #c78b1d 0%, #a3690d 100%);\n      box-shadow: 0 6px 18px rgba(184, 122, 20, 0.4);\n      transform: translateY(-1px);\n    }\n\n    .btn-submit:active:not(:disabled) {\n      transform: translateY(0);\n    }\n\n    .btn-submit:disabled {\n      opacity: 0.65;\n      cursor: not-allowed;\n    }\n\n    .btn-reset {\n      flex: 1;\n      background-color: #f1f5f9;\n      color: #475569;\n      border: 1px solid var(--border);\n    }\n\n    .btn-reset:hover {\n      background-color: #e2e8f0;\n      color: #1e293b;\n    }\n\n    /* Alert Banner */\n    .alert {\n      padding: 12px 14px;\n      border-radius: 10px;\n      font-size: 0.85rem;\n      display: none;\n      align-items: center;\n      gap: 10px;\n      margin-bottom: 16px;\n      animation: fadeIn 0.3s ease;\n    }\n\n    @keyframes fadeIn {\n      from { opacity: 0; transform: translateY(-5px); }\n      to { opacity: 1; transform: translateY(0); }\n    }\n\n    .alert.success {\n      background-color: var(--success-bg);\n      border: 1px solid var(--success-border);\n      color: var(--success-text);\n      display: flex;\n    }\n\n    .alert.error {\n      background-color: var(--error-bg);\n      border: 1px solid var(--error-border);\n      color: var(--error-text);\n      display: flex;\n    }\n\n    /* Last saved card */\n    .last-saved-card {\n      margin-top: 18px;\n      padding: 12px 14px;\n      background: #fafaf9;\n      border: 1px dashed #d6d3d1;\n      border-radius: 10px;\n      font-size: 0.82rem;\n      display: none;\n    }\n\n    .last-saved-card.active {\n      display: block;\n      animation: fadeIn 0.3s ease;\n    }\n\n    .last-saved-title {\n      font-weight: 700;\n      color: #78350f;\n      margin-bottom: 4px;\n      display: flex;\n      align-items: center;\n      gap: 6px;\n    }\n\n    .spinner {\n      width: 16px;\n      height: 16px;\n      border: 2px solid rgba(255, 255, 255, 0.4);\n      border-top-color: #ffffff;\n      border-radius: 50%;\n      animation: spin 0.8s linear infinite;\n      display: none;\n    }\n\n    @keyframes spin {\n      to { transform: rotate(360deg); }\n    }\n\n    .hint {\n      text-align: center;\n      margin-top: 14px;\n      font-size: 0.75rem;\n      color: var(--text-muted);\n    }\n\n    /* Settings Modal */\n    .modal-backdrop {\n      position: fixed;\n      inset: 0;\n      background: rgba(15, 23, 42, 0.45);\n      backdrop-filter: blur(4px);\n      display: none;\n      align-items: center;\n      justify-content: center;\n      padding: 16px;\n      z-index: 100;\n      animation: fadeIn 0.2s ease;\n    }\n\n    .modal-card {\n      background: #ffffff;\n      border-radius: 16px;\n      max-width: 440px;\n      width: 100%;\n      box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.25);\n      padding: 22px;\n      position: relative;\n    }\n\n    .modal-title {\n      font-size: 1.05rem;\n      font-weight: 700;\n      color: var(--text);\n      margin-bottom: 8px;\n      display: flex;\n      align-items: center;\n      gap: 8px;\n    }\n\n    .modal-desc {\n      font-size: 0.82rem;\n      color: var(--text-muted);\n      line-height: 1.5;\n      margin-bottom: 16px;\n    }\n\n    .guide-steps {\n      background: #f8fafc;\n      border: 1px solid var(--border);\n      border-radius: 8px;\n      padding: 10px 12px;\n      font-size: 0.78rem;\n      color: #475569;\n      margin-bottom: 16px;\n      line-height: 1.6;\n    }\n\n    .guide-steps ol {\n      padding-left: 18px;\n    }\n\n    .status-badge {\n      display: inline-flex;\n      align-items: center;\n      gap: 5px;\n      font-size: 0.75rem;\n      padding: 3px 8px;\n      border-radius: 12px;\n      background: #f1f5f9;\n      color: #64748b;\n      margin-top: 6px;\n    }\n\n    .status-badge.connected {\n      background: #dcfce7;\n      color: #15803d;\n    }\n\n    .footer-note {\n      text-align: center;\n      margin-top: 14px;\n      font-size: 0.74rem;\n      color: #94a3b8;\n    }\n  </style>\n</head>\n<body>\n\n  <div class="form-container">\n    <!-- Top Action: Settings Button -->\n    <div class="top-actions">\n      <button class="btn-icon" title="Cài đặt kết nối Google Sheet" onclick="openConfigModal()">\n        ⚙️\n      </button>\n    </div>\n\n    <!-- Header -->\n    <div class="form-header">\n      <h2><span>📋</span> Nhập Dữ Liệu Khách Hàng</h2>\n      <div class="live-clock-badge">\n        <span class="live-dot"></span>\n        <span id="liveClock">Đang tải giờ...</span>\n      </div>\n      <div>\n        <span id="connectionBadge" class="status-badge">\n          <span>⚪</span> Chưa cấu hình Sheet\n        </span>\n      </div>\n    </div>\n\n    <!-- Body -->\n    <div class="form-body">\n      <div id="alertBox" class="alert">\n        <span id="alertIcon"></span>\n        <span id="alertMessage"></span>\n      </div>\n\n      <form id="customerForm" onsubmit="handleSubmit(event)">\n        <!-- Khách hàng -->\n        <div class="form-group">\n          <label class="form-label" for="khachHang">\n            Khách Hàng <span class="req">*</span>\n          </label>\n          <div class="input-wrapper">\n            <span class="input-icon">👤</span>\n            <input \n              type="text" \n              id="khachHang" \n              class="form-control" \n              placeholder="VD: Nguyễn Văn A" \n              required \n              autocomplete="off"\n            />\n          </div>\n        </div>\n\n        <!-- SĐT -->\n        <div class="form-group">\n          <label class="form-label" for="sdt">\n            Số Điện Thoại <span class="req">*</span>\n          </label>\n          <div class="input-wrapper">\n            <span class="input-icon">📞</span>\n            <input \n              type="tel" \n              id="sdt" \n              class="form-control" \n              placeholder="VD: 0987654321" \n              required \n              pattern="[0-9]{9,11}" \n              title="Vui lòng nhập 9 đến 11 chữ số" \n              autocomplete="off"\n            />\n          </div>\n        </div>\n\n        <!-- Sản phẩm khách xem -->\n        <div class="form-group">\n          <label class="form-label" for="sanPham">\n            Sản Phẩm Khách Xem\n          </label>\n          <div class="input-wrapper">\n            <span class="input-icon" style="top: 12px;">🏷️</span>\n            <textarea \n              id="sanPham" \n              class="form-control" \n              rows="2" \n              placeholder="VD: Áo sơ mi nam size L..."\n            ></textarea>\n          </div>\n        </div>\n\n        <!-- Nhân viên -->\n        <div class="form-group">\n          <label class="form-label" for="nhanVien">\n            Nhân Viên Phụ Trách\n          </label>\n          <div class="input-wrapper">\n            <span class="input-icon">👔</span>\n            <input \n              type="text" \n              id="nhanVien" \n              class="form-control" \n              placeholder="VD: Lan Anh" \n              autocomplete="off"\n            />\n          </div>\n        </div>\n\n        <!-- Action buttons -->\n        <div class="button-group">\n          <button type="button" class="btn btn-reset" onclick="resetForm()">\n            🔄 Làm Mới\n          </button>\n          <button type="submit" id="btnSubmit" class="btn btn-submit">\n            <span class="spinner" id="btnSpinner"></span>\n            <span id="btnText">💾 Lưu Thông Tin</span>\n          </button>\n        </div>\n\n        <div class="hint">\n          Mẹo: Nhấn <b>Enter</b> hoặc <b>Ctrl + Enter</b> để lưu nhanh\n        </div>\n      </form>\n\n      <!-- Last Saved Summary -->\n      <div id="lastSavedCard" class="last-saved-card">\n        <div class="last-saved-title">\n          <span>✅</span> Vừa lưu thành công vào Sheet:\n        </div>\n        <div id="lastSavedContent"></div>\n      </div>\n    </div>\n  </div>\n\n  <div class="footer-note">\n    CRM Đơn Hàng • Tự động ghi nhận ngày giờ & STT\n  </div>\n\n  <!-- Modal Cấu hình Web App Google Sheet -->\n  <div id="configModal" class="modal-backdrop">\n    <div class="modal-card">\n      <div class="modal-title">\n        <span>⚙️</span> Cài Đặt Kết Nối Google Sheet\n      </div>\n      <p class="modal-desc">\n        Dán đường link <b>Web App Google Apps Script</b> của trang tính vào đây để form trên GitHub lưu dữ liệu trực tiếp vào Google Sheet của bạn.\n      </p>\n\n      <div class="guide-steps">\n        <b>Cách lấy link Web App (chỉ cần làm 1 lần):</b>\n        <ol>\n          <li>Mở Google Sheet ➔ <b>Tiện ích mở rộng</b> ➔ <b>Apps Script</b>.</li>\n          <li>Góc trên bên phải, bấm nút xanh <b>Triển khai</b> ➔ <b>Triển khai mới</b>.</li>\n          <li>Chọn loại: <b>Ứng dụng web</b>. Mục <i>Ai có quyền truy cập</i> chọn: <b>Bất kỳ ai (Anyone)</b>.</li>\n          <li>Bấm <b>Triển khai</b> và sao chép <b>URL ứng dụng web</b> (dạng: <code>https://script.google.com/macros/s/.../exec</code>).</li>\n        </ol>\n      </div>\n\n      <div class="form-group">\n        <label class="form-label" for="gasUrlInput">URL Web App Google Apps Script:</label>\n        <input \n          type="url" \n          id="gasUrlInput" \n          class="form-control" \n          style="padding-left: 12px;" \n          placeholder="https://script.google.com/macros/s/.../exec"\n        />\n      </div>\n\n      <div class="button-group" style="margin-top: 14px;">\n        <button type="button" class="btn btn-reset" onclick="closeConfigModal()">\n          Đóng\n        </button>\n        <button type="button" class="btn btn-submit" onclick="saveGasUrl()">\n          💾 Lưu Cấu Hình\n        </button>\n      </div>\n    </div>\n  </div>\n\n  <script>\n    const STORAGE_KEY = \'crm_donhang_gas_url\';\n\n    // 1. Đồng hồ thời gian thực\n    function updateClock() {\n      const now = new Date();\n      const options = { \n        timeZone: \'Asia/Ho_Chi_Minh\', \n        day: \'2-digit\', month: \'2-digit\', year: \'numeric\', \n        hour: \'2-digit\', minute: \'2-digit\', second: \'2-digit\', \n        hour12: false \n      };\n      document.getElementById(\'liveClock\').textContent = new Intl.DateTimeFormat(\'vi-VN\', options).format(now);\n    }\n    setInterval(updateClock, 1000);\n    updateClock();\n\n    // 2. Khởi tạo trạng thái kết nối\n    window.addEventListener(\'DOMContentLoaded\', () => {\n      checkConnectionStatus();\n      document.getElementById(\'khachHang\').focus();\n    });\n\n    function getScriptUrl() {\n      return localStorage.getItem(STORAGE_KEY) || \'\';\n    }\n\n    function checkConnectionStatus() {\n      const url = getScriptUrl();\n      const badge = document.getElementById(\'connectionBadge\');\n      if (url && url.includes(\'script.google.com\')) {\n        badge.className = \'status-badge connected\';\n        badge.innerHTML = \'<span>🟢</span> Đã kết nối Google Sheet\';\n      } else {\n        badge.className = \'status-badge\';\n        badge.innerHTML = \'<span>⚪</span> Chưa kết nối Google Sheet\';\n      }\n    }\n\n    function openConfigModal() {\n      document.getElementById(\'gasUrlInput\').value = getScriptUrl();\n      document.getElementById(\'configModal\').style.display = \'flex\';\n    }\n\n    function closeConfigModal() {\n      document.getElementById(\'configModal\').style.display = \'none\';\n    }\n\n    function saveGasUrl() {\n      const val = document.getElementById(\'gasUrlInput\').value.trim();\n      if (!val) {\n        alert(\'Vui lòng nhập đường dẫn URL Web App!\');\n        return;\n      }\n      if (!val.startsWith(\'https://script.google.com/macros/s/\')) {\n        alert(\'Đường dẫn chưa đúng định dạng. Phải bắt đầu bằng: https://script.google.com/macros/s/...\');\n        return;\n      }\n      localStorage.setItem(STORAGE_KEY, val);\n      checkConnectionStatus();\n      closeConfigModal();\n      showAlert(\'success\', \'Đã lưu cấu hình kết nối Google Sheet!\');\n    }\n\n    // 3. Phím tắt\n    document.addEventListener(\'keydown\', (e) => {\n      if ((e.ctrlKey || e.metaKey) && e.key === \'Enter\') {\n        const form = document.getElementById(\'customerForm\');\n        if (form.checkValidity()) handleSubmit(e);\n        else form.reportValidity();\n      }\n    });\n\n    function showAlert(type, message) {\n      const box = document.getElementById(\'alertBox\');\n      box.className = \'alert \' + type;\n      document.getElementById(\'alertIcon\').textContent = type === \'success\' ? \'✅\' : \'⚠️\';\n      document.getElementById(\'alertMessage\').textContent = message;\n      if (type === \'success\') {\n        setTimeout(() => { box.style.display = \'none\'; }, 4500);\n      }\n    }\n\n    function resetForm() {\n      document.getElementById(\'customerForm\').reset();\n      document.getElementById(\'alertBox\').style.display = \'none\';\n      document.getElementById(\'khachHang\').focus();\n    }\n\n    // 4. Xử lý Gửi Dữ Liệu\n    async function handleSubmit(e) {\n      if (e) e.preventDefault();\n\n      const khachHang = document.getElementById(\'khachHang\').value.trim();\n      const sdt = document.getElementById(\'sdt\').value.trim();\n      const sanPham = document.getElementById(\'sanPham\').value.trim();\n      const nhanVien = document.getElementById(\'nhanVien\').value.trim();\n\n      if (!khachHang || !sdt) {\n        showAlert(\'error\', \'Vui lòng điền đủ Tên Khách Hàng và Số Điện Thoại!\');\n        return;\n      }\n\n      const scriptUrl = getScriptUrl();\n\n      // Kiểm tra nếu chưa cấu hình URL trên GitHub Pages\n      if (!scriptUrl && (typeof google === \'undefined\' || !google.script)) {\n        openConfigModal();\n        showAlert(\'error\', \'Vui lòng dán link Web App Google Apps Script để kết nối với Google Sheet!\');\n        return;\n      }\n\n      const btn = document.getElementById(\'btnSubmit\');\n      const spinner = document.getElementById(\'btnSpinner\');\n      const btnText = document.getElementById(\'btnText\');\n\n      btn.disabled = true;\n      spinner.style.display = \'inline-block\';\n      btnText.textContent = \'Đang lưu vào Sheet...\';\n\n      const payload = { khachHang, sdt, sanPham, nhanVien };\n\n      // Trường hợp 1: Chạy trực tiếp trong Google Apps Script (Sidebar/Dialog)\n      if (typeof google !== \'undefined\' && google.script && google.script.run) {\n        google.script.run\n          .withSuccessHandler(onSaveSuccess)\n          .withFailureHandler(onSaveError)\n          .saveCustomerData(payload);\n        return;\n      }\n\n      // Trường hợp 2: Chạy trên GitHub Pages (Gửi qua Webhook POST tới Google Apps Script)\n      try {\n        const response = await fetch(scriptUrl, {\n          method: \'POST\',\n          headers: {\n            \'Content-Type\': \'text/plain;charset=utf-8\'\n          },\n          body: JSON.stringify(payload)\n        });\n\n        const result = await response.json();\n        onSaveSuccess(result);\n      } catch (err) {\n        // Dự phòng: một số trình duyệt chặn CORS ở chế độ POST, thử gửi qua GET fallback\n        try {\n          const params = new URLSearchParams(payload).toString();\n          const fallbackUrl = scriptUrl + (scriptUrl.includes(\'?\') ? \'&\' : \'?\') + params;\n          const fallbackRes = await fetch(fallbackUrl);\n          const fallbackJson = await fallbackRes.json();\n          onSaveSuccess(fallbackJson);\n        } catch (fallbackErr) {\n          onSaveError(err);\n        }\n      }\n    }\n\n    function onSaveSuccess(response) {\n      const btn = document.getElementById(\'btnSubmit\');\n      const spinner = document.getElementById(\'btnSpinner\');\n      const btnText = document.getElementById(\'btnText\');\n\n      btn.disabled = false;\n      spinner.style.display = \'none\';\n      btnText.textContent = \'💾 Lưu Thông Tin\';\n\n      if (response && response.success) {\n        showAlert(\'success\', response.message);\n        \n        document.getElementById(\'lastSavedContent\').innerHTML = `\n          <b>#${response.data.stt}:</b> ${escapeHtml(response.data.khachHang)} (${escapeHtml(response.data.sdt)})<br>\n          <b>Sản phẩm:</b> ${escapeHtml(response.data.sanPham || \'---\')} | <b>Nhân viên:</b> ${escapeHtml(response.data.nhanVien || \'---\')}<br>\n          <i>🕒 ${response.data.thoiGian}</i>\n        `;\n        document.getElementById(\'lastSavedCard\').classList.add(\'active\');\n\n        // Giữ lại tên nhân viên cho khách kế tiếp\n        const nv = document.getElementById(\'nhanVien\').value;\n        document.getElementById(\'customerForm\').reset();\n        document.getElementById(\'nhanVien\').value = nv;\n        document.getElementById(\'khachHang\').focus();\n      } else {\n        showAlert(\'error\', response ? response.message : \'Có lỗi khi lưu vào Sheet!\');\n      }\n    }\n\n    function onSaveError(err) {\n      const btn = document.getElementById(\'btnSubmit\');\n      const spinner = document.getElementById(\'btnSpinner\');\n      const btnText = document.getElementById(\'btnText\');\n\n      btn.disabled = false;\n      spinner.style.display = \'none\';\n      btnText.textContent = \'💾 Lưu Thông Tin\';\n      showAlert(\'error\', \'Lỗi kết nối tới Google Sheet: \' + err.toString());\n    }\n\n    function escapeHtml(str) {\n      if (!str) return \'\';\n      return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");\n    }\n  </script>\n</body>\n</html>\n';
}
