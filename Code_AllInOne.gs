/**
 * =========================================================================================
 * BẢN TẤT CẢ TRONG 1 (ALL-IN-ONE): GOOGLE APPS SCRIPT FORM NHẬP LIỆU KHÁCH HÀNG KÈM NGÀY GIỜ
 * HỖ TRỢ: CHẠY TRỰC TIẾP TRONG GOOGLE SHEETS + KẾT NỐI VỚI GITHUB PAGES (JSONP & POST)
 * =========================================================================================
 */

// Cấu hình tên Sheet: Để trống "" sẽ tự động lấy trang tính đầu tiên
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
 * Xử lý yêu cầu GET từ Web ngoài (GitHub Pages) hoặc mở Web App trực tiếp
 */
function doGet(e) {
  // 1. Kiểm tra kết nối (Ping Test)
  if (e && e.parameter && e.parameter.action === "ping") {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = (SHEET_NAME && SHEET_NAME.trim() !== "") ? ss.getSheetByName(SHEET_NAME.trim()) : ss.getSheets()[0];
    const pingResult = {
      success: true,
      message: "Kết nối thành công tới Google Sheet!",
      sheetName: sheet ? sheet.getName() : "Sheet1"
    };
    return createJsonResponse(pingResult, e.parameter.callback);
  }

  // 2. Nhận dữ liệu gửi qua GET / JSONP từ GitHub Pages
  if (e && e.parameter && (e.parameter.khachHang || e.parameter.sdt)) {
    const result = saveCustomerData(e.parameter);
    return createJsonResponse(result, e.parameter.callback);
  }

  // 3. Mở giao diện Web App trực tiếp
  return HtmlService.createHtmlOutput(getFormHtml())
    .setTitle("Form Nhập Liệu Khách Hàng")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Xử lý yêu cầu POST từ GitHub Pages (nhận cả form-data và json)
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
    if (e.parameter && (e.parameter.khachHang || e.parameter.sdt)) {
      data = e.parameter;
    } else if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (parseErr) {
        data = e.parameter || {};
      }
    }

    const result = saveCustomerData(data);
    return createJsonResponse(result, e.parameter ? e.parameter.callback : null);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: "Lỗi server Apps Script: " + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Tạo phản hồi JSON hỗ trợ cả JSONP (tránh 100% lỗi CORS trình duyệt)
 */
function createJsonResponse(dataObj, callbackName) {
  const jsonStr = JSON.stringify(dataObj);
  if (callbackName) {
    return ContentService.createTextOutput(callbackName + "(" + jsonStr + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(jsonStr)
    .setMimeType(ContentService.MimeType.JSON);
}

function showHelp() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    "HƯỚNG DẪN SỬ DỤNG",
    "1. Bấm 'Mở Form Nhập Liệu' để mở form.\n" +
    "2. Nhập đầy đủ thông tin Khách hàng, SĐT, Sản phẩm và Nhân viên.\n" +
    "3. Bấm 'Lưu Thông Tin', hệ thống tự động tăng STT và lưu ngày giờ chính xác.\n" +
    "4. Số điện thoại được giữ nguyên số 0 ở đầu.",
    ui.ButtonSet.OK
  );
}

/**
 * Hàm ghi dữ liệu vào Google Sheet
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
      // Luôn lấy trang tính đầu tiên để tránh lỗi khi gọi từ Web
      sheet = ss.getSheets()[0];
    }

    if (!sheet) return { success: false, message: "Không tìm thấy trang tính phù hợp!" };

    // 1. Tự động thêm tiêu đề cột F: NGÀY GIỜ LƯU nếu chưa có
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
      message: `Đã lưu thành công khách hàng #${stt}: ${khachHang}`,
      sheetName: sheet.getName(),
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

function getFormHtml() {
  return "<p>Vui lòng sử dụng giao diện trên GitHub Pages hoặc tệp index.html</p>";
}
