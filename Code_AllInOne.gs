/**
 * =========================================================================================
 * BẢN TẤT CẢ TRONG 1 (ALL-IN-ONE): GOOGLE APPS SCRIPT FORM NHẬP LIỆU KHÁCH HÀNG KÈM NGÀY GIỜ
 * TỐI ƯU TỐC ĐỘ LƯU CỰC NHANH (SIÊU TỐC) + GỢI Ý NHÂN VIÊN TỪ TAB "DS NHÂN VIÊN"
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

function showSidebar() {
  const html = HtmlService.createHtmlOutput(getFormHtml())
    .setTitle("Nhập Thông Tin Khách Hàng")
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

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

  // 2. Lấy danh sách nhân viên từ tab "ds nhân viên" để gợi ý
  if (e && e.parameter && e.parameter.action === "getStaff") {
    const staffResult = getStaffList();
    return createJsonResponse(staffResult, e.parameter.callback);
  }

  // 3. Nhận dữ liệu gửi qua GET từ GitHub Pages (nếu có)
  if (e && e.parameter && (e.parameter.khachHang || e.parameter.sdt)) {
    const result = saveCustomerData(e.parameter);
    return createJsonResponse(result, e.parameter.callback);
  }

  // 4. Mở giao diện Web App trực tiếp
  return HtmlService.createHtmlOutput(getFormHtml())
    .setTitle("Form Nhập Liệu Khách Hàng")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Xử lý yêu cầu POST từ GitHub Pages (nhận cả form-data và json) - Tối ưu siêu tốc
 */
function doPost(e) {
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
      message: "Lỗi: " + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Đọc danh sách nhân viên từ tab "ds nhân viên"
 */
function getStaffList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    let staffSheet = null;

    for (let s of sheets) {
      const sName = s.getName().trim().toLowerCase();
      if (sName === "ds nhân viên" || sName === "ds nhan vien" || sName.includes("nhân viên") || sName.includes("nhan vien")) {
        staffSheet = s;
        break;
      }
    }

    if (!staffSheet) {
      return { success: false, staff: [], message: "Chưa tìm thấy trang tính 'ds nhân viên'" };
    }

    const lastRow = staffSheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, staff: [] };
    }

    const lastCol = Math.max(staffSheet.getLastColumn(), 2);
    const data = staffSheet.getRange(1, 1, lastRow, lastCol).getValues();

    let userColIdx = 0;
    let nameColIdx = 1;

    if (data.length > 0) {
      const h0 = (data[0][0] || "").toString().toLowerCase();
      const h1 = (data[0][1] || "").toString().toLowerCase();
      if (h0.includes("tên") || h0.includes("họ")) {
        nameColIdx = 0;
        userColIdx = 1;
      }
    }

    const staff = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const valUser = (row[userColIdx] !== undefined && row[userColIdx] !== null) ? row[userColIdx].toString().trim() : "";
      const valName = (row[nameColIdx] !== undefined && row[nameColIdx] !== null) ? row[nameColIdx].toString().trim() : "";

      if (valUser || valName) {
        const userCode = valUser;
        const fullName = valName || valUser;
        staff.push({
          user: userCode,
          name: fullName,
          label: (userCode && userCode !== fullName) ? `${userCode} - ${fullName}` : fullName
        });
      }
    }

    return { success: true, staff: staff };
  } catch (err) {
    return { success: false, staff: [], message: err.toString() };
  }
}

/**
 * Tạo phản hồi JSON hỗ trợ cả JSONP
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
    "1. Nhập thông tin Khách hàng, SĐT, Sản phẩm và Nhân viên.\n" +
    "2. Bấm 'Lưu Thông Tin' để lưu vào trang tính kèm STT và ngày giờ.\n" +
    "3. Số điện thoại được giữ nguyên số 0 ở đầu.",
    ui.ButtonSet.OK
  );
}

/**
 * Hàm ghi dữ liệu vào Google Sheet - Tối ưu 1 lệnh RPC duy nhất
 */
function saveCustomerData(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = (SHEET_NAME && SHEET_NAME.trim() !== "") ? ss.getSheetByName(SHEET_NAME.trim()) : ss.getSheets()[0];
    if (!sheet) return { success: false, message: "Không tìm thấy trang tính phù hợp!" };

    // 1. Tính STT
    const lastRow = sheet.getLastRow();
    const stt = lastRow >= 1 ? lastRow : 1;

    // 2. Lấy thời gian hiện tại chuẩn giờ Việt Nam (GMT+7)
    const now = new Date();
    const formattedDate = Utilities.formatDate(now, "Asia/Ho_Chi_Minh", "dd/MM/yyyy HH:mm:ss");

    // 3. Chuẩn hóa dữ liệu
    const khachHang = (data.khachHang || "").trim();
    const sdtRaw = (data.sdt || "").trim();
    const sdt = sdtRaw.startsWith("'") ? sdtRaw : `'${sdtRaw}`;
    const sanPham = (data.sanPham || "").trim();
    const nhanVien = (data.nhanVien || "").trim();

    // 4. Ghi dòng mới vào Sheet trong 1 cuộc gọi duy nhất
    sheet.appendRow([stt, khachHang, sdt, sanPham, nhanVien, formattedDate]);

    return {
      success: true,
      message: `Đã lưu thành công khách hàng #${stt}: ${khachHang}`,
      stt: stt,
      khachHang: khachHang,
      thoiGian: formattedDate
    };
  } catch (error) {
    return { success: false, message: "Lỗi: " + error.toString() };
  }
}

function getFormHtml() {
  return "<p>Vui lòng sử dụng giao diện trên GitHub Pages</p>";
}
