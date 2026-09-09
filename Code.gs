/**
 * =========================================================================
 * MÃ NGUỒN GOOGLE APPS SCRIPT - FORM NHẬP LIỆU KHÁCH HÀNG TỰ ĐỘNG VÀO SHEET
 * =========================================================================
 */

// Cấu hình tên Sheet: Để trống "" sẽ tự động lấy trang tính đang mở
const SHEET_NAME = ""; 

/**
 * 1. Tự động tạo Menu "📋 Quản Lý Nhập Liệu" khi mở Google Sheet
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
 * Hiển thị form dưới dạng Sidebar bên phải màn hình
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Nhập Thông Tin Khách Hàng")
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Hiển thị form dưới dạng Hộp thoại (Dialog) ở giữa màn hình
 */
function showModalDialog() {
  const html = HtmlService.createHtmlOutputFromFile("Index")
    .setWidth(460)
    .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, "Nhập Thông Tin Khách Hàng");
}

/**
 * Mở qua link Web App độc lập (nhân viên có thể mở trên điện thoại/máy tính bảng)
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Form Nhập Liệu Khách Hàng")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Thông báo hướng dẫn nhanh
 */
function showHelp() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    "HƯỚNG DẪN SỬ DỤNG",
    "1. Chọn 'Mở Form Nhập Liệu (Bên phải - Sidebar)' để thao tác song song vừa nhập vừa xem trang tính.\n" +
    "2. Điền đầy đủ thông tin Khách hàng, SĐT, Sản phẩm và Nhân viên.\n" +
    "3. Bấm 'Lưu Thông Tin', hệ thống sẽ tự động đánh STT và ghi ngày giờ lưu chính xác (theo giờ Việt Nam).\n" +
    "4. Số điện thoại sẽ luôn được giữ nguyên số 0 ở đầu.",
    ui.ButtonSet.OK
  );
}

/**
 * Hàm nhận dữ liệu từ giao diện Form và lưu vào Trang Tính
 * @param {Object} data - { khachHang, sdt, sanPham, nhanVien }
 */
function saveCustomerData(data) {
  // Sử dụng LockService để tránh trường hợp nhiều người bấm lưu cùng lúc bị đè dòng
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // Chờ tối đa 10s
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

    // 1. Tự động kiểm tra và thêm tiêu đề cột F "NGÀY GIỜ LƯU" nếu chưa có
    const headerF = sheet.getRange(1, 6).getValue();
    if (!headerF || headerF.toString().trim() === "") {
      const headerCell = sheet.getRange(1, 6);
      headerCell.setValue("NGÀY GIỜ LƯU");
      
      // Định dạng cho đẹp mắt đồng bộ với hàng tiêu đề (Cột E)
      const sampleCell = sheet.getRange(1, 5);
      headerCell.setFontWeight("bold")
                .setBackground(sampleCell.getBackground() || "#fce5cd")
                .setFontColor(sampleCell.getFontColor() || "#783f04")
                .setFontFamily(sampleCell.getFontFamily() || "Arial")
                .setHorizontalAlignment("center")
                .setVerticalAlignment("middle");
      sheet.setColumnWidth(6, 170);
    }

    // 2. Tính STT tự động dựa vào dòng cuối cùng
    const lastRow = sheet.getLastRow();
    // Nếu chỉ có 1 dòng tiêu đề (lastRow = 1), STT tiếp theo là 1 (ở dòng 2)
    // Nếu chưa có dòng nào, STT là 1
    const stt = lastRow >= 1 ? lastRow : 1;
    const targetRow = lastRow + 1;

    // 3. Lấy thời gian hiện tại định dạng Việt Nam (GMT+7)
    const now = new Date();
    const formattedDate = Utilities.formatDate(now, "Asia/Ho_Chi_Minh", "dd/MM/yyyy HH:mm:ss");

    // 4. Chuẩn hóa dữ liệu
    const khachHang = (data.khachHang || "").trim();
    // Đặt dấu nháy đơn ' trước số điện thoại để Google Sheet không tự ý xóa số 0 ở đầu
    const sdtRaw = (data.sdt || "").trim();
    const sdt = sdtRaw.startsWith("'") ? sdtRaw : `'${sdtRaw}`;
    const sanPham = (data.sanPham || "").trim();
    const nhanVien = (data.nhanVien || "").trim();

    // 5. Ghi dữ liệu vào dòng mới: [STT, KHÁCH HÀNG, SĐT, SẢN PHẨM KHÁCH XEM, NHÂN VIÊN, NGÀY GIỜ LƯU]
    const rowValues = [
      stt,
      khachHang,
      sdt,
      sanPham,
      nhanVien,
      formattedDate
    ];

    sheet.appendRow(rowValues);

    // Căn chỉnh hiển thị: Cột STT, SĐT, NGÀY GIỜ căn giữa
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
