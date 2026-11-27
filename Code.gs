function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNames = ['TaiKhoan', 'DiTre_DiemDanh', 'DonNghiPhep', 'DatPhong'];

  // Tự động tạo 4 trang tính nếu chưa tồn tại
  sheetNames.forEach(function(name) {
    if (!ss.getSheetByName(name)) {
      ss.insertSheet(name);
    }
  });

  // Nhận dữ liệu từ 1 trong 4 trang (login.html, student.html, parent.html, teacher.html)
  if (e && e.postData && e.postData.contents) {
    var payload = JSON.parse(e.postData.contents);
    if (payload.action === 'SYNC_ALL' && payload.data) {
      writeSheet(ss.getSheetByName('TaiKhoan'), ['ID', 'HoTen', 'VaiTro', 'Lop_DonVi', 'DiemRenLuyen', 'TrangThai'], payload.data.users);
      writeSheet(ss.getSheetByName('DiTre_DiemDanh'), ['MaBanGhi', 'MaHS', 'HoTen', 'Loai', 'ThoiGian', 'LyDo', 'BienDongDiem'], payload.data.attendanceLogs);
      writeSheet(ss.getSheetByName('DonNghiPhep'), ['MaDon', 'MaHS', 'HoTen', 'NgayNghi', 'LoaiNghi', 'LyDo', 'TrangThai', 'ThoiGianGui'], payload.data.leaves);
      writeSheet(ss.getSheetByName('DatPhong'), ['MaDat', 'MaPhong', 'TenPhong', 'Ngay', 'KhungGio', 'NguoiDat', 'MucDich', 'TrangThai'], payload.data.bookings);
    }
  }

  // Trả dữ liệu mới nhất về cho tất cả các trang đang mở
  var output = {
    status: 'success',
    updatedAt: new Date().toISOString(),
    users: readSheet(ss.getSheetByName('TaiKhoan')),
    attendanceLogs: readSheet(ss.getSheetByName('DiTre_DiemDanh')),
    leaves: readSheet(ss.getSheetByName('DonNghiPhep')),
    bookings: readSheet(ss.getSheetByName('DatPhong'))
  };

  return ContentService.createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}

function writeSheet(sheet, headers, rows) {
  sheet.clearContents();
  sheet.appendRow(headers);
  if (rows && rows.length > 0) {
    var matrix = rows.map(function(row) {
      return headers.map(function(h) { return row[h] !== undefined ? row[h] : ''; });
    });
    sheet.getRange(2, 1, matrix.length, headers.length).setValues(matrix);
  }
}

function readSheet(sheet) {
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  return values.slice(1).map(function(r) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = r[i]; });
    return obj;
  });
}