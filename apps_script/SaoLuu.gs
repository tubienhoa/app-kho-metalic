/**
 * SAO LƯU HẰNG ĐÊM file dữ liệu APP KHO sang thư mục Drive "APP_KHO_SAOLUU".
 * Cách cài (làm 1 lần, bằng tài khoản Google của bạn):
 *  1. Mở file Google Sheets dữ liệu → Tiện ích mở rộng → Apps Script.
 *  2. Xóa nội dung có sẵn, dán toàn bộ file này, bấm Lưu.
 *  3. Chọn hàm caiLich → Chạy → cho phép quyền truy cập Drive (lần đầu).
 *  Từ đó mỗi đêm khoảng 1–2 giờ sáng tự tạo 1 bản sao, giữ 30 bản mới nhất.
 */
const THU_MUC = 'APP_KHO_SAOLUU';
const GIU_LAI = 30;

function saoLuu() {
  const file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  const it = DriveApp.getFoldersByName(THU_MUC);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(THU_MUC);
  const ten = 'SAOLUU_' + Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd_HHmm');
  file.makeCopy(ten, folder);
  // xóa bản cũ (đưa vào Thùng rác, khôi phục được 30 ngày)
  const ds = [];
  const files = folder.getFiles();
  while (files.hasNext()) { const f = files.next(); if (f.getName().indexOf('SAOLUU_') === 0) ds.push(f); }
  ds.sort((a, b) => b.getDateCreated() - a.getDateCreated());
  ds.slice(GIU_LAI).forEach(f => f.setTrashed(true));
}

function caiLich() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'saoLuu') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('saoLuu').timeBased().everyDays(1).atHour(1).inTimezone('Asia/Ho_Chi_Minh').create();
  saoLuu();
}
