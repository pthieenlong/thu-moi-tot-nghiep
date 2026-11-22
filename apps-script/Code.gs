/**
 * Sổ lưu bút cho thiệp tốt nghiệp: nhận hồi âm + lời chúc + ảnh, ghi vào Google Sheet, ảnh lưu vào Google Drive.
 *
 * CÀI ĐẶT (làm 1 lần, khoảng 5 phút):
 *  1. Tạo một Google Sheet mới (vd "Lưu bút Anh Thư").
 *  2. Trong Sheet: Tiện ích mở rộng → Apps Script. Xoá code mẫu, dán toàn bộ file này vào, bấm Lưu.
 *  3. Bấm Triển khai → Tùy chọn triển khai mới → chọn loại "Ứng dụng web":
 *       - Thực thi dưới dạng: Tôi
 *       - Người có quyền truy cập: Bất kỳ ai
 *     Bấm Triển khai, cấp quyền (Google sẽ cảnh báo "ứng dụng chưa xác minh" → Nâng cao → Đi tới ... → Cho phép).
 *  4. Copy "URL ứng dụng web" (dạng https://script.google.com/macros/s/.../exec)
 *     và dán vào CFG.sheetUrl trong index.html.
 *
 * Sửa code này sau khi đã triển khai: Triển khai → Quản lý triển khai → bút chì → Phiên bản: "Phiên bản mới" → Triển khai.
 * (Nếu tạo triển khai mới thì URL đổi, phải dán lại vào index.html.)
 *
 * Ẩn một lời chúc khỏi sổ lưu bút trên thiệp: gõ bất kỳ chữ gì (vd "x") vào cột "Ẩn" của dòng đó.
 * Ảnh khách gửi nằm trong thư mục Drive "Lưu bút - ảnh", ai có link mới xem được (để thiệp hiện được ảnh).
 */

const SHEET_NAME = 'Lưu bút';
const FOLDER_NAME = 'Lưu bút - ảnh';
const HEADER = ['Thời gian', 'Tên', 'Tham dự', 'Số người', 'Lời chúc', 'Ảnh', 'Ẩn'];
const MAX_PHOTO_BYTES = 3 * 1024 * 1024; // the page sends ~200KB; anything far bigger isn't from it

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json({ ok: false, error: 'Dữ liệu không hợp lệ' }); }

  // This URL is public, so treat every field as untrusted.
  const name = clean(d.name, 60);
  const message = clean(d.message, 1000);
  const attend = d.attend === 'yes' ? 'Có' : d.attend === 'no' ? 'Không' : '';
  const guests = d.attend === 'yes' ? Math.max(1, Math.min(5, parseInt(d.guests, 10) || 1)) : '';
  if (!name) return json({ ok: false, error: 'Bạn điền tên giúp Thư nha' });
  if (!attend) return json({ ok: false, error: 'Bạn chọn có đến hay không giúp Thư nha' });

  let photoUrl = '';
  if (d.photo) {
    const m = /^data:image\/(jpeg|png|webp);base64,([A-Za-z0-9+/=]+)$/.exec(d.photo);
    if (!m) return json({ ok: false, error: 'Ảnh không hợp lệ' });
    const bytes = Utilities.base64Decode(m[2]);
    if (bytes.length > MAX_PHOTO_BYTES) return json({ ok: false, error: 'Ảnh quá lớn' });
    const file = folder().createFile(Utilities.newBlob(bytes, 'image/' + m[1], Date.now() + '-' + name + '.' + m[1]));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    photoUrl = 'https://drive.google.com/file/d/' + file.getId() + '/view';
  }

  // A leading = + - @ would make Sheets run the text as a formula.
  const safe = v => (typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : v);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    sheet().appendRow([new Date(), safe(name), attend, guests, safe(message), photoUrl, '']);
  } finally {
    lock.releaseLock();
  }
  return json({ ok: true });
}

// Public guestbook: only name, message and photo. Attendance stays private in the Sheet.
function doGet() {
  const rows = sheet().getDataRange().getValues().slice(1);
  const list = rows
    .filter(r => !r[6] && (r[4] || r[5]))
    .reverse()
    .slice(0, 100)
    .map(r => ({ name: String(r[1]), message: String(r[4]), photo: (/\/d\/([\w-]+)/.exec(r[5]) || [])[1] || '' }));
  return json({ ok: true, list });
}

function clean(v, max) {
  return String(v == null ? '' : v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim().slice(0, max);
}

function sheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADER);
    sh.setFrozenRows(1);
  }
  return sh;
}

function folder() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
