/**
 * Nhận đề nghị phân quyền từ trang web và ghi vào Google Sheet.
 * Cách dùng: xem tab "Huong dan" trong Google Sheet.
 */
const SPREADSHEET_ID = '1mgJxu4KLsiPWJ-0yBlFglqrGvFkQJXTFsjVzFr_RsAA';
const SHEET_NAME = 'Phan quyen';
const HEADERS = ['Thời gian nhận', 'Thời gian QL gửi', 'Người đề nghị', 'Email đề nghị', 'Folder / file',
  'Họ tên được cấp', 'Email được cấp', 'Quyền', 'Mail cá nhân thành viên khác (chung)',
  'Ban Giám đốc', 'Trạng thái', 'Mã gửi'];
const MAX_ROWS = 500;

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

/* tránh ô bị hiểu là công thức */
function safe_(v) {
  v = String(v == null ? '' : v);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return out_({ ok: true, msg: 'Web app nhận đề nghị phân quyền đang hoạt động' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const d = JSON.parse(e.postData.contents);
    if (!d || !d.mgr || !Array.isArray(d.rows)) return out_({ ok: false, error: 'Dữ liệu không hợp lệ' });
    if (d.rows.length > MAX_ROWS) return out_({ ok: false, error: 'Quá nhiều dòng' });
    const sh = getSheet_();
    const id = String(d.id || '');
    if (id) { /* chống ghi trùng khi bấm gửi 2 lần */
      const last = sh.getLastRow();
      if (last > 1) {
        const n = Math.min(last - 1, 1000);
        const ids = sh.getRange(last - n + 1, HEADERS.length, n, 1).getValues();
        if (ids.some(r => r[0] === id)) return out_({ ok: true, dup: true });
      }
    }
    const now = new Date();
    const base = [now, safe_(d.time), safe_(d.mgrName), safe_(d.mgr)];
    const tail = [safe_(d.note), 'Mặc định có mặt ở tất cả folder', 'Chờ xử lý', id];
    const rows = d.rows.map(r => base.concat([safe_(r.folder), safe_(r.name), safe_(r.email), safe_(r.perm)], tail));
    if (!rows.length && d.note) rows.push(base.concat(['', '', '', ''], tail));
    if (!rows.length) return out_({ ok: false, error: 'Không có nội dung để ghi' });
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
    return out_({ ok: true, count: rows.length });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}
