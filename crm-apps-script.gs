/**
 * HAMEE CRM – Google Apps Script (miễn phí)
 * Lưu toàn bộ dữ liệu CRM của webApp HAMEE vào 1 Google Sheet để mọi máy Admin dùng chung.
 *
 * CÀI ĐẶT (làm 1 lần, khoảng 5 phút):
 * 1. Tạo Google Sheet mới, đặt tên "HAMEE CRM". Không cần tạo tab, script tự tạo.
 * 2. Tiện ích mở rộng → Apps Script → xoá code mẫu → dán toàn bộ file này.
 * 3. Sửa ADMIN_KEY bên dưới thành một chuỗi bí mật (VD: hamee-2026-xyz). Không chia sẻ chuỗi này.
 * 4. Triển khai → Triển khai mới → Loại: Ứng dụng web
 *      Thực thi với tư cách: Tôi
 *      Người có quyền truy cập: Bất kỳ ai
 *    → Triển khai → cấp quyền → copy link Web App (…/exec).
 * 5. Trên webApp: Admin → ☰ → Cài đặt → dán link Web App + ADMIN_KEY → Lưu.
 *
 * Các tab được tạo: members (hội viên), fees (phiếu thu hội phí), care (nhật ký chăm sóc),
 * events (hoạt động / sự kiện, cột regs chứa danh sách tham dự), applications (phiếu đăng ký gia nhập), settings.
 * Ban thư ký có thể mở Sheet để xem / lọc / in. Nên sửa dữ liệu trên webApp để tránh ghi đè.
 */
// Cài theo cách 2 (script.google.com → Dự án mới): điền ID Google Sheet vào đây
const SHEET_ID = "";
const ss_ = () => SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActive();
const ADMIN_KEY = "DOI-THANH-MA-BI-MAT";
const TABLES = ["members", "fees", "care", "events", "applications", "settings", "site", "news"];

function doGet(e) {
  const p = e.parameter || {};
  // Dữ liệu công khai cho website (ai cũng đọc được): sự kiện (không kèm danh sách người đăng ký), tin tức, nội dung trang
  if (p.action === "public") {
    const ev = readTable("events").map(x => { const r = Array.isArray(x.regs) ? x.regs : []; const o = Object.assign({}, x); delete o.regs; o.regCount = r.length + readTable("public_regs").filter(q => String(q.eventId) === String(x.id) && !r.some(y => y.rid === q.rid)).length; return o; });
    // Danh bạ công khai: chỉ các trường hiển thị trên web (không có ghi chú nội bộ, di động, hội phí)
    const PUB = ["id", "name", "sector", "sectors", "city", "phone", "email", "web", "logo", "url", "intro", "products", "rep", "title", "bday", "founded", "joined", "status", "kcn", "factory"];
    const mem = readTable("members").filter(m => m.status !== "inactive").map(m => { const o = {}; PUB.forEach(k => { if (m[k] !== "" && m[k] !== undefined) o[k] = m[k]; }); return o; });
    return json({ events: ev, news: readTable("news").filter(n => n.pub !== false), site: readTable("site"), members: mem });
  }
  // ---- Đăng nhập hội viên bằng OTP email ----
  const email = String(p.email || "").trim().toLowerCase();
  const cache = CacheService.getScriptCache();
  if (p.action === "otp_send") {
    const m = email && (email === TEST_EMAIL || readTable("members").some(x => String(x.email || "").toLowerCase() === email && x.status !== "inactive"));
    if (!m) return json({ ok: false, error: "khong_phai_hoi_vien" });
    const n = +(cache.get("cnt_" + email) || 0);
    if (n >= 5) return json({ ok: false, error: "qua_nhieu" });
    const code = String(Math.floor(100000 + Math.random() * 900000));
    cache.put("otp_" + email, code, 600); cache.put("cnt_" + email, String(n + 1), 3600);
    MailApp.sendEmail({ to: email, subject: "Mã đăng nhập HAMEE HUB: " + code,
      body: "Mã đăng nhập tài khoản hội viên HAMEE HUB của anh/chị là: " + code + "\n\nMã có hiệu lực trong 10 phút. Không chia sẻ mã này cho người khác.\nNếu anh/chị không yêu cầu đăng nhập, vui lòng bỏ qua email này.\n\nHiệp hội Doanh nghiệp Cơ khí – Điện TP.HCM (HAMEE) · 028 3973 4081" });
    return json({ ok: true });
  }
  if (p.action === "otp_verify") {
    const tries = +(cache.get("try_" + email) || 0);
    if (tries >= 5) return json({ ok: false, error: "qua_nhieu" });
    const code = cache.get("otp_" + email);
    if (!code || code !== String(p.code || "").trim()) { cache.put("try_" + email, String(tries + 1), 600); return json({ ok: false, error: "sai_ma" }); }
    cache.remove("otp_" + email); cache.remove("try_" + email);
    return json({ ok: true, token: makeToken(email), me: myData(email) });
  }
  // Hội viên đã đăng nhập: lấy hội phí, sự kiện, cập nhật hồ sơ của chính mình
  if (p.action === "me") {
    const em = readToken(p.token);
    if (!em) return json({ ok: false, error: "het_han" });
    return json({ ok: true, me: myData(em) });
  }
  if (p.key !== ADMIN_KEY) return json({ error: "sai_ma_khoa" });
  if (p.action === "dump") {
    const out = {};
    TABLES.concat(["public_regs", "public_edits"]).forEach(t => out[t] = readTable(t));
    return json(out);
  }
  return json({ ok: true });
}

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  // Hội viên gửi cập nhật hồ sơ năng lực (cần token đăng nhập) → Admin duyệt trên trang quản trị
  if (d.action === "edit") {
    const em = readToken(d.token);
    if (!em) return json({ ok: false, error: "het_han" });
    const m = readTable("members").find(x => String(x.email || "").toLowerCase() === em);
    if (!m) return json({ ok: false, error: "khong_phai_hoi_vien" });
    const sh = sheetOf("public_edits");
    const keys = ["id", "mid", "date", "status", "intro", "products", "phone", "email", "web", "founded", "bday", "kcn", "factory", "capacity", "certs", "markets", "machines"];
    if (sh.getLastRow() === 0) sh.appendRow(keys);
    const e2 = Object.assign({}, d.edit, { mid: m.id, status: "pending", date: Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "yyyy-MM-dd") });
    sh.appendRow(keys.map(k => { const v = e2[k]; return v === undefined || v === null ? "" : (/^0\d+/.test(String(v)) || /^\d{4}-\d{2}-\d{2}$/.test(String(v)) ? "'" + v : v); }));
    try { MailApp.sendEmail({ to: Session.getEffectiveUser().getEmail(), subject: "[HAMEE HUB] Hội viên gửi cập nhật hồ sơ: " + m.name, body: m.name + " vừa gửi cập nhật hồ sơ năng lực. Vào trang quản trị → Duyệt → Hồ sơ DN để xem và duyệt." }); } catch (err) {}
    return json({ ok: true });
  }
  // Khách / hội viên đăng ký sự kiện trên website (không cần mã khoá)
  if (d.action === "reg") {
    const sh = sheetOf("public_regs");
    const keys = ["rid", "eventId", "eventTitle", "date", "name", "org", "phone", "email", "mid", "qty", "unit", "amount", "paid", "checked"];
    if (sh.getLastRow() === 0) sh.appendRow(keys);
    sh.appendRow(keys.map(k => { const v = d.reg[k]; return v === undefined || v === null ? "" : (/^0\d+/.test(String(v)) ? "'" + v : v); }));
    return json({ ok: true });
  }
  if (d.key !== ADMIN_KEY) return json({ error: "sai_ma_khoa" });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (d.action === "upload") {
      const ss = ss_(), par = DriveApp.getFileById(ss.getId()).getParents();
      const root = par.hasNext() ? par.next() : DriveApp.getRootFolder();
      const it = root.getFoldersByName("Poster sự kiện HAMEE");
      const folder = it.hasNext() ? it.next() : root.createFolder("Poster sự kiện HAMEE");
      const f = folder.createFile(Utilities.newBlob(Utilities.base64Decode(d.data), d.type || "image/jpeg", d.name || "poster.jpg"));
      try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (err) {}
      return json({ ok: true, id: f.getId() });
    }
    if (d.action === "put") {
      Object.keys(d.tables || {}).forEach(t => { if (TABLES.includes(t)) writeTable(t, d.tables[t] || []); });
      backupDaily();
    }
  } finally { lock.releaseLock(); }
  return json({ ok: true });
}

// ---- Token đăng nhập hội viên (ký bằng ADMIN_KEY, hiệu lực 30 ngày) ----
const TEST_EMAIL = "hoicokhidien@gmail.com";
function sign_(s) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(s, ADMIN_KEY)).slice(0, 32); }
function makeToken(email) { const body = email + "|" + (Date.now() + 30 * 864e5); return Utilities.base64EncodeWebSafe(body) + "." + sign_(body); }
function readToken(tok) {
  try {
    const parts = String(tok || "").split(".");
    const body = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString();
    if (sign_(body) !== parts[1]) return null;
    const a = body.split("|");
    return +a[1] > Date.now() ? a[0] : null;
  } catch (err) { return null; }
}
// Dữ liệu riêng của 1 hội viên: hồ sơ, phiếu thu hội phí, sự kiện đã đăng ký / tham gia, cập nhật hồ sơ đã gửi
function myData(email) {
  if (email === TEST_EMAIL) return { test: true };
  const m = readTable("members").find(x => String(x.email || "").toLowerCase() === email);
  if (!m) return null;
  const regs = [];
  readTable("events").forEach(e => (Array.isArray(e.regs) ? e.regs : []).forEach(r => { if (String(r.mid) === String(m.id)) regs.push(Object.assign({ eventId: e.id }, r)); }));
  readTable("public_regs").forEach(r => { if (String(r.mid) === String(m.id) && !regs.some(x => x.rid === r.rid)) regs.push(r); });
  const st = readTable("settings")[0] || {};
  return { member: m, fees: readTable("fees").filter(f => String(f.mid) === String(m.id)),
    regs: regs, edits: readTable("public_edits").filter(x => String(x.mid) === String(m.id)),
    settings: { year: st.year, feeAmount: st.feeAmount, feeDue: st.feeDue } };
}

function sheetOf(name) {
  const ss = ss_();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function readTable(name) {
  const sh = sheetOf(name), v = sh.getDataRange().getValues();
  if (v.length < 2) return [];
  const h = v.shift();
  return v.filter(r => r.some(x => x !== "")).map(r => {
    const o = {};
    h.forEach((k, i) => {
      let x = r[i];
      if (x instanceof Date) x = Utilities.formatDate(x, "Asia/Ho_Chi_Minh", "yyyy-MM-dd");
      if (typeof x === "string" && /^[\[{]/.test(x)) { try { x = JSON.parse(x); } catch (err) {} }
      if (x === "TRUE" || x === true) x = true; else if (x === "FALSE" || x === false) x = false;
      o[k] = x;
    });
    return o;
  });
}

function writeTable(name, rows) {
  const sh = sheetOf(name);
  const keys = [];
  rows.forEach(r => Object.keys(r || {}).forEach(k => { if (!keys.includes(k)) keys.push(k); }));
  sh.clearContents();
  if (!keys.length) return;
  const data = [keys].concat(rows.map(r => keys.map(k => {
    const x = r[k];
    if (x === undefined || x === null) return "";
    if (typeof x === "object") return JSON.stringify(x);
    if (typeof x === "string" && /^\d{4}-\d{2}-\d{2}$/.test(x)) return "'" + x;   // giữ nguyên dạng ngày
    if (typeof x === "string" && /^0\d+/.test(x)) return "'" + x;                 // giữ số 0 đầu SĐT
    return x;
  })));
  sh.getRange(1, 1, data.length, keys.length).setValues(data);
  sh.setFrozenRows(1);
}

// Sao lưu 1 bản mỗi ngày vào thư mục Drive "HAMEE CRM backup"
function backupDaily() {
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "yyyy-MM-dd");
  if (props.getProperty("lastBackup") === today) return;
  const ss = ss_();
  const it = DriveApp.getFoldersByName("HAMEE CRM backup");
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder("HAMEE CRM backup");
  DriveApp.getFileById(ss.getId()).makeCopy("HAMEE CRM " + today, folder);
  props.setProperty("lastBackup", today);
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// Chạy 1 lần trong trình soạn thảo để cấp quyền (Sheet, Drive) trước khi Triển khai
function capQuyen() {
  ss_().getName();
  DriveApp.getRootFolder();
  MailApp.getRemainingDailyQuota();
}
