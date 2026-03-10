// ============================================================
// Code.gs — Village Management System Backend (GAS)
// ============================================================
// วิธีใช้: แก้ไข SHEET_ID ให้ตรงกับ Google Sheet ของคุณ
// แล้ว Deploy เป็น Web App (Execute as: Me, Anyone can access)
// ============================================================

const SHEET_ID = '1MDX7JWY33m1lqtHbGGVQXbz3fPP_Dxllm5U_B5ixOhk';
const DRIVE_FOLDER_ID = '1RF2J9YDSmhg_iGLzvyAGRqzSMQ1FuaHw';

// Sheet names
const SHEETS = {
  USERS: 'Users',
  HOUSES: 'Houses',
  COMMON_FEE: 'CommonFee',
  ANNOUNCEMENTS: 'Announcements',
  NITI_REPORT: 'NitiReport'
};

// ── Token Security ────────────────────────────────────────────
// ⚠️  เปลี่ยน TOKEN_SECRET ให้เป็นค่าลับของคุณก่อน Deploy
const TOKEN_SECRET = 'VMS_SECRET_CHANGE_ME_2024!@#$';
const TOKEN_TTL_MS = 8 * 60 * 60 * 1000; // 8 ชั่วโมง

function signPayload(payload) {
  var key   = Utilities.newBlob(TOKEN_SECRET).getBytes();
  var bytes = Utilities.newBlob(payload).getBytes();
  var sig   = Utilities.computeHmacSha256Signature(bytes, key);
  return sig.slice(0, 8).map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function makeToken(user) {
  var now = Date.now();
  var payload = JSON.stringify({
    user_id: user.user_id, username: user.username,
    role: user.role, house_id: user.house_id,
    iat: now, exp: now + TOKEN_TTL_MS
  });
  var b64 = Utilities.base64Encode(payload);
  var sig = signPayload(b64);
  return b64 + '.' + sig;
}

function verifyToken(token) {
  if (!token) return null;
  try {
    // รองรับ token เก่า (ไม่มี ".") ชั่วคราว
    if (token.indexOf('.') === -1) {
      var decoded = Utilities.base64Decode(token);
      return JSON.parse(Utilities.newBlob(decoded).getDataAsString());
    }
    var parts = token.split('.');
    if (parts.length !== 2) return null;
    var b64 = parts[0], sig = parts[1];
    // ตรวจ signature
    if (signPayload(b64) !== sig) return null;
    // ตรวจ expiry
    var obj = JSON.parse(Utilities.newBlob(Utilities.base64Decode(b64)).getDataAsString());
    if (obj.exp && Date.now() > obj.exp) return null;
    return obj;
  } catch (e) { return null; }
}

// ── Entry Points ──────────────────────────────────────────────

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  // GAS Web App: POST with application/x-www-form-urlencoded lands in e.parameter
  // POST with JSON body lands in e.postData.contents
  return handleRequest(e);
}

function handleRequest(e) {
  let params = e.parameter || {};

  // รองรับ POST ที่ส่ง JSON body (เช่น uploadFile ที่มี base64 ขนาดใหญ่)
  if (e.postData && e.postData.contents) {
    try {
      const body = JSON.parse(e.postData.contents);
      Object.keys(body).forEach(k => { params[k] = body[k]; });
    } catch (err) { /* ignore */ }
  }

  const action = params['action'] || '';
  const data = params;

  let result;
  try {
    switch (action) {
      // Auth
      case 'login':            result = login(data); break;

      // Resident
      case 'getMyInfo':        result = getMyInfo(data); break;
      case 'getMyFees':        result = getMyFees(data); break;

      // Public
      case 'getFeeSummary':    result = getFeeSummary(data); break;
      case 'getAnnouncements': result = getAnnouncements(data); break;
      case 'getNitiReports':   result = getNitiReports(data); break;

      // Admin — Houses
      case 'getHouses':        result = requireAdmin(data, getHouses); break;
      case 'addHouse':         result = requireAdmin(data, addHouse); break;
      case 'updateHouse':      result = requireAdmin(data, updateHouse); break;
      case 'deleteHouse':      result = requireAdmin(data, deleteHouse); break;

      // Admin — Users
      case 'getUsers':         result = requireAdmin(data, getUsers); break;
      case 'addUser':          result = requireAdmin(data, addUser); break;
      case 'updateUser':       result = requireAdmin(data, updateUser); break;

      // Admin — Common Fee
      case 'getFees':          result = requireAdmin(data, getFees); break;
      case 'addFee':           result = requireAdmin(data, addFee); break;
      case 'updateFee':        result = requireAdmin(data, updateFee); break;

      // Admin — Announcements
      case 'addAnnouncement':  result = requireAdmin(data, addAnnouncement); break;
      case 'updateAnnouncement': result = requireAdmin(data, updateAnnouncement); break;
      case 'deleteAnnouncement': result = requireAdmin(data, deleteAnnouncement); break;

      // Admin — Niti Report
      case 'addNitiReport':    result = requireAdmin(data, addNitiReport); break;
      case 'updateNitiReport': result = requireAdmin(data, updateNitiReport); break;
      case 'deleteNitiReport': result = requireAdmin(data, deleteNitiReport); break;

      // Admin — File Upload
      case 'uploadFile':       result = requireAdmin(data, (d) => handleUpload(d)); break;

      default:
        result = { success: false, message: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { success: false, message: err.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Helper ────────────────────────────────────────────────────

function getSheet(name) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
}

function formatDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    const d = val.getDate();
    const m = val.getMonth() + 1;
    // getFullYear() คืนค่า CE (ค.ศ.) เสมอ
    // ✅ ตรวจก่อนบวก: ถ้า year < 2100 ถือเป็น CE → บวก 543 เป็น พ.ศ.
    //    ถ้า year >= 2100 แสดงว่า GAS แปลงผิดพลาดแล้ว หรือเป็น พ.ศ. อยู่แล้ว
    const rawYear = val.getFullYear();
    const y = rawYear < 2100 ? rawYear + 543 : rawYear;
    return (d < 10 ? '0'+d : d) + '/' + (m < 10 ? '0'+m : m) + '/' + y;
  }
  // string เช่น "01/03/2567" — คืนตรงๆ ไม่แตะ
  return String(val);
}

// ── File Upload to Google Drive ───────────────────────────────
function uploadFileToDrive(base64Data, fileName, mimeType, folderId) {
  try {
    var folder = folderId
      ? DriveApp.getFolderById(folderId)
      : DriveApp.getRootFolder();
    var bytes = Utilities.base64Decode(base64Data);
    var blob  = Utilities.newBlob(bytes, mimeType, fileName);
    var file  = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { fileId: file.getId(), fileUrl: 'https://drive.google.com/file/d/' + file.getId() + '/preview' };
  } catch(e) {
    return null;
  }
}

function handleUpload(p) {
  if (!p.base64Data || !p.fileName || !p.mimeType) return { success:false, message:'ข้อมูลไฟล์ไม่ครบ' };
  // Validate mime type — รับทั้ง image/* และ PDF
  var allowed = ['application/pdf','image/jpeg','image/png','image/gif','image/webp','image/jpg'];
  // ถ้า mimeType ขึ้นต้นด้วย image/ ก็รับทั้งหมด (image/heic ฯลฯ)
  if (allowed.indexOf(p.mimeType) < 0 && p.mimeType.indexOf('image/') !== 0) {
    return { success:false, message:'รองรับเฉพาะ PDF และรูปภาพเท่านั้น (got: '+p.mimeType+')' };
  }
  // ✅ BUG FIX: ส่ง folderId จาก parameter จริง (เดิมส่ง null ทำให้ใช้ root folder และ fail)
  var folderId = p.folderId || DRIVE_FOLDER_ID;
  var result = uploadFileToDrive(p.base64Data, p.fileName, p.mimeType, folderId);
  if (!result) return { success:false, message:'อัปโหลดไม่สำเร็จ — ตรวจสอบสิทธิ์ DriveApp ใน Google Apps Script' };
  return { success:true, fileId: result.fileId, fileUrl: result.fileUrl };
}

// ── คอลัมน์ที่ต้องได้รับการปกป้องจาก GAS auto-parse ──────────
// house_no: "9/4" → GAS แปลงเป็น Date  → ต้อง reconstruct เป็น "9/4"
// phone:    "0810012012" → GAS แปลงเป็น 810012012 (ตัด 0 นำ) → ต้องใช้ displayValue
// soi:      เก็บเป็น Number  → คืนค่าตัวเลข (0 ถ้าว่าง)
const HOUSE_TEXT_COLS = new Set(['house_no', 'phone', 'address', 'note',
  'house_id', 'user_id', 'fee_id', 'ann_id', 'report_id', 'owner_name',
  'username', 'full_name', 'category', 'title', 'content',
  // date columns — เก็บเป็น string พ.ศ. ไม่ให้ GAS re-parse เป็น Date แล้วบวก 543 ซ้ำ
  'date', 'created_date', 'h1_date', 'h2_date',
  // URL / file columns — ต้องเป็น string เสมอ
  'photo_urls', 'file_url', 'file_urls']);

function sheetToObjects(sheet) {
  const range    = sheet.getDataRange();
  const values   = range.getValues();          // raw (Date object ถ้า GAS parse)
  const displays = range.getDisplayValues();   // string ที่เห็นในช่อง → ใช้กับ phone

  if (values.length < 2) return [];
  const headers = values[0];

  return values.slice(1).map((row, ri) => {
    const obj = {};
    headers.forEach((h, ci) => {
      const col = String(h);
      const raw = row[ci];

      if (col === 'house_no') {
        // GAS แปลง "9/4" เป็น Date → reconstruct กลับเป็น "d/m"
        if (raw instanceof Date) {
          obj[col] = raw.getDate() + '/' + (raw.getMonth() + 1);
        } else {
          obj[col] = raw === '' || raw === null || raw === undefined ? '' : String(raw);
        }

      } else if (col === 'phone') {
        // ใช้ displayValue เพื่อรักษา leading zero เช่น "0810012012"
        // displayValue ของ phone ที่ถูก parse เป็นตัวเลขจะไม่มี 0 นำ
        // แต่ถ้า cell format เป็น Text จะได้ string ตรง
        if (typeof raw === 'number') {
          // GAS ตัด leading 0 ออก → ใส่กลับตามจำนวนหลักเบอร์ไทย
          // เบอร์มือถือ 10 หลัก: raw=810012012 (9) → "0810012012"
          // เบอร์บ้าน 9 หลัก:    raw=21234567  (8) → "021234567"
          const s = String(raw);
          if (s.length === 9)      obj[col] = '0' + s;   // มือถือ 0x-xxxx-xxxx
          else if (s.length === 8) obj[col] = '0' + s;   // บ้าน 0x-xxx-xxxx
          else if (s.length === 7) obj[col] = '0' + s;   // บ้าน 0x-xx-xxxx
          else                     obj[col] = s;          // รูปแบบอื่น ไม่แตะ
        } else if (raw instanceof Date) {
          // fallback: อ่าน displayValue จาก sheet (ซึ่งเก็บ header ที่ row 0)
          const dispRow = displays[ri + 1]; // ri+1 เพราะ displays มี header row
          obj[col] = dispRow ? String(dispRow[ci]) : String(raw);
        } else {
          obj[col] = raw === '' || raw === null || raw === undefined ? '' : String(raw);
        }

      } else if (col === 'soi') {
        // soi เป็น Number 0-22, ถ้าว่างให้เป็น 0
        const n = parseFloat(String(raw));
        obj[col] = isNaN(n) ? 0 : n;

      } else if (HOUSE_TEXT_COLS.has(col)) {
        obj[col] = raw === '' || raw === null || raw === undefined ? '' : String(raw);

      } else if (raw instanceof Date) {
        obj[col] = formatDate(raw);

      } else if (typeof raw === 'boolean') {
        obj[col] = String(raw).toUpperCase();

      } else {
        obj[col] = raw;
      }
    });
    return obj;
  });
}

function genId(prefix) {
  // ใช้ timestamp+random เพื่อป้องกัน ID ชนเมื่อลบแถว
  var ts  = Date.now();
  var rnd = Math.floor(Math.random() * 900) + 100; // 100-999
  return prefix + ts + rnd;
}

function requireAdmin(data, fn) {
  const user = verifyToken(data.token);
  if (!user) return { success: false, message: 'ไม่มีสิทธิ์เข้าถึง' };
  if (user.role !== 'admin') return { success: false, message: 'ต้องการสิทธิ์ Admin' };
  return fn(data, user);
}



// ── Auth ──────────────────────────────────────────────────────

function login(data) {
  const { username, password } = data;
  if (!username || !password) return { success: false, message: 'กรุณากรอกข้อมูลให้ครบ' };

  const sheet = getSheet(SHEETS.USERS);
  const users = sheetToObjects(sheet);
  const user = users.find(u => u.username == username && u.password == password && String(u.active).toUpperCase() === 'TRUE');

  if (!user) return { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };

  const token = makeToken(user);
  return {
    success: true,
    token,
    role: user.role,
    house_id: user.house_id,
    full_name: user.full_name,
    message: 'เข้าสู่ระบบสำเร็จ'
  };
}

// ── Resident ──────────────────────────────────────────────────

function getMyInfo(data) {
  const user = verifyToken(data.token);
  if (!user) return { success: false, message: 'กรุณาเข้าสู่ระบบ' };

  const houses = sheetToObjects(getSheet(SHEETS.HOUSES));
  const house = houses.find(h => h.house_id == user.house_id);
  if (!house) return { success: false, message: 'ไม่พบข้อมูลบ้าน' };

  return { success: true, data: house };
}

function getMyFees(data) {
  const user = verifyToken(data.token);
  if (!user) return { success: false, message: 'กรุณาเข้าสู่ระบบ' };

  const fees = sheetToObjects(getSheet(SHEETS.COMMON_FEE));
  const myFees = fees.filter(f => f.house_id == user.house_id);
  return { success: true, data: myFees };
}

// ── Public ────────────────────────────────────────────────────

function toNum(v) {
  if (v === null || v === undefined || v === '') return 0;
  var s = String(v).replace(/,/g, '').trim();
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function getFeeSummary(p) {
  const fees   = sheetToObjects(getSheet(SHEETS.COMMON_FEE));
  const houses = sheetToObjects(getSheet(SHEETS.HOUSES));

  // ดึงปีทั้งหมดจาก CommonFee Sheet จริง
  const years = [...new Set(fees.map(f => String(f.year)).filter(Boolean))].sort((a,b) => b - a);
  const year  = p.year || (years[0] || '');
  const filtered = fees.filter(f => String(f.year) === String(year));

  // บ้าน active เท่านั้น + map fee_per_half
  const activeHouses = houses.filter(h => String(h.status).toLowerCase() === 'active');
  const totalHouses  = activeHouses.length;

  // map house_id -> fee record ของปีนี้
  const feeMap = {};
  filtered.forEach(f => { feeMap[String(f.house_id)] = f; });

  // map house_id -> fee_per_half จาก Houses sheet
  const feePerHalfMap = {};
  activeHouses.forEach(h => {
    feePerHalfMap[String(h.house_id)] = toNum(h.fee_per_half);
  });

  // นับแยกครึ่งปี (จำนวนหลัง)
  var h1Paid=0, h1Unpaid=0, h2Paid=0, h2Unpaid=0;
  // นับบ้านชำระครบ/บางส่วน/ค้างชำระ
  var hFullPaid=0, hPartial=0, hUnpaid=0;
  // รวมเงิน: ยอดที่ต้องชำระ = fee_per_half*2, ยอดที่ชำระแล้ว = h1_paid+h2_paid
  var amtDue=0, amtPaid=0;

  activeHouses.forEach(function(h) {
    var hid = String(h.house_id);
    var fph = toNum(h.fee_per_half);
    var f   = feeMap[hid];

    // ยอดที่ต้องชำระ = fee_per_half * 2 (ทั้งปี)
    amtDue += fph * 2;

    if (!f) {
      h1Unpaid++; h2Unpaid++; hUnpaid++;
      return;
    }

    var h1s   = f.h1_status || 'unpaid';
    var h2s   = f.h2_status || 'unpaid';
    var paid1 = toNum(f.h1_paid);
    var paid2 = toNum(f.h2_paid);

    amtPaid += paid1 + paid2;

    if (h1s === 'paid') h1Paid++; else h1Unpaid++;
    if (h2s === 'paid') h2Paid++; else h2Unpaid++;

    if (h1s === 'paid' && h2s === 'paid') hFullPaid++;
    else if (h1s === 'unpaid' && h2s === 'unpaid') hUnpaid++;
    else hPartial++;
  });

  var paidPct      = amtDue > 0 ? Math.round(amtPaid / amtDue * 100) : 0;
  var housePaidPct = totalHouses > 0 ? Math.round(hFullPaid / totalHouses * 100) : 0;

  // แยกตามซอย — ยอดที่ต้องชำระใช้ fee_per_half*2 จาก Houses
  var soiMap = {};
  activeHouses.forEach(function(h) {
    var soi = h.soi || 'ไม่ระบุ';
    if (!soiMap[soi]) soiMap[soi] = {
      total:0, fullPaid:0, partial:0, unpaid:0,
      h1Paid:0, h1Unpaid:0, h2Paid:0, h2Unpaid:0,
      amtDue:0, amtPaid:0
    };
    var s   = soiMap[soi];
    var hid = String(h.house_id);
    var fph = toNum(h.fee_per_half);
    var f   = feeMap[hid];

    s.total++;
    s.amtDue += fph * 2;  // ยอดที่ต้องชำระของซอยนี้

    if (!f) {
      s.unpaid++; s.h1Unpaid++; s.h2Unpaid++;
      return;
    }

    var h1s   = f.h1_status || 'unpaid';
    var h2s   = f.h2_status || 'unpaid';
    var paid1 = toNum(f.h1_paid);
    var paid2 = toNum(f.h2_paid);

    s.amtPaid += paid1 + paid2;

    if (h1s === 'paid') s.h1Paid++; else s.h1Unpaid++;
    if (h2s === 'paid') s.h2Paid++; else s.h2Unpaid++;

    if (h1s === 'paid' && h2s === 'paid') s.fullPaid++;
    else if (h1s === 'unpaid' && h2s === 'unpaid') s.unpaid++;
    else s.partial++;
  });

  return { success: true, data: {
    years, year, totalHouses,
    h1Paid, h1Unpaid, h2Paid, h2Unpaid,
    hFullPaid, hPartial, hUnpaid,
    amtDue, amtPaid, paidPct, housePaidPct,
    bySoi: soiMap
  }};
}

function getAnnouncements(data) {
  const anns = sheetToObjects(getSheet(SHEETS.ANNOUNCEMENTS));
  const active = anns.filter(a => String(a.active) === 'TRUE').reverse();
  return { success: true, data: active };
}

function getNitiReports(data) {
  const reports = sheetToObjects(getSheet(SHEETS.NITI_REPORT));
  // กรองรายงานที่ถูก soft-delete ออก (active = 'FALSE')
  const active = reports.filter(r => String(r.active || 'TRUE').toUpperCase() !== 'FALSE');
  active.sort((a, b) => {
    if (b.year !== a.year) return b.year - a.year;
    return b.month - a.month;
  });
  return { success: true, data: active };
}

// ── Admin: Houses ─────────────────────────────────────────────

function getHouses(data, user) {
  const houses = sheetToObjects(getSheet(SHEETS.HOUSES));
  return { success: true, data: houses };
}

function addHouse(data, user) {
  const sheet    = getSheet(SHEETS.HOUSES);
  const id       = genId('H');
  const areaSqm  = parseFloat(data.area_sqm) || 0;
  const soiNum   = parseFloat(data.soi) >= 0 ? parseFloat(data.soi) : 0;
  const feeRate  = parseFloat(data.fee_rate)  || 0;
  const feePH    = feeRate > 0
    ? Math.round(feeRate * areaSqm * 6)
    : (parseFloat(data.fee_per_half) || 0);
  // phone: เก็บเป็น plain text เพื่อรักษา leading 0
  // หมายเหตุ: ใน Sheet ต้องตั้ง format ของ column phone เป็น "Plain text"
  const phoneVal = String(data.phone || '');
  const houseNoVal = String(data.house_no || '');

  sheet.appendRow([
    id, houseNoVal, data.owner_name, data.address,
    areaSqm, soiNum, data.house_type, phoneVal,
    feePH, 'active', data.note || '', data.account_status || 'ปกติ', feeRate
  ]);
  return { success: true, message: 'เพิ่มข้อมูลบ้านสำเร็จ', house_id: id };
}

function updateHouse(data, user) {
  const sheet    = getSheet(SHEETS.HOUSES);
  const rows     = sheet.getDataRange().getValues();
  const areaSqm  = parseFloat(data.area_sqm) || 0;
  const soiNum   = parseFloat(data.soi) >= 0 ? parseFloat(data.soi) : 0;
  const feeRate  = parseFloat(data.fee_rate)  || 0;
  const feePH    = feeRate > 0
    ? Math.round(feeRate * areaSqm * 6)
    : (parseFloat(data.fee_per_half) || 0);
  const phoneVal   = String(data.phone || '');
  const houseNoVal = String(data.house_no || '');

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] == data.house_id) {
      sheet.getRange(i + 1, 1, 1, 13).setValues([[
        data.house_id, houseNoVal, data.owner_name, data.address,
        areaSqm, soiNum, data.house_type, phoneVal,
        feePH, data.status || 'active', data.note || '', data.account_status || 'ปกติ', feeRate
      ]]);
      return { success: true, message: 'อัปเดตข้อมูลบ้านสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบข้อมูลบ้าน' };
}

function deleteHouse(data, user) {
  const sheet = getSheet(SHEETS.HOUSES);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] == data.house_id) {
      sheet.getRange(i + 1, 10).setValue('inactive');
      return { success: true, message: 'ลบข้อมูลบ้านสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบข้อมูลบ้าน' };
}

// ── Admin: Users ──────────────────────────────────────────────

function getUsers(data, user) {
  const users = sheetToObjects(getSheet(SHEETS.USERS));
  // Hide passwords
  users.forEach(u => { u.password = '***'; });
  return { success: true, data: users };
}

function addUser(data, user) {
  const sheet = getSheet(SHEETS.USERS);
  const id = genId('U');
  sheet.appendRow([
    id, data.username, data.password, data.role,
    data.house_id || '', data.full_name, 'TRUE'
  ]);
  return { success: true, message: 'เพิ่มผู้ใช้สำเร็จ', user_id: id };
}

function updateUser(data, user) {
  const sheet = getSheet(SHEETS.USERS);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] == data.user_id) {
      const newPass = data.password && data.password !== '***' ? data.password : rows[i][2];
      sheet.getRange(i + 1, 1, 1, 7).setValues([[
        data.user_id, data.username, newPass, data.role,
        data.house_id || '', data.full_name, data.active || 'TRUE'
      ]]);
      return { success: true, message: 'อัปเดตผู้ใช้สำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบผู้ใช้' };
}

// ── Admin: Common Fee ─────────────────────────────────────────

function getFees(data, user) {
  const fees = sheetToObjects(getSheet(SHEETS.COMMON_FEE));
  return { success: true, data: fees };
}

function addFee(data, user) {
  const sheet = getSheet(SHEETS.COMMON_FEE);
  // ตรวจซ้ำ: house_id + year ต้องไม่ซ้ำกัน
  const existing = sheetToObjects(sheet);
  const dup = existing.find(f =>
    String(f.house_id) === String(data.house_id) &&
    String(f.year)     === String(data.year)
  );
  if (dup) return {
    success: false,
    message: 'บ้านนี้มีข้อมูลค่าส่วนกลางปี ' + data.year + ' อยู่แล้ว (fee_id: ' + dup.fee_id + ')'
  };
  const id = genId('F');
  sheet.appendRow([
    id, data.house_id, data.year,
    data.h1_amount, data.h1_paid || 0, data.h1_date || '', data.h1_status || 'unpaid',
    data.h2_amount, data.h2_paid || 0, data.h2_date || '', data.h2_status || 'unpaid',
    data.note || ''
  ]);
  return { success: true, message: 'เพิ่มข้อมูลค่าส่วนกลางสำเร็จ', fee_id: id };
}

function updateFee(data, user) {
  const sheet = getSheet(SHEETS.COMMON_FEE);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] == data.fee_id) {
      sheet.getRange(i + 1, 1, 1, 12).setValues([[
        data.fee_id, data.house_id, data.year,
        data.h1_amount, data.h1_paid || 0, data.h1_date || '', data.h1_status || 'unpaid',
        data.h2_amount, data.h2_paid || 0, data.h2_date || '', data.h2_status || 'unpaid',
        data.note || ''
      ]]);
      return { success: true, message: 'อัปเดตค่าส่วนกลางสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบข้อมูล' };
}

// ── Admin: Announcements ──────────────────────────────────────

function addAnnouncement(data, user) {
  const sheet = getSheet(SHEETS.ANNOUNCEMENTS);
  const id = genId('A');
  sheet.appendRow([
    id, data.title, data.content, data.category || 'ทั่วไป',
    data.date || new Date().toLocaleDateString('th-TH'),
    user.username, 'TRUE', data.file_url || ''
  ]);
  return { success: true, message: 'เพิ่มประกาศสำเร็จ', ann_id: id };
}

function updateAnnouncement(data, user) {
  const sheet = getSheet(SHEETS.ANNOUNCEMENTS);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] == data.ann_id) {
      sheet.getRange(i + 1, 1, 1, 8).setValues([[
        data.ann_id, data.title, data.content, data.category || 'ทั่วไป',
        data.date, user.username, data.active || 'TRUE', data.file_url || rows[i][7] || ''
      ]]);
      return { success: true, message: 'อัปเดตประกาศสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบประกาศ' };
}

function deleteAnnouncement(data, user) {
  const sheet = getSheet(SHEETS.ANNOUNCEMENTS);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] == data.ann_id) {
      sheet.getRange(i + 1, 7).setValue('FALSE');
      return { success: true, message: 'ลบประกาศสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบประกาศ' };
}

// ── Admin: Niti Report ────────────────────────────────────────

function addNitiReport(data, user) {
  const sheet = getSheet(SHEETS.NITI_REPORT);
  const id = genId('R');
  sheet.appendRow([
    id, data.month, data.year, data.title, data.content,
    data.income || 0, data.expense || 0,
    user.username, new Date().toLocaleDateString('th-TH'),
    data.photo_urls || ''
  ]);
  return { success: true, message: 'เพิ่มรายงานสำเร็จ', report_id: id };
}

function updateNitiReport(data, user) {
  const sheet   = getSheet(SHEETS.NITI_REPORT);
  const rows    = sheet.getDataRange().getValues();
  const headers = rows[0];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] == data.report_id) {
      const existingPhotos = String(rows[i][9] || '');

      // ── photo_urls logic ──────────────────────────────────────
      // client ส่ง sentinel "__KEEP__" = ให้คงรูปเดิมทั้งหมด
      // client ส่ง string URL = อัปเดตใหม่ (รวม existing ที่ client เลือกไว้แล้ว)
      // client ส่ง "" หรือไม่ส่ง = ล้างรูปทั้งหมด (intentional clear)
      let newPhotos;
      if (data.photo_urls === '__KEEP__' || data.photo_urls === undefined || data.photo_urls === null) {
        newPhotos = existingPhotos;   // คงเดิม
      } else {
        newPhotos = String(data.photo_urls);  // ใช้ค่าจาก client (อาจว่าง = ล้าง)
      }

      // active column (col index 10 ถ้ามี)
      const colCount = headers.length >= 11 ? 11 : 10;
      const vals = [
        String(data.report_id), String(data.month), String(data.year),
        String(data.title || ''), String(data.content || ''),
        parseFloat(data.income)  || 0,
        parseFloat(data.expense) || 0,
        user.username,
        new Date().toLocaleDateString('th-TH'),
        newPhotos
      ];
      if (colCount === 11) vals.push(rows[i][10] || 'TRUE');
      sheet.getRange(i + 1, 1, 1, colCount).setValues([vals]);
      return { success: true, message: 'อัปเดตรายงานสำเร็จ', photo_urls: newPhotos };
    }
  }
  return { success: false, message: 'ไม่พบรายงาน' };
}

function deleteNitiReport(data, user) {
  // soft delete — เปลี่ยน active เป็น FALSE (ไม่ลบแถวจริง ป้องกัน genId ชน)
  const sheet = getSheet(SHEETS.NITI_REPORT);
  const rows  = sheet.getDataRange().getValues();
  const headers = rows[0];
  const activeCol = headers.indexOf('active');  // ถ้ามี column active
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] == data.report_id) {
      if (activeCol >= 0) {
        sheet.getRange(i + 1, activeCol + 1).setValue('FALSE');
      } else {
        // ถ้าไม่มี column active ให้ลบแถวตามเดิม (backward compat)
        sheet.deleteRow(i + 1);
      }
      return { success: true, message: 'ลบรายงานสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบรายงาน' };
}