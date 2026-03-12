// ============================================================
// Code.gs — Village Management System Backend (GAS)
// ============================================================
// วิธีใช้: แก้ไข SHEET_ID ให้ตรงกับ Google Sheet ของคุณ
// แล้ว Deploy เป็น Web App (Execute as: Me, Anyone can access)
// ============================================================

const SHEET_ID = '1MDX7JWY33m1lqtHbGGVQXbz3fPP_Dxllm5U_B5ixOhk';
const DRIVE_FOLDER_ID = '1l4mAepbhZawYphDfE1kybY7bHNHS6IZ-';

// Sheet names
const SHEETS = {
  USERS: 'Users',
  HOUSES: 'Houses',
  COMMON_FEE: 'CommonFee',
  ANNOUNCEMENTS: 'Announcements',
  NITI_REPORT: 'NitiReport',
  CARS: 'Cars',
  CAR_REQUESTS: 'CarRequests'
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
  // ป้องกัน e เป็น undefined เมื่อรันโดยตรงจาก Editor
  if (!e) e = {};
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

      // Cars
      case 'getCars':           result = getCars(data); break;
      case 'addCar':            result = requireAdmin(data, addCar); break;
      case 'updateCar':         result = requireAdmin(data, updateCar); break;
      case 'deleteCar':         result = requireAdmin(data, deleteCar); break;
      // Car Requests (resident)
      case 'submitCarRequest':  result = requireResident(data, submitCarRequest); break;
      case 'getMyCarRequests':  result = requireResident(data, getMyCarRequests); break;
      // Car Requests (admin)
      case 'getCarRequests':    result = requireAdmin(data, getCarRequests); break;
      case 'approveCarRequest': result = requireAdmin(data, approveCarRequest); break;
      case 'rejectCarRequest':  result = requireAdmin(data, rejectCarRequest); break;
      case 'getPendingCount':   result = requireAdmin(data, getPendingCount); break;

      // Admin — Batch Import Fees
      case 'batchImportFees':  result = requireAdmin(data, batchImportFees); break;

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
  var file;
  try {
    var folder = folderId
      ? DriveApp.getFolderById(folderId)
      : DriveApp.getRootFolder();
    var bytes = Utilities.base64Decode(base64Data);
    var blob  = Utilities.newBlob(bytes, mimeType, fileName);
    file = folder.createFile(blob);
  } catch(e) {
    return { ok: false, error: 'createFile failed: ' + e.toString() };
  }

  // setSharing แยก try/catch — Shared Drive อาจ fail แต่ไฟล์ยังใช้ได้
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch(e) {
    Logger.log('setSharing warning (Shared Drive?): ' + e.toString());
    // ไม่ return error — ไฟล์ถูก create แล้ว ใช้ URL ได้
  }

  var fileId  = file.getId();
  var fileUrl = 'https://drive.google.com/file/d/' + fileId + '/view?usp=sharing';
  Logger.log('uploadFileToDrive OK: ' + fileUrl);
  return { ok: true, fileId: fileId, fileUrl: fileUrl };
}

function buildFileName(originalName, prefix) {
  // สร้างชื่อไฟล์ format: prefix_YYYYMMDD_HHmmss_millisec.ext
  // เช่น annou_20250310_143022_456.jpg หรือ niti_20250310_143022_789.png
  var now  = new Date();
  var pad  = function(n,len){ return String(n).padStart(len||2,'0'); };
  var date = now.getFullYear().toString() +
             pad(now.getMonth()+1) +
             pad(now.getDate());
  var time = pad(now.getHours()) +
             pad(now.getMinutes()) +
             pad(now.getSeconds());
  var ms   = pad(now.getMilliseconds(), 3);
  // แยกนามสกุลจากชื่อไฟล์เดิม
  var ext  = '';
  var dot  = originalName.lastIndexOf('.');
  if (dot >= 0) ext = originalName.substring(dot).toLowerCase(); // .jpg .png .pdf
  return (prefix||'file') + '_' + date + '_' + time + '_' + ms + ext;
}

function handleUpload(p) {
  if (!p.base64Data || !p.fileName || !p.mimeType) {
    return { success:false, message:'ข้อมูลไฟล์ไม่ครบ (base64Data/fileName/mimeType)' };
  }
  // Validate mime type
  var allowed = ['application/pdf','image/jpeg','image/png','image/gif','image/webp','image/jpg'];
  if (allowed.indexOf(p.mimeType) < 0 && p.mimeType.indexOf('image/') !== 0) {
    return { success:false, message:'รองรับเฉพาะ PDF และรูปภาพเท่านั้น (got: '+p.mimeType+')' };
  }
  // ── Rename ตาม prefix_YYYYMMDD_HHmmss_ms.ext ──────────────────
  // prefix มาจาก client: 'annou' หรือ 'niti' หรือ 'file'
  var prefix   = p.filePrefix || 'file';
  var newName  = buildFileName(p.fileName, prefix);
  var folderId = p.folderId || DRIVE_FOLDER_ID;
  Logger.log('uploadFile: ' + p.fileName + ' → ' + newName + ' folder=' + folderId);
  var result = uploadFileToDrive(p.base64Data, newName, p.mimeType, folderId);
  if (!result.ok) {
    Logger.log('uploadFile ERROR: ' + result.error);
    return { success:false, message: result.error };
  }
  Logger.log('uploadFile OK: ' + result.fileUrl);
  return { success:true, fileId: result.fileId, fileUrl: result.fileUrl, fileName: newName };
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
        // date columns พิเศษ: ถ้า GAS parse string วันที่เป็น Date object → แปลงกลับเป็น พ.ศ. string
        if (raw instanceof Date && (col === 'date' || col === 'created_date' || col === 'h1_date' || col === 'h2_date')) {
          obj[col] = formatDate(raw);
        } else {
          obj[col] = raw === '' || raw === null || raw === undefined ? '' : String(raw);
        }

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
function requireResident(data, fn) {
  const user = verifyToken(data.token);
  if (!user) return { success: false, message: 'กรุณาเข้าสู่ระบบ' };
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

// ── Admin: Batch Import Fees ─────────────────────────────────────
/**
 * batchImportFees — รับ array ของ fee rows แล้ว insert ทีละแถว
 * rows: [{house_id, year, h1_amount, h1_paid, h1_date, h1_status,
 *          h2_amount, h2_paid, h2_date, h2_status, note}]
 * - ถ้า house_id ไม่มีใน Houses → skip (รายงานใน errors[])
 * - ถ้า house_id+year ซ้ำ → skip (รายงานใน errors[])
 */
function batchImportFees(data, user) {
  const rows = data.rows;
  if (!Array.isArray(rows) || rows.length === 0)
    return { success: false, message: 'ไม่มีข้อมูลที่จะนำเข้า' };

  const feeSheet   = getSheet(SHEETS.COMMON_FEE);
  const houseSheet = getSheet(SHEETS.HOUSES);
  const existing   = sheetToObjects(feeSheet);
  const houses     = sheetToObjects(houseSheet);

  // สร้าง map house_no → house_id เพื่อรองรับการระบุบ้านด้วยเลขบ้านแทน house_id
  const houseMap = {};
  houses.forEach(h => {
    houseMap[String(h.house_id).trim()]   = h.house_id;
    houseMap[String(h.house_no).trim()]   = h.house_id;
  });

  let inserted = 0;
  const errors = [];

  rows.forEach((row, idx) => {
    const lineNo = idx + 2; // row 1 = header
    const rawHouse = String(row.house_id || '').trim();
    const year     = String(row.year || '').trim();

    if (!rawHouse || !year) {
      errors.push(`แถว ${lineNo}: ขาด house_id หรือ year`); return;
    }

    const houseId = houseMap[rawHouse];
    if (!houseId) {
      errors.push(`แถว ${lineNo}: ไม่พบบ้าน "${rawHouse}"`); return;
    }

    const dup = existing.find(f =>
      String(f.house_id) === String(houseId) && String(f.year) === year
    );
    if (dup) {
      errors.push(`แถว ${lineNo}: บ้าน ${rawHouse} ปี ${year} มีข้อมูลอยู่แล้ว`); return;
    }

    const id = genId('F');
    feeSheet.appendRow([
      id, houseId, year,
      row.h1_amount || 0, row.h1_paid || 0, row.h1_date || '', row.h1_status || 'unpaid',
      row.h2_amount || 0, row.h2_paid || 0, row.h2_date || '', row.h2_status || 'unpaid',
      row.note || ''
    ]);
    // เพิ่มเข้า existing เพื่อ prevent dup ภายใน batch เดียวกัน
    existing.push({ house_id: houseId, year });
    inserted++;
  });

  return {
    success: true,
    inserted,
    skipped: rows.length - inserted,
    errors,
    message: `นำเข้าสำเร็จ ${inserted} รายการ${errors.length ? ` (ข้าม ${errors.length} รายการ)` : ''}`
  };
}

// ── Admin: Announcements ──────────────────────────────────────

function addAnnouncement(data, user) {
  const sheet = getSheet(SHEETS.ANNOUNCEMENTS);
  const id = genId('A');
  sheet.appendRow([
    id, data.title, data.content, data.category || 'ทั่วไป',
    cleanDateString(data.date) || new Date().toLocaleDateString('th-TH'),
    user.username, 'TRUE', data.file_url || ''
  ]);
  return { success: true, message: 'เพิ่มประกาศสำเร็จ', ann_id: id };
}

function cleanDateString(d) {
  // รับ string ทุกรูปแบบ → คืน string "dd/mm/yyyy" (พ.ศ.) เพื่อ save ลง Sheet
  if (!d) return '';
  if (d instanceof Date) return formatDate(d);
  var s = String(d);
  // dd/mm/yyyy
  var slash = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slash) {
    var yr = parseInt(slash[3]);
    return slash[1].padStart(2,'0') + '/' + slash[2].padStart(2,'0') + '/' + (yr < 2100 ? yr + 543 : yr);
  }
  // JS Date.toString "Mon Mar 01 3655 ..."
  var jsm = s.match(/([A-Za-z]{3})\s+(\d{1,2})\s+(\d{4})/);
  if (jsm) {
    var months = {Jan:1,Feb:2,Mar:3,Apr:4,May:5,Jun:6,Jul:7,Aug:8,Sep:9,Oct:10,Nov:11,Dec:12};
    var mo = months[jsm[1]] || 1;
    var dy = parseInt(jsm[2]);
    var y  = parseInt(jsm[3]);
    var be = y > 2500 ? y : y + 543;
    return dy.toString().padStart(2,'0') + '/' + mo.toString().padStart(2,'0') + '/' + be;
  }
  return s;
}

function updateAnnouncement(data, user) {
  const sheet = getSheet(SHEETS.ANNOUNCEMENTS);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] == data.ann_id) {
      // cleanDateString: ป้องกัน date toString จาก GAS ถูก save ลง Sheet ผิดรูปแบบ
      const safeDate = cleanDateString(data.date);
      sheet.getRange(i + 1, 1, 1, 8).setValues([[
        data.ann_id, data.title, data.content, data.category || 'ทั่วไป',
        safeDate, user.username, data.active || 'TRUE', data.file_url || rows[i][7] || ''
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

// ── Cars ─────────────────────────────────────────────────────────


// ══════════════════════════════════════════════════════════════
// CAR REQUESTS — ลูกบ้าน submit, Admin approve/reject
// Sheet CarRequests headers (17 cols):
//   req_id | house_id | request_type | car_id |
//   car_type | plate_no | car_brand | car_model | car_color | car_park | car_fee |
//   photo_urls | status | note | submitted_at | reviewed_at | reviewed_by
// ══════════════════════════════════════════════════════════════

var CAR_REQ_HEADERS = [
  'req_id','house_id','request_type','car_id',
  'car_type','plate_no','car_brand','car_model','car_color','car_park','car_fee',
  'photo_urls','status','note','submitted_at','reviewed_at','reviewed_by'
];

function _ensureCarReqSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEETS.CAR_REQUESTS);
  if (!sh) {
    sh = ss.insertSheet(SHEETS.CAR_REQUESTS);
    sh.appendRow(CAR_REQ_HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

// ── ลูกบ้าน ยื่น request ──────────────────────────────────────
function submitCarRequest(data, user) {
  var reqType = data.request_type; // 'add' | 'edit'

  // ตรวจสิทธิ์บ้าน — ต้องเป็นบ้านของตัวเอง (admin ข้ามได้)
  if (user.role !== 'admin' && String(user.house_id) !== String(data.house_id)) {
    return { success: false, message: 'ไม่มีสิทธิ์ยื่นคำขอสำหรับบ้านนี้' };
  }

  // upload photos base64[] → Drive URLs
  var photoUrls = [];
  var photos = data.photos || [];
  if (!Array.isArray(photos)) photos = [];

  if (reqType === 'add' && photos.length === 0) {
    return { success: false, message: 'กรณีขอเพิ่มรถใหม่ต้องแนบรูปอย่างน้อย 1 รูป' };
  }

  // สร้าง timestamp สำหรับชื่อไฟล์ — yyyyMMdd_HHmmss
  var now       = new Date();
  var ymd       = now.getFullYear()
                + ('0'+(now.getMonth()+1)).slice(-2)
                + ('0'+now.getDate()).slice(-2);
  var hms       = ('0'+now.getHours()).slice(-2)
                + ('0'+now.getMinutes()).slice(-2)
                + ('0'+now.getSeconds()).slice(-2);
  // ทำความสะอาดทะเบียน: เอาเฉพาะ a-z0-9ก-ฮ ออก space และอักขระพิเศษ
  var rawPlate  = String(data.plate_no || 'car').replace(/[^a-zA-Z0-9ก-ฮ]/g,'');
  if (!rawPlate) rawPlate = 'car';

  for (var i = 0; i < Math.min(photos.length, 5); i++) {
    var p = photos[i];
    if (!p || !p.base64) continue;
    var ext   = (p.ext || 'jpg').replace(/[^a-z0-9]/gi,'') || 'jpg';
    // ชื่อไฟล์: car_{ทะเบียน}_{yyyyMMdd_HHmmss}_{ลำดับ}.ext
    var fname = 'car_' + rawPlate + '_' + ymd + '_' + hms + '_' + (i+1) + '.' + ext;
    var res   = uploadFileToDrive(p.base64, fname, p.mimeType || 'image/jpeg', DRIVE_FOLDER_ID);
    if (res && res.fileUrl) photoUrls.push(res.fileUrl);
  }

  if (reqType === 'add' && photoUrls.length === 0 && photos.length > 0) {
    return { success: false, message: 'อัปโหลดรูปไม่สำเร็จ กรุณาลองใหม่' };
  }

  var sh = _ensureCarReqSheet();
  var req_id = 'REQ' + Date.now().toString().slice(-8) + Math.floor(Math.random()*90+10);
  sh.appendRow([
    req_id,
    data.house_id || '',
    reqType,
    data.car_id   || '',
    data.car_type  || 'car',
    data.plate_no  || '',
    data.car_brand || '',
    data.car_model || '',
    data.car_color || '',
    data.car_park  || '',
    parseFloat(data.car_fee) || 0,
    photoUrls.join(','),
    'pending',
    data.note || '',
    new Date().toLocaleString('th-TH'),
    '', ''
  ]);
  return { success: true, message: 'ส่งคำขอสำเร็จ รอการอนุมัติจากผู้ดูแล', req_id: req_id };
}

// ── ลูกบ้าน ดู requests ของตัวเอง ────────────────────────────
function getMyCarRequests(data, user) {
  var sh = _ensureCarReqSheet();
  var rows = sh.getDataRange().getValues();
  if (rows.length < 2) return { success: true, data: [] };
  var headers = rows[0].map(String);
  var list = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    var obj = {};
    headers.forEach(function(h,j){ obj[h] = rows[i][j] === null || rows[i][j] === undefined ? '' : String(rows[i][j]); });
    if (String(obj.house_id) !== String(user.house_id)) continue;
    list.push(obj);
  }
  list.sort(function(a,b){ return (b.submitted_at||'').localeCompare(a.submitted_at||''); });
  return { success: true, data: list };
}

// ── Admin ดึง requests ทั้งหมด ──────────────────────────────
function getCarRequests(data, user) {
  var sh = _ensureCarReqSheet();
  var rows = sh.getDataRange().getValues();
  if (rows.length < 2) return { success: true, data: [] };
  var headers = rows[0].map(String);
  var houses = sheetToObjects(getSheet(SHEETS.HOUSES));
  var hMap = {};
  houses.forEach(function(h){ hMap[String(h.house_id)] = h; });
  var list = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    var obj = {};
    headers.forEach(function(h,j){ obj[h] = rows[i][j] === null || rows[i][j] === undefined ? '' : String(rows[i][j]); });
    if (data.status && obj.status !== data.status) continue;
    var hh = hMap[obj.house_id] || {};
    obj.house_no   = hh.house_no   || obj.house_id;
    obj.owner_name = hh.owner_name || '';
    obj.soi        = hh.soi        || '';
    list.push(obj);
  }
  list.sort(function(a,b){
    if (a.status==='pending' && b.status!=='pending') return -1;
    if (a.status!=='pending' && b.status==='pending') return 1;
    return (b.submitted_at||'').localeCompare(a.submitted_at||'');
  });
  return { success: true, data: list };
}

// ── Admin นับ pending ─────────────────────────────────────────
function getPendingCount(data, user) {
  var sh = _ensureCarReqSheet();
  var rows = sh.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][12]) === 'pending') count++;
  }
  return { success: true, count: count };
}

// ── Admin อนุมัติ ──────────────────────────────────────────────
function approveCarRequest(data, user) {
  var sh = _ensureCarReqSheet();
  var rows = sh.getDataRange().getValues();
  var headers = rows[0].map(String);
  var idxReqId = headers.indexOf('req_id');

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][idxReqId]) !== String(data.req_id)) continue;
    if (String(rows[i][12]) !== 'pending') {
      return { success: false, message: 'คำขอนี้ถูกดำเนินการแล้ว' };
    }
    var reqType = String(rows[i][2]);
    // ใช้ข้อมูลที่ Admin แก้ไขได้ก่อน approve
    var carData = {
      house_id:  data.house_id  || String(rows[i][1]),
      car_type:  data.car_type  || String(rows[i][4]),
      plate_no:  data.plate_no  || String(rows[i][5]),
      car_brand: data.car_brand || String(rows[i][6]),
      car_model: data.car_model || String(rows[i][7]),
      car_color: data.car_color || String(rows[i][8]),
      car_park:  data.car_park  || String(rows[i][9]),
      car_fee:   data.car_fee   !== undefined ? data.car_fee : String(rows[i][10]),
      car_id:    data.car_id    || String(rows[i][3])
    };
    var result;
    if (reqType === 'add') {
      result = addCar(carData, user);
    } else if (reqType === 'edit' && carData.car_id) {
      result = updateCar(carData, user);
    } else {
      return { success: false, message: 'ประเภทคำขอไม่ถูกต้อง' };
    }
    if (!result.success) return result;
    // อัปเดต status
    sh.getRange(i+1, 13).setValue('approved');
    sh.getRange(i+1, 16).setValue(new Date().toLocaleString('th-TH'));
    sh.getRange(i+1, 17).setValue(user.username);
    return { success: true, message: 'อนุมัติสำเร็จ' };
  }
  return { success: false, message: 'ไม่พบคำขอ' };
}

// ── Admin ปฏิเสธ ───────────────────────────────────────────────
function rejectCarRequest(data, user) {
  var sh = _ensureCarReqSheet();
  var rows = sh.getDataRange().getValues();
  var headers = rows[0].map(String);
  var idxReqId = headers.indexOf('req_id');
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][idxReqId]) !== String(data.req_id)) continue;
    sh.getRange(i+1, 13).setValue('rejected');
    sh.getRange(i+1, 14).setValue(data.reason || '');
    sh.getRange(i+1, 16).setValue(new Date().toLocaleString('th-TH'));
    sh.getRange(i+1, 17).setValue(user.username);
    return { success: true, message: 'ปฏิเสธคำขอแล้ว' };
  }
  return { success: false, message: 'ไม่พบคำขอ' };
}

function getCars(data) {
  const sheet = getSheet(SHEETS.CARS);
  const rows  = sheet.getDataRange().getValues();
  if (rows.length < 2) return { success: true, data: [] };
  const headers = rows[0].map(String);
  const cars = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row[0] && !row[1]) continue; // skip empty
    const obj = {};
    headers.forEach((h, j) => { obj[h] = row[j] === undefined || row[j] === null ? '' : String(row[j]); });
    // filter by house_id ถ้าส่งมา (สำหรับ dashboard ลูกบ้าน)
    if (data.house_id && obj.house_id !== String(data.house_id)) continue;
    cars.push(obj);
  }
  return { success: true, data: cars };
}

function genCarId() {
  return 'C' + Date.now().toString().slice(-8) + Math.floor(Math.random()*900+100);
}

function addCar(data, user) {
  const sheet  = getSheet(SHEETS.CARS);
  const car_id = genCarId();
  sheet.appendRow([
    data.house_id || '', car_id,
    data.car_type  || 'car',
    data.plate_no  || '',
    data.car_brand || '',
    data.car_model || '',
    data.car_color || '',
    data.car_park  || '',
    parseFloat(data.car_fee) || 0
  ]);
  return { success: true, message: 'เพิ่มรถสำเร็จ', car_id };
}

function updateCar(data, user) {
  const sheet = getSheet(SHEETS.CARS);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) === String(data.car_id)) {
      sheet.getRange(i + 1, 1, 1, 9).setValues([[
        data.house_id || rows[i][0],
        data.car_id,
        data.car_type  || rows[i][2],
        data.plate_no  || rows[i][3],
        data.car_brand || rows[i][4],
        data.car_model || rows[i][5],
        data.car_color || rows[i][6],
        data.car_park  || rows[i][7],
        parseFloat(data.car_fee) || 0
      ]]);
      return { success: true, message: 'อัปเดตรถสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบข้อมูลรถ' };
}

function deleteCar(data, user) {
  const sheet = getSheet(SHEETS.CARS);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) === String(data.car_id)) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'ลบรถสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบข้อมูลรถ' };
}

// ══════════════════════════════════════════════════════════════
// TEST FUNCTIONS — รันจาก Apps Script Editor เพื่อ Grant permission
// ══════════════════════════════════════════════════════════════

/** ✅ ทดสอบ upload แบบละเอียด — หา root cause */
function testUpload() {
  var tiny = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';

  // ── Test 1: DriveApp.getRootFolder() ──────────────────────
  try {
    var root = DriveApp.getRootFolder();
    Logger.log('✅ Test1 Root folder: ' + root.getName() + ' (id=' + root.getId() + ')');

    // ลอง create file ที่ root
    var bytes = Utilities.base64Decode(tiny);
    var blob  = Utilities.newBlob(bytes, 'image/png', 'vms_test_root.png');
    var f = root.createFile(blob);
    Logger.log('✅ Test1 Create file at root: OK — ' + f.getId());
    f.setTrashed(true); // ลบทิ้งหลังทดสอบ
  } catch(e) {
    Logger.log('❌ Test1 Root upload: ' + e.toString());
  }

  // ── Test 2: getFolderById ─────────────────────────────────
  try {
    var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    Logger.log('✅ Test2 Folder found: ' + folder.getName());

    // ตรวจสิทธิ์เจ้าของ folder
    var access = folder.getSharingAccess();
    var perm   = folder.getSharingPermission();
    Logger.log('  Sharing: access=' + access + ' perm=' + perm);

    // ลอง create file ใน folder
    var bytes2 = Utilities.base64Decode(tiny);
    var blob2  = Utilities.newBlob(bytes2, 'image/png', 'vms_test_folder.png');
    var f2 = folder.createFile(blob2);
    Logger.log('✅ Test2 Create file in folder: OK — ' + f2.getId());
    f2.setTrashed(true);
  } catch(e) {
    Logger.log('❌ Test2 Folder upload: ' + e.toString());
  }

  // ── Test 3: ดู email account ที่ run script ───────────────
  try {
    Logger.log('ℹ️  Effective user: ' + Session.getEffectiveUser().getEmail());
  } catch(e) {
    Logger.log('ℹ️  Session info N/A (ANYONE_ANONYMOUS deploy): ' + e.toString());
  }
  Logger.log('✅ DriveApp ทำงานได้ปกติ — ปัญหาอยู่ที่ Web App deployment ต้อง re-deploy ใหม่');
}

/** ทดสอบ step by step เพื่อหา root cause */
function testHandleUpload() {
  var tiny = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';

  // Step A: ทดสอบ uploadFileToDrive โดยตรง (เหมือน testUpload)
  Logger.log('--- Step A: uploadFileToDrive direct ---');
  try {
    var r1 = uploadFileToDrive(tiny, 'stepA.png', 'image/png', DRIVE_FOLDER_ID);
    Logger.log('Step A: ' + JSON.stringify(r1));
    if (r1.ok) DriveApp.getFileById(r1.fileId).setTrashed(true);
  } catch(e) { Logger.log('Step A exception: ' + e); }

  // Step B: ทดสอบ handleUpload โดยตรง
  Logger.log('--- Step B: handleUpload direct ---');
  try {
    var p = { base64Data:tiny, fileName:'stepB.png', mimeType:'image/png', folderId:DRIVE_FOLDER_ID };
    var r2 = handleUpload(p);
    Logger.log('Step B: ' + JSON.stringify(r2));
  } catch(e) { Logger.log('Step B exception: ' + e); }

  // Step C: ทดสอบผ่าน handleRequest (simulate Web App call จริง)
  Logger.log('--- Step C: via handleRequest ---');
  try {
    // login ก่อนเพื่อได้ token จริง
    var loginResult = login({ username:'admin', password:'admin' });
    Logger.log('Login: ' + JSON.stringify(loginResult));
    if (!loginResult.success) { Logger.log('Login failed — skip Step C'); return; }
    var token = loginResult.token;
    var fakeEvent = {
      parameter: {},
      postData: {
        contents: JSON.stringify({
          action: 'uploadFile',
          base64Data: tiny,
          fileName: 'stepC.png',
          mimeType: 'image/png',
          folderId: DRIVE_FOLDER_ID,
          token: token
        })
      }
    };
    var r3 = handleRequest(fakeEvent);
    Logger.log('Step C: ' + r3.getContent());
  } catch(e) { Logger.log('Step C exception: ' + e); }
}

/** ตรวจ account ที่ Script รันอยู่ */
function testWhoAmI() {
  try {
    Logger.log('EffectiveUser: ' + Session.getEffectiveUser().getEmail());
  } catch(e) { Logger.log('EffectiveUser error: ' + e); }
  try {
    Logger.log('ActiveUser: ' + Session.getActiveUser().getEmail());
  } catch(e) { Logger.log('ActiveUser error: ' + e); }
  try {
    var root = DriveApp.getRootFolder();
    Logger.log('Drive root owner: ' + root.getName() + ' id=' + root.getId());
    // list files ใน root เพื่อดูว่าเป็น Drive ของใคร
    var files = root.getFiles();
    var count = 0;
    while(files.hasNext() && count < 3) {
      Logger.log('  file: ' + files.next().getName());
      count++;
    }
  } catch(e) { Logger.log('Drive error: ' + e); }
}

/** ทดสอบว่า DriveApp ทำงานใน context นี้ได้ไหม */
function testDriveContext() {
  var tiny = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
  
  // ทดสอบ 1: DriveApp.getRootFolder
  try {
    var root = DriveApp.getRootFolder();
    Logger.log('✅ getRootFolder: ' + root.getName());
  } catch(e) { Logger.log('❌ getRootFolder: ' + e); return; }

  // ทดสอบ 2: getFolderById
  try {
    var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    Logger.log('✅ getFolderById: ' + folder.getName());
  } catch(e) { Logger.log('❌ getFolderById: ' + e); return; }

  // ทดสอบ 3: createFile ใน folder
  try {
    var bytes = Utilities.base64Decode(tiny);
    var blob  = Utilities.newBlob(bytes, 'image/png', 'ctx_test.png');
    var f = DriveApp.getFolderById(DRIVE_FOLDER_ID).createFile(blob);
    Logger.log('✅ createFile: ' + f.getId());
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    Logger.log('✅ setSharing OK');
    f.setTrashed(true);
    Logger.log('✅ ALL DriveApp OK — ปัญหาอยู่ที่ Web App context ไม่มี OAuth');
  } catch(e) { Logger.log('❌ createFile: ' + e); }
}

/** ✅ รันตัวนี้ก่อนเพื่อ Grant DriveApp + Spreadsheet permission */
function testPermissions() {
  try {
    // ทดสอบ Spreadsheet
    var ss = SpreadsheetApp.openById(SHEET_ID);
    Logger.log('✅ Spreadsheet: ' + ss.getName());

    // ทดสอบ Drive folder
    var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    Logger.log('✅ Drive folder: ' + folder.getName());

    Logger.log('✅ Permission OK — พร้อม Deploy แล้ว');
  } catch(e) {
    Logger.log('❌ Error: ' + e.toString());
  }
}

/** ทดสอบ login action */
function testDoGet() {
  var fakeEvent = {
    parameter: { action: 'login', username: 'admin', password: 'admin' },
    postData: null
  };
  var result = handleRequest(fakeEvent);
  Logger.log(result.getContent());
}

/** ทดสอบ getAnnouncements */
function testGetAnn() {
  var fakeEvent = {
    parameter: { action: 'getAnnouncements' },
    postData: null
  };
  var result = handleRequest(fakeEvent);
  Logger.log(result.getContent());
}
