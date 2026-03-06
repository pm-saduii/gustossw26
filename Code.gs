// ============================================================
// Code.gs — Village Management System Backend (GAS)
// ============================================================
// วิธีใช้: แก้ไข SHEET_ID ให้ตรงกับ Google Sheet ของคุณ
// แล้ว Deploy เป็น Web App (Execute as: Me, Anyone can access)
// ============================================================

const SHEET_ID = '1MDX7JWY33m1lqtHbGGVQXbz3fPP_Dxllm5U_B5ixOhk';

// Sheet names
const SHEETS = {
  USERS: 'Users',
  HOUSES: 'Houses',
  COMMON_FEE: 'CommonFee',
  ANNOUNCEMENTS: 'Announcements',
  NITI_REPORT: 'NitiReport'
};

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
  const params = e.parameter || {};
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
    const y = val.getFullYear() + 543; // Convert AD to BE (พ.ศ.)
    return (d < 10 ? '0'+d : d) + '/' + (m < 10 ? '0'+m : m) + '/' + y;
  }
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
  // Validate mime type
  var allowed = ['application/pdf','image/jpeg','image/png','image/gif','image/webp'];
  if (allowed.indexOf(p.mimeType) < 0) return { success:false, message:'รองรับเฉพาะ PDF และรูปภาพเท่านั้น' };
  var result = uploadFileToDrive(p.base64Data, p.fileName, p.mimeType, null);
  if (!result) return { success:false, message:'อัปโหลดไม่สำเร็จ' };
  return { success:true, fileId: result.fileId, fileUrl: result.fileUrl };
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      const val = row[i];
      if (val instanceof Date) {
        obj[h] = formatDate(val);
      } else if (typeof val === 'boolean') {
        obj[h] = String(val).toUpperCase();
      } else {
        obj[h] = val;
      }
    });
    return obj;
  });
}

function genId(prefix, sheet) {
  const rows = sheet.getDataRange().getValues();
  return prefix + String(rows.length).padStart(3, '0');
}

function requireAdmin(data, fn) {
  const user = verifyToken(data.token);
  if (!user) return { success: false, message: 'ไม่มีสิทธิ์เข้าถึง' };
  if (user.role !== 'admin') return { success: false, message: 'ต้องการสิทธิ์ Admin' };
  return fn(data, user);
}

function verifyToken(token) {
  if (!token) return null;
  try {
    const decoded = Utilities.base64Decode(token);
    const str = Utilities.newBlob(decoded).getDataAsString();
    return JSON.parse(str);
  } catch (e) {
    return null;
  }
}

function makeToken(user) {
  const payload = JSON.stringify({ user_id: user.user_id, username: user.username, role: user.role, house_id: user.house_id });
  return Utilities.base64Encode(payload);
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
  // Sort by year desc, month desc
  reports.sort((a, b) => {
    if (b.year !== a.year) return b.year - a.year;
    return b.month - a.month;
  });
  return { success: true, data: reports };
}

// ── Admin: Houses ─────────────────────────────────────────────

function getHouses(data, user) {
  const houses = sheetToObjects(getSheet(SHEETS.HOUSES));
  return { success: true, data: houses };
}

function addHouse(data, user) {
  const sheet = getSheet(SHEETS.HOUSES);
  const id = genId('H', sheet);
  sheet.appendRow([
    id, data.house_no, data.owner_name, data.address,
    data.area_sqm, data.soi, data.house_type, data.phone,
    data.fee_per_half, 'active', data.note || '', data.account_status || 'ปกติ'
  ]);
  return { success: true, message: 'เพิ่มข้อมูลบ้านสำเร็จ', house_id: id };
}

function updateHouse(data, user) {
  const sheet = getSheet(SHEETS.HOUSES);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] == data.house_id) {
      sheet.getRange(i + 1, 1, 1, 12).setValues([[
        data.house_id, data.house_no, data.owner_name, data.address,
        data.area_sqm, data.soi, data.house_type, data.phone,
        data.fee_per_half, data.status || 'active', data.note || '', data.account_status || 'ปกติ'
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
  const id = genId('U', sheet);
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
  const id = genId('F', sheet);
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
  const id = genId('A', sheet);
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
  const id = genId('R', sheet);
  sheet.appendRow([
    id, data.month, data.year, data.title, data.content,
    data.income || 0, data.expense || 0,
    user.username, new Date().toLocaleDateString('th-TH'),
    data.photo_urls || ''
  ]);
  return { success: true, message: 'เพิ่มรายงานสำเร็จ', report_id: id };
}

function updateNitiReport(data, user) {
  const sheet = getSheet(SHEETS.NITI_REPORT);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] == data.report_id) {
      sheet.getRange(i + 1, 1, 1, 10).setValues([[
        data.report_id, data.month, data.year, data.title, data.content,
        data.income || 0, data.expense || 0,
        user.username, new Date().toLocaleDateString('th-TH'),
        data.photo_urls || rows[i][9] || ''
      ]]);
      return { success: true, message: 'อัปเดตรายงานสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบรายงาน' };
}

function deleteNitiReport(data, user) {
  const sheet = getSheet(SHEETS.NITI_REPORT);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] == data.report_id) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'ลบรายงานสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบรายงาน' };
}
