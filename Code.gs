// ============================================================
// Code.gs — Village Management System (Fully Tuned & Secure)
// ============================================================

const SHEET_ID = '1MDX7JWY33m1lqtHbGGVQXbz3fPP_Dxllm5U_B5ixOhk';
const SECRET_KEY = 'VMS_SECURE_TOKEN_KEY_99'; // แนะนำให้เปลี่ยนเป็นรหัสลับของคุณเอง

const SHEETS = {
  USERS: 'Users',
  HOUSES: 'Houses',
  COMMON_FEE: 'CommonFee',
  ANNOUNCEMENTS: 'Announcements',
  NITI_REPORT: 'NitiReport'
};

// ── SECURITY: HMAC Token Management (ระบบความปลอดภัยใหม่) ──

function generateSecureToken(payload) {
  const data = JSON.stringify({
    ...payload,
    exp: new Date().getTime() + (24 * 60 * 60 * 1000) // หมดอายุใน 24 ชม.
  });
  const dataBase64 = Utilities.base64EncodeWebSafe(data);
  const signature = Utilities.computeHmacSha256Signature(dataBase64, SECRET_KEY);
  const sigBase64 = Utilities.base64EncodeWebSafe(signature);
  return dataBase64 + "." + sigBase64;
}

function validateToken(token) {
  try {
    if (!token) return null;
    const [dataBase64, sigProvided] = token.split('.');
    const sigExpected = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(dataBase64, SECRET_KEY));
    if (sigProvided !== sigExpected) return null;
    const payload = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(dataBase64)).getDataAsString());
    if (new Date().getTime() > payload.exp) return null;
    return payload;
  } catch (e) { return null; }
}

// ── ENTRY POINTS ──────────────────────────────────────────────

function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  const params = e.parameter || {};
  const action = params['action'] || '';
  
  // สิทธิ์การเข้าถึง
  const publicActions = ['login', 'getFeeSummary', 'getAnnouncements', 'getNitiReports'];
  const adminActions = [
    'updateHouse', 'addAnnouncement', 'updateAnnouncement', 'deleteAnnouncement', 
    'addNitiReport', 'updateNitiReport', 'deleteNitiReport', 
    'addFee', 'updateFee', 'deleteFee', 'getUsers', 'getHouses', 'getFees'
  ];

  let user = null;
  if (!publicActions.includes(action)) {
    user = validateToken(params.token);
    if (!user) return createJsonResponse({ error: 'Unauthorized' });
    if (adminActions.includes(action) && user.role !== 'admin') {
      return createJsonResponse({ success: false, message: 'Permission Denied' });
    }
  }

  let result;
  try {
    switch (action) {
      case 'login': result = login(params); break;
      
      // Resident
      case 'getMyInfo':     result = getMyInfo(params, user); break;
      case 'getMyFees':     result = getMyFees(params, user); break;
      
      // Public / Summary (พร้อมระบบ Cache)
      case 'getFeeSummary':    result = getCachedData('fee_sum_'+params.year, () => getFeeSummary(params)); break;
      case 'getAnnouncements': result = getCachedData('ann_list', () => getAnnouncements()); break;
      case 'getNitiReports':   result = getCachedData('niti_list', () => getNitiReports()); break;

      // Admin: Management
      case 'getUsers':         result = getUsers(); break;
      case 'getHouses':        result = getHouses(); break;
      case 'updateHouse':      result = updateHouse(params); break;
      case 'getFees':          result = getFees(params); break;
      case 'addFee':           result = addFee(params); break;
      case 'updateFee':        result = updateFee(params); break;
      case 'deleteFee':        result = deleteFee(params); break;

      // Admin: Announcements
      case 'addAnnouncement':    result = addAnnouncement(params, user); clearCache(['ann_list']); break;
      case 'updateAnnouncement': result = updateAnnouncement(params, user); clearCache(['ann_list']); break;
      case 'deleteAnnouncement': result = deleteAnnouncement(params); clearCache(['ann_list']); break;

      // Admin: Niti Reports
      case 'addNitiReport':      result = addNitiReport(params, user); clearCache(['niti_list']); break;
      case 'updateNitiReport':   result = updateNitiReport(params, user); clearCache(['niti_list']); break;
      case 'deleteNitiReport':   result = deleteNitiReport(params); clearCache(['niti_list']); break;

      case 'uploadFile':       result = uploadFile(params); break;
      
      default: result = { success: false, message: 'Action not found' };
    }
  } catch (err) {
    result = { success: false, message: err.toString() };
  }
  return createJsonResponse(result);
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

// ── PERFORMANCE: Cache Logic ──────────────────────────────────

function getCachedData(key, fetchFn) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(key);
  if (cached) return JSON.parse(cached);
  const data = fetchFn();
  cache.put(key, JSON.stringify(data), 600); // 10 mins
  return data;
}

function clearCache(keys) {
  const cache = CacheService.getScriptCache();
  keys.forEach(k => cache.remove(k));
}

// ── HELPERS ───────────────────────────────────────────────────

function getSheet(name) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function genId(prefix, sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return prefix + '001';
  const lastId = sheet.getRange(lastRow, 1).getValue().toString();
  const num = parseInt(lastId.replace(prefix, '')) + 1;
  return prefix + num.toString().padStart(3, '0');
}

// ── CORE FUNCTIONS (Original Logic with Enhancements) ─────────

function login(data) {
  const rows = sheetToObjects(getSheet(SHEETS.USERS));
  const user = rows.find(r => r.username == data.username && r.password == data.password);
  if (user) {
    const payload = { username: user.username, role: user.role, house_id: user.house_id, full_name: user.full_name };
    return { success: true, token: generateSecureToken(payload), role: user.role, full_name: user.full_name, house_id: user.house_id };
  }
  return { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
}

function getMyInfo(params, user) {
  const houses = sheetToObjects(getSheet(SHEETS.HOUSES));
  const house = houses.find(h => h.house_id == user.house_id);
  return { success: true, data: house };
}

function getMyFees(params, user) {
  const fees = sheetToObjects(getSheet(SHEETS.COMMON_FEE));
  const myFees = fees.filter(f => f.house_id == user.house_id);
  return { success: true, data: myFees };
}

function getFeeSummary(params) {
  const year = params.year || new Date().getFullYear().toString();
  const fees = sheetToObjects(getSheet(SHEETS.COMMON_FEE));
  return { success: true, data: fees.filter(f => f.year == year) };
}

function getAnnouncements() {
  const data = sheetToObjects(getSheet(SHEETS.ANNOUNCEMENTS));
  return { success: true, data: data.reverse() };
}

function getNitiReports() {
  const data = sheetToObjects(getSheet(SHEETS.NITI_REPORT));
  return { success: true, data: data.reverse() };
}

// ── ADMIN FUNCTIONS ───────────────────────────────────────────

function getUsers() { return { success: true, data: sheetToObjects(getSheet(SHEETS.USERS)) }; }
function getHouses() { return { success: true, data: sheetToObjects(getSheet(SHEETS.HOUSES)) }; }

function updateHouse(data) {
  const sheet = getSheet(SHEETS.HOUSES);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] == data.house_id) {
      sheet.getRange(i + 1, 2, 1, 4).setValues([[data.owner_name, data.soi, data.phone, data.email]]);
      return { success: true, message: 'อัปเดตข้อมูลบ้านสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบข้อมูลบ้าน' };
}

function getFees(params) {
  const fees = sheetToObjects(getSheet(SHEETS.COMMON_FEE));
  if (params.year) return { success: true, data: fees.filter(f => f.year == params.year) };
  return { success: true, data: fees };
}

function addFee(data) {
  const sheet = getSheet(SHEETS.COMMON_FEE);
  const id = genId('F', sheet);
  sheet.appendRow([id, data.house_id, data.month, data.year, data.amount, data.paid_amount, data.status, data.paid_date, data.remark]);
  return { success: true, message: 'เพิ่มข้อมูลค่าส่วนกลางสำเร็จ' };
}

function updateFee(data) {
  const sheet = getSheet(SHEETS.COMMON_FEE);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] == data.fee_id) {
      sheet.getRange(i + 1, 2, 1, 8).setValues([[data.house_id, data.month, data.year, data.amount, data.paid_amount, data.status, data.paid_date, data.remark]]);
      return { success: true, message: 'อัปเดตข้อมูลสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบข้อมูล' };
}

function deleteFee(data) {
  const sheet = getSheet(SHEETS.COMMON_FEE);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] == data.fee_id) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'ลบข้อมูลสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบข้อมูล' };
}

function addAnnouncement(data, user) {
  const sheet = getSheet(SHEETS.ANNOUNCEMENTS);
  const id = genId('A', sheet);
  sheet.appendRow([id, data.date, data.category, data.title, data.content, user.full_name, data.file_url || '']);
  return { success: true, message: 'เพิ่มประกาศสำเร็จ' };
}

function updateAnnouncement(data, user) {
  const sheet = getSheet(SHEETS.ANNOUNCEMENTS);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] == data.ann_id) {
      sheet.getRange(i + 1, 2, 1, 6).setValues([[data.date, data.category, data.title, data.content, user.full_name, data.file_url || rows[i][6]]]);
      return { success: true, message: 'อัปเดตประกาศสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบประกาศ' };
}

function deleteAnnouncement(data) {
  const sheet = getSheet(SHEETS.ANNOUNCEMENTS);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] == data.ann_id) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'ลบประกาศสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบประกาศ' };
}

function addNitiReport(data, user) {
  const sheet = getSheet(SHEETS.NITI_REPORT);
  const id = genId('R', sheet);
  sheet.appendRow([id, data.month, data.year, data.title, data.content, data.income || 0, data.expense || 0, user.username, new Date().toLocaleDateString('th-TH'), data.photo_urls || '']);
  return { success: true, message: 'เพิ่มรายงานสำเร็จ' };
}

function updateNitiReport(data, user) {
  const sheet = getSheet(SHEETS.NITI_REPORT);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] == data.report_id) {
      sheet.getRange(i + 1, 1, 1, 10).setValues([[data.report_id, data.month, data.year, data.title, data.content, data.income || 0, data.expense || 0, user.username, new Date().toLocaleDateString('th-TH'), data.photo_urls || rows[i][9]]]);
      return { success: true, message: 'อัปเดตรายงานสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบรายงาน' };
}

function deleteNitiReport(data) {
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

function uploadFile(data) {
  try {
    const folderId = '1RF2J9YDSmhg_iGLzvyAGRqzSMQ1FuaHw'; // ใส่ Folder ID ของคุณที่นี่
    const folder = DriveApp.getFolderById(folderId);
    const blob = Utilities.newBlob(Utilities.base64Decode(data.base64Data), data.mimeType, data.fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, fileUrl: file.getUrl() };
  } catch (e) { return { success: false, message: e.toString() }; }
}