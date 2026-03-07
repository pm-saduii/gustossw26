// ════════════════════════════════════════════════════════════════
//  VMS — Google Apps Script Backend  (Code.gs)
//  แก้ไข: uploadFile ส่งไป Google Drive Folder, NitiReport ครบ columns
// ════════════════════════════════════════════════════════════════

const SHEET_ID        = '1MDX7JWY33m1lqtHbGGVQXbz3fPP_Dxllm5U_B5ixOhk';
const DRIVE_FOLDER_ID = '1RF2J9YDSmhg_iGLzvyAGRqzSMQ1FuaHw';

// ── Main entry point ────────────────────────────────────────────
function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  const params = {};

  // 1. Query string / form params
  if (e.parameter) {
    Object.keys(e.parameter).forEach(k => { params[k] = e.parameter[k]; });
  }
  // 2. JSON POST body (apiPost() ในฝั่ง frontend ส่ง JSON body)
  if (e.postData && e.postData.contents) {
    try {
      const body = JSON.parse(e.postData.contents);
      Object.keys(body).forEach(k => { params[k] = body[k]; });
    } catch (err) { /* ignore */ }
  }

  const action = params.action || '';
  let result;

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    switch (action) {
      case 'getHouses':          result = getHouses(ss, params);          break;
      case 'addHouse':           result = addHouse(ss, params);           break;
      case 'updateHouse':        result = updateHouse(ss, params);        break;
      case 'deleteHouse':        result = deleteHouse(ss, params);        break;
      case 'getFees':            result = getFees(ss, params);            break;
      case 'addFee':             result = addFee(ss, params);             break;
      case 'updateFee':          result = updateFee(ss, params);          break;
      case 'getFeeSummary':      result = getFeeSummary(ss, params);      break;
      case 'getUsers':           result = getUsers(ss, params);           break;
      case 'addUser':            result = addUser(ss, params);            break;
      case 'updateUser':         result = updateUser(ss, params);         break;
      case 'login':              result = login(ss, params);              break;
      case 'getAnnouncements':   result = getAnnouncements(ss, params);   break;
      case 'addAnnouncement':    result = addAnnouncement(ss, params);    break;
      case 'updateAnnouncement': result = updateAnnouncement(ss, params); break;
      case 'deleteAnnouncement': result = deleteAnnouncement(ss, params); break;
      case 'getNitiReports':     result = getNitiReports(ss, params);     break;
      case 'addNitiReport':      result = addNitiReport(ss, params);      break;
      case 'updateNitiReport':   result = updateNitiReport(ss, params);   break;
      case 'deleteNitiReport':   result = deleteNitiReport(ss, params);   break;
      case 'uploadFile':         result = uploadFile(params);             break;
      default: result = { success: false, message: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { success: false, message: err.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════════
//  FILE UPLOAD  ← แก้ไขหลัก
//  - รับ folderId จาก params (frontend ส่ง DRIVE_FOLDER_ID มาด้วย)
//  - ตั้งค่า sharing = ANYONE_WITH_LINK VIEW
//  - return fileUrl เป็น /file/d/{id}/view
// ════════════════════════════════════════════════════════════════
function uploadFile(params) {
  const base64Data = params.base64Data;
  const fileName   = params.fileName;
  const mimeType   = params.mimeType || 'application/octet-stream';
  const folderId   = params.folderId || DRIVE_FOLDER_ID;

  if (!base64Data || !fileName) {
    return { success: false, message: 'Missing base64Data or fileName' };
  }

  try {
    const folder  = DriveApp.getFolderById(folderId);
    const decoded = Utilities.base64Decode(base64Data);
    const blob    = Utilities.newBlob(decoded, mimeType, fileName);
    const file    = folder.createFile(blob);

    // ✅ Set public read access
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const fileId  = file.getId();
    const fileUrl = 'https://drive.google.com/file/d/' + fileId + '/view';

    return { success: true, fileId: fileId, fileUrl: fileUrl, fileName: fileName };

  } catch (err) {
    return { success: false, message: 'Upload error: ' + err.toString() };
  }
}

// ════════════════════════════════════════════════════════════════
//  NITI REPORTS
//  Sheet "NitiReport" columns:
//  report_id | month | year | title | content | income | expense |
//  created_by | created_date | photo_urls
// ════════════════════════════════════════════════════════════════
function getNitiSheet(ss) { return ss.getSheetByName('NitiReport'); }

function sheetHeaders(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = {};
  headers.forEach((h, i) => { idx[String(h).trim().toLowerCase()] = i; });
  return idx;
}

function getNitiReports(ss, params) {
  const sheet = getNitiSheet(ss);
  if (!sheet) return { success: false, message: 'Sheet NitiReport not found' };

  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const idx     = {};
  headers.forEach((h, i) => { idx[h] = i; });

  const reports = data.slice(1)
    .filter(row => row[idx['report_id'] ?? 0])
    .map(row => ({
      report_id:    String(row[idx['report_id']    ?? 0] || ''),
      month:        Number(row[idx['month']        ?? 1]) || 0,
      year:         Number(row[idx['year']         ?? 2]) || 0,
      title:        String(row[idx['title']        ?? 3] || ''),
      content:      String(row[idx['content']      ?? 4] || ''),
      income:       Number(row[idx['income']       ?? 5]) || 0,
      expense:      Number(row[idx['expense']      ?? 6]) || 0,
      created_by:   String(row[idx['created_by']   ?? 7] || ''),
      created_date: String(row[idx['created_date'] ?? 8] || ''),
      // ✅ photo_urls — return as-is (comma-separated Drive URLs)
      photo_urls:   String(row[idx['photo_urls']   ?? 9] || ''),
    }));

  // Sort: newest year+month first
  reports.sort((a, b) => b.year !== a.year ? b.year - a.year : b.month - a.month);

  return { success: true, data: reports };
}

function addNitiReport(ss, params) {
  const sheet = getNitiSheet(ss);
  if (!sheet) return { success: false, message: 'Sheet NitiReport not found' };

  const reportId   = 'NR' + Date.now();
  const now        = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');
  const createdBy  = params.created_by || 'admin';

  // ✅ photo_urls รับค่า comma-separated URLs จาก frontend โดยตรง
  const photoUrls  = String(params.photo_urls || '');

  sheet.appendRow([
    reportId,
    Number(params.month)   || 0,
    Number(params.year)    || 0,
    params.title           || '',
    params.content         || '',
    Number(params.income)  || 0,
    Number(params.expense) || 0,
    createdBy,
    now,
    photoUrls,              // ← บันทึก photo_urls เข้า column สุดท้าย
  ]);

  return { success: true, message: 'บันทึกรายงานสำเร็จ', report_id: reportId };
}

function updateNitiReport(ss, params) {
  const sheet = getNitiSheet(ss);
  if (!sheet) return { success: false, message: 'Sheet NitiReport not found' };

  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const idx     = {};
  headers.forEach((h, i) => { idx[h] = i; });

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx['report_id'] ?? 0]) === String(params.report_id)) {
      const r = i + 1;
      const set = (col, val) => {
        if (val !== undefined && idx[col] !== undefined) {
          sheet.getRange(r, idx[col] + 1).setValue(val);
        }
      };
      set('month',      Number(params.month) || 0);
      set('year',       Number(params.year)  || 0);
      set('title',      params.title);
      set('content',    params.content);
      set('income',     Number(params.income)  || 0);
      set('expense',    Number(params.expense) || 0);
      // photo_urls — อัปเดตถ้ามีส่งมา
      if (params.photo_urls !== undefined) set('photo_urls', params.photo_urls);

      return { success: true, message: 'อัปเดตรายงานสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบรายงาน: ' + params.report_id };
}

function deleteNitiReport(ss, params) {
  const sheet = getNitiSheet(ss);
  if (!sheet) return { success: false, message: 'Sheet NitiReport not found' };

  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const idx     = {};
  headers.forEach((h, i) => { idx[h] = i; });

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx['report_id'] ?? 0]) === String(params.report_id)) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'ลบรายงานสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบรายงาน' };
}

// ════════════════════════════════════════════════════════════════
//  HOUSES
// ════════════════════════════════════════════════════════════════
function getHouses(ss, params) {
  const sheet = ss.getSheetByName('Houses');
  if (!sheet) return { success: false, message: 'Sheet Houses not found' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const houses = data.slice(1).filter(r => r[0]).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i] !== undefined ? row[i] : ''; });
    return obj;
  });
  return { success: true, data: houses };
}

function addHouse(ss, params) {
  const sheet = ss.getSheetByName('Houses');
  if (!sheet) return { success: false, message: 'Sheet Houses not found' };
  const houseId = 'H' + Date.now();
  sheet.appendRow([houseId, params.house_no||'', params.owner_name||'', params.address||'',
    params.area_sqm||'', params.soi||'', params.house_type||'บ้านเดี่ยว',
    params.phone||'', Number(params.fee_per_half)||0, 'active',
    params.account_status||'ปกติ', params.note||'']);
  return { success: true, message: 'เพิ่มบ้านสำเร็จ', house_id: houseId };
}

function updateHouse(ss, params) {
  const sheet = ss.getSheetByName('Houses');
  if (!sheet) return { success: false, message: 'Sheet Houses not found' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim().toLowerCase());
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(params.house_id)) {
      const u = { house_no:params.house_no, owner_name:params.owner_name, address:params.address,
        area_sqm:params.area_sqm, soi:params.soi, house_type:params.house_type,
        phone:params.phone, fee_per_half:params.fee_per_half,
        status:params.status, account_status:params.account_status, note:params.note };
      headers.forEach((h, ci) => { if (u[h]!==undefined) sheet.getRange(i+1,ci+1).setValue(u[h]); });
      return { success: true, message: 'อัปเดตบ้านสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบบ้าน' };
}

function deleteHouse(ss, params) {
  const sheet = ss.getSheetByName('Houses');
  if (!sheet) return { success: false, message: 'Sheet Houses not found' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const si = headers.indexOf('status');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(params.house_id)) {
      if (si >= 0) sheet.getRange(i+1,si+1).setValue('inactive');
      return { success: true, message: 'ปิดการใช้งานบ้านสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบบ้าน' };
}

// ════════════════════════════════════════════════════════════════
//  FEES
// ════════════════════════════════════════════════════════════════
function getFees(ss, params) {
  const sheet = ss.getSheetByName('CommonFee');
  if (!sheet) return { success: false, message: 'Sheet CommonFee not found' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const fees = data.slice(1).filter(r => r[0]).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i] !== undefined ? row[i] : ''; });
    return obj;
  });
  return { success: true, data: fees };
}

function addFee(ss, params) {
  const sheet = ss.getSheetByName('CommonFee');
  if (!sheet) return { success: false, message: 'Sheet CommonFee not found' };
  const feeId = 'F' + Date.now();
  sheet.appendRow([feeId, params.house_id||'', params.year||'',
    Number(params.h1_amount)||0, Number(params.h1_paid)||0, params.h1_date||'', params.h1_status||'unpaid',
    Number(params.h2_amount)||0, Number(params.h2_paid)||0, params.h2_date||'', params.h2_status||'unpaid',
    params.note||'']);
  return { success: true, message: 'บันทึกค่าส่วนกลางสำเร็จ', fee_id: feeId };
}

function updateFee(ss, params) {
  const sheet = ss.getSheetByName('CommonFee');
  if (!sheet) return { success: false, message: 'Sheet CommonFee not found' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim().toLowerCase());
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(params.fee_id)) {
      const u = { h1_amount:params.h1_amount, h1_paid:params.h1_paid,
        h1_date:params.h1_date, h1_status:params.h1_status,
        h2_amount:params.h2_amount, h2_paid:params.h2_paid,
        h2_date:params.h2_date, h2_status:params.h2_status, note:params.note };
      headers.forEach((h, ci) => { if (u[h]!==undefined) sheet.getRange(i+1,ci+1).setValue(u[h]); });
      return { success: true, message: 'อัปเดตค่าส่วนกลางสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบรายการ' };
}

function getFeeSummary(ss, params) {
  try {
    const feesRes   = getFees(ss, params);
    const housesRes = getHouses(ss, params);
    const fees   = feesRes.data   || [];
    const houses = housesRes.data || [];
    const currentBE = new Date().getFullYear() + 543;
    const yearsSet = new Set(fees.map(f => Number(f.year)).filter(Boolean));
    if (!yearsSet.size) yearsSet.add(currentBE);
    const years = [...yearsSet].sort((a,b) => b-a);
    const year  = params.year ? Number(params.year) : years[0] || currentBE;
    const yearFees = fees.filter(f => Number(f.year) === year);
    const activeHouses = houses.filter(h => h.status !== 'inactive');
    const totalHouses  = activeHouses.length || houses.length;
    let hFullPaid=0, hPartial=0, hUnpaid=0;
    let h1Paid=0, h1Unpaid=0, h2Paid=0, h2Unpaid=0;
    let amtDue=0, amtPaid=0;
    yearFees.forEach(f => {
      const h1s=f.h1_status||'unpaid', h2s=f.h2_status||'unpaid';
      if (h1s==='paid') h1Paid++; else h1Unpaid++;
      if (h2s==='paid') h2Paid++; else h2Unpaid++;
      if (h1s==='paid'&&h2s==='paid') hFullPaid++;
      else if (h1s!=='unpaid'||h2s!=='unpaid') hPartial++;
      else hUnpaid++;
      amtDue  += Number(f.h1_amount||0) + Number(f.h2_amount||0);
      amtPaid += Number(f.h1_paid||0)   + Number(f.h2_paid||0);
    });
    const paidPct      = amtDue>0 ? Math.round(amtPaid/amtDue*100) : 0;
    const housePaidPct = totalHouses>0 ? Math.round(hFullPaid/totalHouses*100) : 0;
    const bySoi = {};
    activeHouses.forEach(h => {
      const soi = h.soi || 'อื่นๆ';
      if (!bySoi[soi]) bySoi[soi]={total:0,fullPaid:0,partial:0,unpaid:0,h1Paid:0,h1Unpaid:0,h2Paid:0,h2Unpaid:0};
      bySoi[soi].total++;
      const fee = yearFees.find(f => String(f.house_id)===String(h.house_id));
      if (!fee) { bySoi[soi].unpaid++; bySoi[soi].h1Unpaid++; bySoi[soi].h2Unpaid++; return; }
      const h1s=fee.h1_status||'unpaid', h2s=fee.h2_status||'unpaid';
      if (h1s==='paid') bySoi[soi].h1Paid++; else bySoi[soi].h1Unpaid++;
      if (h2s==='paid') bySoi[soi].h2Paid++; else bySoi[soi].h2Unpaid++;
      if (h1s==='paid'&&h2s==='paid') bySoi[soi].fullPaid++;
      else if (h1s!=='unpaid'||h2s!=='unpaid') bySoi[soi].partial++;
      else bySoi[soi].unpaid++;
    });
    return { success:true, data:{ year, years, totalHouses,
      hFullPaid, hPartial, hUnpaid, h1Paid, h1Unpaid, h2Paid, h2Unpaid,
      amtDue, amtPaid, paidPct, housePaidPct, bySoi }};
  } catch(err) {
    return { success:false, message:err.toString() };
  }
}

// ════════════════════════════════════════════════════════════════
//  USERS
// ════════════════════════════════════════════════════════════════
function getUsers(ss, params) {
  const sheet = ss.getSheetByName('Users');
  if (!sheet) return { success:false, message:'Sheet Users not found' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const users = data.slice(1).filter(r=>r[0]).map(row => {
    const obj={};
    headers.forEach((h,i)=>{ obj[h]=row[i]!==undefined?row[i]:''; });
    delete obj.password;
    return obj;
  });
  return { success:true, data:users };
}

function addUser(ss, params) {
  const sheet = ss.getSheetByName('Users');
  if (!sheet) return { success:false, message:'Sheet Users not found' };
  const userId='U'+Date.now();
  sheet.appendRow([userId,params.username||'',params.password||'',
    params.role||'resident',params.house_id||'',params.full_name||'','TRUE']);
  return { success:true, message:'เพิ่มผู้ใช้สำเร็จ', user_id:userId };
}

function updateUser(ss, params) {
  const sheet = ss.getSheetByName('Users');
  if (!sheet) return { success:false, message:'Sheet Users not found' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim().toLowerCase());
  for (let i=1; i<data.length; i++) {
    if (String(data[i][0])===String(params.user_id)) {
      const u={full_name:params.full_name, house_id:params.house_id, active:params.active};
      if (params.password && params.password!=='***') u.password=params.password;
      headers.forEach((h,ci)=>{ if(u[h]!==undefined) sheet.getRange(i+1,ci+1).setValue(u[h]); });
      return { success:true, message:'อัปเดตผู้ใช้สำเร็จ' };
    }
  }
  return { success:false, message:'ไม่พบผู้ใช้' };
}

function login(ss, params) {
  const sheet = ss.getSheetByName('Users');
  if (!sheet) return { success:false, message:'ระบบไม่พร้อม' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const idx={};
  headers.forEach((h,i)=>{ idx[h]=i; });
  for (let i=1; i<data.length; i++) {
    const row=data[i];
    if (String(row[idx.username??1])===String(params.username) &&
        String(row[idx.password??2])===String(params.password) &&
        String(row[idx.active??6])==='TRUE') {
      const payload={
        user_id:row[idx.user_id??0], username:row[idx.username??1],
        role:row[idx.role??3], house_id:row[idx.house_id??4], full_name:row[idx.full_name??5]
      };
      const token=Utilities.base64Encode(JSON.stringify(payload));
      return { success:true, token, role:payload.role, name:payload.full_name, house_id:payload.house_id };
    }
  }
  return { success:false, message:'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
}

// ════════════════════════════════════════════════════════════════
//  ANNOUNCEMENTS
// ════════════════════════════════════════════════════════════════
function getAnnouncements(ss, params) {
  const sheet = ss.getSheetByName('Announcements');
  if (!sheet) return { success:false, message:'Sheet Announcements not found' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const anns = data.slice(1).filter(r=>r[0]).map(row => {
    const obj={};
    headers.forEach((h,i)=>{ obj[h]=row[i]!==undefined?row[i]:''; });
    return obj;
  }).filter(a => a.active==='TRUE' || a.active===true || params.token);
  anns.sort((a,b) => String(b.date||'').localeCompare(String(a.date||'')));
  return { success:true, data:anns };
}

function addAnnouncement(ss, params) {
  const sheet = ss.getSheetByName('Announcements');
  if (!sheet) return { success:false, message:'Sheet Announcements not found' };
  const annId='A'+Date.now();
  sheet.appendRow([annId,params.title||'',params.content||'',
    params.category||'ทั่วไป',params.date||'','TRUE',params.file_url||'']);
  return { success:true, message:'เพิ่มประกาศสำเร็จ', ann_id:annId };
}

function updateAnnouncement(ss, params) {
  const sheet = ss.getSheetByName('Announcements');
  if (!sheet) return { success:false, message:'Sheet Announcements not found' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim().toLowerCase());
  for (let i=1; i<data.length; i++) {
    if (String(data[i][0])===String(params.ann_id)) {
      const u={title:params.title,content:params.content,category:params.category,date:params.date,active:params.active};
      headers.forEach((h,ci)=>{ if(u[h]!==undefined) sheet.getRange(i+1,ci+1).setValue(u[h]); });
      return { success:true, message:'อัปเดตประกาศสำเร็จ' };
    }
  }
  return { success:false, message:'ไม่พบประกาศ' };
}

function deleteAnnouncement(ss, params) {
  const sheet = ss.getSheetByName('Announcements');
  if (!sheet) return { success:false, message:'Sheet Announcements not found' };
  const data = sheet.getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if (String(data[i][0])===String(params.ann_id)) {
      sheet.deleteRow(i+1);
      return { success:true, message:'ลบประกาศสำเร็จ' };
    }
  }
  return { success:false, message:'ไม่พบประกาศ' };
}
