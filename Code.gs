const SPREADSHEET_ID = '1xuu1ibhlTuEmS2fhvtvdW9FY2lh2Ic_F9qsSHgyvYlc';
const MANPOWER_DB_ID = '1Wn1gnstzG_2Wi_Tc95cw0rIC1dIoHJvHbghlxceJy9A';
const TARGET_SHEET_NAME = 'Target';
const WEEKLY_UPDATE_SHEET_NAME = 'Weekly_Update';
const USER_SHEET_NAME = 'Users';

const WEEKLY_HEADERS = [
  'Timestamp', 'Action Date', 'Store No', 'Branch', 'Branch TH', 'Region',
  'Main Vendor', 'Update Vendor', 'Type', 'Position', 'Interview Passed',
  'Started Work', 'Remark', 'Updated By'
];

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';

  if (!action) {
    return HtmlService.createHtmlOutput(getAppHtml_())
      .setTitle('AEC Tracking System (Pro)')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  let result;
  try {
    switch (action) {
      case 'login':
        result = loginUserOnly(e.parameter.email);
        break;
      case 'initApp':
        result = initApp(
          e.parameter.email,
          e.parameter.filterStart || '',
          e.parameter.filterEnd || '',
          e.parameter.forceRefresh === 'true',
          e.parameter.gridIntDate || '',
          e.parameter.gridWorkDate || ''
        );
        break;
      case 'getGridData':
        result = getExistingGridDataAPI(e.parameter.email, e.parameter.type, e.parameter.date);
        break;
      case 'exportAdmin':
        result = { url: exportAdminExcelAPI(e.parameter.email, e.parameter.filterStart || '', e.parameter.filterEnd || '') };
        break;
      default:
        result = { success: false, error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { success: false, error: err.message };
  }

  const json = JSON.stringify(result);

  // JSONP support — เรียกผ่าน <script> tag เพื่อข้าม CORS จาก github.io
  const callback = e && e.parameter && e.parameter.callback;
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────
//  ดึงหน้าเว็บ (index.html) จาก GitHub อัตโนมัติ
//  → deploy Code.gs ครั้งเดียว ทุกครั้งที่แก้ index.html บน GitHub จะอัปเดตเอง
//  → ไม่ต้อง copy index มาวางใน Apps Script อีกต่อไป
// ─────────────────────────────────────────────
function getAppHtml_() {
  const HTML_URL = 'https://raw.githubusercontent.com/recruitmakrocareer/AEC/main/index.html';
  const cache = CacheService.getScriptCache();
  const CKEY = 'AEC_INDEX_HTML_V1';

  // cache 60 วิ เพื่อให้อัปเดตไว แต่ไม่ต้อง fetch ทุกครั้ง
  const cached = cache.get(CKEY);
  if (cached) return cached;

  try {
    const res = UrlFetchApp.fetch(HTML_URL, { muteHttpExceptions: true, followRedirects: true });
    if (res.getResponseCode() === 200) {
      const html = res.getContentText();
      try { if (html.length < 100000) cache.put(CKEY, html, 60); } catch (e) {}
      return html;
    }
  } catch (e) {}

  // สำรอง: ถ้าดึงจาก GitHub ไม่ได้ ใช้ไฟล์ index ที่ฝังใน project (ถ้ามี)
  try {
    return HtmlService.createHtmlOutputFromFile('index').getContent();
  } catch (e) {
    return '<h2 style="font-family:sans-serif;text-align:center;margin-top:40px;">โหลดหน้าเว็บไม่สำเร็จ กรุณาลองใหม่อีกครั้ง</h2>';
  }
}

function doPost(e) {
  let result;
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    switch (action) {
      case 'submitData':
        result = submitCombinedData(payload.data);
        break;
      default:
        result = { success: false, error: 'Unknown action' };
    }
  } catch (err) {
    result = { success: false, error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────
//  UTILS
//  Store_No ใน External DB เป็นตัวเลข 1,2,3
//  Local DB ใช้ padded "001","002","003"
//  ต้อง normalize ให้ตรงกันก่อน compare
// ─────────────────────────────────────────────
function padStoreNo_(val) {
  if (!val && val !== 0) return '';
  return String(val).trim().replace(/\D/g, '').padStart(3, '0');
}

function normalizeDate_(val) {
  if (!val) return '';
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(val).trim();
}

// ─────────────────────────────────────────────
//  AUTH
// ─────────────────────────────────────────────
function loginUserOnly(email) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const cleanEmail = String(email || '').trim().toLowerCase();
    const usersSheet = ss.getSheetByName(USER_SHEET_NAME);
    if (!usersSheet) throw new Error('ไม่พบชีต Users กรุณาติดต่อ Admin');
    const users = usersSheet.getDataRange().getValues().slice(1);
    const userRow = users.find(row => String(row[0] || '').trim().toLowerCase() === cleanEmail);
    if (!userRow) throw new Error('อีเมลนี้ไม่มีสิทธิ์เข้าใช้งาน กรุณาตรวจสอบให้ถูกต้อง');
    return {
      user: {
        email: cleanEmail,
        role: String(userRow[1] || '').trim().toLowerCase(),
        vendor: String(userRow[2] || '').trim(),
        isAdmin: String(userRow[1] || '').trim().toLowerCase() === 'admin'
      }
    };
  } catch (e) {
    throw new Error(e.message);
  }
}

// ─────────────────────────────────────────────
//  INIT APP — single bundled call after login
// ─────────────────────────────────────────────
function initApp(email, filterStart, filterEnd, forceRefresh, gridIntDate, gridWorkDate) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const userObj = loginUserOnly(email).user;
    const targets = getTargets_(ss);

    // ใช้ cache เสมอถ้าไม่ได้ force refresh
    // เพื่อป้องกัน timeout จากการอ่าน External DB ซ้ำๆ
    const extDB = getExternalDBData_(forceRefresh === true);

    if (!extDB) {
      return { success: false, error: 'ไม่สามารถโหลด External DB ได้ กรุณาลองใหม่' };
    }

    const dashboard = getDashboardDataCore_(ss, userObj, targets, extDB, filterStart || '', filterEnd || '');
    const stores = targets.map(row => ({ storeNo: row.storeNo, branch: row.branch, region: row.region }));

    // รวมข้อมูลตารางกรอก (Interview/StartWork) มาด้วยในครั้งเดียว เพื่อลดจำนวน request → โหลดเร็วขึ้น
    const gridInt  = gridIntDate  ? getExistingGridDataCore_(ss, userObj.vendor, 'Interview', gridIntDate)  : null;
    const gridWork = gridWorkDate ? getExistingGridDataCore_(ss, userObj.vendor, 'StartWork', gridWorkDate) : null;

    return { success: true, user: userObj, dashboard: dashboard, stores: stores, gridInt: gridInt, gridWork: gridWork };
  } catch (e) {
    // ต้อง return object ไม่ใช่ throw เพื่อให้ frontend จัดการได้
    return { success: false, error: e.message };
  }
}

// ─────────────────────────────────────────────
//  TARGETS (Local DB — Target sheet)
//  Col: A=Store_No, B=Store_Name, C=Store_Name_TH,
//       D=Format, E=Region, F=Vendor, G=Target
// ─────────────────────────────────────────────
function getTargets_(ss) {
  const sheet = ss.getSheetByName(TARGET_SHEET_NAME);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  // skip header row (row[0] = 'Store_No' text)
  return values.filter(r => r[0] !== '' && !isNaN(Number(r[0])) && String(r[0]).trim() !== '').map(r => ({
    storeNo: padStoreNo_(r[0]),          // "001"
    branch: String(r[1] || ''),
    branchTH: String(r[2] || ''),
    region: String(r[4] || ''),
    mainVendor: String(r[5] || '').trim(),
    totalTarget: Number(String(r[6] || '0').replace(/,/g, '')) || 0
  }));
}

// ─────────────────────────────────────────────
//  EXTERNAL DB WITH CACHE
//  Master_Target_AEC   : Store_No | Department | Target
//  Manpower_Status_AEC : Store_No | Store_Name | Format | Subregion | OD |
//                        Regional CEO | Store_Focus | Department | Target | Active
// ─────────────────────────────────────────────
function getExternalDBData_(forceRefresh) {
  const cache = CacheService.getScriptCache();
  const CACHE_KEY = 'AEC_EXT_DB_V8';

  if (forceRefresh === true) cache.remove(CACHE_KEY);

  const cached = cache.get(CACHE_KEY);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }

  const result = {
    manpower: { success: false, map: {}, roles: {}, totals: { active: 0 }, error: '' },
    master:   { success: false, map: {}, error: '' }
  };

  try {
    const ssDB = SpreadsheetApp.openById(MANPOWER_DB_ID);

    // ── Master_Target_AEC ──
    // Headers: Store_No | Department | Target
    const mtSheet = ssDB.getSheetByName('Master_Target_AEC');
    if (mtSheet) {
      const rows = mtSheet.getDataRange().getValues();
      // row[0] = headers → skip
      for (let i = 1; i < rows.length; i++) {
        const rawStore = rows[i][0];  // Store_No (number)
        const target   = Number(rows[i][2]) || 0;  // Target col index 2
        if (!rawStore && rawStore !== 0) continue;
        const storeNo = padStoreNo_(rawStore);
        result.master.map[storeNo] = (result.master.map[storeNo] || 0) + target;
      }
      result.master.success = true;
    } else {
      result.master.error = 'ไม่พบชีต Master_Target_AEC';
    }

    // ── Manpower_Status_AEC ──
    // Headers: Store_No(0) | Store_Name(1) | Format(2) | Subregion(3) | OD(4) |
    //          Regional CEO(5) | Store_Focus(6) | Department(7) | Target(8) | Active(9)
    const mpSheet = ssDB.getSheetByName('Manpower_Status_AEC');
    if (mpSheet) {
      const rows = mpSheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        const rawStore = rows[i][0];  // Store_No
        const dept     = String(rows[i][7] || 'Other').trim();  // Department
        const active   = Number(rows[i][9]) || 0;               // Active
        if (!rawStore && rawStore !== 0) continue;
        const storeNo = padStoreNo_(rawStore);
        if (!result.manpower.map[storeNo]) result.manpower.map[storeNo] = { active: 0 };
        result.manpower.map[storeNo].active += active;
        if (active > 0) result.manpower.roles[dept] = (result.manpower.roles[dept] || 0) + active;
        result.manpower.totals.active += active;
      }
      result.manpower.success = true;
    } else {
      result.manpower.error = 'ไม่พบชีต Manpower_Status_AEC';
    }

    // Cache result (30 นาที) — ลดการอ่าน External Spreadsheet ซ้ำที่ทำให้โหลดช้า
    // กด "Apply Filter (Refresh)" หรือ login ใหม่ จะ force refresh ดึงข้อมูลสดอยู่แล้ว
    try {
      const str = JSON.stringify(result);
      if (str.length < 100000) cache.put(CACHE_KEY, str, 1800);
    } catch (e) {}

  } catch (e) {
    result.master.error   = 'DB Error: ' + e.message;
    result.manpower.error = 'DB Error: ' + e.message;
  }

  return result;
}

// ─────────────────────────────────────────────
//  DASHBOARD CORE
// ─────────────────────────────────────────────
function getDashboardDataCore_(ss, currentUser, targets, extDB, filterStart, filterEnd) {
  const manpowerMap    = extDB.manpower.map;
  const masterTargetMap = extDB.master.map;

  const dbErrors = [];
  if (extDB.manpower.error) dbErrors.push(extDB.manpower.error);
  if (extDB.master.error)   dbErrors.push(extDB.master.error);

  // ── Read Weekly_Update ──
  let allUpdates = [];
  const weeklySheet = ss.getSheetByName(WEEKLY_UPDATE_SHEET_NAME)
    || ss.getSheets().find(s => s.getName().toLowerCase() === WEEKLY_UPDATE_SHEET_NAME.toLowerCase());

  if (weeklySheet) {
    const values = weeklySheet.getDataRange().getValues();
    if (values.length >= 2) {
      const headers = values[0].map(h => String(h).trim().replace(/\s+/g, ' ').toLowerCase());
      allUpdates = values.slice(1).map(row => {
        const obj = {};
        headers.forEach((h, i) => obj[h] = row[i]);
        return obj;
      });
    }
  }

  // ── Date filter (admin only) ──
  let filteredUpdates = allUpdates;
  if (filterStart && filterEnd && currentUser.isAdmin) {
    filteredUpdates = allUpdates.filter(r => {
      const d = normalizeDate_(r['action date']);
      return d >= filterStart && d <= filterEnd;
    });
  }

  const scopedUpdates = currentUser.isAdmin
    ? filteredUpdates
    : filteredUpdates.filter(r => String(r['update vendor']).trim() === currentUser.vendor);

  // ── Aggregate weekly per store (use filteredUpdates for interview/started counts) ──
  const summaryMap = {};
  filteredUpdates.forEach(r => {
    const s = padStoreNo_(r['store no']);
    if (!summaryMap[s]) summaryMap[s] = { intPassed: 0, started: 0, lastIntDate: '', lastStartDate: '' };
    const intVal   = Number(r['interview passed'] || r['interview_passed'] || 0);
    const startVal = Number(r['started work']     || r['started_work']     || 0);
    summaryMap[s].intPassed += intVal;
    summaryMap[s].started   += startVal;
    const rDate = normalizeDate_(r['action date']);
    const rType = String(r['type'] || '').toLowerCase();
    if (rType.includes('interview') && intVal > 0) {
      if (!summaryMap[s].lastIntDate || rDate > summaryMap[s].lastIntDate) summaryMap[s].lastIntDate = rDate;
    }
    if (rType.includes('start') && startVal > 0) {
      if (!summaryMap[s].lastStartDate || rDate > summaryMap[s].lastStartDate) summaryMap[s].lastStartDate = rDate;
    }
  });

  // ── Build branch rows ──
  const allBranchRows = [];
  const myBranchRows  = [];
  const allTotals = { target: 0, interviewPassed: 0, startedWork: 0, remaining: 0, activeDB: 0 };
  const myTotals  = { target: 0, interviewPassed: 0, startedWork: 0, remaining: 0, activeDB: 0 };

  targets.forEach(t => {
    const localSum = summaryMap[t.storeNo] || { intPassed: 0, started: 0, lastIntDate: '', lastStartDate: '' };

    // Override target from Master_Target_AEC if available
    let totalTarget = t.totalTarget;
    if (extDB.master.success && masterTargetMap[t.storeNo] !== undefined) {
      totalTarget = masterTargetMap[t.storeNo];
    }

    // Active from Manpower_Status_AEC, fallback to local started
    let activeDB = 0;
    if (extDB.manpower.success && manpowerMap[t.storeNo]) {
      activeDB = manpowerMap[t.storeNo].active;
    } else {
      activeDB = localSum.started;
    }

    // Core calculations — DO NOT CHANGE
    const trueRemaining = Math.max(totalTarget - activeDB, 0);
    const progress      = totalTarget ? Math.round((activeDB / totalTarget) * 100) : 0;

    const rowData = {
      storeNo: t.storeNo, branch: t.branch, region: t.region, mainVendor: t.mainVendor,
      totalTarget,
      interviewPassed: localSum.intPassed,
      startedWork:     localSum.started,
      remaining:       trueRemaining,
      progress,
      activeDB,
      lastIntDate:   localSum.lastIntDate,
      lastStartDate: localSum.lastStartDate
    };

    const maskedRow = Object.assign({}, rowData);
    if (!currentUser.isAdmin && maskedRow.mainVendor !== currentUser.vendor) maskedRow.mainVendor = '-';

    allBranchRows.push(maskedRow);
    allTotals.target          += totalTarget;
    allTotals.interviewPassed += localSum.intPassed;
    allTotals.startedWork     += localSum.started;
    allTotals.remaining       += trueRemaining;
    allTotals.activeDB        += activeDB;

    if (t.mainVendor === currentUser.vendor) {
      myBranchRows.push(rowData);
      myTotals.target          += totalTarget;
      myTotals.interviewPassed += localSum.intPassed;
      myTotals.startedWork     += localSum.started;
      myTotals.remaining       += trueRemaining;
      myTotals.activeDB        += activeDB;
    }
  });

  allTotals.progress = allTotals.target ? Math.round((allTotals.activeDB / allTotals.target) * 100) : 0;
  myTotals.progress  = myTotals.target  ? Math.round((myTotals.activeDB  / myTotals.target)  * 100) : 0;

  // ── History ──
  const historyRows = scopedUpdates.map(r => ({
    actionDate: normalizeDate_(r['action date']),
    storeNo:    padStoreNo_(r['store no']),
    branch:     String(r['branch']        || ''),
    vendor:     String(r['update vendor'] || ''),
    type:       String(r['type']          || 'Interview'),
    position:   String(r['position']      || 'Other'),
    qty:        Number(r['interview passed'] || 0) + Number(r['started work'] || 0),
    remark:     String(r['remark']        || '')
  })).reverse();

  // ── Region & Vendor rows (admin only) ──
  let regionRows = [];
  let vendorRows = [];

  if (currentUser.isAdmin) {
    const regMap = {};
    allBranchRows.forEach(r => {
      if (!regMap[r.region]) regMap[r.region] = { target: 0, passed: 0, started: 0, remaining: 0, active: 0 };
      regMap[r.region].target   += r.totalTarget;
      regMap[r.region].passed   += r.interviewPassed;
      regMap[r.region].started  += r.startedWork;
      regMap[r.region].remaining += r.remaining;
      regMap[r.region].active   += r.activeDB;
    });
    regionRows = Object.keys(regMap).sort().map(k => ({
      region:    k,
      target:    regMap[k].target,
      passed:    regMap[k].passed,
      started:   regMap[k].started,
      remaining: regMap[k].remaining,
      progress:  regMap[k].target ? Math.round((regMap[k].active / regMap[k].target) * 100) : 0
    }));

    const venMap = {};
    filteredUpdates.forEach(r => {
      const v = String(r['update vendor'] || '').trim() || 'Unknown';
      if (!venMap[v]) venMap[v] = { intLam: 0, intBak: 0, intOth: 0, wrkLam: 0, wrkBak: 0, wrkOth: 0 };
      const pos  = String(r['position'] || '');
      const intP = Number(r['interview passed'] || 0);
      const strW = Number(r['started work']     || 0);
      if (String(r['type'] || '').toLowerCase().includes('interview')) {
        if (pos === 'ล่าม') venMap[v].intLam += intP;
        else if (pos === 'Bakery') venMap[v].intBak += intP;
        else venMap[v].intOth += intP;
      } else if (String(r['type'] || '').toLowerCase().includes('start')) {
        if (pos === 'ล่าม') venMap[v].wrkLam += strW;
        else if (pos === 'Bakery') venMap[v].wrkBak += strW;
        else venMap[v].wrkOth += strW;
      }
    });
    vendorRows = Object.keys(venMap).sort().map(k => ({
      vendor: k,
      intLam: venMap[k].intLam, intBak: venMap[k].intBak, intOth: venMap[k].intOth,
      intTot: venMap[k].intLam + venMap[k].intBak + venMap[k].intOth,
      wrkLam: venMap[k].wrkLam, wrkBak: venMap[k].wrkBak, wrkOth: venMap[k].wrkOth,
      wrkTot: venMap[k].wrkLam + venMap[k].wrkBak + venMap[k].wrkOth
    }));
  }

  return {
    allTotals, myTotals, allBranchRows, myBranchRows, historyRows, regionRows, vendorRows,
    rawUpdates: currentUser.isAdmin ? filteredUpdates : [],
    dbError:  dbErrors.join(' | '),
    dbRoles:  extDB.manpower.roles,
    dbTotals: extDB.manpower.totals
  };
}

// ─────────────────────────────────────────────
//  GRID DATA
// ─────────────────────────────────────────────
function getExistingGridDataAPI(email, type, date) {
  const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  const userObj = loginUserOnly(email).user;
  return getExistingGridDataCore_(ss, userObj.vendor, type, date);
}

function getExistingGridDataCore_(ss, vendor, type, targetDate) {
  const sheet = ss.getSheetByName(WEEKLY_UPDATE_SHEET_NAME)
    || ss.getSheets().find(s => s.getName().toLowerCase() === WEEKLY_UPDATE_SHEET_NAME.toLowerCase());
  if (!sheet) return { rows: [], remark: '' };

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return { rows: [], remark: '' };

  const storeMap  = {};
  let lastRemark  = '';

  for (let i = 1; i < values.length; i++) {
    const r       = values[i];
    const rDate   = normalizeDate_(r[1]);
    const rVendor = String(r[7]).trim();
    const rType   = String(r[8]).trim();
    if (rVendor !== vendor || rType !== type || rDate !== targetDate) continue;

    const storeNo = padStoreNo_(r[2]);
    const pos     = String(r[9]).trim();
    const qty     = type === 'Interview' ? Number(r[10]) : Number(r[11]);
    if (r[12]) lastRemark = String(r[12]);
    if (!storeMap[storeNo]) storeMap[storeNo] = { lam: 0, bak: 0, oth: 0 };
    if (pos === 'ล่าม')         storeMap[storeNo].lam += qty;
    else if (pos === 'Bakery')  storeMap[storeNo].bak += qty;
    else                        storeMap[storeNo].oth += qty;
  }

  return {
    rows: Object.keys(storeMap).map(s => ({ storeNo: s, ...storeMap[s] })),
    remark: lastRemark
  };
}

// ─────────────────────────────────────────────
//  SUBMIT DATA
// ─────────────────────────────────────────────
function submitCombinedData(payload) {
  const ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
  const currentUser = loginUserOnly(payload.userEmail).user;
  const vendor      = currentUser.vendor;
  const targets     = getTargets_(ss);
  const timestamp   = new Date();

  const sheet = ss.getSheetByName(WEEKLY_UPDATE_SHEET_NAME)
    || ss.getSheets().find(s => s.getName().toLowerCase() === WEEKLY_UPDATE_SHEET_NAME.toLowerCase());

  const intDate  = payload.int  && payload.int.date  ? normalizeDate_(payload.int.date)  : null;
  const workDate = payload.work && payload.work.date ? normalizeDate_(payload.work.date) : null;

  // Delete existing rows for same vendor+date+type (upsert)
  const values = sheet.getDataRange().getValues();
  for (let i = values.length - 1; i >= 1; i--) {
    const rDate   = normalizeDate_(values[i][1]);
    const rVendor = String(values[i][7]).trim();
    const rType   = String(values[i][8]).trim();
    const del = (payload.int  && intDate  && rVendor === vendor && rType === 'Interview' && rDate === intDate)
             || (payload.work && workDate && rVendor === vendor && rType === 'StartWork' && rDate === workDate);
    if (del) sheet.deleteRow(i + 1);
  }

  // Save uploaded file to Drive
  if (payload.uploadFile && payload.uploadFile.data) {
    try {
      const folder   = DriveApp.getFolderById('1kEkuRP2QMs_Pws4hIGEJs6tYyrlmGiHG');
      const fileBlob = Utilities.newBlob(
        Utilities.base64Decode(payload.uploadFile.data),
        payload.uploadFile.type,
        payload.uploadFile.name
      );
      fileBlob.setName(vendor + '_' + Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyyMMdd') + '_' + payload.uploadFile.name);
      folder.createFile(fileBlob);
    } catch (e) {}
  }

  const rowsToAppend = [];

  function processRows(dataObj, type) {
    if (!dataObj || !dataObj.rows || dataObj.rows.length === 0) return;
    const actionDate = normalizeDate_(dataObj.date);
    dataObj.rows.forEach(r => {
      const storeNo   = padStoreNo_(r.storeNo);
      const targetRow = targets.find(t => t.storeNo === storeNo);
      if (!targetRow || !storeNo) return;
      [{ pos: 'ล่าม', qty: r.lam }, { pos: 'Bakery', qty: r.bak }, { pos: 'Other', qty: r.oth }].forEach(p => {
        if (p.qty > 0) {
          rowsToAppend.push([
            timestamp, actionDate, storeNo, targetRow.branch, targetRow.branchTH,
            targetRow.region, targetRow.mainVendor, vendor, type,
            p.pos,
            type === 'Interview'  ? p.qty : 0,
            type === 'StartWork'  ? p.qty : 0,
            dataObj.remark || '', currentUser.email
          ]);
        }
      });
    });
  }

  if (payload.int)  processRows(payload.int,  'Interview');
  if (payload.work) processRows(payload.work, 'StartWork');

  if (rowsToAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
  }
  return { success: true };
}

// ─────────────────────────────────────────────
//  DASHBOARD REFRESH (filter apply)
// ─────────────────────────────────────────────
function getDashboardDataAPI(email, filterStart, filterEnd, forceRefresh) {
  const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  const userObj = loginUserOnly(email).user;
  const targets = getTargets_(ss);
  const extDB   = getExternalDBData_(forceRefresh === true);
  return getDashboardDataCore_(ss, userObj, targets, extDB, filterStart, filterEnd);
}

// ─────────────────────────────────────────────
//  EXPORT
// ─────────────────────────────────────────────
function exportAdminExcelAPI(userEmail, filterStart, filterEnd) {
  const data       = getDashboardDataAPI(userEmail, filterStart, filterEnd, false);
  const exportFile = SpreadsheetApp.create('AEC_Admin_Report_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss'));

  const s1 = exportFile.getActiveSheet();
  s1.setName('All Branches Summary');
  const h1 = ['Store No','Branch','Region','Main Vendor','Target','Passed','Started','Remaining','Progress','Last Interview Date','Last Start Date'];
  const r1 = data.allBranchRows.map(r => [r.storeNo,r.branch,r.region,r.mainVendor,r.totalTarget,r.interviewPassed,r.startedWork,r.remaining,r.progress,r.lastIntDate,r.lastStartDate]);
  s1.getRange(1,1,1,h1.length).setValues([h1]);
  if (r1.length > 0) s1.getRange(2,1,r1.length,h1.length).setValues(r1);

  if (data.regionRows && data.regionRows.length > 0) {
    const s2 = exportFile.insertSheet('Region Summary');
    const h2 = ['Region','Target','Passed','Started','Remaining','Progress'];
    const r2 = data.regionRows.map(r => [r.region,r.target,r.passed,r.started,r.remaining,r.progress]);
    s2.getRange(1,1,1,h2.length).setValues([h2]);
    s2.getRange(2,1,r2.length,h2.length).setValues(r2);
  }

  if (data.vendorRows && data.vendorRows.length > 0) {
    const s3 = exportFile.insertSheet('Vendor Performance');
    const h3 = ['Vendor','Interview_Lam','Interview_Bakery','Interview_Other','Interview_Total','Started_Lam','Started_Bakery','Started_Other','Started_Total'];
    const r3 = data.vendorRows.map(v => [v.vendor,v.intLam,v.intBak,v.intOth,v.intTot,v.wrkLam,v.wrkBak,v.wrkOth,v.wrkTot]);
    s3.getRange(1,1,1,h3.length).setValues([h3]);
    s3.getRange(2,1,r3.length,h3.length).setValues(r3);
  }

  if (data.rawUpdates && data.rawUpdates.length > 0) {
    const s4 = exportFile.insertSheet('Raw Data Logs');
    const h4 = ['Action Date','Store No','Branch','Region','Update Vendor','Type','Position','Interview Passed','Started Work','Remark'];
    const r4 = data.rawUpdates.map(r => [
      normalizeDate_(r['action date']), r['store no'], r['branch'], r['region'], r['update vendor'],
      r['type'], r['position'], Number(r['interview passed']||0), Number(r['started work']||0), String(r['remark']||'')
    ]);
    s4.getRange(1,1,1,h4.length).setValues([h4]);
    s4.getRange(2,1,r4.length,h4.length).setValues(r4);
  }

  try { DriveApp.getFileById(exportFile.getId()).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
  return `https://docs.google.com/spreadsheets/d/${exportFile.getId()}/export?format=xlsx`;
}

function exportExcelAPI(userEmail) {
  const data       = getDashboardDataAPI(userEmail, '', '', false);
  const exportFile = SpreadsheetApp.create('AEC_Export_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss'));
  const s1         = exportFile.getActiveSheet();
  s1.setName('All Branches');
  const h1 = ['Store No','Branch','Region','Main Vendor','Target','Passed','Started','Remaining','Progress'];
  const r1 = data.allBranchRows.map(r => [r.storeNo,r.branch,r.region,r.mainVendor,r.totalTarget,r.interviewPassed,r.startedWork,r.remaining,r.progress]);
  s1.getRange(1,1,1,h1.length).setValues([h1]);
  if (r1.length > 0) s1.getRange(2,1,r1.length,h1.length).setValues(r1);
  try { DriveApp.getFileById(exportFile.getId()).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
  return `https://docs.google.com/spreadsheets/d/${exportFile.getId()}/export?format=xlsx`;
}
