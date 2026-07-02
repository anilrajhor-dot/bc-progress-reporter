// ═══════════════════════════════════════════════════════════════
//  BC PROGRESS REPORTER — Google Apps Script Backend
//  Paste this ENTIRE file into Google Apps Script editor
//  Then Deploy → New Deployment → Web App → Anyone
// ═══════════════════════════════════════════════════════════════

const SHEET_NAME_REPORTS  = 'Reports';
const SHEET_NAME_USERS    = 'Users';
const SHEET_NAME_PARAMS   = 'Parameters';
const SHEET_NAME_TEMPLATES= 'Templates';

// ── MAIN ENTRY POINT ──
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    let result;

    if      (action === 'login')         result = handleLogin(body);
    else if (action === 'submitReport')  result = handleSubmitReport(body);
    else if (action === 'getReports')    result = handleGetReports(body);
    else if (action === 'getUsers')      result = handleGetUsers(body);
    else if (action === 'addUser')       result = handleAddUser(body);
    else if (action === 'getParams')     result = handleGetParams(body);
    else if (action === 'saveParams')    result = handleSaveParams(body);
    else if (action === 'getTemplates')  result = handleGetTemplates(body);
    else if (action === 'saveTemplates') result = handleSaveTemplates(body);
    else result = { error: 'Unknown action' };

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'BC Progress Reporter API running' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── SHEET HELPERS ──
function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    initSheet(sheet, name);
  }
  return sheet;
}

function initSheet(sheet, name) {
  if (name === SHEET_NAME_USERS) {
    sheet.appendRow(['id','name','role','pass','mobile','email','district','location','createdAt']);
    // Seed default users
    const now = new Date().toISOString();
    sheet.appendRow(['ADMIN','Admin','ADMIN','admin123','9999900000','admin@corporatebc.in','','HO',now]);
    sheet.appendRow(['DC01','Anand Mishra','DC','dc123','9811100001','anand@bc.in','Varanasi','Varanasi HQ',now]);
    sheet.appendRow(['DC02','Kavita Sharma','DC','dc123','9811100002','kavita@bc.in','Mirzapur','Mirzapur HQ',now]);
    sheet.appendRow(['BC001','Ramesh Kumar','BC','pass123','9876543210','ramesh@bc.in','Varanasi','Sarnath',now]);
    sheet.appendRow(['BC002','Sunita Devi','BC','pass123','9876543211','sunita@bc.in','Varanasi','Ramnagar',now]);
    sheet.appendRow(['BC003','Priya Singh','BC','pass123','9876543212','priya@bc.in','Mirzapur','Chunar',now]);
  }
  if (name === SHEET_NAME_REPORTS) {
    sheet.appendRow(['bcId','bcName','bcDistrict','date','submittedAt','remarks',
                     'accounts','pmjjby','pmsby','apy','rekyc',
                     'p6','p7','p8','p9','p10']); // extra param cols
    sheet.getRange(1,1,1,16).setFontWeight('bold').setBackground('#1A3A5C').setFontColor('#FFFFFF');
  }
  if (name === SHEET_NAME_PARAMS) {
    sheet.appendRow(['id','label','icon','active']);
    sheet.appendRow(['accounts','Accounts Opened','🏦','TRUE']);
    sheet.appendRow(['pmjjby','PMJJBY Enrolled','🛡️','TRUE']);
    sheet.appendRow(['pmsby','PMSBY Enrolled','⛑️','TRUE']);
    sheet.appendRow(['apy','APY Enrolled','🏛️','TRUE']);
    sheet.appendRow(['rekyc','Re-KYC Done','🪪','TRUE']);
  }
  if (name === SHEET_NAME_TEMPLATES) {
    sheet.appendRow(['key','value']);
    sheet.appendRow(['wa','Dear {name} (ID: {id}), your daily progress report for {date} is pending. Please submit. — Corporate BC']);
    sheet.appendRow(['email','Dear {name},\n\nYour daily report for {date} is pending. Please login and submit.\n\nRegards,\nCorporate BC Management']);
  }
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

// ── LOGIN ──
function handleLogin(body) {
  const sheet = getOrCreateSheet(SHEET_NAME_USERS);
  const users = sheetToObjects(sheet);
  const user = users.find(u => u.id === body.id && u.pass === body.pass);
  if (!user) return { error: 'Invalid ID or Password' };
  const { pass, ...safeUser } = user; // don't send password back
  return { success: true, user: safeUser };
}

// ── SUBMIT REPORT ──
function handleSubmitReport(body) {
  const sheet = getOrCreateSheet(SHEET_NAME_REPORTS);
  const params = handleGetParams({}).params;
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Find existing row for this BC + date
  let existingRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === body.bcId && String(data[i][3]) === String(body.date)) {
      existingRow = i + 1; // 1-indexed
      break;
    }
  }

  const rowData = [
    body.bcId, body.bcName || '', body.bcDistrict || '',
    body.date, new Date().toISOString(), body.remarks || '',
    ...params.map(p => body.data[p.id] || 0),
    ...Array(10 - params.length).fill(0) // pad to 10 param cols
  ];

  if (existingRow > 0) {
    sheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }

  // Apply alternating row colors for readability
  const lastRow = sheet.getLastRow();
  if (lastRow % 2 === 0) {
    sheet.getRange(lastRow, 1, 1, 16).setBackground('#F1F5F9');
  }

  return { success: true };
}

// ── GET REPORTS ──
function handleGetReports(body) {
  const sheet = getOrCreateSheet(SHEET_NAME_REPORTS);
  const params = handleGetParams({}).params;
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { reports: [] };
  const headers = data[0];
  let reports = data.slice(1).map(row => {
    const paramData = {};
    params.forEach((p, i) => { paramData[p.id] = row[6 + i] || 0; });
    return {
      bcId: row[0], bcName: row[1], bcDistrict: row[2],
      date: row[3], submittedAt: row[4], remarks: row[5],
      data: paramData
    };
  });

  // Filters
  if (body.date)     reports = reports.filter(r => String(r.date) === String(body.date));
  if (body.bcId)     reports = reports.filter(r => r.bcId === body.bcId);
  if (body.district) reports = reports.filter(r => r.bcDistrict === body.district);

  return { reports };
}

// ── GET USERS ──
function handleGetUsers(body) {
  const sheet = getOrCreateSheet(SHEET_NAME_USERS);
  const users = sheetToObjects(sheet).map(u => { const {pass,...s}=u; return s; });
  let filtered = users;
  if (body.role) filtered = filtered.filter(u => u.role === body.role);
  if (body.district) filtered = filtered.filter(u => u.district === body.district);
  return { users: filtered };
}

// ── ADD USER ──
function handleAddUser(body) {
  const sheet = getOrCreateSheet(SHEET_NAME_USERS);
  const users = sheetToObjects(sheet);
  if (users.find(u => u.id === body.user.id)) return { error: 'ID already exists' };
  const u = body.user;
  sheet.appendRow([u.id, u.name, u.role, u.pass||'pass123', u.mobile||'', u.email||'', u.district||'', u.location||'', new Date().toISOString()]);
  return { success: true };
}

// ── PARAMS ──
function handleGetParams(body) {
  const sheet = getOrCreateSheet(SHEET_NAME_PARAMS);
  const rows = sheetToObjects(sheet).filter(r => String(r.active) === 'TRUE');
  return { params: rows };
}

function handleSaveParams(body) {
  const sheet = getOrCreateSheet(SHEET_NAME_PARAMS);
  sheet.clearContents();
  sheet.appendRow(['id','label','icon','active']);
  body.params.forEach(p => sheet.appendRow([p.id, p.label, p.icon, 'TRUE']));
  return { success: true };
}

// ── TEMPLATES ──
function handleGetTemplates(body) {
  const sheet = getOrCreateSheet(SHEET_NAME_TEMPLATES);
  const rows = sheetToObjects(sheet);
  const tpl = {};
  rows.forEach(r => { tpl[r.key] = r.value; });
  return { templates: tpl };
}

function handleSaveTemplates(body) {
  const sheet = getOrCreateSheet(SHEET_NAME_TEMPLATES);
  sheet.clearContents();
  sheet.appendRow(['key','value']);
  Object.entries(body.templates).forEach(([k,v]) => sheet.appendRow([k,v]));
  return { success: true };
}

// ── DAILY SUMMARY EMAIL (Optional — run via Time Trigger) ──
function sendDailySummaryEmail() {
  const adminSheet = getOrCreateSheet(SHEET_NAME_USERS);
  const admins = sheetToObjects(adminSheet).filter(u => u.role === 'ADMIN' && u.email);
  const bcSheet = sheetToObjects(getOrCreateSheet(SHEET_NAME_USERS)).filter(u => u.role === 'BC');
  const today = new Date().toISOString().split('T')[0];
  const reports = handleGetReports({ date: today }).reports;
  const submitted = reports.length;
  const pending = bcSheet.length - submitted;

  const subject = `BC Progress Summary — ${today}`;
  const body = `Daily Report Summary for ${today}\n\nTotal BCs: ${bcSheet.length}\nSubmitted: ${submitted}\nPending: ${pending}\nSubmission Rate: ${Math.round(submitted/bcSheet.length*100)}%\n\nLogin to view details.`;

  admins.forEach(admin => {
    try { MailApp.sendEmail(admin.email, subject, body); } catch(e) {}
  });
}
