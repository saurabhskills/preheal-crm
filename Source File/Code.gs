/**
 * SERVE HTML
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Preheal CRM')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * DATABASE CONNECTION HELPER
 */
function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

/**
 * USERS: Email | Password | Role | Status | Name
 */
function loginUser(email, password) {
  const sheet = getSheet('Users');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] == email && row[1] == password) {
      const status = row[3];
      if (status === 'PENDING') {
        return { success: false, type: 'warning', message: 'Your account is pending admin approval.' };
      }
      if (status === 'INACTIVE') {
        return { success: false, message: 'Account deactivated. Contact admin.' };
      }
      if (status === 'REJECTED') {
        return { success: false, type: 'error', message: 'Your account was rejected. Contact admin.' };
      }
      if (status === 'APPROVED') {
        return {
          success: true,
          email: row[0],
          role: row[2],
          name: row[4]
        };
      }
      return { success: false, type: 'error', message: 'Account inactive. Contact admin.' };
    }
  }
  return { success: false, type: 'error', message: 'Invalid email or password.' };
}

function registerUser(form) {
  const sheet = getSheet('Users');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === form.email) {
      return { success: false, type: 'error', message: 'User already exists.' };
    }
  }

  // Email | Password | Role | Status | Name
  sheet.appendRow([form.email, form.password, form.role, 'PENDING', form.name]);
  return { success: true, type: 'info', message: 'Registered. Waiting for admin approval.' };
}

/**
 *  LeadID generation logic:
 */
function leadDateKeyDDMMYY_() {
  const tz = Session.getScriptTimeZone();
  return Utilities.formatDate(new Date(), tz, 'ddMMyy'); // DDMMYY [web:87]
}

function nextLeadIds_(count) {
  const n = Number(count || 1);
  if (n <= 0) return [];

  const lock = LockService.getDocumentLock();
  lock.waitLock(20000); // throws if can't acquire [web:151]

  try {
    const props = PropertiesService.getScriptProperties(); // shared store [web:180]
    const keyDate = leadDateKeyDDMMYY_();

    const lastDate = props.getProperty('LEAD_SEQ_DATE') || '';
    let seq = Number(props.getProperty('LEAD_SEQ_NUM') || '0');

    if (lastDate !== keyDate) {
      seq = 0; // reset counter each day
    }

    if (seq + n > 999) {
      throw new Error('Daily LeadID limit exceeded (999). Please contact admin.');
    }

    const ids = [];
    for (let i = 1; i <= n; i++) {
      const num = seq + i;
      const xxx = String(num).padStart(3, '0');
      ids.push(`L0${keyDate}0${xxx}`); // L0DDMMYY0XXX
    }

    props.setProperty('LEAD_SEQ_DATE', keyDate);
    props.setProperty('LEAD_SEQ_NUM', String(seq + n));

    SpreadsheetApp.flush(); // commit while holding lock (recommended) [web:151]
    return ids;
  } finally {
    lock.releaseLock();
  }
}



/**
 * ADMIN: user approval
 */
function getPendingUsers() {
  const sheet = getSheet('Users');
  const data = sheet.getDataRange().getValues();
  const pending = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][3] === 'PENDING') {
      pending.push({
        row: i + 1,
        email: data[i][0],
        role: data[i][2],
        name: data[i][4]
      });
    }
  }
  return pending;
}

function approveUser(rowNumber) {
  const sheet = getSheet('Users');
  sheet.getRange(rowNumber, 4).setValue('APPROVED'); // Status column (1-based)
  return { success: true };
}

// Telecaller Approval by Executive

function isTeleCallerRole_(role) {
  const r = String(role || '').toUpperCase().replace(/[^A-Z]/g, '');
  return r === 'TELECALLER';
}

function getUserByEmail_(email) {
  const sheet = getSheet('Users');
  if (!sheet) return null;
  const values = sheet.getDataRange().getValues(); // Email | Password | Role | Status | Name
  const emailNorm = String(email || '').trim().toLowerCase();

  for (let i = 1; i < values.length; i++) {
    const e = String(values[i][0] || '').trim().toLowerCase();
    if (e === emailNorm) {
      return { row: i + 1, email: values[i][0], role: values[i][2], status: values[i][3], name: values[i][4] };
    }
  }
  return null;
}


function getApprovedUsersForAdmin() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Users');
  const data = sh.getDataRange().getValues();
  const headers = data.shift().map(h => String(h).trim());

  const H = name => headers.indexOf(name);

  const out = [];
  data.forEach((r, i) => {
    const status = String(r[H('Status')] || '').toUpperCase();
    if (status === 'APPROVED') {
      out.push({
        row: i + 2,
        Email: r[H('Email')],
        Name: r[H('Name')],
        Role: r[H('Role')],
        Status: status
      });
    }
  });

  return out;
}


function deactivateApprovedUserByAdmin(row, adminEmail, reason) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Users');

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
    .getValues()[0].map(h => String(h).trim());
  const H = name => headers.indexOf(name) + 1;

  // Block: can't deactivate Admin users
  const role = String(sh.getRange(row, H('Role')).getValue() || '').trim();
  if (role === 'Admin') {
    return { success: false, message: 'Admin account cannot be deactivated.' };
  }

  // Block: can't deactivate self (recommended)
  const targetEmail = String(sh.getRange(row, H('Email')).getValue() || '').trim().toLowerCase();
  if (targetEmail === String(adminEmail || '').trim().toLowerCase()) {
    return { success: false, message: 'You cannot deactivate your own account.' };
  }

  const statusCell = sh.getRange(row, H('Status'));
  const currentStatus = String(statusCell.getValue() || '').toUpperCase();
  if (currentStatus !== 'APPROVED') {
    return { success: false, message: 'Only APPROVED users can be deactivated.' };
  }

  statusCell.setValue('INACTIVE');

  // If you DID NOT add Deactivated* columns, store the audit in Decision* columns instead
  if (H('DeactivatedAt') > 0) {
    sh.getRange(row, H('DeactivatedAt')).setValue(new Date());
    if (H('DeactivatedBy') > 0) sh.getRange(row, H('DeactivatedBy')).setValue(adminEmail);
    if (H('DeactivationReason') > 0) sh.getRange(row, H('DeactivationReason')).setValue(reason || '');
  } else {
    // Fallback to existing columns you already have
    if (H('DecisionAt') > 0) sh.getRange(row, H('DecisionAt')).setValue(new Date());
    if (H('DecisionBy') > 0) sh.getRange(row, H('DecisionBy')).setValue(adminEmail);
    if (H('DecisionReason') > 0) sh.getRange(row, H('DecisionReason')).setValue('DEACTIVATED: ' + (reason || ''));
  }

  return { success: true, message: 'User deactivated (history preserved).' };
}


/**
  Deactive users by EXECUTIVE:

*/
function getApprovedTeleCallersForExecutive(execEmail) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Users');
  const data = sh.getDataRange().getValues();
  const headers = data.shift().map(h => String(h).trim());
  const H = name => headers.indexOf(name);

  const out = [];
  data.forEach((r, i) => {
    const role = String(r[H('Role')] || '').trim();
    const status = String(r[H('Status')] || '').toUpperCase();

    if (role === 'TeleCaller' && status === 'APPROVED') {
      out.push({
        row: i + 2,
        Email: r[H('Email')],
        Name: r[H('Name')],
        Role: role,
        Status: status
      });
    }
  });

  return out;
}

function deactivateApprovedTeleCallerByExecutive(row, execEmail, reason) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Users');
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const H = name => headers.indexOf(name) + 1;

  const role = String(sh.getRange(row, H('Role')).getValue() || '').trim();
  const statusCell = sh.getRange(row, H('Status'));
  const status = String(statusCell.getValue() || '').toUpperCase();

  if (role !== 'TeleCaller') return { success: false, message: 'Only TeleCaller can be deactivated here.' };
  if (status !== 'APPROVED') return { success: false, message: 'Only APPROVED TeleCallers can be deactivated.' };

  statusCell.setValue('INACTIVE');

  // Audit (re-using Decision columns)
  sh.getRange(row, H('DecisionAt')).setValue(new Date());
  sh.getRange(row, H('DecisionBy')).setValue(execEmail);
  sh.getRange(row, H('DecisionReason')).setValue(reason || '');

  return { success: true, message: 'TeleCaller deactivated (history preserved).' };
}



/**
 * EXECUTIVE: view only pending TeleCallers
 */
function getPendingTeleCallersForExecutive(executiveEmail) {
  const exec = getUserByEmail_(executiveEmail);
  if (!exec || String(exec.status).toUpperCase() !== 'APPROVED' || String(exec.role).toUpperCase() !== 'EXECUTIVE') {
    return []; // or throw new Error('Not allowed');
  }

  const sheet = getSheet('Users');
  const data = sheet.getDataRange().getValues();
  const pending = [];

  for (let i = 1; i < data.length; i++) {
    const role = data[i][2];
    const status = String(data[i][3] || '').toUpperCase();
    if (status === 'PENDING' && isTeleCallerRole_(role)) {
      pending.push({
        row: i + 1,
        email: data[i][0],
        role: data[i][2],
        name: data[i][4]
      });
    }
  }
  return pending;
}

/**
 * EXECUTIVE: approve only pending TeleCallers
 */
function approveTeleCallerByExecutive(rowNumber, executiveEmail) {
  const exec = getUserByEmail_(executiveEmail);
  if (!exec || String(exec.status).toUpperCase() !== 'APPROVED' || String(exec.role).toUpperCase() !== 'EXECUTIVE') {
    return { success: false, message: 'Not allowed.' };
  }

  const sheet = getSheet('Users');
  const role = sheet.getRange(rowNumber, 3).getValue();   // Role col
  const status = String(sheet.getRange(rowNumber, 4).getValue() || '').toUpperCase(); // Status col

  if (!isTeleCallerRole_(role)) return { success: false, message: 'Only TeleCallers can be approved here.' };
  if (status !== 'PENDING') return { success: false, message: 'User is not pending.' };

  sheet.getRange(rowNumber, 4).setValue('APPROVED');
  return { success: true, message: 'TeleCaller approved.' };
}


/**
 * LEADS SHEET HEADERS (A..M):
 * 0 LeadID | 1 LeadName | 2 Phone | 3 LeadSource | 4 AssignedTo | 5 CallNote | 6 Issue
 * 7 FollowUpDate | 8 FollowUpStatus | 9 LeadStatus | 10 CreatedBy | 11 CreatedDate | 12 LastUpdatedDate
 */

/**
 * Add a lead (AssignedTo stays empty because your intake form doesn't include it)
 * FollowUpDate stored as string DD/MM/YYYY (coming from UI conversion).
 */
function addLead(form, createdBy) {
  const sheet = getSheet('Leads');
  const id = nextLeadIds_(1)[0];
  const now = new Date();

  sheet.appendRow([
    id,                    // LeadID
    form.name || '',       // LeadName
    form.phone || '',      // Phone
    form.source || '',     // LeadSource
    form.assignedTo || '', // AssignedTo
    form.callNotes || '',  // CallNote
    form.callNotes || '',  // Issue (kept same so your UI "Issue" column shows something)
    form.followUpDate || '', // FollowUpDate (DD/MM/YYYY string)
    'Pending',             // FollowUpStatus
    'New',                 // LeadStatus
    createdBy || '',       // CreatedBy
    now,                   // CreatedDate
    now                    // LastUpdatedDate
  ]);

  return { success: true, type: 'success', message: 'Lead created.' };
}

/**
 * Bulk import - FollowUpDate expected as DD/MM/YYYY (string)
 */
function bulkUploadLeads(dataArray, createdBy) {
  const sheet = getSheet('Leads');
  const now = new Date();

  // 1) Keep only valid rows
  const cleaned = (dataArray || []).filter(r => r && r.name);
  const count = cleaned.length;

  if (!count) {
    return { success: true, type: 'success', count: 0, message: '0 leads imported.' };
  }

  // 2) Generate sequential LeadIDs: L0DDMMYY0XXX
  const ids = nextLeadIds_(count);

  // 3) Build rows for a single batch write
  const rows = [];
  for (let i = 0; i < cleaned.length; i++) {
    const row = cleaned[i];
    const id = ids[i];

    rows.push([
      id,
      row.name || '',
      row.phone || '',
      row.source || 'Bulk Import',
      '', // AssignedTo
      row.callNotes || '', // CallNote
      row.callNotes || '', // Issue
      row.followUpDate || '', // FollowUpDate (DD/MM/YYYY string)
      'Pending',
      'New',
      createdBy || '',
      now,
      now
    ]);
  }

  // 4) Write to sheet in one call
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows); // batch write [web:191]

  return { success: true, type: 'success', count, message: count + ' leads imported.' };
}

function importLeadsFromExcel(fileObj, createdBy) {
  if (!fileObj || !fileObj.base64) {
    return { success: false, type: 'error', message: 'No file data received.' };
  }

  const bytes = Utilities.base64Decode(fileObj.base64);
  const blob = Utilities.newBlob(bytes, fileObj.mimeType || 'application/octet-stream', fileObj.fileName || 'upload');

  // Upload + convert to Google Sheet (Advanced Drive API v2)
  const inserted = Drive.Files.insert(
    { title: fileObj.fileName || ('BulkImport_' + Date.now()), mimeType: blob.getContentType() },
    blob,
    { convert: true }
  ); // convert parameter is supported by files.insert [web:261]

  const tempSs = SpreadsheetApp.openById(inserted.id);
  const srcSheet = tempSs.getSheets()[0];
  const values = srcSheet.getDataRange().getValues();

  if (!values || values.length < 2) {
    DriveApp.getFileById(inserted.id).setTrashed(true);
    return { success: false, type: 'warning', message: 'No data rows found in uploaded file.' };
  }

  // Build header map
  const header = values[0].map(h => String(h || '').trim().toLowerCase());
  const idx = (name) => header.indexOf(String(name).toLowerCase());

  const iName = idx('name');
  const iPhone = idx('phone');
  const iSource = idx('source');
  const iNotes = idx('callnotes');
  const iFollowUp = idx('followupdate');
  const iAssigned = idx('assignedto');

  if (iName === -1) {
    DriveApp.getFileById(inserted.id).setTrashed(true);
    return { success: false, type: 'error', message: 'Column "Name" is required in the header row.' };
  }

  // Prepare rows
  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const name = String(row[iName] || '').trim();
    if (!name) continue;

    const phone = iPhone >= 0 ? String(row[iPhone] || '').trim() : '';
    const source = iSource >= 0 ? String(row[iSource] || '').trim() : 'Bulk Import';
    const notes = iNotes >= 0 ? String(row[iNotes] || '').trim() : '';
    const followUpRaw = iFollowUp >= 0 ? row[iFollowUp] : '';
    const assignedTo = iAssigned >= 0 ? String(row[iAssigned] || '').trim() : '';

    const followUp = normalizeFollowUpDate_(followUpRaw); // you already added this earlier

    rows.push({
      name, phone, source,
      callNotes: notes,
      followUpDate: followUp,
      assignedTo
    });
  }

  if (!rows.length) {
    DriveApp.getFileById(inserted.id).setTrashed(true);
    return { success: false, type: 'warning', message: 'No valid leads found (blank names).' };
  }

  // Generate LeadIDs using your L0DDMMYY0XXX logic
  const ids = nextLeadIds_(rows.length);

  const leadsSheet = getSheet('Leads');
  const now = new Date();

  const out = rows.map((r, k) => ([
    ids[k],
    r.name,
    r.phone,
    r.source,
    r.assignedTo || '', // ✅ AssignedTo (E)
    r.callNotes || '',
    r.callNotes || '',
    r.followUpDate || '',
    'Pending',
    'New',
    createdBy || '',
    now,
    now
  ]));

  const startRow = leadsSheet.getLastRow() + 1;
  leadsSheet.getRange(startRow, 1, out.length, out[0].length).setValues(out); // batch write [web:191]

  // Cleanup temp converted sheet
  DriveApp.getFileById(inserted.id).setTrashed(true);

  return { success: true, type: 'success', count: out.length, message: out.length + ' leads imported from file.' };
}


/**
 * Update lead (your existing function kept for backward compatibility)
 */
function updateLead(update) {
  const sheet = getSheet('Leads');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === update.leadId) {
      const rowIndex = i + 1;

      if (update.assignedTo !== undefined) {
        sheet.getRange(rowIndex, 5).setValue(update.assignedTo); // AssignedTo (E)
      }
      if (update.callNotes !== undefined) {
        sheet.getRange(rowIndex, 6).setValue(update.callNotes); // CallNote (F)
      }
      if (update.issue !== undefined) {
        sheet.getRange(rowIndex, 7).setValue(update.issue); // Issue (G)
      }
      if (update.followUpDate !== undefined) {
        sheet.getRange(rowIndex, 8).setValue(update.followUpDate); // FollowUpDate (H) string
      }
      if (update.followUpStatus !== undefined) {
        sheet.getRange(rowIndex, 9).setValue(update.followUpStatus); // FollowUpStatus (I)
      }
      if (update.leadStatus !== undefined) {
        sheet.getRange(rowIndex, 10).setValue(update.leadStatus); // LeadStatus (J)
      }

      sheet.getRange(rowIndex, 13).setValue(new Date()); // LastUpdatedDate (M)
      return { success: true, message: 'Lead updated.' };
    }
  }

  return { success: false, message: 'Lead not found.' };
}

/**
 * Helper: build role lookup from Users sheet
 */
function getApprovedUserInfoMap_() {
  const sheet = getSheet('Users');
  if (!sheet) return {};

  const values = sheet.getDataRange().getValues(); // Email | Password | Role | Status | Name
  const map = {};

  for (let i = 1; i < values.length; i++) {
    const email = values[i][0];
    const role = values[i][2];
    const status = values[i][3];
    const name = values[i][4];
    if (email && status === 'APPROVED') {
      map[String(email).trim().toLowerCase()] = { role: String(role || ''), name: String(name || '') };
    }
  }
  return map;
}

/**
 * GET LEADS
 * Uses getDisplayValues to return strings (safe for HTML)
 */
function getLeads(role, userEmail) {
  const sheet = getSheet('Leads');
  const data = sheet.getDataRange().getDisplayValues();

  const roleNorm = String(role || '').toUpperCase();
  const emailNorm = String(userEmail || '').trim().toLowerCase();

  const result = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue; // skip empty rows

    const assignedTo = String(row[4] || '').trim();
    const createdBy = String(row[10] || '').trim();

    // Filtering:
    // - Admin/Manager: all leads
    // - Executive/TeleCaller: leads assigned to them OR created by them
    if (roleNorm === 'EXECUTIVE' || roleNorm === 'TELECALLER') {
      const a = assignedTo.toLowerCase();
      const c = createdBy.toLowerCase();
      if (a !== emailNorm && c !== emailNorm) continue;
    }

    result.push({
      id: row[0],
      name: row[1],
      phone: row[2],
      source: row[3],
      assignedTo: assignedTo || 'Unassigned',
      callNotes: row[5],
      issue: row[6],
      followUpDate: row[7],
      followUpStatus: row[8],
      status: row[9],
      createdBy: row[10],
      createdDate: row[11],
      lastUpdatedDate: row[12]
    });
  }

  return result.reverse();
}

/**
 * ADMIN DASHBOARD DATA
 */
function getAdminDashboard() {
  const leadsSheet = getSheet('Leads');

  if (!leadsSheet) {
    return { totalLeads: 0, leadsByStatus: {}, executivePerformance: {}, teleCallerPerformance: {} };
  }

  const userMap = getApprovedUserInfoMap_();
  const leadData = leadsSheet.getDataRange().getValues();

  let totalLeads = 0;
  const leadsByStatus = {};
  const executivePerformance = {};
  const teleCallerPerformance = {};

  for (let i = 1; i < leadData.length; i++) {
    const row = leadData[i];
    if (!row[0]) continue;

    totalLeads++;

    const status = row[9] ? String(row[9]).trim() : 'New';
    leadsByStatus[status] = (leadsByStatus[status] || 0) + 1;

    const isConverted = (status === 'Converted' || status === 'Closed Won');

    // Executive performance: CreatedBy
    const createdBy = row[10] ? String(row[10]).trim().toLowerCase() : '';
    if (createdBy) {
      const info = userMap[createdBy] || { role: '', name: '' };
      const userName = String(info.name || '').trim();

      if (!executivePerformance[createdBy]) {
        executivePerformance[createdBy] = { total: 0, converted: 0, name: userName };
      }
      executivePerformance[createdBy].total++;
      if (isConverted) executivePerformance[createdBy].converted++;
    }

    // Telecaller performance: AssignedTo
    const assignedTo = row[4] ? String(row[4]).trim().toLowerCase() : '';
    if (assignedTo) {
      const info = userMap[assignedTo] || { role: '', name: '' };
      const userRole = String(info.role || '').toUpperCase();
      const userName = String(info.name || '').trim();

      if (userRole === 'TELECALLER' || userRole === 'TELE-CALLER' || userRole === 'TELE CALLER') {
        if (!teleCallerPerformance[assignedTo]) {
          teleCallerPerformance[assignedTo] = { total: 0, converted: 0, name: userName };
        }
        teleCallerPerformance[assignedTo].total++;
        if (isConverted) teleCallerPerformance[assignedTo].converted++;
      }
    }
  }

  return { totalLeads, leadsByStatus, executivePerformance, teleCallerPerformance };
}

/** =========================
 *  HISTORY (NEW)
 *  ========================= */
function ensureSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
  } else if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
  }
  return sh;
}

function ensureLeadsHistorySheet_() {
  const headers = [
    'LoggedAt', 'UpdatedBy', 'LeadID', 'LeadRowNumber', 'ChangeJSON',
    // Snapshot BEFORE update (same order as Leads sheet)
    'LeadID_snap', 'LeadName_snap', 'Phone_snap', 'LeadSource_snap', 'AssignedTo_snap',
    'CallNote_snap', 'Issue_snap', 'FollowUpDate_snap', 'FollowUpStatus_snap', 'LeadStatus_snap',
    'CreatedBy_snap', 'CreatedDate_snap', 'LastUpdatedDate_snap'
  ];
  return ensureSheet_('LeadsHistory', headers);
}

function logLeadHistory_(leadId, rowNumber, oldRowValues, updatedBy, changeObj) {
  const hist = ensureLeadsHistorySheet_();
  hist.appendRow([
    new Date(),
    updatedBy || '',
    leadId || '',
    rowNumber || '',
    JSON.stringify(changeObj || {}),
    ...oldRowValues
  ]);
}

/** Convert DD/MM/YYYY -> YYYY-MM-DD (for HTML date input) */
function dmyToISO_(dmy) {
  if (!dmy) return '';
  const s = String(dmy).trim();
  const m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!m) return '';
  return `${m[3]}-${m[2]}-${m[1]}`;
}

/**
 * TELECALLER/EXECUTIVE editable leads list (NEW)
 * - TeleCaller: only assigned to them
 * - Executive: assigned to them OR created by them (same as earlier filtering)
 * Returns followUpDateISO for <input type="date">.
 */
function getLeadsForUpdate(role, userEmail) {
  const sheet = getSheet('Leads');
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  const roleNorm = String(role || '').toUpperCase();
  const emailNorm = String(userEmail || '').trim().toLowerCase();

  const out = [];

  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (!r[0]) continue;

    const assignedTo = String(r[4] || '').trim().toLowerCase();
    const createdBy = String(r[10] || '').trim().toLowerCase();

    if (roleNorm === 'ADMIN') {
      // Admin sees all leads
    } else if (roleNorm === 'EXECUTIVE') {
      // Executive sees only leads they created
      if (createdBy !== emailNorm) continue;
    } else if (roleNorm === 'TELECALLER' || roleNorm === 'TELE-CALLER' || roleNorm === 'TELE CALLER') {
      // TeleCaller sees only leads assigned to them
      if (assignedTo !== emailNorm) continue;
    } else {
      // default safe behavior
      continue;
    }

    const followUpDateISO = dmyToISO_(r[7]); // r[7] is DD/MM/YYYY string in your system

    out.push({
      id: String(r[0] || ''),
      name: String(r[1] || ''),
      phone: String(r[2] || ''),
      source: String(r[3] || ''),
      assignedTo: String(r[4] || ''),
      callNotes: String(r[5] || ''),
      issue: String(r[6] || ''),
      followUpDateISO,
      followUpStatus: String(r[8] || ''),
      status: String(r[9] || '')
    });
  }

  return out.reverse();
}

/**
 * UPDATE WITH HISTORY (NEW)
 * - Stores a snapshot to LeadsHistory BEFORE update
 * - FollowUpDate stored as DD/MM/YYYY string (no sheet formatting)
 * - CallNotes appended (keeps older notes)
 */
function updateLeadWithHistory(update, userEmail, userRole, options) {
  const lock = LockService.getDocumentLock(); // reduces concurrent write collisions [web:21]
  lock.waitLock(20000);

  try {
    const sheet = getSheet('Leads');
    if (!sheet) return { success: false, message: 'Leads sheet not found.' };

    const data = sheet.getDataRange().getValues();
    const leadId = String(update.leadId || '').trim();
    if (!leadId) return { success: false, message: 'Missing leadId.' };

    const roleNorm = String(userRole || '').toUpperCase();
    const emailNorm = String(userEmail || '').trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() !== leadId) continue;

      const rowIndex = i + 1;
      const oldRow = data[i];

      // TeleCaller: strict - only assigned to them
      const assignedTo = String(oldRow[4] || '').trim().toLowerCase();
      if (roleNorm === 'TELECALLER' || roleNorm === 'TELE-CALLER' || roleNorm === 'TELE CALLER') {
        if (assignedTo !== emailNorm) {
          return { success: false, message: 'Not allowed: lead is not assigned to you.' };
        }
      }

      // History snapshot BEFORE changes
      logLeadHistory_(leadId, rowIndex, oldRow, userEmail, update);

      // Apply updates
      if (update.issue !== undefined) {
        sheet.getRange(rowIndex, 7).setValue(update.issue); // Issue (G)
      }

      if (update.followUpDate !== undefined) {
        sheet.getRange(rowIndex, 8).setValue(String(update.followUpDate || '')); // FollowUpDate (H) DD/MM/YYYY string
      }

      if (update.followUpStatus !== undefined) {
        sheet.getRange(rowIndex, 9).setValue(update.followUpStatus); // FollowUpStatus (I)
      }

      if (update.leadStatus !== undefined) {
        sheet.getRange(rowIndex, 10).setValue(update.leadStatus); // LeadStatus (J)
      }

      // Append call note
      if (update.callNotes !== undefined) {
        const newNote = String(update.callNotes || '').trim();
        if (newNote) {
          const oldNotes = String(oldRow[5] || '').trim();
          const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
          const merged = (oldNotes ? oldNotes + '\n\n' : '') + `[${stamp} by ${userEmail}] ${newNote}`;
          sheet.getRange(rowIndex, 6).setValue(merged); // CallNote (F)
        }
      }

      sheet.getRange(rowIndex, 13).setValue(new Date()); // LastUpdatedDate (M)
      return { success: true, message: 'Lead updated (history saved).' };
    }

    return { success: false, message: 'Lead not found.' };
  } finally {
    lock.releaseLock();
  }
}

// Function to download CSV file.

function csvEscape_(v) {
  const s = String(v == null ? '' : v);
  // Quote fields that contain comma, quote, or newline
  if (/[",\n\r]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}

function exportLeadsCsv(role, userEmail) {
  // Use your existing filtering + display-string system
  const leads = getLeads(role, userEmail) || [];

  const headers = [
    'LeadID',
    'Name',
    'Phone',
    'LeadSource',
    'AssignedTo',
    'CallNotes',
    'Issue',
    'FollowUpDate',
    'FollowUpStatus',
    'LeadStatus',
    'CreatedBy',
    'CreatedDate',
    'LastUpdatedDate',
  ];

  const lines = [];
  lines.push(headers.map(csvEscape_).join(','));

  for (let i = 0; i < leads.length; i++) {
    const l = leads[i];
    const row = [
      l.id || '',
      l.name || '',
      l.phone || '',
      l.source || '',
      l.assignedTo || '',
      l.callNotes || '',
      l.issue || '',
      l.followUpDate || '',
      l.followUpStatus || '',
      l.status || '',
      l.createdBy || '',
      l.createdDate || '',
      l.lastUpdatedDate || ''      
    ];
    lines.push(row.map(csvEscape_).join(','));
  }

  return lines.join('\r\n');
}

function getApprovedTeleCallers_() {
  const users = getSheet('Users');
  if (!users) return [];

  const values = users.getDataRange().getValues(); // Email | Password | Role | Status | Name
  const out = [];

  for (let i = 1; i < values.length; i++) {
    const email = String(values[i][0] || '').trim();
    const role = String(values[i][2] || '').trim().toUpperCase();
    const status = String(values[i][3] || '').trim().toUpperCase();
    const name = String(values[i][4] || '').trim();

    if (email && status === 'APPROVED' && (role === 'TELECALLER' || role === 'TELE-CALLER' || role === 'TELE CALLER')) {
      out.push({ email, name });
    }
  }
  // Sort by name then email
  out.sort((a, b) => (a.name || '').localeCompare(b.name || '') || a.email.localeCompare(b.email));
  return out;
}

function getApprovedTelecallers() {
  return getApprovedTeleCallers_();
}


function getAssignmentDataForAdmin() {
  const leadsSheet = getSheet('Leads');
  if (!leadsSheet) return { teleCallers: [], leads: [] };

  const teleCallers = getApprovedTeleCallers_();

  const rows = leadsSheet.getDataRange().getDisplayValues();
  const leads = [];

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;

    leads.push({
      id: r[0],
      name: r[1],
      phone: r[2],
      source: r[3],
      assignedTo: r[4],
      followUpDate: r[7],
      status: r[9]
    });
  }

  return { teleCallers, leads };
}

function assignLeadToTeleCaller(leadId, teleCallerEmail, adminEmail) {
  // Recommended: use your history-enabled updater so old AssignedTo is preserved
  if (typeof updateLeadWithHistory === 'function') {
    return updateLeadWithHistory({ leadId: leadId, assignedTo: teleCallerEmail }, adminEmail);
  }

  // Fallback: simple overwrite (no history)
  return updateLead({ leadId: leadId, assignedTo: teleCallerEmail });
}


/** ---------- USERS approval helpers (NEW) ---------- */

function normalizeRole_(role) {
  return String(role || '').trim().toUpperCase();
}

function isTeleCallerRole_(role) {
  const r = normalizeRole_(role).replace(/[^A-Z]/g, '');
  return r === 'TELECALLER';
}

function ensureUserDecisionCols_() {
  // Users headers: A Email | B Password | C Role | D Status | E Name
  // New: F DecisionAt | G DecisionBy | H DecisionReason
  const sh = getSheet('Users');
  if (!sh) throw new Error('Users sheet not found.');

  const headers = sh.getRange(1, 1, 1, Math.max(5, sh.getLastColumn())).getValues()[0] || [];
  const need = ['DecisionAt', 'DecisionBy', 'DecisionReason'];

  // Ensure at least 5 cols exist
  if (sh.getLastColumn() < 5) sh.insertColumnsAfter(sh.getLastColumn(), 5 - sh.getLastColumn());

  // Put missing headers in F..H
  for (let i = 0; i < need.length; i++) {
    const col = 6 + i; // F=6
    const current = String(sh.getRange(1, col).getValue() || '').trim();
    if (!current) sh.getRange(1, col).setValue(need[i]);
  }
}

function getUserByEmail_(email) {
  const sh = getSheet('Users');
  if (!sh) return null;

  const values = sh.getDataRange().getValues(); // Email | Password | Role | Status | Name | ...
  const emailNorm = String(email || '').trim().toLowerCase();

  for (let i = 1; i < values.length; i++) {
    const e = String(values[i][0] || '').trim().toLowerCase();
    if (e === emailNorm) {
      return {
        row: i + 1,
        email: values[i][0],
        role: values[i][2],
        status: values[i][3],
        name: values[i][4]
      };
    }
  }
  return null;
}

function canAdminApprove_(actorRole) {
  const r = normalizeRole_(actorRole);
  return r === 'ADMIN' || r === 'MANAGER';
}

function canExecutiveApprove_(actorRole) {
  return normalizeRole_(actorRole) === 'EXECUTIVE';
}

/** ---------- ADMIN: list pending users (existing) ----------
 * You already have getPendingUsers(); keep it.
 * We'll add secure approve/reject that check actor email.
 */

function approveUserByAdmin(rowNumber, adminEmail) {
  ensureUserDecisionCols_();

  const actor = getUserByEmail_(adminEmail);
  if (!actor || normalizeRole_(actor.role) === '' || String(actor.status).toUpperCase() !== 'APPROVED' || !canAdminApprove_(actor.role)) {
    return { success: false, message: 'Not allowed.' };
  }

  const sh = getSheet('Users');
  const status = String(sh.getRange(rowNumber, 4).getValue() || '').toUpperCase();
  if (status !== 'PENDING') return { success: false, message: 'User is not pending.' };

  sh.getRange(rowNumber, 4).setValue('APPROVED');   // Status (D)
  sh.getRange(rowNumber, 6).setValue(new Date());   // DecisionAt (F)
  sh.getRange(rowNumber, 7).setValue(adminEmail);   // DecisionBy (G)
  sh.getRange(rowNumber, 8).setValue('');           // DecisionReason (H)

  return { success: true, message: 'User approved.' };
}

function rejectUserByAdmin(rowNumber, adminEmail, reason) {
  ensureUserDecisionCols_();

  const actor = getUserByEmail_(adminEmail);
  if (!actor || String(actor.status).toUpperCase() !== 'APPROVED' || !canAdminApprove_(actor.role)) {
    return { success: false, message: 'Not allowed.' };
  }

  const sh = getSheet('Users');
  const status = String(sh.getRange(rowNumber, 4).getValue() || '').toUpperCase();
  if (status !== 'PENDING') return { success: false, message: 'User is not pending.' };

  sh.getRange(rowNumber, 4).setValue('REJECTED');
  sh.getRange(rowNumber, 6).setValue(new Date());
  sh.getRange(rowNumber, 7).setValue(adminEmail);
  sh.getRange(rowNumber, 8).setValue(String(reason || '').trim());

  return { success: true, message: 'User rejected.' };
}

/** ---------- EXECUTIVE: list pending telecallers + approve/reject ---------- */

function getPendingTeleCallersForExecutive(executiveEmail) {
  const actor = getUserByEmail_(executiveEmail);
  if (!actor || String(actor.status).toUpperCase() !== 'APPROVED' || !canExecutiveApprove_(actor.role)) {
    return [];
  }

  const sh = getSheet('Users');
  const data = sh.getDataRange().getValues();

  const out = [];
  for (let i = 1; i < data.length; i++) {
    const status = String(data[i][3] || '').toUpperCase();
    const role = data[i][2];
    if (status === 'PENDING' && isTeleCallerRole_(role)) {
      out.push({ row: i + 1, email: data[i][0], role: data[i][2], name: data[i][4] });
    }
  }
  return out;
}

function approveTeleCallerByExecutive(rowNumber, executiveEmail) {
  ensureUserDecisionCols_();

  const actor = getUserByEmail_(executiveEmail);
  if (!actor || String(actor.status).toUpperCase() !== 'APPROVED' || !canExecutiveApprove_(actor.role)) {
    return { success: false, message: 'Not allowed.' };
  }

  const sh = getSheet('Users');
  const role = sh.getRange(rowNumber, 3).getValue();
  const status = String(sh.getRange(rowNumber, 4).getValue() || '').toUpperCase();

  if (!isTeleCallerRole_(role)) return { success: false, message: 'Only TeleCallers can be approved here.' };
  if (status !== 'PENDING') return { success: false, message: 'User is not pending.' };

  sh.getRange(rowNumber, 4).setValue('APPROVED');
  sh.getRange(rowNumber, 6).setValue(new Date());
  sh.getRange(rowNumber, 7).setValue(executiveEmail);
  sh.getRange(rowNumber, 8).setValue('');

  return { success: true, message: 'TeleCaller approved.' };
}

function rejectTeleCallerByExecutive(rowNumber, executiveEmail, reason) {
  ensureUserDecisionCols_();

  const actor = getUserByEmail_(executiveEmail);
  if (!actor || String(actor.status).toUpperCase() !== 'APPROVED' || !canExecutiveApprove_(actor.role)) {
    return { success: false, message: 'Not allowed.' };
  }

  const sh = getSheet('Users');
  const role = sh.getRange(rowNumber, 3).getValue();
  const status = String(sh.getRange(rowNumber, 4).getValue() || '').toUpperCase();

  if (!isTeleCallerRole_(role)) return { success: false, message: 'Only TeleCallers can be rejected here.' };
  if (status !== 'PENDING') return { success: false, message: 'User is not pending.' };

  sh.getRange(rowNumber, 4).setValue('REJECTED');
  sh.getRange(rowNumber, 6).setValue(new Date());
  sh.getRange(rowNumber, 7).setValue(executiveEmail);
  sh.getRange(rowNumber, 8).setValue(String(reason || '').trim());

  return { success: true, message: 'TeleCaller rejected.' };
}

// CSV file upload
function importLeadsFromCsv(csvText, createdBy) {
  if (!csvText || !String(csvText).trim()) {
    return { success: false, type: 'error', message: 'Empty CSV file.' };
  }

  // Parse CSV -> 2D array [web:35]
  const rows = Utilities.parseCsv(String(csvText));

  if (!rows || rows.length < 2) {
    return { success: false, type: 'warning', message: 'CSV has no data rows.' };
  }

  // Header map (first row)
  const header = rows[0].map(h => String(h || '').trim().toLowerCase());
  const idx = (name) => header.indexOf(String(name).toLowerCase());

  const iName = idx('name');
  if (iName === -1) {
    return { success: false, type: 'error', message: 'Header "Name" is required.' };
  }

  const iPhone = idx('phone');
  const iSource = idx('source');
  const iNotes = idx('callnotes');
  const iFollowUp = idx('followupdate');
  const iAssigned = idx('assignedto');

  // Build normalized objects
  const data = [];
  for (let r = 1; r < rows.length; r++) {
    const line = rows[r] || [];
    const name = String(line[iName] || '').trim();
    if (!name) continue;

    const phone = iPhone >= 0 ? String(line[iPhone] || '').trim() : '';
    const source = iSource >= 0 ? String(line[iSource] || '').trim() : 'Bulk Import';
    const notes = iNotes >= 0 ? String(line[iNotes] || '').trim() : '';
    const followUpRaw = iFollowUp >= 0 ? line[iFollowUp] : '';
    const assignedTo = iAssigned >= 0 ? String(line[iAssigned] || '').trim() : '';

    // If you already have normalizeFollowUpDate_() in your project, use it
    const followUpDate = (typeof normalizeFollowUpDate_ === 'function')
      ? normalizeFollowUpDate_(followUpRaw)
      : String(followUpRaw || '').trim();

    data.push({ name, phone, source, notes, followUpDate, assignedTo });
  }

  if (!data.length) {
    return { success: false, type: 'warning', message: 'No valid rows found (blank Name).' };
  }

  // Lead IDs: L0DDMMYY0XXX (uses your helper)
  const ids = nextLeadIds_(data.length);
  const sheet = getSheet('Leads');
  const now = new Date();

  const out = data.map((r, k) => ([
    ids[k],
    r.name,
    r.phone,
    r.source,
    r.assignedTo || '',       // AssignedTo (E)
    r.notes || '',            // CallNote
    r.notes || '',            // Issue (same as your system)
    r.followUpDate || '',     // FollowUpDate (DD/MM/YYYY string)
    'Pending',
    'New',
    createdBy || '',
    now,
    now
  ]));

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, out.length, out[0].length).setValues(out);

  return { success: true, type: 'success', count: out.length, message: out.length + ' leads imported from CSV.' };
}

