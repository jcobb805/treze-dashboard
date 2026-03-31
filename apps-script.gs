// ═══════════════════════════════════════════════════════════════════════════
// Treze Alcove — Loan Servicing Dashboard Backend v2
// Google Apps Script — Deploy as Web App ("Anyone" access)
// ═══════════════════════════════════════════════════════════════════════════

const VALID_USERS = ['Josh', 'Kristen', 'Kinzie', 'Jill', 'Lauren'];

// Sheet names
const SH = {
  NOTES: 'Notes',
  PAYMENTS: 'Payments',
  INSURANCE: 'Insurance',
  CONTACTS: 'Contacts',
  STATUS: 'NoteStatus',
  AUDIT: 'AuditLog'
};

// ─── GET: Return data from requested sheet(s) ─────────────────────────────
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'all';
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === 'all') {
      return jsonResponse({
        notes: readSheet(ss, SH.NOTES),
        payments: readSheet(ss, SH.PAYMENTS),
        insurance: readSheet(ss, SH.INSURANCE),
        contacts: readSheet(ss, SH.CONTACTS),
        statuses: readSheet(ss, SH.STATUS)
      });
    }

    // Single sheet request: ?action=notes, ?action=payments, etc.
    const sheetMap = {
      notes: SH.NOTES, payments: SH.PAYMENTS, insurance: SH.INSURANCE,
      contacts: SH.CONTACTS, statuses: SH.STATUS, audit: SH.AUDIT
    };
    const sheetName = sheetMap[action];
    if (!sheetName) return jsonResponse({ error: 'Unknown action: ' + action });
    return jsonResponse(readSheet(ss, sheetName));
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// ─── POST: Handle write operations ───────────────────────────────────────
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;

    if (action === 'mark_collected') return markCollected(payload);
    if (action === 'undo_collected') return undoCollected(payload);
    if (action === 'save_status') return saveStatus(payload);
    if (action === 'save_insurance') return saveInsurance(payload);
    if (action === 'save_contact') return saveContact(payload);
    if (action === 'log_action') return logAction(payload);
    if (action === 'bulk_import') return bulkImport(payload);

    return jsonResponse({ error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ error: err.message });
  } finally {
    lock.releaseLock();
  }
}

// ─── Mark payment as collected ───────────────────────────────────────────
function markCollected(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SH.PAYMENTS);
  if (!sheet) return jsonResponse({ error: 'Payments sheet not found' });

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const noteIdCol = headers.indexOf('noteId');
  const dateCol = headers.indexOf('date');
  const collectedCol = headers.indexOf('collected');
  const collectedByCol = headers.indexOf('collectedBy');
  const collectedAtCol = headers.indexOf('collectedAt');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][noteIdCol]) === String(p.noteId) && String(data[i][dateCol]) === String(p.date)) {
      sheet.getRange(i + 1, collectedCol + 1).setValue('TRUE');
      sheet.getRange(i + 1, collectedByCol + 1).setValue(p.updatedBy || '');
      sheet.getRange(i + 1, collectedAtCol + 1).setValue(new Date().toISOString());
      addAuditRow(ss, p.updatedBy, 'payment_collected', p.noteId, 'Date: ' + p.date + ', Amount: ' + (p.amount || ''));
      return jsonResponse({ success: true });
    }
  }
  return jsonResponse({ error: 'Payment row not found' });
}

// ─── Undo collected status ───────────────────────────────────────────────
function undoCollected(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SH.PAYMENTS);
  if (!sheet) return jsonResponse({ error: 'Payments sheet not found' });

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const noteIdCol = headers.indexOf('noteId');
  const dateCol = headers.indexOf('date');
  const collectedCol = headers.indexOf('collected');
  const collectedByCol = headers.indexOf('collectedBy');
  const collectedAtCol = headers.indexOf('collectedAt');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][noteIdCol]) === String(p.noteId) && String(data[i][dateCol]) === String(p.date)) {
      sheet.getRange(i + 1, collectedCol + 1).setValue('FALSE');
      sheet.getRange(i + 1, collectedByCol + 1).setValue('');
      sheet.getRange(i + 1, collectedAtCol + 1).setValue('');
      addAuditRow(ss, p.updatedBy, 'payment_uncollected', p.noteId, 'Date: ' + p.date);
      return jsonResponse({ success: true });
    }
  }
  return jsonResponse({ error: 'Payment row not found' });
}

// ─── Save status field (key-value) ───────────────────────────────────────
function saveStatus(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SH.STATUS);
  if (!sheet) return jsonResponse({ error: 'Status sheet not found' });

  const data = sheet.getDataRange().getValues();
  const now = new Date().toISOString();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(p.noteId) && String(data[i][1]) === String(p.fieldName)) {
      sheet.getRange(i + 1, 3).setValue(p.fieldValue);
      sheet.getRange(i + 1, 4).setValue(p.updatedBy || '');
      sheet.getRange(i + 1, 5).setValue(now);
      return jsonResponse({ success: true });
    }
  }
  sheet.appendRow([String(p.noteId), String(p.fieldName), String(p.fieldValue), p.updatedBy || '', now]);
  return jsonResponse({ success: true });
}

// ─── Save insurance record ───────────────────────────────────────────────
function saveInsurance(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SH.INSURANCE);
  if (!sheet) return jsonResponse({ error: 'Insurance sheet not found' });

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const addrCol = headers.indexOf('address');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][addrCol]).trim() === String(p.address).trim()) {
      // Update fields that were provided
      if (p.insured !== undefined) sheet.getRange(i + 1, headers.indexOf('insured') + 1).setValue(p.insured);
      if (p.expiration !== undefined) sheet.getRange(i + 1, headers.indexOf('expiration') + 1).setValue(p.expiration);
      if (p.carrier !== undefined) sheet.getRange(i + 1, headers.indexOf('carrier') + 1).setValue(p.carrier);
      if (p.policyNum !== undefined) sheet.getRange(i + 1, headers.indexOf('policyNum') + 1).setValue(p.policyNum);
      if (p.notes !== undefined) sheet.getRange(i + 1, headers.indexOf('notes') + 1).setValue(p.notes);
      addAuditRow(ss, p.updatedBy, 'insurance_updated', p.address, JSON.stringify(p));
      return jsonResponse({ success: true });
    }
  }
  // Insert new
  sheet.appendRow([p.address, p.insured || '', p.expiration || '', p.carrier || '', p.policyNum || '', p.notes || '']);
  addAuditRow(ss, p.updatedBy, 'insurance_added', p.address, '');
  return jsonResponse({ success: true });
}

// ─── Save contact record ─────────────────────────────────────────────────
function saveContact(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SH.CONTACTS);
  if (!sheet) return jsonResponse({ error: 'Contacts sheet not found' });

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const addrCol = headers.indexOf('address');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][addrCol]).trim() === String(p.address).trim()) {
      ['contact','ownerName','email','mailingAddress','paymentMethod','phone'].forEach(f => {
        if (p[f] !== undefined) sheet.getRange(i + 1, headers.indexOf(f) + 1).setValue(p[f]);
      });
      addAuditRow(ss, p.updatedBy, 'contact_updated', p.address, '');
      return jsonResponse({ success: true });
    }
  }
  sheet.appendRow([p.address, p.contact||'', p.ownerName||'', p.email||'', p.mailingAddress||'', p.paymentMethod||'', p.phone||'']);
  return jsonResponse({ success: true });
}

// ─── Log action ──────────────────────────────────────────────────────────
function logAction(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  addAuditRow(ss, p.updatedBy, p.actionType, p.noteId || '', p.details || '');
  return jsonResponse({ success: true });
}

// ─── Bulk import (for migration) ─────────────────────────────────────────
function bulkImport(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = p.sheet;
  const rows = p.rows; // Array of objects

  if (!sheetName || !rows || !rows.length) return jsonResponse({ error: 'Missing sheet or rows' });

  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  const headers = Object.keys(rows[0]);
  const data = [headers, ...rows.map(r => headers.map(h => r[h] !== null && r[h] !== undefined ? r[h] : ''))];

  sheet.clear();
  sheet.getRange(1, 1, data.length, headers.length).setValues(data);
  sheet.getRange('1:1').setFontWeight('bold');
  sheet.setFrozenRows(1);

  return jsonResponse({ success: true, imported: rows.length });
}

// ─── Helpers ─────────────────────────────────────────────────────────────
function readSheet(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = row[i] instanceof Date ? row[i].toISOString().split('T')[0] : row[i];
    });
    return obj;
  });
}

function addAuditRow(ss, user, action, noteId, details) {
  let sheet = ss.getSheetByName(SH.AUDIT);
  if (!sheet) {
    sheet = ss.insertSheet(SH.AUDIT);
    sheet.getRange('A1:E1').setValues([['timestamp', 'user', 'action', 'noteId', 'details']]);
    sheet.getRange('1:1').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([new Date().toISOString(), user || '', action, noteId, details]);
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── Run once: Create all sheets with headers ────────────────────────────
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheets = {
    [SH.NOTES]: ['noteId','address','closing','borrower','loanAmt','rate','salesPrice','gain','gpPct','amortMonths','balloonAt','beginDate','monthlyPmt','balloonDate','noteStatus'],
    [SH.PAYMENTS]: ['noteId','month','date','payment','interest','principal','balance','collected','collectedBy','collectedAt'],
    [SH.INSURANCE]: ['address','insured','expiration','carrier','policyNum','notes'],
    [SH.CONTACTS]: ['address','contact','ownerName','email','mailingAddress','paymentMethod','phone'],
    [SH.STATUS]: ['note_id','field_name','field_value','updated_by','updated_at'],
    [SH.AUDIT]: ['timestamp','user','action','noteId','details']
  };

  Object.entries(sheets).forEach(([name, headers]) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange('1:1').setFontWeight('bold');
    sheet.setFrozenRows(1);
  });

  Logger.log('Setup complete. All sheets created.');
}
