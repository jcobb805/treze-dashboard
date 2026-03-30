// ═══════════════════════════════════════════════════════════════════════════
// Treze Alcove — Loan Servicing Dashboard Backend
// Google Apps Script — Deploy as Web App ("Anyone" access)
// ═══════════════════════════════════════════════════════════════════════════

const SHEET_NAME = 'NoteStatus';

const VALID_USERS = ['Josh', 'Kristen', 'Kinzie', 'Jill', 'Lauren'];

const VALID_FIELDS = [
  'collection_status',
  'collection_notes',
  'insurance_expiration',
  'insurance_carrier',
  'insurance_policy_num',
  'general_notes'
];

// ─── GET: Return all note statuses ─────────────────────────────────────
function doGet(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return jsonResponse([]);

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return jsonResponse([]);

    const headers = data[0];
    const rows = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        obj[h] = row[i] instanceof Date ? row[i].toISOString() : row[i];
      });
      return obj;
    });

    return jsonResponse(rows);
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// ─── POST: Upsert a note field value ──────────────────────────────────
function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);
    const payload = JSON.parse(e.postData.contents);
    const { noteId, fieldName, fieldValue, updatedBy } = payload;

    // Validate
    if (!noteId || !fieldName) {
      return jsonResponse({ error: 'Missing required fields (noteId, fieldName)' });
    }
    if (!VALID_FIELDS.includes(fieldName)) {
      return jsonResponse({ error: 'Invalid field: ' + fieldName });
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return jsonResponse({ error: 'Sheet not found. Run setup() first.' });

    const data = sheet.getDataRange().getValues();
    const now = new Date().toISOString();
    let found = false;

    // Look for existing row with same note_id + field_name
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(noteId) && String(data[i][1]) === String(fieldName)) {
        sheet.getRange(i + 1, 3).setValue(fieldValue);
        sheet.getRange(i + 1, 4).setValue(updatedBy || '');
        sheet.getRange(i + 1, 5).setValue(now);
        found = true;
        break;
      }
    }

    // Insert new row if not found
    if (!found) {
      sheet.appendRow([String(noteId), String(fieldName), String(fieldValue), updatedBy || '', now]);
    }

    return jsonResponse({ success: true, noteId, fieldName });
  } catch (err) {
    return jsonResponse({ error: err.message });
  } finally {
    lock.releaseLock();
  }
}

// ─── Helper: JSON response ─────────────────────────────────────────────
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── Run once: Create the NoteStatus sheet with headers ────────────────
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  // Set headers
  sheet.getRange('A1:E1').setValues([['note_id', 'field_name', 'field_value', 'updated_by', 'updated_at']]);
  sheet.getRange('1:1').setFontWeight('bold');
  sheet.setFrozenRows(1);

  // Force text format to prevent auto-conversion
  sheet.getRange('A:C').setNumberFormat('@');

  // Auto-size columns
  sheet.autoResizeColumns(1, 5);

  Logger.log('Setup complete. Sheet "NoteStatus" is ready.');
}
