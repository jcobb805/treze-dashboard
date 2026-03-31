// ═══════════════════════════════════════════════════════════════════════════
// Treze Alcove — Seed Google Sheet from Excel
// Extracts data from Excel, then POSTs to Apps Script bulk_import endpoint
// Run: node seed.js
// ═══════════════════════════════════════════════════════════════════════════

const ExcelJS = require('exceljs');
const path = require('path');

const API_URL = 'https://script.google.com/macros/s/AKfycby7VS75jgpKKmq_SOzGq05v6J34wn_yMClm34h-bPAEolDOgGayAV-o0YZQVfVcw4a_/exec';

const BASE = path.join(__dirname, '..', 'OneDrive', 'Treze Alcove');
const INPUT = path.join(BASE, 'Treze Alcove - Preze Enterprises - Amortization Schedule.xlsx');
const INS_FILE = path.join(BASE, 'Treze Alcove - Proof of Insurance.xlsx');
const CONTACT_FILE = path.join(BASE, 'Treze Alcove - Property Contact Info..xlsx');

const GREEN_ARGB = 'FF92D050';

function isGreenFill(cell) {
  const fill = cell.fill || (cell.style && cell.style.fill);
  if (!fill || fill.pattern !== 'solid') return false;
  return fill.fgColor && fill.fgColor.argb === GREEN_ARGB;
}

async function postBulk(sheetName, rows) {
  console.log(`  Importing ${rows.length} rows to ${sheetName}...`);
  const res = await fetch(API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ action: 'bulk_import', sheet: sheetName, rows })
  });
  const data = await res.json();
  if (data.error) throw new Error(`${sheetName}: ${data.error}`);
  console.log(`  Done: ${data.imported} rows imported`);
}

async function seed() {
  console.log('Reading Excel workbook...');
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(INPUT);

  const notesRows = [];
  const paymentsRows = [];

  for (const ws of wb.worksheets) {
    const getVal = (row, col) => {
      let v = ws.getRow(row).getCell(col).value;
      if (v && typeof v === 'object' && (v.formula || v.sharedFormula)) v = v.result;
      return v;
    };

    const borrower = String(getVal(1, 1) || '').trim();
    const address = String(getVal(1, 6) || ws.name.split('(')[0].trim()).trim();
    const loanAmt = getVal(2, 2);
    const rate = getVal(3, 2);
    const salesPrice = getVal(3, 5);
    const amortMonths = getVal(4, 2);
    const balloonAt = getVal(5, 2);
    const beginDateRaw = getVal(6, 2);
    const monthlyPmt = getVal(8, 2);
    const beginDate = beginDateRaw instanceof Date ? beginDateRaw : new Date(beginDateRaw);
    const gain = getVal(2, 5);
    const gpPct = getVal(4, 5);

    if (!loanAmt || !borrower) continue;

    // Compute balloon date
    const balloonDate = new Date(beginDate);
    balloonDate.setMonth(balloonDate.getMonth() + (balloonAt || 120));

    // Determine closing number from sheet name
    const closingMatch = ws.name.match(/#(\d+)/);
    const closing = closingMatch ? parseInt(closingMatch[1]) : 0;

    const noteId = ws.name.trim();

    // Parse header row
    const headerRow = ws.getRow(10);
    const headers = {};
    for (let c = 1; c <= 15; c++) {
      const h = headerRow.getCell(c).value;
      if (h) headers[String(h).toLowerCase().trim()] = c;
    }
    const dateCol = headers['date'] || 2;
    const pmtCol = headers['payment'] || 3;
    const interestCol = headers['interest'] || 4;
    const principalCol = headers['principal'] || 5;
    let balanceCol = -1;
    for (let c = 1; c <= 15; c++) {
      const h = headerRow.getCell(c).value;
      if (h && String(h).toLowerCase().trim() === 'balance') { balanceCol = c; break; }
    }

    let hasCollected = false;
    let hasScheduled = false;
    const payments = [];

    for (let r = 11; r <= Math.min(ws.rowCount, 500); r++) {
      const dateCell = ws.getRow(r).getCell(dateCol);
      const dateVal = dateCell.value;
      if (dateVal === null || dateVal === undefined || dateVal === '') continue;
      let rawDate = dateVal;
      if (rawDate && typeof rawDate === 'object' && (rawDate.formula || rawDate.sharedFormula)) rawDate = rawDate.result;
      const date = rawDate instanceof Date ? rawDate : new Date(rawDate);
      if (isNaN(date.getTime())) continue;
      const pmt = getVal(r, pmtCol) || 0;
      const interest = getVal(r, interestCol) || 0;
      const principal = getVal(r, principalCol) || 0;
      const balance = balanceCol > 0 ? (getVal(r, balanceCol) || 0) : null;
      const month = getVal(r, 1);
      if (month === 'Loan' || month === 'loan') continue;
      const collected = isGreenFill(dateCell);

      if (collected) hasCollected = true;
      else hasScheduled = true;

      payments.push({
        noteId,
        month: typeof month === 'number' ? month : '',
        date: date.toISOString().split('T')[0],
        payment: typeof pmt === 'number' ? Math.round(pmt * 100) / 100 : 0,
        interest: typeof interest === 'number' ? Math.round(interest * 100) / 100 : 0,
        principal: typeof principal === 'number' ? Math.round(principal * 100) / 100 : 0,
        balance: typeof balance === 'number' ? Math.round(balance * 100) / 100 : '',
        collected: collected ? 'TRUE' : 'FALSE',
        collectedBy: '',
        collectedAt: ''
      });
    }

    if (!hasCollected || !hasScheduled) continue;

    notesRows.push({
      noteId,
      address,
      closing,
      borrower,
      loanAmt: Math.round(loanAmt * 100) / 100,
      rate,
      salesPrice: salesPrice ? Math.round(salesPrice * 100) / 100 : '',
      gain: gain ? Math.round(gain * 100) / 100 : '',
      gpPct: gpPct || '',
      amortMonths: amortMonths || 360,
      balloonAt: balloonAt || 120,
      beginDate: beginDate.toISOString().split('T')[0],
      monthlyPmt: Math.round(monthlyPmt * 100) / 100,
      balloonDate: balloonDate.toISOString().split('T')[0],
      noteStatus: 'active'
    });

    paymentsRows.push(...payments);
    console.log(`  ${noteId}: ${payments.length} payments`);
  }

  // Read insurance file
  let insRows = [];
  try {
    const insWb = new ExcelJS.Workbook();
    await insWb.xlsx.readFile(INS_FILE);
    const insWs = insWb.worksheets[0];
    if (insWs) {
      for (let r = 2; r <= insWs.rowCount; r++) {
        const addr = String(insWs.getRow(r).getCell(1).value || '').trim();
        if (!addr) continue;
        const getV = (c) => { let v = insWs.getRow(r).getCell(c).value; return v instanceof Date ? v.toISOString().split('T')[0] : (v || ''); };
        insRows.push({
          address: addr,
          insured: String(getV(2)),
          expiration: String(getV(3)),
          carrier: String(getV(4)),
          policyNum: String(getV(5)),
          notes: String(getV(6))
        });
      }
    }
  } catch (e) { console.log('  No insurance file found, skipping'); }

  // Read contacts file
  let contactRows = [];
  try {
    const conWb = new ExcelJS.Workbook();
    await conWb.xlsx.readFile(CONTACT_FILE);
    const conWs = conWb.worksheets[0];
    if (conWs) {
      for (let r = 2; r <= conWs.rowCount; r++) {
        const addr = String(conWs.getRow(r).getCell(1).value || '').trim();
        if (!addr) continue;
        const getV = (c) => String(conWs.getRow(r).getCell(c).value || '');
        contactRows.push({
          address: addr,
          contact: getV(2),
          ownerName: getV(3),
          email: getV(4),
          mailingAddress: getV(5),
          paymentMethod: getV(6),
          phone: getV(7)
        });
      }
    }
  } catch (e) { console.log('  No contacts file found, skipping'); }

  console.log(`\nFound ${notesRows.length} notes, ${paymentsRows.length} payments, ${insRows.length} insurance, ${contactRows.length} contacts`);
  console.log('\nSeeding Google Sheet...');

  await postBulk('Notes', notesRows);
  await postBulk('Payments', paymentsRows);
  if (insRows.length > 0) await postBulk('Insurance', insRows);
  if (contactRows.length > 0) await postBulk('Contacts', contactRows);

  console.log('\nDone! Google Sheet is seeded.');
}

seed().catch(err => { console.error('Error:', err.message); process.exit(1); });