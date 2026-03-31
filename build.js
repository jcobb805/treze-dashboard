const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const BASE = path.join(__dirname, '..', 'OneDrive', 'Treze Alcove');
const INPUT = path.join(BASE, 'Treze Alcove - Preze Enterprises - Amortization Schedule.xlsx');
const INS_FILE = path.join(BASE, 'Treze Alcove - Proof of Insurance.xlsx');
const CONTACT_FILE = path.join(BASE, 'Treze Alcove - Property Contact Info..xlsx');
const OUTPUT = path.join(__dirname, 'index.html');
const TODAY = new Date();
const LATE_FEE_RATE = 0.05;

const GREEN_ARGB = 'FF92D050';

function isGreenFill(cell) {
  const fill = cell.fill || (cell.style && cell.style.fill);
  if (!fill || fill.pattern !== 'solid') return false;
  return fill.fgColor && fill.fgColor.argb === GREEN_ARGB;
}

async function build() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(INPUT);

  const notes = [];

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

    const collectedPayments = [];
    const scheduledPayments = [];
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
      const entry = {
        month: typeof month === 'number' ? month : null,
        date: date.toISOString().split('T')[0],
        payment: typeof pmt === 'number' ? Math.round(pmt * 100) / 100 : 0,
        interest: typeof interest === 'number' ? Math.round(interest * 100) / 100 : 0,
        principal: typeof principal === 'number' ? Math.round(principal * 100) / 100 : 0,
        balance: typeof balance === 'number' ? Math.round(balance * 100) / 100 : null,
        collected
      };
      if (collected) collectedPayments.push(entry);
      else scheduledPayments.push(entry);
    }

    if (collectedPayments.length === 0) continue;
    if (scheduledPayments.length === 0) continue;

    const lastCollected = collectedPayments[collectedPayments.length - 1];
    const lastCollectedDate = new Date(lastCollected.date);
    const payoffThreshold = loanAmt * 0.5;
    const payoffCheck = [...collectedPayments, ...scheduledPayments.slice(0, 3)];
    const hasPayoff = payoffCheck.some(p => p.payment > payoffThreshold);
    if (hasPayoff) continue;
    const daysSinceLastGreen = Math.floor((TODAY - lastCollectedDate) / (1000 * 60 * 60 * 24));
    if (daysSinceLastGreen > 180) continue;

    const nextScheduled = scheduledPayments[0];
    const balloonDate = new Date(beginDate);
    balloonDate.setMonth(balloonDate.getMonth() + balloonAt);
    const nextDueDate = new Date(nextScheduled.date);
    const daysSinceLastPmt = Math.floor((TODAY - lastCollectedDate) / (1000 * 60 * 60 * 24));
    const daysPastDue = Math.max(0, Math.floor((TODAY - nextDueDate) / (1000 * 60 * 60 * 24)));

    let status = 'Current';
    if (daysPastDue > 60) status = 'Severely Delinquent';
    else if (daysPastDue > 30) status = 'Demand Letter';
    else if (daysPastDue > 15) status = 'Late Notice';
    else if (daysPastDue > 0) status = 'Past Due';

    const closingMatch = ws.name.match(/\(#(\d+)\)/);
    const closing = closingMatch ? parseInt(closingMatch[1]) : null;

    let currentBalance = lastCollected.balance;
    if (!currentBalance || typeof currentBalance !== 'number') {
      const totalPrincipal = collectedPayments.reduce((s, p) => s + p.principal, 0);
      currentBalance = Math.max(0, loanAmt - totalPrincipal);
    }

    notes.push({
      id: ws.name, address, closing, borrower, loanAmt, rate, salesPrice,
      gain: typeof gain === 'number' ? Math.round(gain * 100) / 100 : null,
      gpPct: typeof gpPct === 'number' ? gpPct : null,
      amortMonths, balloonAt,
      beginDate: beginDate.toISOString().split('T')[0],
      monthlyPmt,
      balloonDate: balloonDate.toISOString().split('T')[0],
      currentBalance: Math.round(currentBalance * 100) / 100,
      lastPmtDate: lastCollected.date,
      lastPmtAmount: lastCollected.payment,
      nextDueDate: nextDueDate.toISOString().split('T')[0],
      daysSinceLastPmt, daysPastDue, status,
      totalCollected: collectedPayments.length,
      remainingPayments: scheduledPayments.length,
      collectedPayments, scheduledPayments
    });
  }

  const statusOrder = { 'Severely Delinquent': 0, 'Demand Letter': 1, 'Late Notice': 2, 'Past Due': 3, 'Current': 4 };
  notes.sort((a, b) => (statusOrder[a.status] - statusOrder[b.status]) || (b.daysPastDue - a.daysPastDue));

  // Read insurance
  const insMap = {};
  try {
    const iwb = new ExcelJS.Workbook();
    await iwb.xlsx.readFile(INS_FILE);
    const iws = iwb.worksheets[0];
    for (let r = 2; r <= iws.rowCount; r++) {
      let addr = iws.getRow(r).getCell(1).value;
      if (!addr) continue;
      addr = String(addr).trim();
      let insured = iws.getRow(r).getCell(2).value;
      let expDate = iws.getRow(r).getCell(3).value;
      let insNotes = iws.getRow(r).getCell(4).value;
      if (expDate instanceof Date) expDate = expDate.toISOString().split('T')[0];
      insMap[addr] = { insured: insured ? String(insured).trim() : '', expiration: expDate || '', insNotes: insNotes ? String(insNotes).trim() : '' };
    }
    console.log(`Loaded ${Object.keys(insMap).length} insurance records`);
  } catch(e) { console.warn('Could not read insurance file:', e.message); }

  // Read contacts
  const contactMap = {};
  try {
    const cwb = new ExcelJS.Workbook();
    await cwb.xlsx.readFile(CONTACT_FILE);
    const cws = cwb.worksheets[0];
    for (let r = 2; r <= cws.rowCount; r++) {
      const getC = (col) => {
        let v = cws.getRow(r).getCell(col).value;
        if (v && typeof v === 'object' && v.richText) v = v.richText.map(rt => rt.text).join('');
        if (v && typeof v === 'object' && v.text) v = v.text;
        return v ? String(v).trim() : '';
      };
      const addr = getC(1);
      if (!addr) continue;
      contactMap[addr] = { contact: getC(2), ownerName: getC(3), email: getC(4), mailingAddress: getC(5), paymentMethod: getC(6), phone: getC(7) };
    }
    console.log(`Loaded ${Object.keys(contactMap).length} contact records`);
  } catch(e) { console.warn('Could not read contact file:', e.message); }

  // Merge
  for (const n of notes) {
    const addrKey = n.address.trim();
    const ins = insMap[addrKey] || {};
    n.insured = ins.insured || '';
    n.insuranceExpiration = ins.expiration || '';
    n.insuranceNotes = ins.insNotes || '';
    const ct = contactMap[addrKey] || {};
    n.contact = ct.contact || '';
    n.ownerName = ct.ownerName || '';
    n.email = ct.email || '';
    n.mailingAddress = ct.mailingAddress || '';
    n.paymentMethod = ct.paymentMethod || '';
    n.phone = ct.phone || '';
    n.lateFee = Math.round(n.monthlyPmt * LATE_FEE_RATE * 100) / 100;
    n.totalOwed = Math.round((n.monthlyPmt + n.lateFee) * 100) / 100;
  }

  console.log(`Extracted ${notes.length} active notes`);
  notes.forEach(n => console.log(`  ${n.id} | ${n.status} | ${n.daysPastDue} days past due | Bal: $${n.currentBalance.toLocaleString()}`));

  const html = generateHTML(notes);
  fs.writeFileSync(OUTPUT, html, 'utf8');
  console.log(`\nDashboard written to ${OUTPUT}`);
}

function generateHTML(notes) {
  const dataJSON = JSON.stringify(notes);
  const todayStr = TODAY.toISOString().split('T')[0];
  const todayDisplay = TODAY.toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' });

return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Treze Alcove - Loan Servicing Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.2/dist/chart.umd.min.js"><\/script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.2/jspdf.umd.min.js"><\/script>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Segoe UI',Arial,sans-serif;background:#0f1117;color:#e0e0e0;min-height:100vh}
.header{background:linear-gradient(135deg,#1a1d2e 0%,#252940 100%);padding:18px 30px;border-bottom:2px solid #c9952b;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px}
.header h1{font-size:22px;font-weight:600;color:#fff}
.header h1 span{color:#c9952b}
.hdr-right{display:flex;align-items:center;gap:16px}
.hdr-right .date{font-size:13px;color:#8890a4}
.user-sel{display:flex;align-items:center;gap:6px}
.user-sel label{font-size:11px;color:#8890a4;text-transform:uppercase;letter-spacing:.5px}
.user-sel select{background:#252940;color:#c9952b;border:1px solid #3a3d50;border-radius:6px;padding:5px 10px;font-size:13px;font-weight:600}
.sync-status{font-size:11px;color:#555;margin-left:4px}
.sync-status.ok{color:#2ecc71}
.sync-status.err{color:#e74c3c}

.tabs{display:flex;background:#181b28;border-bottom:1px solid #2a2d3e;padding:0 20px;overflow-x:auto}
.tab{padding:12px 22px;cursor:pointer;font-size:13px;font-weight:500;color:#8890a4;border-bottom:2px solid transparent;transition:all .2s;white-space:nowrap}
.tab:hover{color:#c9952b}
.tab.active{color:#c9952b;border-bottom-color:#c9952b}
.tab .badge{background:#e74c3c;color:#fff;font-size:10px;padding:1px 6px;border-radius:8px;margin-left:6px}

.content{padding:20px 30px}
.tab-content{display:none}
.tab-content.active{display:block}

.kpi-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:14px;margin-bottom:22px}
.kpi{background:#1a1d2e;border-radius:8px;padding:16px;border:1px solid #2a2d3e;cursor:default;transition:border-color .2s}
.kpi.clickable{cursor:pointer}
.kpi.clickable:hover{border-color:#c9952b}
.kpi .label{font-size:11px;text-transform:uppercase;color:#8890a4;letter-spacing:.5px;margin-bottom:5px}
.kpi .value{font-size:24px;font-weight:700;color:#fff}
.kpi .value.green{color:#2ecc71}.kpi .value.yellow{color:#f1c40f}.kpi .value.orange{color:#e67e22}.kpi .value.red{color:#e74c3c}
.kpi .sub{font-size:11px;color:#8890a4;margin-top:3px}

table{width:100%;border-collapse:collapse;font-size:13px}
thead th{background:#1a1d2e;color:#8890a4;text-transform:uppercase;font-size:11px;letter-spacing:.5px;padding:10px 12px;text-align:left;border-bottom:1px solid #2a2d3e;cursor:pointer;user-select:none;position:sticky;top:0;z-index:1}
thead th:hover{color:#c9952b}
tbody tr{border-bottom:1px solid #1e2130;transition:background .15s;cursor:pointer}
tbody tr:hover{background:#1e2235}
tbody td{padding:10px 12px}
tbody td.num{text-align:right;font-variant-numeric:tabular-nums}

.status{display:inline-block;padding:3px 10px;border-radius:12px;font-size:11px;font-weight:600;white-space:nowrap}
.status.current{background:rgba(46,204,113,.15);color:#2ecc71}
.status.past-due{background:rgba(241,196,15,.15);color:#f1c40f}
.status.late-notice{background:rgba(230,126,34,.15);color:#e67e22}
.status.demand-letter{background:rgba(231,76,60,.15);color:#e74c3c}
.status.severely-delinquent{background:rgba(231,76,60,.25);color:#ff6b6b}

.chart-row{display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-bottom:22px}
.chart-card{background:#1a1d2e;border-radius:8px;padding:18px;border:1px solid #2a2d3e}
.chart-card h3{font-size:14px;color:#fff;margin-bottom:12px;font-weight:600}
.chart-wrap{position:relative;height:280px}

.filter-bar{display:flex;gap:12px;margin-bottom:18px;flex-wrap:wrap;align-items:center}
.filter-bar select,.filter-bar input{background:#1a1d2e;color:#e0e0e0;border:1px solid #2a2d3e;border-radius:6px;padding:8px 12px;font-size:13px}
.filter-bar select:focus,.filter-bar input:focus{outline:none;border-color:#c9952b}
.filter-bar label{font-size:12px;color:#8890a4}
.active-filter{display:inline-flex;align-items:center;gap:6px;background:#c9952b22;border:1px solid #c9952b55;color:#c9952b;padding:4px 12px;border-radius:16px;font-size:12px;font-weight:500}
.active-filter .clear{cursor:pointer;font-weight:700;margin-left:4px}
.active-filter .clear:hover{color:#fff}

.table-wrap{max-height:600px;overflow-y:auto;border-radius:8px;border:1px solid #2a2d3e}
.table-wrap::-webkit-scrollbar{width:6px}
.table-wrap::-webkit-scrollbar-track{background:#1a1d2e}
.table-wrap::-webkit-scrollbar-thumb{background:#3a3d4e;border-radius:3px}

tr.action-needed{background:rgba(231,76,60,.05)}
.balloon-warn{color:#e67e22;font-weight:600}

.edit-field{background:#151825;border:1px solid #2a2d3e;border-radius:4px;color:#e0e0e0;padding:4px 8px;font-size:12px;width:100%}
.edit-field:focus{border-color:#c9952b;outline:none}
select.edit-field{cursor:pointer}
textarea.edit-field{resize:vertical;min-height:32px;font-family:inherit}

.coll-status{padding:3px 8px;border-radius:10px;font-size:11px;font-weight:600;border:none;cursor:pointer}
.coll-status.monitoring{background:rgba(241,196,15,.15);color:#f1c40f}
.coll-status.late_notice_sent{background:rgba(230,126,34,.15);color:#e67e22}
.coll-status.demand_letter_sent{background:rgba(231,76,60,.15);color:#e74c3c}
.coll-status.in_collections{background:rgba(231,76,60,.25);color:#ff6b6b}
.coll-status.payment_plan{background:rgba(52,152,219,.15);color:#3498db}
.coll-status.resolved{background:rgba(46,204,113,.15);color:#2ecc71}

.ins-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:12px}
.ins-card{background:#1a1d2e;border-radius:8px;padding:15px;border:1px solid #2a2d3e}
.ins-card .ins-addr{font-weight:600;color:#fff;font-size:14px}
.ins-card .ins-borrower{font-size:12px;color:#8890a4;margin-top:2px}
.ins-card .ins-status{margin-top:10px;display:flex;align-items:center;gap:8px}
.ins-dot{width:10px;height:10px;border-radius:50%;display:inline-block}
.ins-dot.valid{background:#2ecc71}.ins-dot.expiring{background:#f1c40f}.ins-dot.lapsed{background:#e74c3c}.ins-dot.unknown{background:#555}
.ins-card .ins-fields{margin-top:10px;display:grid;grid-template-columns:1fr 1fr;gap:6px}
.ins-card .ins-fields label{font-size:10px;color:#8890a4;text-transform:uppercase}
.ins-card .ins-warn{margin-top:8px;color:#e74c3c;font-size:12px;font-weight:600}

.amort-overlay{display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,.75);z-index:200}
.amort-overlay.show{display:block}
.amort-panel{position:fixed;top:0;right:0;width:95vw;max-width:1200px;height:100vh;background:#12141f;border-left:2px solid #c9952b;overflow-y:auto;z-index:201;transform:translateX(100%);transition:transform .3s ease;padding:25px 30px}
.amort-overlay.show .amort-panel{transform:translateX(0)}
.amort-panel .close-btn{position:absolute;top:15px;right:20px;color:#8890a4;cursor:pointer;font-size:24px;z-index:10}
.amort-panel .close-btn:hover{color:#fff}
.amort-panel h2{font-size:20px;color:#fff;margin-bottom:4px}
.amort-panel .amort-sub{font-size:13px;color:#8890a4;margin-bottom:20px}
.amort-panel .detail-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:20px}
.amort-panel .detail-item .dl{font-size:11px;color:#8890a4;text-transform:uppercase}
.amort-panel .detail-item .dv{font-size:15px;font-weight:600;color:#fff;margin-top:2px}
.amort-panel .amort-charts{display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-bottom:22px}
.amort-panel .amort-chart-card{background:#1a1d2e;border-radius:8px;padding:16px;border:1px solid #2a2d3e}
.amort-panel .amort-chart-card h3{font-size:13px;color:#fff;margin-bottom:10px}
.amort-panel .amort-chart-wrap{height:240px;position:relative}
.amort-panel .notes-section{margin-top:20px;background:#1a1d2e;border-radius:8px;padding:16px;border:1px solid #2a2d3e}
.amort-panel .notes-section h3{font-size:13px;color:#fff;margin-bottom:10px}
.amort-table{font-size:12px}
.amort-table td,.amort-table th{padding:6px 10px}
.amort-table tr.collected{background:rgba(46,204,113,.06)}
.amort-table tr.scheduled{color:#8890a4}
.amort-table tr.divider td{border-top:2px solid #c9952b;font-weight:600;color:#c9952b;font-size:11px;text-transform:uppercase;padding:8px 10px}

/* Buttons */
.btn{display:inline-flex;align-items:center;gap:6px;padding:7px 16px;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer;border:none;transition:all .15s}
.btn-gold{background:#c9952b;color:#fff}.btn-gold:hover{background:#d4a73a}
.btn-red{background:#e74c3c;color:#fff}.btn-red:hover{background:#c0392b}
.btn-outline{background:transparent;border:1px solid #3a3d50;color:#8890a4}.btn-outline:hover{border-color:#c9952b;color:#c9952b}
.btn-sm{padding:4px 10px;font-size:11px}
.btn-group{display:flex;gap:8px;margin-bottom:16px;flex-wrap:wrap}

/* Mark collected button in amort table */
.collect-btn{background:#2ecc71;color:#fff;border:none;border-radius:4px;padding:2px 8px;font-size:11px;cursor:pointer;font-weight:600}
.collect-btn:hover{background:#27ae60}
.undo-btn{background:transparent;border:1px solid #555;color:#888;border-radius:4px;padding:2px 8px;font-size:11px;cursor:pointer}
.undo-btn:hover{border-color:#e74c3c;color:#e74c3c}

/* Toast */
.toast{position:fixed;bottom:20px;right:20px;padding:10px 20px;border-radius:8px;font-size:13px;font-weight:500;z-index:300;opacity:0;transition:opacity .3s;pointer-events:none}
.toast.show{opacity:1}
.toast.saving{background:#1a1d2e;color:#c9952b;border:1px solid #c9952b}
.toast.saved{background:#1a1d2e;color:#2ecc71;border:1px solid #2ecc71}
.toast.error{background:#1a1d2e;color:#e74c3c;border:1px solid #e74c3c}

@media(max-width:900px){
  .chart-row,.amort-panel .amort-charts{grid-template-columns:1fr}
  .kpi-row{grid-template-columns:repeat(2,1fr)}
  .amort-panel .detail-grid{grid-template-columns:1fr 1fr}
  .amort-panel{width:100vw}
}
</style>
</head>
<body>

<div class="header">
  <div>
    <h1>Treze Alcove <span>Loan Servicing</span></h1>
  </div>
  <div class="hdr-right">
    <div class="user-sel">
      <label>User</label>
      <select id="userSelect">
        <option value="">Select...</option>
        <option>Josh</option><option>Kristen</option><option>Kinzie</option><option>Jill</option><option>Lauren</option>
      </select>
    </div>
    <div>
      <div class="date">As of ${todayDisplay}</div>
      <div class="sync-status" id="syncStatus"></div>
    </div>
  </div>
</div>

<div class="tabs">
  <div class="tab active" data-tab="overview">Overview</div>
  <div class="tab" data-tab="payments">Payment Status</div>
  <div class="tab" data-tab="collections">Collections<span class="badge" id="collectBadge">0</span></div>
  <div class="tab" data-tab="insurance">Insurance</div>
  <div class="tab" data-tab="cashflow">Cash Flow</div>
</div>

<div class="content">
  <div class="tab-content active" id="tab-overview">
    <div class="kpi-row" id="kpiRow"></div>
    <div class="chart-row">
      <div class="chart-card"><h3>Portfolio Status</h3><div class="chart-wrap"><canvas id="statusChart"></canvas></div></div>
      <div class="chart-card"><h3>Aging Distribution (Days Past Due)</h3><div class="chart-wrap"><canvas id="agingChart"></canvas></div></div>
    </div>
    <div class="chart-row">
      <div class="chart-card"><h3>Balance by Closing</h3><div class="chart-wrap"><canvas id="closingChart"></canvas></div></div>
      <div class="chart-card"><h3>Balloon Maturity Timeline</h3><div class="chart-wrap"><canvas id="balloonChart"></canvas></div></div>
    </div>
  </div>

  <div class="tab-content" id="tab-payments">
    <div id="pmtFilterInfo"></div>
    <div class="filter-bar">
      <div><label>Status</label><br><select id="filterStatus"><option value="all">All Statuses</option><option value="Current">Current</option><option value="Past Due">Past Due</option><option value="Late Notice">Late Notice</option><option value="Demand Letter">Demand Letter</option><option value="Severely Delinquent">Severely Delinquent</option></select></div>
      <div><label>Closing</label><br><select id="filterClosing"><option value="all">All Closings</option></select></div>
      <div><label>Search</label><br><input type="text" id="filterSearch" placeholder="Address or borrower..."></div>
    </div>
    <div class="table-wrap">
      <table id="paymentTable"><thead><tr>
        <th data-sort="address">Address</th><th data-sort="closing">Closing</th><th data-sort="borrower">Borrower</th>
        <th data-sort="monthlyPmt">Payment</th><th data-sort="lastPmtDate">Last Payment</th><th data-sort="nextDueDate">Next Due</th>
        <th data-sort="daysPastDue">Days Past Due</th><th data-sort="currentBalance">Balance</th><th data-sort="status">Status</th>
      </tr></thead><tbody id="paymentBody"></tbody></table>
    </div>
  </div>

  <div class="tab-content" id="tab-collections">
    <div class="kpi-row" id="collectKpi"></div>
    <div class="btn-group">
      <button class="btn btn-gold" onclick="window.generateAllLateNotices()">Generate All Late Notices (PDF)</button>
      <button class="btn btn-red" onclick="window.generateAllDemandLetters()">Generate All Demand Letters (PDF)</button>
    </div>
    <h3 style="color:#fff;margin-bottom:15px">Action Items</h3>
    <div class="table-wrap">
      <table><thead><tr>
        <th>Address</th><th>Borrower</th><th>Days Past Due</th><th>Status</th><th>Collection Status</th><th>Action</th><th>Notes</th><th>Balance</th>
      </tr></thead><tbody id="collectBody"></tbody></table>
    </div>
  </div>

  <div class="tab-content" id="tab-insurance">
    <p style="color:#8890a4;margin-bottom:15px;font-size:13px">Insurance validity tracking. Click a card to edit policy details. Changes sync to Google Sheets.</p>
    <div class="kpi-row" id="insKpi"></div>
    <div class="filter-bar">
      <div><label>Status</label><br><select id="filterInsStatus"><option value="all">All</option><option value="valid">Valid</option><option value="expiring">Expiring Soon</option><option value="lapsed">Lapsed</option><option value="unknown">Unknown</option></select></div>
    </div>
    <div class="ins-grid" id="insGrid"></div>
  </div>

  <div class="tab-content" id="tab-cashflow">
    <div class="kpi-row" id="cfKpi"></div>
    <div class="chart-row">
      <div class="chart-card"><h3>Expected Monthly Collections</h3><div class="chart-wrap"><canvas id="cfChart"></canvas></div></div>
      <div class="chart-card"><h3>Balloon Payoffs by Year</h3><div class="chart-wrap"><canvas id="balloonPayoffChart"></canvas></div></div>
    </div>
    <h3 style="color:#fff;margin-bottom:15px">Upcoming Balloon Maturities</h3>
    <div class="table-wrap">
      <table><thead><tr><th>Address</th><th>Borrower</th><th>Balloon Date</th><th>Months Remaining</th><th>Est. Balance at Maturity</th><th>Monthly Payment</th></tr></thead><tbody id="balloonBody"></tbody></table>
    </div>
  </div>
</div>

<!-- Amortization schedule panel -->
<div class="amort-overlay" id="amortOverlay">
  <div class="amort-panel" id="amortPanel"></div>
</div>

<div class="toast" id="toast"></div>

<script>
const NOTES = ${dataJSON};
const TODAY = new Date('${todayStr}');
const API_URL = ''; // Set after deploying Apps Script

// ═══════ STATE ═══════
let chartFilter = null;
let sortCol = 'daysPastDue', sortDir = -1;
let statusData = {};
let amortCharts = [];
const COLL_STATUSES = [
  {v:'monitoring',l:'Monitoring'},{v:'late_notice_sent',l:'Late Notice Sent'},{v:'demand_letter_sent',l:'Demand Letter Sent'},
  {v:'in_collections',l:'In Collections'},{v:'payment_plan',l:'Payment Plan'},{v:'resolved',l:'Resolved'}
];

// ═══════ HELPERS ═══════
const fmt = n => n != null ? '$' + n.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2}) : '--';
const fmtK = n => '$' + (n/1000).toFixed(0) + 'K';
const fmtMM = n => '$' + (n/1000000).toFixed(1) + 'MM';
const fmtDate = d => d ? new Date(d+'T00:00:00').toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'}) : '--';
const fmtPct = n => (n*100).toFixed(2) + '%';
const statusClass = s => s.toLowerCase().replace(/\\s+/g,'-');
const getUser = () => document.getElementById('userSelect').value;

function getStatus(noteId, field, fallback) {
  return statusData[noteId] && statusData[noteId][field] !== undefined ? statusData[noteId][field] : (fallback || '');
}

// ═══════ TOAST ═══════
function showToast(msg, type, dur) {
  const t = document.getElementById('toast');
  t.textContent = msg; t.className = 'toast show ' + type;
  clearTimeout(t._tm);
  t._tm = setTimeout(() => t.classList.remove('show'), dur || 2000);
}

// ═══════ GOOGLE SHEETS API ═══════
async function loadStatuses() {
  if (!API_URL) return;
  try {
    document.getElementById('syncStatus').textContent = 'syncing...';
    document.getElementById('syncStatus').className = 'sync-status';
    const res = await fetch(API_URL + '?action=statuses');
    const rows = await res.json();
    statusData = {};
    if (Array.isArray(rows)) {
      rows.forEach(r => {
        if (!statusData[r.note_id]) statusData[r.note_id] = {};
        statusData[r.note_id][r.field_name] = r.field_value;
      });
    }
    document.getElementById('syncStatus').textContent = 'synced';
    document.getElementById('syncStatus').className = 'sync-status ok';
  } catch(e) {
    document.getElementById('syncStatus').textContent = 'offline';
    document.getElementById('syncStatus').className = 'sync-status err';
  }
}

async function postAPI(payload) {
  if (!API_URL) { showToast('API not configured -- using local only', 'error'); return; }
  showToast('Saving...', 'saving');
  try {
    await fetch(API_URL, { method: 'POST', body: JSON.stringify(payload) });
    showToast('Saved', 'saved');
  } catch(e) { showToast('Save failed', 'error', 4000); }
}

async function saveStatus(noteId, field, value) {
  if (!statusData[noteId]) statusData[noteId] = {};
  statusData[noteId][field] = value;
  await postAPI({ action: 'save_status', noteId, fieldName: field, fieldValue: value, updatedBy: getUser() });
}

// ═══════ TABS ═══════
document.querySelectorAll('.tab').forEach(tab => {
  tab.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(tc => tc.classList.remove('active'));
    tab.classList.add('active');
    document.getElementById('tab-'+tab.dataset.tab).classList.add('active');
  });
});

// ═══════ USER SELECTOR ═══════
const savedUser = localStorage.getItem('treze_user');
if (savedUser) document.getElementById('userSelect').value = savedUser;
document.getElementById('userSelect').addEventListener('change', function() {
  localStorage.setItem('treze_user', this.value);
});

// ═══════ CHART CLICK HELPERS ═══════
function setChartFilter(type, value) {
  chartFilter = {type, value};
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(tc => tc.classList.remove('active'));
  document.querySelector('[data-tab="payments"]').classList.add('active');
  document.getElementById('tab-payments').classList.add('active');
  renderPayments();
}
function clearChartFilter() { chartFilter = null; renderPayments(); }

function applyChartFilter(notes) {
  if (!chartFilter) return notes;
  const f = chartFilter;
  if (f.type === 'status') return notes.filter(n => n.status === f.value);
  if (f.type === 'closing') return notes.filter(n => '#'+n.closing === f.value);
  if (f.type === 'aging') {
    return notes.filter(n => {
      if (f.value === '0 (Current)') return n.daysPastDue === 0;
      if (f.value === '1-15') return n.daysPastDue >= 1 && n.daysPastDue <= 15;
      if (f.value === '16-30') return n.daysPastDue >= 16 && n.daysPastDue <= 30;
      if (f.value === '31-60') return n.daysPastDue >= 31 && n.daysPastDue <= 60;
      if (f.value === '60+') return n.daysPastDue > 60;
      return true;
    });
  }
  if (f.type === 'balloon') {
    return notes.filter(n => {
      const d = new Date(n.balloonDate);
      return d.getFullYear()+' Q'+(Math.floor(d.getMonth()/3)+1) === f.value;
    });
  }
  return notes;
}

// ═══════ OVERVIEW TAB ═══════
let overviewCharts = [];

function renderOverview() {
  overviewCharts.forEach(c => c.destroy());
  overviewCharts = [];

  const totalBal = NOTES.reduce((s,n) => s+n.currentBalance, 0);
  const totalMonthly = NOTES.reduce((s,n) => s+n.monthlyPmt, 0);
  const currentN = NOTES.filter(n => n.status==='Current').length;
  const actionN = NOTES.filter(n => n.status!=='Current').length;
  const avgDPD = NOTES.length ? Math.round(NOTES.reduce((s,n)=>s+n.daysPastDue,0)/NOTES.length) : 0;
  const soon = NOTES.filter(n => { const m=(new Date(n.balloonDate)-TODAY)/(1000*60*60*24*30); return m<=12&&m>0; }).length;

  document.getElementById('kpiRow').innerHTML = [
    {label:'Active Notes',value:NOTES.length,cls:''},
    {label:'Total Balance',value:fmt(totalBal),cls:''},
    {label:'Expected Monthly',value:fmt(totalMonthly),cls:''},
    {label:'Current',value:currentN,cls:'green',click:"setChartFilter('status','Current')"},
    {label:'Action Needed',value:actionN,cls:actionN>0?'red':'green',click:actionN>0?"setChartFilter('status','Past Due')":null},
    {label:'Avg Days Past Due',value:avgDPD,cls:avgDPD>15?'orange':'green'},
    {label:'Balloons in 12 Mo',value:soon,cls:soon>0?'yellow':''}
  ].map(k => '<div class="kpi'+(k.click?' clickable" onclick="'+k.click+'"':'"')+'><div class="label">'+k.label+'</div><div class="value '+k.cls+'">'+k.value+'</div></div>').join('');

  const statusCounts = {};
  NOTES.forEach(n => { statusCounts[n.status]=(statusCounts[n.status]||0)+1; });
  const sLabels = Object.keys(statusCounts);
  const sColors = sLabels.map(s=>({'Current':'#2ecc71','Past Due':'#f1c40f','Late Notice':'#e67e22','Demand Letter':'#e74c3c','Severely Delinquent':'#ff6b6b'}[s]||'#555'));
  overviewCharts.push(new Chart(document.getElementById('statusChart'),{
    type:'doughnut',
    data:{labels:sLabels,datasets:[{data:sLabels.map(s=>statusCounts[s]),backgroundColor:sColors,borderWidth:0}]},
    options:{responsive:true,maintainAspectRatio:false,onClick(e,els){if(els.length){setChartFilter('status',sLabels[els[0].index])}},plugins:{legend:{position:'right',labels:{color:'#8890a4',font:{size:12}}}}}
  }));

  const buckets={'0 (Current)':0,'1-15':0,'16-30':0,'31-60':0,'60+':0};
  NOTES.forEach(n=>{if(n.daysPastDue===0)buckets['0 (Current)']++;else if(n.daysPastDue<=15)buckets['1-15']++;else if(n.daysPastDue<=30)buckets['16-30']++;else if(n.daysPastDue<=60)buckets['31-60']++;else buckets['60+']++});
  overviewCharts.push(new Chart(document.getElementById('agingChart'),{
    type:'bar',
    data:{labels:Object.keys(buckets),datasets:[{data:Object.values(buckets),backgroundColor:['#2ecc71','#a8d854','#f1c40f','#e67e22','#e74c3c'],borderRadius:4}]},
    options:{responsive:true,maintainAspectRatio:false,onClick(e,els){if(els.length){setChartFilter('aging',Object.keys(buckets)[els[0].index])}},plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>ctx.raw+' notes'}}},scales:{y:{ticks:{color:'#8890a4',stepSize:1},grid:{color:'#2a2d3e'}},x:{ticks:{color:'#8890a4'},grid:{display:false}}}}
  }));

  const closingBal={};
  NOTES.forEach(n=>{const k='#'+n.closing;closingBal[k]=(closingBal[k]||0)+n.currentBalance});
  const cKeys=Object.keys(closingBal).sort((a,b)=>parseInt(a.slice(1))-parseInt(b.slice(1)));
  overviewCharts.push(new Chart(document.getElementById('closingChart'),{
    type:'bar',
    data:{labels:cKeys,datasets:[{label:'Balance',data:cKeys.map(k=>closingBal[k]),backgroundColor:'#c9952b',borderRadius:4}]},
    options:{responsive:true,maintainAspectRatio:false,onClick(e,els){if(els.length){setChartFilter('closing',cKeys[els[0].index])}},plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>fmt(ctx.raw)}}},scales:{y:{ticks:{color:'#8890a4',callback:v=>fmtK(v)},grid:{color:'#2a2d3e'}},x:{ticks:{color:'#8890a4'},grid:{display:false}}}}
  }));

  const balloonByQ={};
  NOTES.forEach(n=>{const d=new Date(n.balloonDate);const q=d.getFullYear()+' Q'+(Math.floor(d.getMonth()/3)+1);balloonByQ[q]=(balloonByQ[q]||0)+1});
  const bKeys=Object.keys(balloonByQ).sort();
  overviewCharts.push(new Chart(document.getElementById('balloonChart'),{
    type:'bar',
    data:{labels:bKeys,datasets:[{label:'Notes Maturing',data:bKeys.map(k=>balloonByQ[k]),backgroundColor:'#3498db',borderRadius:4}]},
    options:{responsive:true,maintainAspectRatio:false,onClick(e,els){if(els.length){setChartFilter('balloon',bKeys[els[0].index])}},plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>ctx.raw+' notes'}}},scales:{y:{ticks:{color:'#8890a4',stepSize:1},grid:{color:'#2a2d3e'}},x:{ticks:{color:'#8890a4',font:{size:10}},grid:{display:false}}}}
  }));
}

// ═══════ PAYMENT STATUS TAB ═══════
function renderPayments() {
  const info = document.getElementById('pmtFilterInfo');
  if (chartFilter) {
    info.innerHTML = '<div class="active-filter">Filtered by '+chartFilter.type+': <strong>'+chartFilter.value+'</strong> <span class="clear" onclick="clearChartFilter()">&times;</span></div>';
    info.style.marginBottom = '12px';
  } else { info.innerHTML = ''; info.style.marginBottom = '0'; }

  const statusF = document.getElementById('filterStatus').value;
  const closingF = document.getElementById('filterClosing').value;
  const searchF = document.getElementById('filterSearch').value.toLowerCase();

  let filtered = NOTES.filter(n => {
    if (statusF !== 'all' && n.status !== statusF) return false;
    if (closingF !== 'all' && n.closing !== parseInt(closingF)) return false;
    if (searchF && !n.address.toLowerCase().includes(searchF) && !n.borrower.toLowerCase().includes(searchF)) return false;
    return true;
  });
  filtered = applyChartFilter(filtered);
  filtered.sort((a,b) => {
    let va=a[sortCol],vb=b[sortCol];
    if(typeof va==='string') return sortDir*va.localeCompare(vb);
    return sortDir*((va||0)-(vb||0));
  });

  document.getElementById('paymentBody').innerHTML = filtered.map(n =>
    '<tr onclick="window.showAmortSchedule(\\''+n.id.replace(/'/g,"\\\\'")+'\\')">'+
    '<td>'+n.address+'</td>'+
    '<td class="num">#'+n.closing+'</td>'+
    '<td>'+n.borrower.substring(0,35)+(n.borrower.length>35?'...':'')+'</td>'+
    '<td class="num">'+fmt(n.monthlyPmt)+'</td>'+
    '<td>'+fmtDate(n.lastPmtDate)+'</td>'+
    '<td>'+fmtDate(n.nextDueDate)+'</td>'+
    '<td class="num">'+(n.daysPastDue>0?'<span style="color:'+(n.daysPastDue>30?'#e74c3c':n.daysPastDue>15?'#e67e22':'#f1c40f')+'">'+n.daysPastDue+'</span>':'0')+'</td>'+
    '<td class="num">'+fmt(n.currentBalance)+'</td>'+
    '<td><span class="status '+statusClass(n.status)+'">'+n.status+'</span></td>'+
    '</tr>'
  ).join('');
}

const closings=[...new Set(NOTES.map(n=>n.closing))].sort((a,b)=>a-b);
closings.forEach(c=>{document.getElementById('filterClosing').innerHTML+='<option value="'+c+'">#'+c+'</option>'});
document.getElementById('filterStatus').addEventListener('change',renderPayments);
document.getElementById('filterClosing').addEventListener('change',renderPayments);
document.getElementById('filterSearch').addEventListener('input',renderPayments);
document.querySelectorAll('#paymentTable th[data-sort]').forEach(th=>{
  th.addEventListener('click',()=>{
    const col=th.dataset.sort;
    if(sortCol===col)sortDir*=-1;else{sortCol=col;sortDir=col==='daysPastDue'?-1:1}
    renderPayments();
  });
});

// ═══════ COLLECTIONS TAB ═══════
function renderCollections() {
  const actionNotes = NOTES.filter(n => n.daysPastDue > 0);
  const lateNotice = actionNotes.filter(n => n.daysPastDue > 0 && n.daysPastDue <= 30);
  const demandLetter = actionNotes.filter(n => n.daysPastDue > 30);
  const totalPastDueBal = actionNotes.reduce((s,n) => s+n.currentBalance, 0);

  document.getElementById('collectBadge').textContent = actionNotes.length;
  document.getElementById('collectKpi').innerHTML = [
    {label:'Total Past Due',value:actionNotes.length,cls:actionNotes.length>0?'red':'green'},
    {label:'Late Notice (1-30 days)',value:lateNotice.length,cls:lateNotice.length>0?'orange':'green'},
    {label:'Demand Letter (30+ days)',value:demandLetter.length,cls:demandLetter.length>0?'red':'green'},
    {label:'Past Due Balance',value:fmt(totalPastDueBal),cls:'red'}
  ].map(k=>'<div class="kpi"><div class="label">'+k.label+'</div><div class="value '+k.cls+'">'+k.value+'</div></div>').join('');

  document.getElementById('collectBody').innerHTML = actionNotes
    .sort((a,b)=>b.daysPastDue-a.daysPastDue)
    .map(n => {
      const nid = n.id.replace(/'/g,"\\\\'");
      const cs = getStatus(n.id, 'collection_status', 'monitoring');
      const cn = getStatus(n.id, 'collection_notes', '');

      let actionBtns = '';
      if (n.daysPastDue > 30) {
        actionBtns = '<button class="btn btn-red btn-sm" onclick="event.stopPropagation();window.generateDemandLetter(\\''+nid+'\\')">Demand Letter</button>';
      } else if (n.daysPastDue > 0) {
        actionBtns = '<button class="btn btn-gold btn-sm" onclick="event.stopPropagation();window.generateLateNotice(\\''+nid+'\\')">Late Notice</button>';
      }

      return '<tr class="action-needed" onclick="window.showAmortSchedule(\\''+nid+'\\')">'+
        '<td>'+n.address+'</td>'+
        '<td>'+n.borrower.substring(0,30)+'</td>'+
        '<td class="num" style="color:'+(n.daysPastDue>30?'#e74c3c':'#e67e22')+';font-weight:600">'+n.daysPastDue+'</td>'+
        '<td><span class="status '+statusClass(n.status)+'">'+n.status+'</span></td>'+
        '<td onclick="event.stopPropagation()"><select class="edit-field coll-status '+cs+'" onchange="window.saveCollStatus(\\''+nid+'\\',this.value,this)">'+
          COLL_STATUSES.map(s=>'<option value="'+s.v+'"'+(s.v===cs?' selected':'')+'>'+s.l+'</option>').join('')+
        '</select></td>'+
        '<td onclick="event.stopPropagation()">'+actionBtns+'</td>'+
        '<td onclick="event.stopPropagation()"><input class="edit-field" value="'+cn.replace(/"/g,'&quot;')+'" placeholder="Add note..." onblur="window.saveCollNote(\\''+nid+'\\',this.value)"></td>'+
        '<td class="num">'+fmt(n.currentBalance)+'</td>'+
        '</tr>';
    }).join('');
}

window.saveCollStatus = function(id, val, el) {
  el.className = 'edit-field coll-status ' + val;
  saveStatus(id, 'collection_status', val);
};
window.saveCollNote = function(id, val) {
  saveStatus(id, 'collection_notes', val);
};

// ═══════ PDF LETTER GENERATION (jsPDF) ═══════
const LETTERHEAD = {
  company: 'TREZE ALCOVE',
  address1: '810 TEXAS AVE.',
  cityStateZip: 'LUBBOCK, TX 79401',
  payTo: 'Treze Alcove, LLC',
  payAddr1: 'PO Box 2078',
  payAddr2: 'Abilene, TX 79604',
  signer: 'Kristen A. Gil, CPA',
  signerTitle: 'Chief Financial Officer',
  signerEntity: 'Treze Alcove, LLC',
  contactEmail: 'kristen@platformenergy.com',
  contactPhone: '(325) 232-7813'
};

function buildLetterHeader(doc) {
  // Letterhead
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(22);
  doc.text(LETTERHEAD.company, 105, 25, { align: 'center' });
  doc.setFontSize(13);
  doc.text('- ' + LETTERHEAD.address1 + ' -', 105, 33, { align: 'center' });
  doc.text(LETTERHEAD.cityStateZip, 105, 40, { align: 'center' });

  // Gold line
  doc.setDrawColor(201, 149, 43);
  doc.setLineWidth(1);
  doc.line(20, 45, 190, 45);

  return 55; // y position after header
}

function fmtDateLong(dateStr) {
  return new Date(dateStr + 'T00:00:00').toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' });
}

function fmtMoney(n) {
  return '$' + n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

window.generateLateNotice = function(noteId) {
  const n = NOTES.find(x => x.id === noteId);
  if (!n) return;

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: 'mm', format: 'letter' });

  let y = buildLetterHeader(doc);

  // Date (right aligned)
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(11);
  const todayLong = new Date().toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' });
  doc.text(todayLong, 190, y, { align: 'right' });
  y += 14;

  // Recipient
  const recipientName = n.ownerName || n.contact || n.borrower;
  const recipientCompany = n.borrower !== recipientName ? n.borrower : '';
  if (recipientCompany) { doc.text(recipientCompany, 20, y); y += 5; }
  doc.text('Attn: ' + recipientName, 20, y); y += 5;
  if (n.mailingAddress) {
    const addrLines = n.mailingAddress.split(',').map(s => s.trim());
    addrLines.forEach(line => { doc.text(line, 20, y); y += 5; });
  }
  y += 5;

  // RE line
  doc.text('RE: ' + n.address + ', Wolfforth, TX 79382', 20, y);
  y += 10;

  // Salutation
  const saluteName = recipientName.split(' ').pop();
  doc.text('Dear ' + recipientName + ',', 20, y);
  y += 10;

  // Body
  const dueDate = fmtDateLong(n.nextDueDate);
  const lateFee = n.lateFee;
  const totalOwed = n.totalOwed;

  const body1 = 'Your payment of ' + fmtMoney(n.monthlyPmt) + ' due on ' + dueDate + ' is past due.';
  doc.text(body1, 20, y, { maxWidth: 170 });
  y += 8;

  doc.text('Please make arrangements to pay your outstanding balance plus a late fee of ' + fmtMoney(lateFee) + '.', 20, y, { maxWidth: 170 });
  y += 6;

  doc.setFont('helvetica', 'bold');
  doc.text('The total you owe is ' + fmtMoney(totalOwed) + '.', 20, y, { maxWidth: 170 });
  doc.setFont('helvetica', 'normal');
  y += 6;

  // Payment deadline — end of current month
  const endOfMonth = new Date(TODAY.getFullYear(), TODAY.getMonth() + 1, 0);
  const deadline = endOfMonth.toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' });
  doc.text('You may make this payment with a check, wire, or ACH. Payment is due no later than ' + deadline + ',', 20, y, { maxWidth: 170 });
  y += 6;
  doc.text('or your loan will be in default and additional fees will be incurred.', 20, y, { maxWidth: 170 });
  y += 12;

  // Payment address
  doc.text('Please forward payment to the following address:', 20, y);
  y += 8;
  doc.text(LETTERHEAD.payTo, 20, y); y += 5;
  doc.text(LETTERHEAD.payAddr1, 20, y); y += 5;
  doc.text(LETTERHEAD.payAddr2, 20, y); y += 12;

  // Contact
  doc.text('Should you have any questions, please contact me at ' + LETTERHEAD.contactEmail + ' or ' + LETTERHEAD.contactPhone + '.', 20, y, { maxWidth: 170 });
  y += 18;

  // Signature
  doc.setFont('helvetica', 'italic');
  doc.text(LETTERHEAD.signer.split(' ')[0], 20, y);
  y += 6;
  doc.setFont('helvetica', 'normal');
  doc.text(LETTERHEAD.signer, 20, y); y += 5;
  doc.text(LETTERHEAD.signerTitle, 20, y); y += 5;
  doc.text(LETTERHEAD.signerEntity, 20, y);

  // Save
  const filename = 'Late Notice - ' + n.address.replace(/[^a-zA-Z0-9 ]/g, '') + '.pdf';
  doc.save(filename);
  showToast('Late Notice generated: ' + n.address, 'saved', 3000);

  // Log action
  postAPI({ action: 'log_action', updatedBy: getUser(), actionType: 'late_notice_generated', noteId: n.id, details: 'Total owed: ' + fmtMoney(totalOwed) });
};

window.generateDemandLetter = function(noteId) {
  const n = NOTES.find(x => x.id === noteId);
  if (!n) return;

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: 'mm', format: 'letter' });

  let y = buildLetterHeader(doc);

  // Date
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(11);
  const todayLong = new Date().toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' });
  doc.text(todayLong, 190, y, { align: 'right' });
  y += 14;

  // Recipient
  const recipientName = n.ownerName || n.contact || n.borrower;
  const recipientCompany = n.borrower !== recipientName ? n.borrower : '';
  if (recipientCompany) { doc.text(recipientCompany, 20, y); y += 5; }
  doc.text('Attn: ' + recipientName, 20, y); y += 5;
  if (n.mailingAddress) {
    const addrLines = n.mailingAddress.split(',').map(s => s.trim());
    addrLines.forEach(line => { doc.text(line, 20, y); y += 5; });
  }
  y += 5;

  // RE
  doc.text('RE: ' + n.address + ', Wolfforth, TX 79382', 20, y);
  y += 10;

  // Title
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(12);
  doc.text('NOTICE OF DEFAULT AND INTENT TO ACCELERATE', 105, y, { align: 'center' });
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(11);
  y += 10;

  // Body
  doc.text('Dear ' + recipientName + ',', 20, y);
  y += 8;

  const noteDate = fmtDateLong(n.beginDate);
  doc.text('This letter is regarding the promissory note dated ' + noteDate + ' in the original principal', 20, y, { maxWidth: 170 });
  y += 5;
  doc.text('amount of ' + fmtMoney(n.loanAmt) + ' (the "Note"), secured by a Deed of Trust on the property located at:', 20, y, { maxWidth: 170 });
  y += 10;

  doc.setFont('helvetica', 'bold');
  doc.text(n.address + ', Wolfforth, TX 79382', 30, y);
  doc.setFont('helvetica', 'normal');
  y += 10;

  doc.text('You are in default under the terms of the Note for failure to make your regularly scheduled', 20, y, { maxWidth: 170 });
  y += 5;
  doc.text('monthly payment(s). As of the date of this letter, your account is ' + n.daysPastDue + ' days past due.', 20, y, { maxWidth: 170 });
  y += 10;

  // Arrearage calculation
  const missedMonths = Math.ceil(n.daysPastDue / 30);
  const missedPmts = missedMonths * n.monthlyPmt;
  const lateFees = missedMonths * n.lateFee;
  const totalArrearage = missedPmts + lateFees;

  doc.text('Arrearage Summary:', 20, y);
  y += 6;
  doc.text('  Missed payments (' + missedMonths + ' month' + (missedMonths > 1 ? 's' : '') + '):   ' + fmtMoney(missedPmts), 25, y); y += 5;
  doc.text('  Late fees:                             ' + fmtMoney(lateFees), 25, y); y += 5;
  doc.setFont('helvetica', 'bold');
  doc.text('  Total amount due:                 ' + fmtMoney(totalArrearage), 25, y);
  doc.setFont('helvetica', 'normal');
  y += 10;

  // Cure period
  const cureDate = new Date();
  cureDate.setDate(cureDate.getDate() + 30);
  const cureDateStr = cureDate.toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' });

  doc.text('You have the right to cure the default on your loan. In order to cure the default, you must pay', 20, y, { maxWidth: 170 });
  y += 5;
  doc.text('the above amount to Treze Alcove, LLC on or before ' + cureDateStr + '.', 20, y, { maxWidth: 170 });
  y += 10;

  doc.setFont('helvetica', 'bold');
  doc.text('If the amount set forth above is not received in full by the date set forth above, the loan', 20, y, { maxWidth: 170 });
  y += 5;
  doc.text('will be accelerated and formal legal proceedings to foreclose the property may be initiated.', 20, y, { maxWidth: 170 });
  doc.setFont('helvetica', 'normal');
  y += 10;

  // Payment instructions
  doc.text('Payments must be in the form of certified funds, cashier\\'s check, or wire transfer.', 20, y, { maxWidth: 170 });
  y += 8;
  doc.text('Send payments to:', 20, y); y += 6;
  doc.setFont('helvetica', 'bold');
  doc.text(LETTERHEAD.payTo, 25, y); y += 5;
  doc.text('749 Gateway Street, Suite 203', 25, y); y += 5;
  doc.text('Abilene, TX 79602', 25, y);
  doc.setFont('helvetica', 'normal');
  y += 10;

  // Contact
  doc.text('Should you have any questions or wish to make arrangements for payment, please contact', 20, y, { maxWidth: 170 });
  y += 5;
  doc.text(LETTERHEAD.contactEmail + ' or ' + LETTERHEAD.contactPhone + '.', 20, y);
  y += 15;

  // Signature
  doc.text('Sincerely,', 20, y); y += 10;
  doc.setFont('helvetica', 'italic');
  doc.text(LETTERHEAD.signer.split(' ')[0], 20, y);
  y += 6;
  doc.setFont('helvetica', 'normal');
  doc.text(LETTERHEAD.signer, 20, y); y += 5;
  doc.text(LETTERHEAD.signerTitle, 20, y); y += 5;
  doc.text(LETTERHEAD.signerEntity, 20, y);

  const filename = 'Demand Letter - ' + n.address.replace(/[^a-zA-Z0-9 ]/g, '') + '.pdf';
  doc.save(filename);
  showToast('Demand Letter generated: ' + n.address, 'saved', 3000);

  postAPI({ action: 'log_action', updatedBy: getUser(), actionType: 'demand_letter_generated', noteId: n.id, details: 'Arrearage: ' + fmtMoney(totalArrearage) });
};

// Batch PDF generation
window.generateAllLateNotices = function() {
  const lateNotes = NOTES.filter(n => n.daysPastDue > 0 && n.daysPastDue <= 30);
  if (lateNotes.length === 0) { showToast('No notes qualify for late notice', 'error'); return; }
  lateNotes.forEach(n => window.generateLateNotice(n.id));
  showToast('Generated ' + lateNotes.length + ' late notice(s)', 'saved', 3000);
};

window.generateAllDemandLetters = function() {
  const demandNotes = NOTES.filter(n => n.daysPastDue > 30);
  if (demandNotes.length === 0) { showToast('No notes qualify for demand letter', 'error'); return; }
  demandNotes.forEach(n => window.generateDemandLetter(n.id));
  showToast('Generated ' + demandNotes.length + ' demand letter(s)', 'saved', 3000);
};

// ═══════ INSURANCE TAB ═══════
function getInsuranceStatus(n) {
  const exp = getStatus(n.id, 'insurance_expiration') || n.insuranceExpiration;
  const carrier = getStatus(n.id, 'insurance_carrier', '');
  const policy = getStatus(n.id, 'insurance_policy_num', '');
  if (!exp) return {status:'unknown',label:'No data on file',daysUntil:null,carrier,policy,exp:''};
  const expDate = new Date(exp);
  const daysUntil = Math.floor((expDate-TODAY)/(1000*60*60*24));
  if (daysUntil < 0) return {status:'lapsed',label:'Lapsed '+Math.abs(daysUntil)+' days ago',daysUntil,carrier,policy,exp};
  if (daysUntil <= 30) return {status:'expiring',label:'Expires in '+daysUntil+' days',daysUntil,carrier,policy,exp};
  return {status:'valid',label:'Valid through '+fmtDate(exp),daysUntil,carrier,policy,exp};
}

function renderInsurance() {
  const statuses = NOTES.map(n => ({...n, ins: getInsuranceStatus(n)}));
  const valid=statuses.filter(s=>s.ins.status==='valid').length;
  const expiring=statuses.filter(s=>s.ins.status==='expiring').length;
  const lapsed=statuses.filter(s=>s.ins.status==='lapsed').length;
  const unknown=statuses.filter(s=>s.ins.status==='unknown').length;

  document.getElementById('insKpi').innerHTML = [
    {label:'Valid',value:valid,cls:'green'},
    {label:'Expiring Soon',value:expiring,cls:'yellow'},
    {label:'Lapsed',value:lapsed,cls:lapsed>0?'red':'green'},
    {label:'Unknown',value:unknown,cls:unknown>0?'orange':''}
  ].map(k=>'<div class="kpi"><div class="label">'+k.label+'</div><div class="value '+k.cls+'">'+k.value+'</div></div>').join('');

  const filterVal=document.getElementById('filterInsStatus').value;
  const filtered=filterVal==='all'?statuses:statuses.filter(s=>s.ins.status===filterVal);

  document.getElementById('insGrid').innerHTML = filtered.map(s => {
    const nid = s.id.replace(/'/g,"\\\\'");
    return '<div class="ins-card">'+
      '<div style="display:flex;justify-content:space-between;align-items:start">'+
        '<div><div class="ins-addr" style="cursor:pointer" onclick="window.showAmortSchedule(\\''+nid+'\\')">'+s.address+'</div>'+
        '<div class="ins-borrower">'+s.borrower.substring(0,45)+' | Closing #'+s.closing+'</div></div>'+
        '<span class="status '+statusClass(s.status)+'">'+s.status+'</span>'+
      '</div>'+
      '<div class="ins-status"><span class="ins-dot '+s.ins.status+'"></span><span style="font-size:12px">'+s.ins.label+'</span></div>'+
      '<div class="ins-fields">'+
        '<div><label>Expiration</label><br><input type="date" class="edit-field" value="'+s.ins.exp+'" onchange="window.saveIns(\\''+nid+'\\',\\'insurance_expiration\\',this.value)"></div>'+
        '<div><label>Carrier</label><br><input class="edit-field" value="'+s.ins.carrier.replace(/"/g,'&quot;')+'" placeholder="Carrier name..." onblur="window.saveIns(\\''+nid+'\\',\\'insurance_carrier\\',this.value)"></div>'+
        '<div><label>Policy #</label><br><input class="edit-field" value="'+s.ins.policy.replace(/"/g,'&quot;')+'" placeholder="Policy number..." onblur="window.saveIns(\\''+nid+'\\',\\'insurance_policy_num\\',this.value)"></div>'+
      '</div>'+
      (s.ins.status==='lapsed'&&s.ins.daysUntil!==null&&Math.abs(s.ins.daysUntil)>30?'<div class="ins-warn">30+ days lapsed -- send demand letter for insurance</div>':'')+
    '</div>';
  }).join('');
}

window.saveIns = function(id, field, val) {
  saveStatus(id, field, val);
  renderInsurance();
};
document.getElementById('filterInsStatus').addEventListener('change', renderInsurance);

// ═══════ CASH FLOW TAB ═══════
let cfCharts = [];

function renderCashFlow() {
  cfCharts.forEach(c=>c.destroy());
  cfCharts=[];

  const totalMonthly=NOTES.reduce((s,n)=>s+n.monthlyPmt,0);
  const annualExpected=totalMonthly*12;
  const totalBal=NOTES.reduce((s,n)=>s+n.currentBalance,0);

  const months=[];
  for(let i=0;i<12;i++){
    const d=new Date(TODAY);d.setMonth(d.getMonth()+i);
    const label=d.toLocaleDateString('en-US',{month:'short',year:'2-digit'});
    let expected=0,balloonAmt=0;
    NOTES.forEach(n=>{
      const bd=new Date(n.balloonDate);
      if(bd.getFullYear()===d.getFullYear()&&bd.getMonth()===d.getMonth()) balloonAmt+=n.currentBalance;
      else if(bd>d) expected+=n.monthlyPmt;
    });
    months.push({label,expected,balloon:balloonAmt});
  }

  document.getElementById('cfKpi').innerHTML = [
    {label:'Monthly Collections',value:fmt(totalMonthly)},{label:'Annual Expected',value:fmt(annualExpected)},
    {label:'Total Outstanding',value:fmt(totalBal)},{label:'Active Notes',value:NOTES.length}
  ].map(k=>'<div class="kpi"><div class="label">'+k.label+'</div><div class="value">'+k.value+'</div></div>').join('');

  cfCharts.push(new Chart(document.getElementById('cfChart'),{
    type:'bar',
    data:{labels:months.map(m=>m.label),datasets:[
      {label:'Monthly Payments',data:months.map(m=>m.expected),backgroundColor:'#2ecc71',borderRadius:4},
      {label:'Balloon Payoffs',data:months.map(m=>m.balloon),backgroundColor:'#c9952b',borderRadius:4}
    ]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#8890a4'}}},scales:{y:{ticks:{color:'#8890a4',callback:v=>fmtK(v)},grid:{color:'#2a2d3e'}},x:{ticks:{color:'#8890a4'},grid:{display:false}}}}
  }));

  const balloonYears={};
  NOTES.forEach(n=>{const y=new Date(n.balloonDate).getFullYear();balloonYears[y]=(balloonYears[y]||0)+n.currentBalance});
  const yKeys=Object.keys(balloonYears).sort();
  cfCharts.push(new Chart(document.getElementById('balloonPayoffChart'),{
    type:'bar',
    data:{labels:yKeys,datasets:[{label:'Balloon Balance',data:yKeys.map(k=>balloonYears[k]),backgroundColor:'#3498db',borderRadius:4}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>fmt(ctx.raw)}}},scales:{y:{ticks:{color:'#8890a4',callback:v=>fmtMM(v)},grid:{color:'#2a2d3e'}},x:{ticks:{color:'#8890a4'},grid:{display:false}}}}
  }));

  const sorted=[...NOTES].sort((a,b)=>new Date(a.balloonDate)-new Date(b.balloonDate));
  document.getElementById('balloonBody').innerHTML = sorted.map(n=>{
    const monthsLeft=Math.round((new Date(n.balloonDate)-TODAY)/(1000*60*60*24*30.44));
    const warn=monthsLeft<=12;
    return '<tr onclick="window.showAmortSchedule(\\''+n.id.replace(/'/g,"\\\\'")+'\\')">'+
      '<td>'+n.address+'</td><td>'+n.borrower.substring(0,35)+'</td>'+
      '<td'+(warn?' class="balloon-warn"':'')+'>'+fmtDate(n.balloonDate)+'</td>'+
      '<td class="num"'+(warn?' style="color:#e67e22;font-weight:600"':'')+'>'+monthsLeft+'</td>'+
      '<td class="num">'+fmt(n.currentBalance)+'</td><td class="num">'+fmt(n.monthlyPmt)+'</td></tr>';
  }).join('');
}

// ═══════ AMORTIZATION SCHEDULE PANEL ═══════
window.showAmortSchedule = function(id) {
  const n = NOTES.find(x => x.id === id);
  if (!n) return;

  amortCharts.forEach(c => c.destroy());
  amortCharts = [];

  const all = [...n.collectedPayments, ...n.scheduledPayments];
  const generalNotes = getStatus(n.id, 'general_notes', '');
  const nid = n.id.replace(/'/g, "\\\\'");

  const panel = document.getElementById('amortPanel');
  panel.innerHTML =
    '<div class="close-btn" onclick="document.getElementById(\\'amortOverlay\\').classList.remove(\\'show\\')">&times;</div>'+
    '<h2>'+n.address+' <span style="color:#c9952b">#'+n.closing+'</span></h2>'+
    '<div class="amort-sub">'+n.borrower+'</div>'+

    // Action buttons in panel
    '<div class="btn-group">'+
      (n.daysPastDue > 30 ? '<button class="btn btn-red btn-sm" onclick="window.generateDemandLetter(\\''+nid+'\\')">Generate Demand Letter</button>' : '')+
      (n.daysPastDue > 0 && n.daysPastDue <= 30 ? '<button class="btn btn-gold btn-sm" onclick="window.generateLateNotice(\\''+nid+'\\')">Generate Late Notice</button>' : '')+
    '</div>'+

    '<div class="detail-grid">'+
      '<div class="detail-item"><div class="dl">Loan Amount</div><div class="dv">'+fmt(n.loanAmt)+'</div></div>'+
      '<div class="detail-item"><div class="dl">Interest Rate</div><div class="dv">'+fmtPct(n.rate)+'</div></div>'+
      '<div class="detail-item"><div class="dl">Monthly Payment</div><div class="dv">'+fmt(n.monthlyPmt)+'</div></div>'+
      '<div class="detail-item"><div class="dl">Current Balance</div><div class="dv">'+fmt(n.currentBalance)+'</div></div>'+
      '<div class="detail-item"><div class="dl">Begin Date</div><div class="dv">'+fmtDate(n.beginDate)+'</div></div>'+
      '<div class="detail-item"><div class="dl">Balloon Date</div><div class="dv">'+fmtDate(n.balloonDate)+'</div></div>'+
      '<div class="detail-item"><div class="dl">Last Payment</div><div class="dv">'+fmtDate(n.lastPmtDate)+'</div></div>'+
      '<div class="detail-item"><div class="dl">Days Past Due</div><div class="dv" style="color:'+(n.daysPastDue>30?'#e74c3c':n.daysPastDue>0?'#e67e22':'#2ecc71')+'">'+n.daysPastDue+'</div></div>'+
      '<div class="detail-item"><div class="dl">Status</div><div class="dv"><span class="status '+statusClass(n.status)+'">'+n.status+'</span></div></div>'+
      '<div class="detail-item"><div class="dl">Payments Made</div><div class="dv">'+n.totalCollected+' of '+(n.totalCollected+n.remainingPayments)+'</div></div>'+
      '<div class="detail-item"><div class="dl">Gain on Sale</div><div class="dv">'+(n.gain?fmt(n.gain):'--')+'</div></div>'+
      '<div class="detail-item"><div class="dl">GP %</div><div class="dv">'+(n.gpPct?fmtPct(n.gpPct):'--')+'</div></div>'+
    '</div>'+

    // Contact info section
    '<div style="background:#1a1d2e;border-radius:8px;padding:16px;border:1px solid #2a2d3e;margin-bottom:18px">'+
      '<h3 style="font-size:13px;color:#fff;margin-bottom:10px">Contact Information</h3>'+
      '<div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;font-size:12px">'+
        '<div><span style="color:#8890a4">Owner:</span> '+(n.ownerName||'--')+'</div>'+
        '<div><span style="color:#8890a4">Contact:</span> '+(n.contact||'--')+'</div>'+
        '<div><span style="color:#8890a4">Email:</span> '+(n.email||'--')+'</div>'+
        '<div><span style="color:#8890a4">Phone:</span> '+(n.phone||'--')+'</div>'+
        '<div><span style="color:#8890a4">Payment Method:</span> '+(n.paymentMethod||'--')+'</div>'+
        '<div><span style="color:#8890a4">Mailing:</span> '+(n.mailingAddress||'--')+'</div>'+
      '</div>'+
    '</div>'+

    '<div class="amort-charts">'+
      '<div class="amort-chart-card"><h3>Payment Breakdown (Interest vs Principal)</h3><div class="amort-chart-wrap"><canvas id="amortBreakdown"></canvas></div></div>'+
      '<div class="amort-chart-card"><h3>Remaining Balance</h3><div class="amort-chart-wrap"><canvas id="amortBalance"></canvas></div></div>'+
    '</div>'+

    '<div class="notes-section">'+
      '<h3>Notes</h3>'+
      '<textarea class="edit-field" rows="3" placeholder="General notes about this note..." onblur="window.saveGeneralNote(\\''+nid+'\\',this.value)">'+generalNotes.replace(/</g,'&lt;')+'</textarea>'+
    '</div>'+

    '<h3 style="color:#fff;margin:20px 0 10px">Full Amortization Schedule</h3>'+
    '<div class="table-wrap" style="max-height:400px">'+
      '<table class="amort-table"><thead><tr>'+
        '<th>#</th><th>Date</th><th>Payment</th><th>Interest</th><th>Principal</th><th>Balance</th><th>Status</th>'+
      '</tr></thead><tbody>'+
      n.collectedPayments.map((p,i) =>
        '<tr class="collected"><td class="num">'+(p.month||i+1)+'</td><td>'+fmtDate(p.date)+'</td>'+
        '<td class="num">'+fmt(p.payment)+'</td><td class="num">'+fmt(p.interest)+'</td>'+
        '<td class="num">'+fmt(p.principal)+'</td><td class="num">'+(p.balance!=null?fmt(p.balance):'--')+'</td>'+
        '<td><span style="color:#2ecc71;font-size:11px">Collected</span></td></tr>'
      ).join('')+
      '<tr class="divider"><td colspan="7">Scheduled (Remaining)</td></tr>'+
      n.scheduledPayments.slice(0,3).map((p,i) =>
        '<tr class="scheduled"><td class="num">'+(p.month||(n.totalCollected+i+1))+'</td><td>'+fmtDate(p.date)+'</td>'+
        '<td class="num">'+fmt(p.payment)+'</td><td class="num">'+fmt(p.interest)+'</td>'+
        '<td class="num">'+fmt(p.principal)+'</td><td class="num">'+(p.balance!=null?fmt(p.balance):'--')+'</td>'+
        '<td><button class="collect-btn" onclick="event.stopPropagation();window.markCollected(\\''+nid+'\\',\\''+p.date+'\\','+p.payment+')">Mark Collected</button></td></tr>'
      ).join('')+
      n.scheduledPayments.slice(3).map((p,i) =>
        '<tr class="scheduled"><td class="num">'+(p.month||(n.totalCollected+i+4))+'</td><td>'+fmtDate(p.date)+'</td>'+
        '<td class="num">'+fmt(p.payment)+'</td><td class="num">'+fmt(p.interest)+'</td>'+
        '<td class="num">'+fmt(p.principal)+'</td><td class="num">'+(p.balance!=null?fmt(p.balance):'--')+'</td>'+
        '<td style="color:#555;font-size:11px">Scheduled</td></tr>'
      ).join('')+
    '</tbody></table></div>';

  // Charts
  const labels = all.map((p,i) => i+1);
  const interestData = all.map(p => p.interest);
  const principalData = all.map(p => p.principal);
  const balanceData = all.map(p => p.balance);
  const todayIdx = n.collectedPayments.length;

  amortCharts.push(new Chart(document.getElementById('amortBreakdown'), {
    type: 'line',
    data: {
      labels,
      datasets: [
        {label:'Interest',data:interestData,backgroundColor:'rgba(231,76,60,.3)',borderColor:'#e74c3c',fill:true,pointRadius:0,tension:.3},
        {label:'Principal',data:principalData,backgroundColor:'rgba(46,204,113,.3)',borderColor:'#2ecc71',fill:true,pointRadius:0,tension:.3}
      ]
    },
    options: {
      responsive:true,maintainAspectRatio:false,
      plugins:{legend:{labels:{color:'#8890a4'}}},
      scales:{
        y:{stacked:true,ticks:{color:'#8890a4',callback:v=>fmt(v)},grid:{color:'#2a2d3e'}},
        x:{ticks:{color:'#8890a4',maxTicksLimit:12,callback:function(v){return 'Mo '+this.getLabelForValue(v)}},grid:{display:false}}
      }
    },
    plugins: [{
      id: 'todayLine',
      afterDraw(chart) {
        if (todayIdx <= 0) return;
        const xScale = chart.scales.x;
        const x = xScale.getPixelForValue(todayIdx - 1);
        const ctx = chart.ctx;
        ctx.save();
        ctx.beginPath(); ctx.strokeStyle = '#c9952b'; ctx.lineWidth = 2; ctx.setLineDash([5, 3]);
        ctx.moveTo(x, chart.chartArea.top); ctx.lineTo(x, chart.chartArea.bottom); ctx.stroke();
        ctx.fillStyle = '#c9952b'; ctx.font = '11px Segoe UI'; ctx.fillText('Today', x + 4, chart.chartArea.top + 12);
        ctx.restore();
      }
    }]
  }));

  amortCharts.push(new Chart(document.getElementById('amortBalance'), {
    type: 'line',
    data: {
      labels,
      datasets: [{
        label: 'Balance', data: balanceData, borderColor: '#3498db', backgroundColor: 'rgba(52,152,219,.1)',
        fill: true, pointRadius: 0, tension: .3,
        segment: { borderColor: ctx => ctx.p0DataIndex < todayIdx ? '#3498db' : '#555', borderDash: ctx => ctx.p0DataIndex < todayIdx ? [] : [5, 3] }
      }]
    },
    options: {
      responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>fmt(ctx.raw)}}},
      scales:{
        y:{ticks:{color:'#8890a4',callback:v=>fmtK(v)},grid:{color:'#2a2d3e'}},
        x:{ticks:{color:'#8890a4',maxTicksLimit:12,callback:function(v){return 'Mo '+this.getLabelForValue(v)}},grid:{display:false}}
      }
    },
    plugins: [{
      id: 'todayLine2',
      afterDraw(chart) {
        if (todayIdx <= 0) return;
        const xScale = chart.scales.x;
        const x = xScale.getPixelForValue(todayIdx - 1);
        const ctx = chart.ctx;
        ctx.save();
        ctx.beginPath(); ctx.strokeStyle = '#c9952b'; ctx.lineWidth = 2; ctx.setLineDash([5, 3]);
        ctx.moveTo(x, chart.chartArea.top); ctx.lineTo(x, chart.chartArea.bottom); ctx.stroke();
        ctx.fillStyle = '#c9952b'; ctx.font = '11px Segoe UI'; ctx.fillText('Today', x + 4, chart.chartArea.top + 12);
        ctx.restore();
      }
    }]
  }));

  document.getElementById('amortOverlay').classList.add('show');
};

// ═══════ MARK COLLECTED (from amort panel) ═══════
window.markCollected = function(noteId, date, amount) {
  const user = getUser();
  if (!user) { showToast('Please select a user first', 'error'); return; }

  // Update local data
  const n = NOTES.find(x => x.id === noteId);
  if (!n) return;
  const idx = n.scheduledPayments.findIndex(p => p.date === date);
  if (idx >= 0) {
    const pmt = n.scheduledPayments.splice(idx, 1)[0];
    pmt.collected = true;
    n.collectedPayments.push(pmt);
    n.totalCollected++;
    n.remainingPayments--;
    n.lastPmtDate = pmt.date;
    n.lastPmtAmount = pmt.payment;
    if (pmt.balance != null) n.currentBalance = pmt.balance;

    // Recalculate status
    const next = n.scheduledPayments[0];
    if (next) {
      n.nextDueDate = next.date;
      const nextDue = new Date(next.date);
      n.daysPastDue = Math.max(0, Math.floor((TODAY - nextDue) / (1000*60*60*24)));
      if (n.daysPastDue > 60) n.status = 'Severely Delinquent';
      else if (n.daysPastDue > 30) n.status = 'Demand Letter';
      else if (n.daysPastDue > 15) n.status = 'Late Notice';
      else if (n.daysPastDue > 0) n.status = 'Past Due';
      else n.status = 'Current';
    }
  }

  // Re-render
  window.showAmortSchedule(noteId);
  renderPayments();
  renderCollections();
  renderOverview();

  // Sync to API
  postAPI({ action: 'mark_collected', noteId, date, amount, updatedBy: user });
  showToast('Payment marked as collected by ' + user, 'saved', 3000);
};

window.saveGeneralNote = function(id, val) {
  saveStatus(id, 'general_notes', val);
};

// Close amort panel
document.getElementById('amortOverlay').addEventListener('click', function(e) {
  if (e.target === this) this.classList.remove('show');
});
document.addEventListener('keydown', e => {
  if (e.key === 'Escape') document.getElementById('amortOverlay').classList.remove('show');
});

// ═══════ INIT ═══════
async function init() {
  await loadStatuses();
  renderOverview();
  renderPayments();
  renderCollections();
  renderInsurance();
  renderCashFlow();
}
init();
<\/script>
</body>
</html>`;
}

build().catch(e => { console.error(e); process.exit(1); });
