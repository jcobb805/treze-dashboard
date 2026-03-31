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
<script src="https://unpkg.com/jspdf@2.5.2/dist/jspdf.umd.min.js"></script>
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
  <div class="tab" data-tab="analysis">Analysis</div>
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
    <div class="btn-group" style="display:flex;gap:12px;flex-wrap:wrap;align-items:center;margin-bottom:16px">
      <button class="btn btn-gold" onclick="window.generateSelectedLateNotices()">Generate Late Notice (PDF)</button>
      <button class="btn btn-red" onclick="window.generateSelectedDemandLetters()">Generate Demand Letter (PDF)</button>
      <span id="collectSelCount" style="font-size:12px;color:#8890a4;margin-left:8px"></span>
    </div>
    <h3 style="color:#fff;margin-bottom:15px">Action Items</h3>
    <div class="table-wrap">
      <table><thead><tr>
        <th style="width:36px"><input type="checkbox" id="collectSelectAll" onchange="window.toggleCollectAll(this.checked)" title="Select all"></th><th>Address</th><th>Borrower</th><th>Days Past Due</th><th>Status</th><th>Collection Status</th><th>Action</th><th>Notes</th><th>Balance</th>
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

  <div class="tab-content" id="tab-analysis">
    <h3 style="color:#fff;margin-bottom:5px">Treze Alcove — Net Cash Flow Analysis</h3>
    <p style="color:#8890a4;margin-bottom:18px;font-size:13px">Notes Receivable income vs. First United Bank debt service. Adjust sliders to model scenarios.</p>
    <div class="kpi-row" id="analysisKpi"></div>
    <div style="display:flex;gap:24px;flex-wrap:wrap;margin-bottom:20px">
      <div style="flex:1;min-width:280px;background:#1a1d2e;padding:16px;border-radius:8px;border:1px solid #2a2d3e">
        <label style="font-size:11px;text-transform:uppercase;color:#8890a4;letter-spacing:.5px">Additional Monthly Principal Paydown</label>
        <div style="display:flex;align-items:center;gap:12px;margin-top:8px">
          <input type="range" id="extraPaydown" min="0" max="10000" step="100" value="0" style="flex:1" oninput="window.renderAnalysis()">
          <span id="extraPaydownVal" style="color:#c9952b;font-weight:600;min-width:70px;text-align:right">$0</span>
        </div>
      </div>
      <div style="flex:1;min-width:280px;background:#1a1d2e;padding:16px;border-radius:8px;border:1px solid #2a2d3e">
        <label style="font-size:11px;text-transform:uppercase;color:#8890a4;letter-spacing:.5px">Assumed Floating Rate (after Apr 2027)</label>
        <div style="display:flex;align-items:center;gap:12px;margin-top:8px">
          <input type="range" id="floatingRate" min="4.25" max="12" step="0.25" value="8.0" style="flex:1" oninput="window.renderAnalysis()">
          <span id="floatingRateVal" style="color:#c9952b;font-weight:600;min-width:55px;text-align:right">8.00%</span>
        </div>
        <div style="font-size:11px;color:#555;margin-top:4px">Floor: 4.25% | Currently: prime + 0.50%</div>
      </div>
    </div>
    <div class="chart-row">
      <div class="chart-card" style="flex:2"><h3>Monthly Net Cash Flow</h3><div class="chart-wrap"><canvas id="analysisCfChart"></canvas></div></div>
      <div class="chart-card" style="flex:1"><h3>FUB Loan Balance</h3><div class="chart-wrap"><canvas id="analysisBal"></canvas></div></div>
    </div>
    <h3 style="color:#fff;margin:20px 0 15px">Monthly Projection</h3>
    <div class="table-wrap" style="max-height:500px;overflow-y:auto">
      <table><thead><tr>
        <th>Month</th><th>Active Notes</th><th>N/R Income</th><th>FUB Interest</th><th>FUB Principal</th><th>Extra Paydown</th><th>Total Debt Service</th><th>Net Cash Flow</th><th>FUB Balance</th>
      </tr></thead><tbody id="analysisBody"></tbody></table>
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
const API_URL = 'https://script.google.com/macros/s/AKfycby7VS75jgpKKmq_SOzGq05v6J34wn_yMClm34h-bPAEolDOgGayAV-o0YZQVfVcw4a_/exec';

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
  if (f.type === 'status') {
    if (f.value === 'Past Due') return notes.filter(n => n.daysPastDue > 0);
    return notes.filter(n => n.status === f.value);
  }
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
  const lateNotice = actionNotes.filter(n => n.daysPastDue >= 15 && n.daysPastDue <= 30);
  const demandLetter = actionNotes.filter(n => n.daysPastDue > 30);
  const totalPastDueBal = actionNotes.reduce((s,n) => s+n.currentBalance, 0);

  document.getElementById('collectBadge').textContent = actionNotes.length;
  document.getElementById('collectKpi').innerHTML = [
    {label:'Total Past Due',value:actionNotes.length,cls:actionNotes.length>0?'red':'green'},
    {label:'Late Notice (15-30 days)',value:lateNotice.length,cls:lateNotice.length>0?'orange':'green'},
    {label:'Demand Letter (30+ days, 30-day cure)',value:demandLetter.length,cls:demandLetter.length>0?'red':'green'},
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
      } else if (n.daysPastDue >= 15) {
        actionBtns = '<button class="btn btn-gold btn-sm" onclick="event.stopPropagation();window.generateLateNotice(\\''+nid+'\\')">Late Notice</button>';
      }

      return '<tr class="action-needed" onclick="window.showAmortSchedule(\\''+nid+'\\')">'+
        '<td onclick="event.stopPropagation()"><input type="checkbox" class="collect-check" data-note-id="'+n.id.replace(/"/g,'&quot;')+'" onchange="window.updateCollectSelCount()"></td>'+
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

const SIGNATURE_IMG = 'data:image/png;base64,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';
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

  // Letter date: 15 days after due date, skip to Monday if weekend
  const dueDate15 = new Date(n.nextDueDate + 'T00:00:00');
  dueDate15.setDate(dueDate15.getDate() + 15);
  const dow = dueDate15.getDay();
  if (dow === 0) dueDate15.setDate(dueDate15.getDate() + 1); // Sun -> Mon
  if (dow === 6) dueDate15.setDate(dueDate15.getDate() + 2); // Sat -> Mon
  const letterDateLong = dueDate15.toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' });

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(11);
  doc.text(letterDateLong, 190, y, { align: 'right' });
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

  // Payment deadline — end of the month the letter is dated
  const endOfMonth = new Date(dueDate15.getFullYear(), dueDate15.getMonth() + 1, 0);
  const deadline = endOfMonth.toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' });
  doc.text('You may make this payment with a check, wire, or ACH. Payment is due no later than ' + deadline + ', or your loan will be in default and additional fees will be incurred.', 20, y, { maxWidth: 170 });
  y += 18;

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
  doc.addImage(SIGNATURE_IMG, 'PNG', 20, y - 8, 38, 16);
  y += 12;
  doc.setFont('helvetica', 'normal');
  doc.text(LETTERHEAD.signer, 20, y); y += 5;
  doc.text(LETTERHEAD.signerTitle, 20, y); y += 5;
  doc.text(LETTERHEAD.signerEntity, 20, y);

  // Save
  const filename = 'Late Notice - ' + n.address.replace(/[^a-zA-Z0-9 ]/g, '') + '.pdf';
  var a=document.createElement('a');a.href=doc.output('bloburl');a.target='_blank';a.rel='noopener';document.body.appendChild(a);a.click();document.body.removeChild(a);
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
  doc.addImage(SIGNATURE_IMG, 'PNG', 20, y - 8, 38, 16);
  y += 12;
  doc.setFont('helvetica', 'normal');
  doc.text(LETTERHEAD.signer, 20, y); y += 5;
  doc.text(LETTERHEAD.signerTitle, 20, y); y += 5;
  doc.text(LETTERHEAD.signerEntity, 20, y);

  const filename = 'Demand Letter - ' + n.address.replace(/[^a-zA-Z0-9 ]/g, '') + '.pdf';
  var a=document.createElement('a');a.href=doc.output('bloburl');a.target='_blank';a.rel='noopener';document.body.appendChild(a);a.click();document.body.removeChild(a);
  showToast('Demand Letter generated: ' + n.address, 'saved', 3000);

  postAPI({ action: 'log_action', updatedBy: getUser(), actionType: 'demand_letter_generated', noteId: n.id, details: 'Arrearage: ' + fmtMoney(totalArrearage) });
};

// Checkbox helpers
function getCheckedNoteIds() {
  return Array.from(document.querySelectorAll('.collect-check:checked')).map(cb => cb.dataset.noteId);
}

window.toggleCollectAll = function(checked) {
  document.querySelectorAll('.collect-check').forEach(cb => { cb.checked = checked; });
  window.updateCollectSelCount();
};

window.updateCollectSelCount = function() {
  const checked = document.querySelectorAll('.collect-check:checked').length;
  const total = document.querySelectorAll('.collect-check').length;
  const el = document.getElementById('collectSelCount');
  el.textContent = checked > 0 ? checked + ' of ' + total + ' selected' : 'Check rows to select, or generates all qualifying';
  document.getElementById('collectSelectAll').checked = checked === total && total > 0;
};

window.generateSelectedLateNotices = function() {
  const checkedIds = getCheckedNoteIds();
  let targets;
  if (checkedIds.length > 0) {
    targets = NOTES.filter(n => checkedIds.includes(n.id) && n.daysPastDue >= 15 && n.daysPastDue <= 30);
    if (targets.length === 0) { showToast('None of the selected notes qualify for late notice (1-30 days past due)', 'error'); return; }
  } else {
    targets = NOTES.filter(n => n.daysPastDue >= 15 && n.daysPastDue <= 30);
    if (targets.length === 0) { showToast('No notes qualify for late notice', 'error'); return; }
  }
  targets.forEach(n => window.generateLateNotice(n.id));
  showToast('Generated ' + targets.length + ' late notice(s)', 'saved', 3000);
};

window.generateSelectedDemandLetters = function() {
  const checkedIds = getCheckedNoteIds();
  let targets;
  if (checkedIds.length > 0) {
    targets = NOTES.filter(n => checkedIds.includes(n.id) && n.daysPastDue > 30);
    if (targets.length === 0) { showToast('None of the selected notes qualify for demand letter (30+ days past due)', 'error'); return; }
  } else {
    targets = NOTES.filter(n => n.daysPastDue > 30);
    if (targets.length === 0) { showToast('No notes qualify for demand letter', 'error'); return; }
  }
  targets.forEach(n => window.generateDemandLetter(n.id));
  showToast('Generated ' + targets.length + ' demand letter(s)', 'saved', 3000);
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
      (n.daysPastDue >= 15 && n.daysPastDue <= 30 ? '<button class="btn btn-gold btn-sm" onclick="window.generateLateNotice(\\''+nid+'\\')">Generate Late Notice</button>' : '')+
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

// ═══════ ANALYSIS TAB ═══════
const FUB = {
  balanceAsOf: 2594885.02,
  balanceDate: '2026-02-28',
  fixedRate: 0.0725,
  fixedUntil: '2027-04-18',
  floatingSpread: 0.005, // prime + 0.5%
  floor: 0.0425,
  monthlyPmt: 23600,
  maturity: '2030-03-18'
};

let analysisCfChart = null;
let analysisBalChart = null;

window.renderAnalysis = function() {
  const extraPaydown = parseFloat(document.getElementById('extraPaydown').value) || 0;
  const floatingRate = parseFloat(document.getElementById('floatingRate').value) / 100;
  document.getElementById('extraPaydownVal').textContent = '$' + extraPaydown.toLocaleString();
  document.getElementById('floatingRateVal').textContent = (floatingRate * 100).toFixed(2) + '%';

  // Build month-by-month projection from March 2026 to March 2030
  const startDate = new Date(2026, 2, 1); // Mar 2026
  const fixedEnd = new Date(2027, 3, 18); // Apr 18, 2027
  const maturity = new Date(2030, 2, 18); // Mar 18, 2030
  let fubBal = FUB.balanceAsOf;

  // Pre-compute which notes are active each month (based on balloon dates)
  const rows = [];
  const labels = [];
  const netCfData = [];
  const nrIncomeData = [];
  const debtServiceData = [];
  const fubBalData = [];

  let totalNetCf = 0;
  let totalNrIncome = 0;
  let totalDebtService = 0;
  let totalInterest = 0;

  for (let m = 0; m < 49; m++) { // ~4 years
    const monthDate = new Date(startDate.getFullYear(), startDate.getMonth() + m, 1);
    if (monthDate > maturity) break;
    const monthEnd = new Date(monthDate.getFullYear(), monthDate.getMonth() + 1, 0);
    const monthLabel = monthDate.toLocaleDateString('en-US', { month: 'short', year: 'numeric' });

    // Count active notes and income this month
    const activeNotes = NOTES.filter(n => {
      const balloon = new Date(n.balloonDate);
      return balloon > monthDate;
    });
    const nrIncome = activeNotes.reduce((s, n) => s + n.monthlyPmt, 0);

    // FUB interest calc
    const rate = monthDate < fixedEnd ? FUB.fixedRate : Math.max(floatingRate, FUB.floor);
    const monthlyInterest = fubBal * rate / 12;
    const basePrincipal = Math.max(0, FUB.monthlyPmt - monthlyInterest);
    const totalPrincipal = basePrincipal + extraPaydown;
    const actualPrincipal = Math.min(totalPrincipal, fubBal);
    const actualExtra = Math.min(extraPaydown, Math.max(0, fubBal - basePrincipal));
    const totalDs = monthlyInterest + actualPrincipal;
    const netCf = nrIncome - monthlyInterest - basePrincipal - actualExtra;

    totalNetCf += netCf;
    totalNrIncome += nrIncome;
    totalDebtService += totalDs;
    totalInterest += monthlyInterest;

    rows.push({
      label: monthLabel,
      activeCount: activeNotes.length,
      nrIncome,
      fubInterest: monthlyInterest,
      fubPrincipal: basePrincipal,
      extraPd: actualExtra,
      totalDs,
      netCf,
      fubBal
    });

    labels.push(monthLabel);
    netCfData.push(Math.round(netCf));
    nrIncomeData.push(Math.round(nrIncome));
    debtServiceData.push(Math.round(-totalDs));
    fubBalData.push(Math.round(fubBal));

    fubBal = Math.max(0, fubBal - actualPrincipal);
  }

  // KPIs
  const avgNetCf = totalNetCf / rows.length;
  const totalPrincipalPaid = FUB.balanceAsOf - fubBal;
  document.getElementById('analysisKpi').innerHTML = [
    {label:'Avg Monthly Net Cash Flow',value:'$'+Math.round(avgNetCf).toLocaleString(),cls:avgNetCf>=0?'green':'red'},
    {label:'Total N/R Income (Proj.)',value:'$'+Math.round(totalNrIncome/1000).toLocaleString()+'K',cls:'green'},
    {label:'Total Interest to FUB',value:'$'+Math.round(totalInterest/1000).toLocaleString()+'K',cls:'red'},
    {label:'FUB Balance at Maturity',value:'$'+Math.round(fubBal).toLocaleString(),cls:fubBal>0?'red':'green'},
    {label:'Total Principal Paid',value:'$'+Math.round(totalPrincipalPaid/1000).toLocaleString()+'K',cls:'green'}
  ].map(k=>'<div class="kpi"><div class="label">'+k.label+'</div><div class="value '+k.cls+'">'+k.value+'</div></div>').join('');

  // Net CF chart
  const cfCtx = document.getElementById('analysisCfChart');
  if (analysisCfChart) analysisCfChart.destroy();
  analysisCfChart = new Chart(cfCtx, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        {label:'N/R Income',data:nrIncomeData,backgroundColor:'rgba(46,204,113,0.7)',stack:'stack'},
        {label:'Debt Service',data:debtServiceData,backgroundColor:'rgba(231,76,60,0.7)',stack:'stack'},
        {label:'Net Cash Flow',data:netCfData,type:'line',borderColor:'#c9952b',backgroundColor:'transparent',borderWidth:2,pointRadius:1,tension:0.3}
      ]
    },
    options: {
      responsive:true,maintainAspectRatio:false,
      plugins:{legend:{labels:{color:'#8890a4',font:{size:11}}}},
      scales:{
        x:{ticks:{color:'#8890a4',font:{size:10},maxRotation:45},grid:{color:'#1e2130'}},
        y:{ticks:{color:'#8890a4',callback:v=>'$'+Math.round(v/1000)+'K'},grid:{color:'#1e2130'}}
      }
    }
  });

  // Balance chart
  const balCtx = document.getElementById('analysisBal');
  if (analysisBalChart) analysisBalChart.destroy();
  analysisBalChart = new Chart(balCtx, {
    type: 'line',
    data: {
      labels,
      datasets: [{
        label:'FUB Loan Balance',
        data:fubBalData,
        borderColor:'#e74c3c',
        backgroundColor:'rgba(231,76,60,0.1)',
        fill:true,
        borderWidth:2,
        pointRadius:1,
        tension:0.3
      }]
    },
    options: {
      responsive:true,maintainAspectRatio:false,
      plugins:{legend:{labels:{color:'#8890a4',font:{size:11}}}},
      scales:{
        x:{ticks:{color:'#8890a4',font:{size:10},maxRotation:45},grid:{color:'#1e2130'}},
        y:{ticks:{color:'#8890a4',callback:v=>'$'+(v/1000000).toFixed(1)+'MM'},grid:{color:'#1e2130'}}
      }
    }
  });

  // Table
  document.getElementById('analysisBody').innerHTML = rows.map((r,i) => {
    const rateLabel = new Date(2026, 2 + i, 1) < new Date(2027, 3, 18) ? '(fixed ' + (FUB.fixedRate*100).toFixed(2) + '%)' : '(float ' + (Math.max(floatingRate, FUB.floor)*100).toFixed(2) + '%)';
    return '<tr>'+
      '<td>'+r.label+' <span style="color:#555;font-size:10px">'+rateLabel+'</span></td>'+
      '<td class="num">'+r.activeCount+'</td>'+
      '<td class="num" style="color:#2ecc71">'+fmt(r.nrIncome)+'</td>'+
      '<td class="num" style="color:#e74c3c">'+fmt(r.fubInterest)+'</td>'+
      '<td class="num">'+fmt(r.fubPrincipal)+'</td>'+
      '<td class="num" style="color:#c9952b">'+(r.extraPd>0?fmt(r.extraPd):'-')+'</td>'+
      '<td class="num" style="color:#e74c3c">'+fmt(r.totalDs)+'</td>'+
      '<td class="num" style="color:'+(r.netCf>=0?'#2ecc71':'#e74c3c')+';font-weight:600">'+fmt(r.netCf)+'</td>'+
      '<td class="num">'+fmt(r.fubBal)+'</td>'+
    '</tr>';
  }).join('');
};

// ═══════ INIT ═══════
async function init() {
  await loadStatuses();
  renderOverview();
  renderPayments();
  renderCollections();
  renderInsurance();
  renderCashFlow();
  window.renderAnalysis();
}
init();
<\/script>
</body>
</html>`;
}

build().catch(e => { console.error(e); process.exit(1); });
