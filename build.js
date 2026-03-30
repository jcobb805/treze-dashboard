const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const INPUT = path.join(__dirname, '..', 'OneDrive', 'Treze Alcove', 'Treze Alcove - Preze Enterprises - Amortization Schedule.xlsx');
const OUTPUT = path.join(__dirname, 'index.html');
const TODAY = new Date();

// Green fill color used to mark collected payments in Excel
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

    // Determine header row columns (row 10)
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
    const lateFeeCol = headers['late fee/attor. fee'] || headers['late/attor fee'] || headers['late fee/attor fee'] || null;

    // Find balance column reliably
    let balanceCol = -1;
    for (let c = 1; c <= 15; c++) {
      const h = headerRow.getCell(c).value;
      if (h && String(h).toLowerCase().trim() === 'balance') { balanceCol = c; break; }
    }

    // Extract payment rows -- use green fill on date cell to determine collected vs scheduled
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

      // Skip the initial "Loan" row (down payment)
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

    // Skip notes with no collected payments (empty/unused sheets)
    if (collectedPayments.length === 0) continue;

    // Determine if note is active: must have scheduled future payments remaining
    // If no scheduled payments remain, the note is paid off or matured
    if (scheduledPayments.length === 0) continue;

    const lastCollected = collectedPayments[collectedPayments.length - 1];
    const lastCollectedDate = new Date(lastCollected.date);

    // Payoff detection: check collected payments AND first 3 scheduled rows for lump sum
    // Threshold: > 50% of original loan amount (catches payoffs, ignores extra principal payments)
    const payoffThreshold = loanAmt * 0.5;
    const payoffCheck = [...collectedPayments, ...scheduledPayments.slice(0, 3)];
    const hasPayoff = payoffCheck.some(p => p.payment > payoffThreshold);
    if (hasPayoff) continue;

    // Inactive detection: if last green payment is > 180 days ago, note is inactive/written off
    const daysSinceLastGreen = Math.floor((TODAY - lastCollectedDate) / (1000 * 60 * 60 * 24));
    if (daysSinceLastGreen > 180) continue;
    const nextScheduled = scheduledPayments[0]; // First non-green row = next expected payment

    // Calculate balloon date
    const balloonDate = new Date(beginDate);
    balloonDate.setMonth(balloonDate.getMonth() + balloonAt);

    // Next due date = first scheduled (non-green) payment date
    const nextDueDate = new Date(nextScheduled.date);

    // Days since last collected payment
    const daysSinceLastPmt = Math.floor((TODAY - lastCollectedDate) / (1000 * 60 * 60 * 24));

    // Days past due (from next due date -- if today is past the next scheduled date)
    const daysPastDue = Math.max(0, Math.floor((TODAY - nextDueDate) / (1000 * 60 * 60 * 24)));

    // Determine status
    let status = 'Current';
    if (daysPastDue > 60) status = 'Severely Delinquent';
    else if (daysPastDue > 30) status = 'Demand Letter';
    else if (daysPastDue > 15) status = 'Late Notice';
    else if (daysPastDue > 0) status = 'Past Due';

    // Closing number from sheet name
    const closingMatch = ws.name.match(/\(#(\d+)\)/);
    const closing = closingMatch ? parseInt(closingMatch[1]) : null;

    // Current balance from the last collected payment's balance, or estimate
    let currentBalance = lastCollected.balance;
    if (!currentBalance || typeof currentBalance !== 'number') {
      const totalPrincipal = collectedPayments.reduce((s, p) => s + p.principal, 0);
      currentBalance = Math.max(0, loanAmt - totalPrincipal);
    }

    notes.push({
      id: ws.name,
      address,
      closing,
      borrower,
      loanAmt,
      rate,
      amortMonths,
      balloonAt,
      beginDate: beginDate.toISOString().split('T')[0],
      monthlyPmt,
      balloonDate: balloonDate.toISOString().split('T')[0],
      currentBalance: Math.round(currentBalance * 100) / 100,
      lastPmtDate: lastCollected.date,
      lastPmtAmount: lastCollected.payment,
      nextDueDate: nextDueDate.toISOString().split('T')[0],
      daysSinceLastPmt,
      daysPastDue,
      status,
      totalCollected: collectedPayments.length,
      remainingPayments: scheduledPayments.length,
      payments: collectedPayments.slice(-12) // Last 12 collected for detail view
    });
  }

  // Sort by status severity, then days past due
  const statusOrder = { 'Severely Delinquent': 0, 'Demand Letter': 1, 'Late Notice': 2, 'Past Due': 3, 'Current': 4 };
  notes.sort((a, b) => (statusOrder[a.status] - statusOrder[b.status]) || (b.daysPastDue - a.daysPastDue));

  console.log(`Extracted ${notes.length} active notes`);
  notes.forEach(n => console.log(`  ${n.id} | ${n.status} | ${n.daysPastDue} days past due | Bal: $${n.currentBalance.toLocaleString()}`));

  // Generate HTML
  const html = generateHTML(notes);
  fs.writeFileSync(OUTPUT, html, 'utf8');
  console.log(`\nDashboard written to ${OUTPUT}`);
}

function generateHTML(notes) {
  const dataJSON = JSON.stringify(notes);

  // Summary stats
  const totalBalance = notes.reduce((s, n) => s + n.currentBalance, 0);
  const totalMonthly = notes.reduce((s, n) => s + n.monthlyPmt, 0);
  const currentCount = notes.filter(n => n.status === 'Current').length;
  const pastDueCount = notes.filter(n => n.status === 'Past Due').length;
  const lateNoticeCount = notes.filter(n => n.status === 'Late Notice').length;
  const demandCount = notes.filter(n => n.status === 'Demand Letter').length;
  const severeCount = notes.filter(n => n.status === 'Severely Delinquent').length;

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Treze Alcove - Loan Servicing Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.2/dist/chart.umd.min.js"><\/script>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: 'Segoe UI', Arial, sans-serif; background: #0f1117; color: #e0e0e0; min-height: 100vh; }
.header { background: linear-gradient(135deg, #1a1d2e 0%, #252940 100%); padding: 20px 30px; border-bottom: 2px solid #c9952b; display: flex; align-items: center; justify-content: space-between; }
.header h1 { font-size: 22px; font-weight: 600; color: #fff; }
.header h1 span { color: #c9952b; }
.header .subtitle { font-size: 13px; color: #8890a4; margin-top: 2px; }
.header .date { font-size: 13px; color: #8890a4; }

/* Tabs */
.tabs { display: flex; background: #181b28; border-bottom: 1px solid #2a2d3e; padding: 0 20px; }
.tab { padding: 12px 24px; cursor: pointer; font-size: 13px; font-weight: 500; color: #8890a4; border-bottom: 2px solid transparent; transition: all 0.2s; }
.tab:hover { color: #c9952b; }
.tab.active { color: #c9952b; border-bottom-color: #c9952b; }
.tab .badge { background: #e74c3c; color: #fff; font-size: 10px; padding: 1px 6px; border-radius: 8px; margin-left: 6px; }

/* Content */
.content { padding: 20px 30px; }
.tab-content { display: none; }
.tab-content.active { display: block; }

/* KPI Cards */
.kpi-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 15px; margin-bottom: 25px; }
.kpi { background: #1a1d2e; border-radius: 8px; padding: 18px; border: 1px solid #2a2d3e; }
.kpi .label { font-size: 11px; text-transform: uppercase; color: #8890a4; letter-spacing: 0.5px; margin-bottom: 6px; }
.kpi .value { font-size: 24px; font-weight: 700; color: #fff; }
.kpi .value.green { color: #2ecc71; }
.kpi .value.yellow { color: #f1c40f; }
.kpi .value.orange { color: #e67e22; }
.kpi .value.red { color: #e74c3c; }
.kpi .sub { font-size: 11px; color: #8890a4; margin-top: 4px; }

/* Tables */
table { width: 100%; border-collapse: collapse; font-size: 13px; }
thead th { background: #1a1d2e; color: #8890a4; text-transform: uppercase; font-size: 11px; letter-spacing: 0.5px; padding: 10px 12px; text-align: left; border-bottom: 1px solid #2a2d3e; cursor: pointer; user-select: none; position: sticky; top: 0; z-index: 1; }
thead th:hover { color: #c9952b; }
thead th .sort-arrow { margin-left: 4px; font-size: 9px; }
tbody tr { border-bottom: 1px solid #1e2130; transition: background 0.15s; }
tbody tr:hover { background: #1e2235; }
tbody td { padding: 10px 12px; }
tbody td.num { text-align: right; font-variant-numeric: tabular-nums; }

/* Status badges */
.status { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 11px; font-weight: 600; white-space: nowrap; }
.status.current { background: rgba(46,204,113,0.15); color: #2ecc71; }
.status.past-due { background: rgba(241,196,15,0.15); color: #f1c40f; }
.status.late-notice { background: rgba(230,126,34,0.15); color: #e67e22; }
.status.demand-letter { background: rgba(231,76,60,0.15); color: #e74c3c; }
.status.severely-delinquent { background: rgba(231,76,60,0.25); color: #ff6b6b; }

/* Charts */
.chart-row { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 25px; }
.chart-card { background: #1a1d2e; border-radius: 8px; padding: 20px; border: 1px solid #2a2d3e; }
.chart-card h3 { font-size: 14px; color: #fff; margin-bottom: 15px; font-weight: 600; }
.chart-wrap { position: relative; height: 280px; }

/* Filter bar */
.filter-bar { display: flex; gap: 12px; margin-bottom: 20px; flex-wrap: wrap; align-items: center; }
.filter-bar select, .filter-bar input { background: #1a1d2e; color: #e0e0e0; border: 1px solid #2a2d3e; border-radius: 6px; padding: 8px 12px; font-size: 13px; }
.filter-bar select:focus, .filter-bar input:focus { outline: none; border-color: #c9952b; }
.filter-bar label { font-size: 12px; color: #8890a4; }

/* Detail modal */
.modal-overlay { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.7); z-index: 100; justify-content: center; align-items: center; }
.modal-overlay.show { display: flex; }
.modal { background: #1a1d2e; border-radius: 10px; border: 1px solid #2a2d3e; width: 90%; max-width: 900px; max-height: 85vh; overflow-y: auto; padding: 25px; }
.modal h2 { font-size: 18px; color: #fff; margin-bottom: 5px; }
.modal .modal-sub { font-size: 13px; color: #8890a4; margin-bottom: 20px; }
.modal .close-btn { position: absolute; top: 15px; right: 20px; color: #8890a4; cursor: pointer; font-size: 20px; }
.modal-header { display: flex; justify-content: space-between; align-items: flex-start; position: relative; }
.modal .detail-grid { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 12px; margin-bottom: 20px; }
.modal .detail-item { }
.modal .detail-item .dl { font-size: 11px; color: #8890a4; text-transform: uppercase; }
.modal .detail-item .dv { font-size: 15px; font-weight: 600; color: #fff; margin-top: 2px; }

/* Insurance section */
.ins-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 12px; }
.ins-card { background: #1a1d2e; border-radius: 8px; padding: 15px; border: 1px solid #2a2d3e; }
.ins-card .ins-addr { font-weight: 600; color: #fff; font-size: 14px; }
.ins-card .ins-borrower { font-size: 12px; color: #8890a4; margin-top: 2px; }
.ins-card .ins-status { margin-top: 10px; display: flex; align-items: center; gap: 8px; }
.ins-dot { width: 10px; height: 10px; border-radius: 50%; display: inline-block; }
.ins-dot.valid { background: #2ecc71; }
.ins-dot.expiring { background: #f1c40f; }
.ins-dot.lapsed { background: #e74c3c; }
.ins-dot.unknown { background: #555; }
.ins-card .ins-label { font-size: 12px; }

/* Scrollable table wrapper */
.table-wrap { max-height: 600px; overflow-y: auto; border-radius: 8px; border: 1px solid #2a2d3e; }
.table-wrap::-webkit-scrollbar { width: 6px; }
.table-wrap::-webkit-scrollbar-track { background: #1a1d2e; }
.table-wrap::-webkit-scrollbar-thumb { background: #3a3d4e; border-radius: 3px; }

/* Collections queue highlight */
tr.action-needed { background: rgba(231,76,60,0.05); }

/* Balloon warning */
.balloon-warn { color: #e67e22; font-weight: 600; }

@media (max-width: 900px) {
  .chart-row { grid-template-columns: 1fr; }
  .kpi-row { grid-template-columns: repeat(2, 1fr); }
  .modal .detail-grid { grid-template-columns: 1fr 1fr; }
}
</style>
</head>
<body>

<div class="header">
  <div>
    <h1>Treze Alcove <span>Loan Servicing</span></h1>
    <div class="subtitle">Preze Enterprises -- Notes Receivable Portfolio</div>
  </div>
  <div class="date">As of ${TODAY.toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' })}</div>
</div>

<div class="tabs">
  <div class="tab active" data-tab="overview">Overview</div>
  <div class="tab" data-tab="payments">Payment Status</div>
  <div class="tab" data-tab="collections">Collections<span class="badge" id="collectBadge">0</span></div>
  <div class="tab" data-tab="insurance">Insurance</div>
  <div class="tab" data-tab="cashflow">Cash Flow</div>
</div>

<div class="content">

  <!-- OVERVIEW TAB -->
  <div class="tab-content active" id="tab-overview">
    <div class="kpi-row" id="kpiRow"></div>
    <div class="chart-row">
      <div class="chart-card">
        <h3>Portfolio Status</h3>
        <div class="chart-wrap"><canvas id="statusChart"></canvas></div>
      </div>
      <div class="chart-card">
        <h3>Aging Distribution (Days Past Due)</h3>
        <div class="chart-wrap"><canvas id="agingChart"></canvas></div>
      </div>
    </div>
    <div class="chart-row">
      <div class="chart-card">
        <h3>Balance by Closing</h3>
        <div class="chart-wrap"><canvas id="closingChart"></canvas></div>
      </div>
      <div class="chart-card">
        <h3>Balloon Maturity Timeline</h3>
        <div class="chart-wrap"><canvas id="balloonChart"></canvas></div>
      </div>
    </div>
  </div>

  <!-- PAYMENT STATUS TAB -->
  <div class="tab-content" id="tab-payments">
    <div class="filter-bar">
      <div>
        <label>Status</label><br>
        <select id="filterStatus">
          <option value="all">All Statuses</option>
          <option value="Current">Current</option>
          <option value="Past Due">Past Due</option>
          <option value="Late Notice">Late Notice</option>
          <option value="Demand Letter">Demand Letter</option>
          <option value="Severely Delinquent">Severely Delinquent</option>
        </select>
      </div>
      <div>
        <label>Closing</label><br>
        <select id="filterClosing"><option value="all">All Closings</option></select>
      </div>
      <div>
        <label>Search</label><br>
        <input type="text" id="filterSearch" placeholder="Address or borrower...">
      </div>
    </div>
    <div class="table-wrap">
      <table id="paymentTable">
        <thead>
          <tr>
            <th data-sort="address">Address <span class="sort-arrow"></span></th>
            <th data-sort="closing">Closing <span class="sort-arrow"></span></th>
            <th data-sort="borrower">Borrower <span class="sort-arrow"></span></th>
            <th data-sort="monthlyPmt">Payment <span class="sort-arrow"></span></th>
            <th data-sort="lastPmtDate">Last Payment <span class="sort-arrow"></span></th>
            <th data-sort="nextDueDate">Next Due <span class="sort-arrow"></span></th>
            <th data-sort="daysPastDue">Days Past Due <span class="sort-arrow"></span></th>
            <th data-sort="currentBalance">Balance <span class="sort-arrow"></span></th>
            <th data-sort="status">Status <span class="sort-arrow"></span></th>
          </tr>
        </thead>
        <tbody id="paymentBody"></tbody>
      </table>
    </div>
  </div>

  <!-- COLLECTIONS TAB -->
  <div class="tab-content" id="tab-collections">
    <div class="kpi-row" id="collectKpi"></div>
    <h3 style="color:#fff; margin-bottom:15px;">Action Items</h3>
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Address</th>
            <th>Borrower</th>
            <th>Days Past Due</th>
            <th>Status</th>
            <th>Action Required</th>
            <th>Balance</th>
          </tr>
        </thead>
        <tbody id="collectBody"></tbody>
      </table>
    </div>
  </div>

  <!-- INSURANCE TAB -->
  <div class="tab-content" id="tab-insurance">
    <p style="color:#8890a4; margin-bottom:15px; font-size:13px;">Insurance validity tracking for all active notes. Update expiration dates in the build data to reflect current policy status.</p>
    <div class="kpi-row" id="insKpi"></div>
    <div class="filter-bar">
      <div>
        <label>Status</label><br>
        <select id="filterInsStatus">
          <option value="all">All</option>
          <option value="valid">Valid</option>
          <option value="expiring">Expiring Soon (30 days)</option>
          <option value="lapsed">Lapsed</option>
          <option value="unknown">Unknown</option>
        </select>
      </div>
    </div>
    <div class="ins-grid" id="insGrid"></div>
  </div>

  <!-- CASH FLOW TAB -->
  <div class="tab-content" id="tab-cashflow">
    <div class="kpi-row" id="cfKpi"></div>
    <div class="chart-row">
      <div class="chart-card">
        <h3>Expected Monthly Collections</h3>
        <div class="chart-wrap"><canvas id="cfChart"></canvas></div>
      </div>
      <div class="chart-card">
        <h3>Balloon Payoffs by Year</h3>
        <div class="chart-wrap"><canvas id="balloonPayoffChart"></canvas></div>
      </div>
    </div>
    <h3 style="color:#fff; margin-bottom:15px;">Upcoming Balloon Maturities</h3>
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Address</th>
            <th>Borrower</th>
            <th>Balloon Date</th>
            <th>Months Remaining</th>
            <th>Est. Balance at Maturity</th>
            <th>Monthly Payment</th>
          </tr>
        </thead>
        <tbody id="balloonBody"></tbody>
      </table>
    </div>
  </div>

</div>

<!-- Detail Modal -->
<div class="modal-overlay" id="modalOverlay">
  <div class="modal" id="noteModal"></div>
</div>

<script>
const NOTES = ${dataJSON};
const TODAY = new Date('${TODAY.toISOString().split('T')[0]}');

const fmt = (n) => n != null ? '$' + n.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}) : '--';
const fmtDate = (d) => d ? new Date(d + 'T00:00:00').toLocaleDateString('en-US', {month: 'short', day: 'numeric', year: 'numeric'}) : '--';
const fmtPct = (n) => (n * 100).toFixed(2) + '%';
const statusClass = (s) => s.toLowerCase().replace(/\\s+/g, '-');

// ========== TABS ==========
document.querySelectorAll('.tab').forEach(tab => {
  tab.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(tc => tc.classList.remove('active'));
    tab.classList.add('active');
    document.getElementById('tab-' + tab.dataset.tab).classList.add('active');
  });
});

// ========== OVERVIEW ==========
function renderOverview() {
  const totalBal = NOTES.reduce((s, n) => s + n.currentBalance, 0);
  const totalMonthly = NOTES.reduce((s, n) => s + n.monthlyPmt, 0);
  const currentN = NOTES.filter(n => n.status === 'Current').length;
  const actionN = NOTES.filter(n => n.status !== 'Current').length;
  const avgDPD = NOTES.length ? Math.round(NOTES.reduce((s, n) => s + n.daysPastDue, 0) / NOTES.length) : 0;

  // Balloon within 12 months
  const soon = NOTES.filter(n => {
    const bd = new Date(n.balloonDate);
    const diff = (bd - TODAY) / (1000*60*60*24*30);
    return diff <= 12 && diff > 0;
  }).length;

  document.getElementById('kpiRow').innerHTML = [
    {label: 'Active Notes', value: NOTES.length, cls: ''},
    {label: 'Total Balance', value: fmt(totalBal), cls: ''},
    {label: 'Expected Monthly', value: fmt(totalMonthly), cls: ''},
    {label: 'Current', value: currentN, cls: 'green'},
    {label: 'Action Needed', value: actionN, cls: actionN > 0 ? 'red' : 'green'},
    {label: 'Avg Days Past Due', value: avgDPD, cls: avgDPD > 15 ? 'orange' : 'green'},
    {label: 'Balloons in 12 Mo', value: soon, cls: soon > 0 ? 'yellow' : ''}
  ].map(k => '<div class="kpi"><div class="label">' + k.label + '</div><div class="value ' + k.cls + '">' + k.value + '</div></div>').join('');

  // Status doughnut
  const statusCounts = {};
  NOTES.forEach(n => { statusCounts[n.status] = (statusCounts[n.status] || 0) + 1; });
  const statusLabels = Object.keys(statusCounts);
  const statusColors = statusLabels.map(s => ({
    'Current': '#2ecc71', 'Past Due': '#f1c40f', 'Late Notice': '#e67e22',
    'Demand Letter': '#e74c3c', 'Severely Delinquent': '#ff6b6b'
  }[s] || '#555'));

  new Chart(document.getElementById('statusChart'), {
    type: 'doughnut',
    data: { labels: statusLabels, datasets: [{ data: statusLabels.map(s => statusCounts[s]), backgroundColor: statusColors, borderWidth: 0 }] },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'right', labels: { color: '#8890a4', font: { size: 12 } } } } }
  });

  // Aging bar
  const buckets = { '0 (Current)': 0, '1-15': 0, '16-30': 0, '31-60': 0, '60+': 0 };
  NOTES.forEach(n => {
    if (n.daysPastDue === 0) buckets['0 (Current)']++;
    else if (n.daysPastDue <= 15) buckets['1-15']++;
    else if (n.daysPastDue <= 30) buckets['16-30']++;
    else if (n.daysPastDue <= 60) buckets['31-60']++;
    else buckets['60+']++;
  });
  new Chart(document.getElementById('agingChart'), {
    type: 'bar',
    data: { labels: Object.keys(buckets), datasets: [{ data: Object.values(buckets), backgroundColor: ['#2ecc71','#a8d854','#f1c40f','#e67e22','#e74c3c'], borderRadius: 4 }] },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { y: { ticks: { color: '#8890a4', stepSize: 1 }, grid: { color: '#2a2d3e' } }, x: { ticks: { color: '#8890a4' }, grid: { display: false } } } }
  });

  // Balance by closing
  const closingBal = {};
  NOTES.forEach(n => {
    const key = '#' + n.closing;
    closingBal[key] = (closingBal[key] || 0) + n.currentBalance;
  });
  const closingKeys = Object.keys(closingBal).sort((a,b) => parseInt(a.slice(1)) - parseInt(b.slice(1)));
  new Chart(document.getElementById('closingChart'), {
    type: 'bar',
    data: { labels: closingKeys, datasets: [{ label: 'Balance', data: closingKeys.map(k => closingBal[k]), backgroundColor: '#c9952b', borderRadius: 4 }] },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { y: { ticks: { color: '#8890a4', callback: v => '$' + (v/1000).toFixed(0) + 'K' }, grid: { color: '#2a2d3e' } }, x: { ticks: { color: '#8890a4' }, grid: { display: false } } } }
  });

  // Balloon timeline
  const balloonByQ = {};
  NOTES.forEach(n => {
    const d = new Date(n.balloonDate);
    const q = d.getFullYear() + ' Q' + (Math.floor(d.getMonth()/3)+1);
    balloonByQ[q] = (balloonByQ[q] || 0) + 1;
  });
  const bKeys = Object.keys(balloonByQ).sort();
  new Chart(document.getElementById('balloonChart'), {
    type: 'bar',
    data: { labels: bKeys, datasets: [{ label: 'Notes Maturing', data: bKeys.map(k => balloonByQ[k]), backgroundColor: '#3498db', borderRadius: 4 }] },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { y: { ticks: { color: '#8890a4', stepSize: 1 }, grid: { color: '#2a2d3e' } }, x: { ticks: { color: '#8890a4', font: { size: 11 } }, grid: { display: false } } } }
  });
}

// ========== PAYMENT STATUS ==========
let sortCol = 'daysPastDue', sortDir = -1;

function renderPayments() {
  const statusF = document.getElementById('filterStatus').value;
  const closingF = document.getElementById('filterClosing').value;
  const searchF = document.getElementById('filterSearch').value.toLowerCase();

  let filtered = NOTES.filter(n => {
    if (statusF !== 'all' && n.status !== statusF) return false;
    if (closingF !== 'all' && n.closing !== parseInt(closingF)) return false;
    if (searchF && !n.address.toLowerCase().includes(searchF) && !n.borrower.toLowerCase().includes(searchF)) return false;
    return true;
  });

  filtered.sort((a, b) => {
    let va = a[sortCol], vb = b[sortCol];
    if (typeof va === 'string') return sortDir * va.localeCompare(vb);
    return sortDir * ((va || 0) - (vb || 0));
  });

  document.getElementById('paymentBody').innerHTML = filtered.map(n =>
    '<tr style="cursor:pointer" onclick="window.showDetail(\\'' + n.id.replace(/'/g, "\\\\'") + '\\')">' +
    '<td>' + n.address + '</td>' +
    '<td class="num">#' + n.closing + '</td>' +
    '<td>' + n.borrower.substring(0, 35) + (n.borrower.length > 35 ? '...' : '') + '</td>' +
    '<td class="num">' + fmt(n.monthlyPmt) + '</td>' +
    '<td>' + fmtDate(n.lastPmtDate) + '</td>' +
    '<td>' + fmtDate(n.nextDueDate) + '</td>' +
    '<td class="num">' + (n.daysPastDue > 0 ? '<span style="color:' + (n.daysPastDue > 30 ? '#e74c3c' : n.daysPastDue > 15 ? '#e67e22' : '#f1c40f') + '">' + n.daysPastDue + '</span>' : '0') + '</td>' +
    '<td class="num">' + fmt(n.currentBalance) + '</td>' +
    '<td><span class="status ' + statusClass(n.status) + '">' + n.status + '</span></td>' +
    '</tr>'
  ).join('');
}

// Populate closing filter
const closings = [...new Set(NOTES.map(n => n.closing))].sort((a,b) => a-b);
closings.forEach(c => {
  document.getElementById('filterClosing').innerHTML += '<option value="' + c + '">#' + c + '</option>';
});

document.getElementById('filterStatus').addEventListener('change', renderPayments);
document.getElementById('filterClosing').addEventListener('change', renderPayments);
document.getElementById('filterSearch').addEventListener('input', renderPayments);

// Sort
document.querySelectorAll('#paymentTable th[data-sort]').forEach(th => {
  th.addEventListener('click', () => {
    const col = th.dataset.sort;
    if (sortCol === col) sortDir *= -1;
    else { sortCol = col; sortDir = col === 'daysPastDue' ? -1 : 1; }
    renderPayments();
  });
});

// ========== COLLECTIONS ==========
function renderCollections() {
  const actionNotes = NOTES.filter(n => n.daysPastDue > 0);
  const lateNotice = actionNotes.filter(n => n.daysPastDue > 0 && n.daysPastDue <= 30);
  const demandLetter = actionNotes.filter(n => n.daysPastDue > 30);
  const totalPastDueBal = actionNotes.reduce((s, n) => s + n.currentBalance, 0);

  document.getElementById('collectBadge').textContent = actionNotes.length;
  document.getElementById('collectKpi').innerHTML = [
    {label: 'Total Past Due', value: actionNotes.length, cls: actionNotes.length > 0 ? 'red' : 'green'},
    {label: 'Late Notice (1-30 days)', value: lateNotice.length, cls: lateNotice.length > 0 ? 'orange' : 'green'},
    {label: 'Demand Letter (30+ days)', value: demandLetter.length, cls: demandLetter.length > 0 ? 'red' : 'green'},
    {label: 'Past Due Balance', value: fmt(totalPastDueBal), cls: 'red'}
  ].map(k => '<div class="kpi"><div class="label">' + k.label + '</div><div class="value ' + k.cls + '">' + k.value + '</div></div>').join('');

  document.getElementById('collectBody').innerHTML = actionNotes
    .sort((a, b) => b.daysPastDue - a.daysPastDue)
    .map(n => {
      let action = '';
      if (n.daysPastDue > 30) action = '<span style="color:#e74c3c;font-weight:600;">Coordinate demand letter with Ty</span>';
      else if (n.daysPastDue > 15) action = '<span style="color:#e67e22;font-weight:600;">Mail late notice</span>';
      else action = '<span style="color:#f1c40f;">Monitor -- approaching late notice threshold</span>';
      return '<tr class="action-needed">' +
        '<td>' + n.address + '</td>' +
        '<td>' + n.borrower.substring(0, 40) + '</td>' +
        '<td class="num" style="color:' + (n.daysPastDue > 30 ? '#e74c3c' : '#e67e22') + ';font-weight:600;">' + n.daysPastDue + '</td>' +
        '<td><span class="status ' + statusClass(n.status) + '">' + n.status + '</span></td>' +
        '<td>' + action + '</td>' +
        '<td class="num">' + fmt(n.currentBalance) + '</td>' +
        '</tr>';
    }).join('');
}

// ========== INSURANCE ==========
// Insurance data placeholder -- update this object with actual expiration dates
// Format: { "note id": "YYYY-MM-DD" } or null for unknown
const INSURANCE = {};

function renderInsurance() {
  const statuses = NOTES.map(n => {
    const exp = INSURANCE[n.id];
    if (!exp) return { ...n, insStatus: 'unknown', insDate: null, insLabel: 'No data on file' };
    const expDate = new Date(exp);
    const daysUntil = Math.floor((expDate - TODAY) / (1000*60*60*24));
    if (daysUntil < 0) return { ...n, insStatus: 'lapsed', insDate: exp, insLabel: 'Lapsed ' + Math.abs(daysUntil) + ' days ago' };
    if (daysUntil <= 30) return { ...n, insStatus: 'expiring', insDate: exp, insLabel: 'Expires in ' + daysUntil + ' days' };
    return { ...n, insStatus: 'valid', insDate: exp, insLabel: 'Valid through ' + fmtDate(exp) };
  });

  const valid = statuses.filter(s => s.insStatus === 'valid').length;
  const expiring = statuses.filter(s => s.insStatus === 'expiring').length;
  const lapsed = statuses.filter(s => s.insStatus === 'lapsed').length;
  const unknown = statuses.filter(s => s.insStatus === 'unknown').length;

  document.getElementById('insKpi').innerHTML = [
    {label: 'Valid', value: valid, cls: 'green'},
    {label: 'Expiring Soon', value: expiring, cls: 'yellow'},
    {label: 'Lapsed', value: lapsed, cls: lapsed > 0 ? 'red' : 'green'},
    {label: 'Unknown', value: unknown, cls: unknown > 0 ? 'orange' : ''}
  ].map(k => '<div class="kpi"><div class="label">' + k.label + '</div><div class="value ' + k.cls + '">' + k.value + '</div></div>').join('');

  const filterVal = document.getElementById('filterInsStatus').value;
  const filtered = filterVal === 'all' ? statuses : statuses.filter(s => s.insStatus === filterVal);

  document.getElementById('insGrid').innerHTML = filtered.map(s =>
    '<div class="ins-card">' +
    '<div class="ins-addr">' + s.address + '</div>' +
    '<div class="ins-borrower">' + s.borrower.substring(0, 45) + '</div>' +
    '<div class="ins-status"><span class="ins-dot ' + s.insStatus + '"></span><span class="ins-label">' + s.insLabel + '</span></div>' +
    (s.insStatus === 'lapsed' && Math.abs(Math.floor((new Date(s.insDate) - TODAY)/(1000*60*60*24))) > 30 ? '<div style="margin-top:8px;color:#e74c3c;font-size:12px;font-weight:600;">Coordinate demand letter with Ty</div>' : '') +
    '</div>'
  ).join('');
}
document.getElementById('filterInsStatus').addEventListener('change', renderInsurance);

// ========== CASH FLOW ==========
function renderCashFlow() {
  const totalMonthly = NOTES.reduce((s, n) => s + n.monthlyPmt, 0);
  const annualExpected = totalMonthly * 12;
  const totalBal = NOTES.reduce((s, n) => s + n.currentBalance, 0);

  // Monthly expected for next 12 months (accounting for balloons)
  const months = [];
  for (let i = 0; i < 12; i++) {
    const d = new Date(TODAY);
    d.setMonth(d.getMonth() + i);
    const label = d.toLocaleDateString('en-US', { month: 'short', year: '2-digit' });
    let expected = 0;
    let balloonAmt = 0;
    NOTES.forEach(n => {
      const bd = new Date(n.balloonDate);
      if (bd.getFullYear() === d.getFullYear() && bd.getMonth() === d.getMonth()) {
        balloonAmt += n.currentBalance;
      } else if (bd > d) {
        expected += n.monthlyPmt;
      }
    });
    months.push({ label, expected, balloon: balloonAmt });
  }

  document.getElementById('cfKpi').innerHTML = [
    {label: 'Monthly Collections', value: fmt(totalMonthly), cls: ''},
    {label: 'Annual Expected', value: fmt(annualExpected), cls: ''},
    {label: 'Total Outstanding', value: fmt(totalBal), cls: ''},
    {label: 'Active Notes', value: NOTES.length, cls: ''}
  ].map(k => '<div class="kpi"><div class="label">' + k.label + '</div><div class="value ' + k.cls + '">' + k.value + '</div></div>').join('');

  new Chart(document.getElementById('cfChart'), {
    type: 'bar',
    data: {
      labels: months.map(m => m.label),
      datasets: [
        { label: 'Monthly Payments', data: months.map(m => m.expected), backgroundColor: '#2ecc71', borderRadius: 4 },
        { label: 'Balloon Payoffs', data: months.map(m => m.balloon), backgroundColor: '#c9952b', borderRadius: 4 }
      ]
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { labels: { color: '#8890a4' } } }, scales: { y: { stacked: false, ticks: { color: '#8890a4', callback: v => '$' + (v/1000).toFixed(0) + 'K' }, grid: { color: '#2a2d3e' } }, x: { ticks: { color: '#8890a4' }, grid: { display: false } } } }
  });

  // Balloon payoffs by year
  const balloonYears = {};
  NOTES.forEach(n => {
    const y = new Date(n.balloonDate).getFullYear();
    balloonYears[y] = (balloonYears[y] || 0) + n.currentBalance;
  });
  const yKeys = Object.keys(balloonYears).sort();
  new Chart(document.getElementById('balloonPayoffChart'), {
    type: 'bar',
    data: { labels: yKeys, datasets: [{ label: 'Balloon Balance', data: yKeys.map(k => balloonYears[k]), backgroundColor: '#3498db', borderRadius: 4 }] },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { y: { ticks: { color: '#8890a4', callback: v => '$' + (v/1000000).toFixed(1) + 'MM' }, grid: { color: '#2a2d3e' } }, x: { ticks: { color: '#8890a4' }, grid: { display: false } } } }
  });

  // Balloon table
  const sorted = [...NOTES].sort((a, b) => new Date(a.balloonDate) - new Date(b.balloonDate));
  document.getElementById('balloonBody').innerHTML = sorted.map(n => {
    const bd = new Date(n.balloonDate);
    const monthsLeft = Math.round((bd - TODAY) / (1000*60*60*24*30.44));
    const warn = monthsLeft <= 12;
    return '<tr>' +
      '<td>' + n.address + '</td>' +
      '<td>' + n.borrower.substring(0, 35) + '</td>' +
      '<td' + (warn ? ' class="balloon-warn"' : '') + '>' + fmtDate(n.balloonDate) + '</td>' +
      '<td class="num"' + (warn ? ' style="color:#e67e22;font-weight:600;"' : '') + '>' + monthsLeft + '</td>' +
      '<td class="num">' + fmt(n.currentBalance) + '</td>' +
      '<td class="num">' + fmt(n.monthlyPmt) + '</td>' +
      '</tr>';
  }).join('');
}

// ========== NOTE DETAIL MODAL ==========
window.showDetail = function(id) {
  const n = NOTES.find(x => x.id === id);
  if (!n) return;
  const modal = document.getElementById('noteModal');
  modal.innerHTML =
    '<div class="modal-header"><div><h2>' + n.address + '</h2><div class="modal-sub">' + n.borrower + ' | Closing #' + n.closing + '</div></div><div class="close-btn" onclick="document.getElementById(\\'modalOverlay\\').classList.remove(\\'show\\')">&times;</div></div>' +
    '<div class="detail-grid">' +
    '<div class="detail-item"><div class="dl">Loan Amount</div><div class="dv">' + fmt(n.loanAmt) + '</div></div>' +
    '<div class="detail-item"><div class="dl">Interest Rate</div><div class="dv">' + fmtPct(n.rate) + '</div></div>' +
    '<div class="detail-item"><div class="dl">Monthly Payment</div><div class="dv">' + fmt(n.monthlyPmt) + '</div></div>' +
    '<div class="detail-item"><div class="dl">Current Balance</div><div class="dv">' + fmt(n.currentBalance) + '</div></div>' +
    '<div class="detail-item"><div class="dl">Begin Date</div><div class="dv">' + fmtDate(n.beginDate) + '</div></div>' +
    '<div class="detail-item"><div class="dl">Balloon Date</div><div class="dv">' + fmtDate(n.balloonDate) + '</div></div>' +
    '<div class="detail-item"><div class="dl">Last Payment</div><div class="dv">' + fmtDate(n.lastPmtDate) + '</div></div>' +
    '<div class="detail-item"><div class="dl">Days Past Due</div><div class="dv" style="color:' + (n.daysPastDue > 30 ? '#e74c3c' : n.daysPastDue > 0 ? '#e67e22' : '#2ecc71') + '">' + n.daysPastDue + '</div></div>' +
    '<div class="detail-item"><div class="dl">Status</div><div class="dv"><span class="status ' + statusClass(n.status) + '">' + n.status + '</span></div></div>' +
    '</div>' +
    '<h3 style="color:#fff;margin-bottom:10px;">Recent Payments</h3>' +
    '<table><thead><tr><th>Date</th><th>Payment</th><th>Interest</th><th>Principal</th><th>Balance</th></tr></thead><tbody>' +
    n.payments.map(p =>
      '<tr><td>' + fmtDate(p.date) + '</td><td class="num">' + fmt(p.payment) + '</td><td class="num">' + fmt(p.interest) + '</td><td class="num">' + fmt(p.principal) + '</td><td class="num">' + (p.balance != null ? fmt(p.balance) : '--') + '</td></tr>'
    ).join('') +
    '</tbody></table>';
  document.getElementById('modalOverlay').classList.add('show');
};

document.getElementById('modalOverlay').addEventListener('click', function(e) {
  if (e.target === this) this.classList.remove('show');
});
document.addEventListener('keydown', e => { if (e.key === 'Escape') document.getElementById('modalOverlay').classList.remove('show'); });

// ========== INIT ==========
renderOverview();
renderPayments();
renderCollections();
renderInsurance();
renderCashFlow();
<\/script>
</body>
</html>`;
}

build().catch(e => { console.error(e); process.exit(1); });
