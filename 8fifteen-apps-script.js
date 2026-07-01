// ════════════════════════════════════════════════════════════
//  8fifteen Wellness — Google Apps Script Backend
//  Handles: CRUD for all data + Monthly Email Reports
// ════════════════════════════════════════════════════════════

const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

function doGet(e) {
  const action = e.parameter.action;
  if (action === 'getAllData') return ok(getAllData());
  if (action === 'ping')       return ok({pong: true});
  return ok({error: 'Unknown action'});
}

function doPost(e) {
  try {
    const action = e.parameter.action;
    const data   = JSON.parse(e.parameter.data || '{}');

    if (action === 'addProduct')     return ok(addRow('Products', productRow(data)));
    if (action === 'updateProduct')  return ok(updateRow('Products', data.id, productRow(data)));
    if (action === 'deleteProduct')  return ok(deleteRow('Products', data.id));
    if (action === 'updateStock')    return ok(updateStock(data));
    if (action === 'addSale')        return ok(addSale(data));
    if (action === 'logStock')       return ok(addRow('StockLog', [data.id,data.productId,data.name,data.qty,data.reason,data.time,data.date]));
    if (action === 'addExpense')     return ok(addRow('Expenses', [data.id,data.date,data.description,data.category,data.amount]));
    if (action === 'deleteExpense')  return ok(deleteRow('Expenses', data.id));
    if (action === 'addUser')        return ok(addRow('Users', [data.id,data.name,data.email,data.role]));
    if (action === 'deleteUser')     return ok(deleteRow('Users', data.id));
    if (action === 'addAttendance')  return ok(addRow('Attendance', [
      data.id, data.staffName, data.staffEmail||'', data.type,
      data.date, data.time, data.lat||'', data.lng||'',
      data.distance||'', data.withinFence===null?'manual':data.withinFence?'yes':'no',
      data.manual?'yes':'no'
    ]));

    // ── Email actions ──
    if (action === 'sendTestEmail') {
      const result = sendMonthlyReport(data.email, 0); // current month
      return ok({sent: true, message: result});
    }
    if (action === 'setupEmailTrigger') {
      setupMonthlyTrigger(data.email);
      return ok({setup: true});
    }
    if (action === 'removeEmailTrigger') {
      removeMonthlyTrigger();
      return ok({removed: true});
    }

    return ok({error: 'Unknown action: ' + action});
  } catch(err) {
    return ok({ok: false, error: err.message});
  }
}

function ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify({ok: true, ...data}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Sheet helpers ──
function getSheet(name) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
function getAllData() {
  return {
    products: readSheet('Products', ['id','sku','name','category','cost','price','stock','low','subcategory']),
    sales:    readSales(),
    expenses: readSheet('Expenses', ['id','date','description','category','amount']),
    users:    readSheet('Users',    ['id','name','email','role']),
    attendance: readSheet('Attendance', ['id','staffName','staffEmail','type','date','time','lat','lng','distance','withinFence','manual']),
  };
}
function readSheet(sheetName, cols) {
  const sh = getSheet(sheetName);
  const rows = sh.getDataRange().getValues();
  if (rows.length <= 1) return [];
  return rows.slice(1).filter(r => r[0]).map(r =>
    Object.fromEntries(cols.map((c, i) => [c, r[i] !== undefined ? r[i] : '']))
  );
}
function readSales() {
  const sh = getSheet('sales');
  const rows = sh.getDataRange().getValues();
  if (rows.length <= 1) return [];
  const sales = {};
  rows.slice(1).filter(r => r[0]).forEach(r => {
    const saleId = r[1];
    if (!sales[saleId]) sales[saleId] = {
      id: saleId, receipt: r[2], date: r[3], time: r[4],
      payment: r[5], staff: r[6], total: parseFloat(r[7])||0,
      cost: parseFloat(r[8])||0, profit: parseFloat(r[9])||0,
      discount: parseFloat(r[10])||0, customer: r[11]||'', items: []
    };
    if (r[12]) sales[saleId].items.push({
      id: r[0], productId: r[12], name: r[13], category: r[14],
      qty: parseInt(r[15])||0, price: parseFloat(r[16])||0, cost: parseFloat(r[17])||0
    });
  });
  return Object.values(sales);
}
function addRow(sheetName, rowData) {
  getSheet(sheetName).appendRow(rowData);
  return {added: true};
}
function updateRow(sheetName, id, rowData) {
  const sh = getSheet(sheetName);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] == id) { sh.getRange(i+1, 1, 1, rowData.length).setValues([rowData]); return {updated:true}; }
  }
  return {notFound: true};
}
function deleteRow(sheetName, id) {
  const sh = getSheet(sheetName);
  const rows = sh.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (rows[i][0] == id) { sh.deleteRow(i+1); return {deleted:true}; }
  }
  return {notFound: true};
}
function productRow(d) {
  return [d.id, d.sku||'', d.name||'', d.category||'', d.cost||0, d.price||0, d.stock||0, d.low||5, d.subcategory||''];
}
function addSale(sale) {
  const sh = getSheet('sales');
  if (!sale.items || !sale.items.length) {
    sh.appendRow([uid(), sale.id, sale.receipt, sale.date, sale.time, sale.payment, sale.staff,
      sale.total, sale.cost, sale.profit, sale.discount||0, sale.customer||'', '', '', '', '', '', '']);
    return {added: true};
  }
  sale.items.forEach(item => {
    sh.appendRow([item.id||uid(), sale.id, sale.receipt, sale.date, sale.time, sale.payment, sale.staff,
      sale.total, sale.cost, sale.profit, sale.discount||0, sale.customer||'',
      item.productId, item.name, item.category, item.qty, item.price, item.cost]);
  });
  return {added: true};
}
function updateStock(data) {
  const sh = getSheet('Products');
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] == data.id) { sh.getRange(i+1, 7).setValue(data.stock); return {updated:true}; }
  }
  return {notFound: true};
}
function uid() { return 'r_' + Math.random().toString(36).slice(2,9); }

// ════════════════════════════════════════════════════════════
//  MONTHLY EMAIL REPORT
// ════════════════════════════════════════════════════════════

/**
 * Send the monthly report email.
 * monthOffset: 0 = current month (for test), 1 = last month (for auto trigger)
 */
function sendMonthlyReport(emailTo, monthOffset) {
  if (!emailTo) throw new Error('No recipient email provided');

  const now    = new Date();
  const target = new Date(now.getFullYear(), now.getMonth() - (monthOffset||0), 1);
  const year   = target.getFullYear();
  const month  = target.getMonth(); // 0-indexed
  const monthName = target.toLocaleString('en-US', {month:'long'});

  // Date range for the target month
  const startStr = `${year}-${String(month+1).padStart(2,'0')}-01`;
  const endDate  = new Date(year, month+1, 0);
  const endStr   = `${year}-${String(month+1).padStart(2,'0')}-${String(endDate.getDate()).padStart(2,'0')}`;

  // Load data
  const allSales    = readSales();
  const allProducts = readSheet('Products', ['id','sku','name','category','cost','price','stock','low','subcategory']);
  const allExpenses = readSheet('Expenses', ['id','date','description','category','amount']);

  // Filter to target month
  const monthSales = allSales.filter(s => s.date >= startStr && s.date <= endStr);
  const monthExp   = allExpenses.filter(e => e.date >= startStr && e.date <= endStr);

  // Previous month for comparison
  const prevStart = new Date(year, month-1, 1);
  const prevEnd   = new Date(year, month, 0);
  const prevStartStr = formatDate(prevStart);
  const prevEndStr   = formatDate(prevEnd);
  const prevSales = allSales.filter(s => s.date >= prevStartStr && s.date <= prevEndStr);

  // Aggregates
  const revenue = monthSales.reduce((s,x)=>s+x.total, 0);
  const cogs    = monthSales.reduce((s,x)=>s+x.cost,  0);
  const gross   = revenue - cogs;
  const expAmt  = monthExp.reduce((s,e)=>s+(parseFloat(e.amount)||0), 0);
  const net     = gross - expAmt;
  const txCount = monthSales.length;
  const avgOrd  = txCount ? Math.round(revenue/txCount) : 0;

  const prevRev = prevSales.reduce((s,x)=>s+x.total, 0);
  const revDelta = prevRev > 0 ? ((revenue-prevRev)/prevRev*100) : null;

  // Top products
  const prodMap = {};
  monthSales.forEach(s=>s.items.forEach(l=>{
    if(!prodMap[l.name]) prodMap[l.name]={name:l.name,qty:0,rev:0};
    prodMap[l.name].qty += parseInt(l.qty)||0;
    prodMap[l.name].rev += (parseFloat(l.price)||0) * (parseInt(l.qty)||0);
  }));
  const topProds = Object.values(prodMap).sort((a,b)=>b.rev-a.rev).slice(0,5);

  // Staff performance
  const staffMap = {};
  monthSales.forEach(s=>{
    const n=s.staff||'Unknown';
    if(!staffMap[n]) staffMap[n]={name:n,revenue:0,transactions:0};
    staffMap[n].revenue+=s.total; staffMap[n].transactions++;
  });
  const staffList = Object.values(staffMap).sort((a,b)=>b.revenue-a.revenue);

  // Low stock
  const lowStock = allProducts.filter(p=>(parseInt(p.stock)||0)<=(parseInt(p.low)||5) && (parseInt(p.stock)||0)>0);
  const outStock  = allProducts.filter(p=>(parseInt(p.stock)||0)===0);

  // Expense breakdown
  const expCats = {};
  monthExp.forEach(e=>{expCats[e.category]=(expCats[e.category]||0)+(parseFloat(e.amount)||0);});

  // Build HTML email
  const html = buildEmailHtml({
    monthName, year, revenue, cogs, gross, expAmt, net, txCount, avgOrd,
    revDelta, topProds, staffList, lowStock, outStock, expCats,
    startStr, endStr
  });

  GmailApp.sendEmail(emailTo, `8fifteen Wellness — ${monthName} ${year} Report`, '', {
    htmlBody: html,
    name: '8fifteen Wellness'
  });

  return `Report sent to ${emailTo} for ${monthName} ${year}`;
}

function formatDate(d) {
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}

function fmtN(n) { return '₦' + Math.round(n||0).toLocaleString('en-US'); }
function fmtPct(n) { return (isFinite(n)?n:0).toFixed(1) + '%'; }

function buildEmailHtml(d) {
  const deltaHtml = d.revDelta !== null
    ? `<span style="color:${d.revDelta>=0?'#3E6B4A':'#9B4B3A'};font-size:12px;margin-left:8px">${d.revDelta>=0?'↑':'↓'}${Math.abs(d.revDelta).toFixed(0)}% vs last month</span>`
    : '';

  const gm = d.revenue > 0 ? (d.gross/d.revenue*100) : 0;
  const nm = d.revenue > 0 ? (d.net/d.revenue*100)   : 0;

  return `<!DOCTYPE html>
<html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<style>
  body{margin:0;padding:0;background:#F1EBDD;font-family:'Helvetica Neue',Arial,sans-serif;color:#1A1814;}
  .wrap{max-width:600px;margin:0 auto;background:#F1EBDD;}
  .header{background:#131210;padding:36px 32px;text-align:center;color:#EFE9DA;}
  .logo-word{font-size:26px;font-weight:900;letter-spacing:.15em;color:#C9A766;}
  .logo-sub{font-size:9px;letter-spacing:.22em;color:#948C7A;margin-top:5px;}
  .month-title{font-size:14px;color:#948C7A;margin-top:12px;}
  .body{padding:28px 32px;}
  .kpi-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:12px;margin-bottom:24px;}
  .kpi{background:#fff;border:1px solid #E2DAC6;border-radius:10px;padding:16px;text-align:center;}
  .kpi-label{font-size:10px;text-transform:uppercase;letter-spacing:.07em;color:#8A8170;font-weight:600;margin-bottom:6px;}
  .kpi-value{font-size:22px;font-weight:700;color:#1A1814;}
  .kpi-value.green{color:#3E6B4A;}
  .kpi-value.red{color:#9B4B3A;}
  .section{margin-bottom:24px;}
  .section-title{font-size:11px;text-transform:uppercase;letter-spacing:.08em;color:#8A8170;font-weight:700;border-bottom:2px solid #E2DAC6;padding-bottom:8px;margin-bottom:14px;}
  table{width:100%;border-collapse:collapse;font-size:13px;}
  th{font-size:10px;text-transform:uppercase;letter-spacing:.06em;color:#8A8170;font-weight:600;padding:7px 0;border-bottom:1px solid #E2DAC6;text-align:left;}
  td{padding:9px 0;border-bottom:1px solid #F1EBDD;color:#1A1814;}
  td.right,th.right{text-align:right;}
  td.mono{font-family:'Courier New',monospace;font-size:12px;}
  .pl-row{display:flex;justify-content:space-between;padding:8px 0;font-size:13px;border-bottom:1px dashed #E2DAC6;}
  .pl-row.total{border-bottom:none;border-top:2px solid #1A1814;padding-top:12px;font-weight:700;margin-top:4px;}
  .badge{display:inline-block;padding:2px 9px;border-radius:999px;font-size:10px;font-weight:700;}
  .badge-warn{background:#F2E6C9;color:#8A6A2E;}
  .badge-danger{background:#F3DFD8;color:#9B4B3A;}
  .footer{background:#131210;padding:24px 32px;text-align:center;color:#948C7A;font-size:11px;line-height:1.8;}
  @media(max-width:480px){.kpi-grid{grid-template-columns:1fr 1fr;}.body{padding:20px 18px;}}
</style></head>
<body><div class="wrap">

<!-- Header -->
<div class="header">
  <div class="logo-word">8fifteen WELLNESS</div>
  <div class="logo-sub">TEA · SKIN · SCENT</div>
  <div class="month-title">Monthly Report — ${d.monthName} ${d.year}</div>
</div>

<div class="body">

  <!-- KPI grid -->
  <div class="kpi-grid">
    <div class="kpi">
      <div class="kpi-label">Revenue ${deltaHtml}</div>
      <div class="kpi-value">${fmtN(d.revenue)}</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Gross Profit</div>
      <div class="kpi-value green">${fmtN(d.gross)}</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Net Profit</div>
      <div class="kpi-value ${d.net>=0?'green':'red'}">${fmtN(d.net)}</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Transactions</div>
      <div class="kpi-value">${d.txCount}</div>
    </div>
  </div>

  <!-- P&L summary -->
  <div class="section">
    <div class="section-title">Profit &amp; Loss Summary</div>
    <div class="pl-row"><span>Revenue</span><span class="mono">${fmtN(d.revenue)}</span></div>
    <div class="pl-row"><span style="color:#8A8170;padding-left:14px">Cost of goods sold</span><span class="mono" style="color:#9B4B3A">(${fmtN(d.cogs)})</span></div>
    <div class="pl-row"><span style="font-weight:600">Gross profit</span><span class="mono" style="font-weight:600">${fmtN(d.gross)} <small style="color:#8A8170">${fmtPct(gm)}</small></span></div>
    <div class="pl-row"><span style="color:#8A8170;padding-left:14px">Operating expenses</span><span class="mono" style="color:#9B4B3A">(${fmtN(d.expAmt)})</span></div>
    <div class="pl-row total"><span>Net profit</span><span class="mono" style="color:${d.net>=0?'#3E6B4A':'#9B4B3A'}">${fmtN(d.net)} <small style="font-weight:400;color:#8A8170">${fmtPct(nm)}</small></span></div>
    <div style="margin-top:10px;font-size:12px;color:#8A8170">Avg order value: <b style="color:#1A1814">${fmtN(d.avgOrd)}</b></div>
  </div>

  <!-- Top products -->
  ${d.topProds.length > 0 ? `
  <div class="section">
    <div class="section-title">Top products — ${d.monthName}</div>
    <table>
      <thead><tr><th>#</th><th>Product</th><th class="right">Units</th><th class="right">Revenue</th></tr></thead>
      <tbody>
        ${d.topProds.map((p,i)=>`<tr><td style="color:#8A8170">${i+1}</td><td>${p.name}</td><td class="mono right">${p.qty}</td><td class="mono right">${fmtN(p.rev)}</td></tr>`).join('')}
      </tbody>
    </table>
  </div>` : ''}

  <!-- Staff performance -->
  ${d.staffList.length > 0 ? `
  <div class="section">
    <div class="section-title">Staff performance</div>
    <table>
      <thead><tr><th>Staff member</th><th class="right">Transactions</th><th class="right">Revenue</th></tr></thead>
      <tbody>
        ${d.staffList.map((s,i)=>`<tr>
          <td>${i===0?'🏆 ':''}<b>${s.name}</b></td>
          <td class="mono right">${s.transactions}</td>
          <td class="mono right" style="color:#3E6B4A">${fmtN(s.revenue)}</td>
        </tr>`).join('')}
      </tbody>
    </table>
  </div>` : ''}

  <!-- Expenses -->
  ${Object.keys(d.expCats).length > 0 ? `
  <div class="section">
    <div class="section-title">Operating expenses breakdown</div>
    <table>
      <thead><tr><th>Category</th><th class="right">Amount</th></tr></thead>
      <tbody>
        ${Object.entries(d.expCats).sort((a,b)=>b[1]-a[1]).map(([cat,amt])=>`<tr><td>${cat}</td><td class="mono right">${fmtN(amt)}</td></tr>`).join('')}
        <tr><td style="font-weight:700;border-top:1px solid #E2DAC6;padding-top:8px">Total</td><td class="mono right" style="font-weight:700;border-top:1px solid #E2DAC6;padding-top:8px">${fmtN(d.expAmt)}</td></tr>
      </tbody>
    </table>
  </div>` : ''}

  <!-- Stock alerts -->
  ${(d.outStock.length > 0 || d.lowStock.length > 0) ? `
  <div class="section">
    <div class="section-title">⚠ Stock alerts</div>
    ${d.outStock.map(p=>`<div style="display:flex;justify-content:space-between;padding:8px 0;border-bottom:1px solid #F1EBDD;font-size:13px"><span>${p.name}</span><span class="badge badge-danger">Out of stock</span></div>`).join('')}
    ${d.lowStock.map(p=>`<div style="display:flex;justify-content:space-between;padding:8px 0;border-bottom:1px solid #F1EBDD;font-size:13px"><span>${p.name}</span><span class="badge badge-warn">${p.stock} left</span></div>`).join('')}
  </div>` : ''}

</div><!-- /body -->

<!-- Footer -->
<div class="footer">
  8fifteen Wellness — Smart Inventory &amp; Sales<br/>
  ${d.startStr} to ${d.endStr}<br/>
  <span style="font-size:10px;opacity:.6">This is an automated monthly report.</span>
</div>

</div></body></html>`;
}

// ── Monthly trigger setup ──
function setupMonthlyTrigger(emailTo) {
  // Remove existing triggers for this function
  removeMonthlyTrigger();
  // Create new monthly trigger on the 1st of every month at 8am
  ScriptApp.newTrigger('sendAutoMonthlyReport')
    .timeBased()
    .onMonthDay(1)
    .atHour(8)
    .nearMinute(0)
    .create();
  // Store the email address
  PropertiesService.getScriptProperties().setProperty('REPORT_EMAIL', emailTo);
  Logger.log('Monthly trigger set up for: ' + emailTo);
}

function removeMonthlyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendAutoMonthlyReport')
    .forEach(t => ScriptApp.deleteTrigger(t));
  PropertiesService.getScriptProperties().deleteProperty('REPORT_EMAIL');
}

// Called automatically by the time-based trigger
function sendAutoMonthlyReport() {
  const emailTo = PropertiesService.getScriptProperties().getProperty('REPORT_EMAIL');
  if (!emailTo) {
    Logger.log('No report email configured. Run setupMonthlyTrigger(email) first.');
    return;
  }
  try {
    const result = sendMonthlyReport(emailTo, 1); // 1 = last month
    Logger.log(result);
  } catch(err) {
    Logger.log('Monthly report error: ' + err.message);
    GmailApp.sendEmail(emailTo, '8fifteen — Report Error', 'Monthly report failed: ' + err.message);
  }
}
