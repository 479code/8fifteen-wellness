/**
 * 8fifteen Wellness — Google Apps Script Backend
 * ─────────────────────────────────────────────
 * SETUP INSTRUCTIONS (takes ~3 minutes):
 *
 * 1. Open your Google Sheet
 * 2. Click  Extensions > Apps Script
 * 3. Delete all existing code and paste this entire file
 * 4. Click  Deploy > New deployment
 * 5. Type:   Web app
 * 6. Set     "Execute as"  →  Me
 * 7. Set     "Who has access" → Anyone
 * 8. Click   Deploy  →  Authorise  →  Allow
 * 9. Copy the Web App URL  (looks like: https://script.google.com/macros/s/AKfy.../exec)
 * 10. Paste that URL into the app's  Settings > Google Sheets URL
 *
 * IMPORTANT: After any code change you must click
 * "Deploy > Manage deployments > Edit (pencil) > New version > Deploy"
 * to update the live endpoint.
 * ─────────────────────────────────────────────
 */

// ── Sheet tab names (must match your spreadsheet exactly) ──
const TAB = {
  Products : 'Products',
  Sales    : 'sales',
  SaleItems: 'saleItems',
  Users    : 'Users',
  Expenses : 'Expenses',   // auto-created if missing
  StockLog : 'StockLog',   // auto-created if missing
};

// ── Column headers for each tab ──
const COLS = {
  Products : ['id','sku','name','category','cost','price','stock','low'],
  Sales    : ['id','receipt','date','time','payment','staff','total','cost','profit'],
  SaleItems: ['id','saleId','productId','name','category','qty','price','cost'],
  Users    : ['id','name','email','role'],
  Expenses : ['id','date','description','category','amount'],
  StockLog : ['id','productId','name','qty','reason','time','date'],
};

// ── Spreadsheet reference ──
function SS() { return SpreadsheetApp.getActiveSpreadsheet(); }

// ── Get or create a sheet, adding header row if brand new ──
function getSheet(tabKey) {
  const name = TAB[tabKey];
  let sheet = SS().getSheetByName(name);
  if (!sheet) {
    sheet = SS().insertSheet(name);
    sheet.appendRow(COLS[tabKey]);
    sheet.getRange(1, 1, 1, COLS[tabKey].length)
         .setFontWeight('bold')
         .setBackground('#1A1814')
         .setFontColor('#F1EBDD');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── Read all rows of a tab as an array of objects ──
function readAll(tabKey) {
  const sheet = getSheet(tabKey);
  const data  = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

// ── Append a new row ──
function appendRow(tabKey, obj) {
  const sheet = getSheet(tabKey);
  sheet.appendRow(COLS[tabKey].map(h => obj[h] !== undefined ? obj[h] : ''));
}

// ── Update a row matched by id field ──
function updateRow(tabKey, obj) {
  const sheet   = getSheet(tabKey);
  const values  = sheet.getDataRange().getValues();
  const idIdx   = values[0].indexOf('id');
  const rowIdx  = values.findIndex((r, i) => i > 0 && String(r[idIdx]) === String(obj.id));
  if (rowIdx < 1) return false;
  sheet.getRange(rowIdx + 1, 1, 1, COLS[tabKey].length)
       .setValues([COLS[tabKey].map(h => obj[h] !== undefined ? obj[h] : '')]);
  return true;
}

// ── Delete a row matched by id ──
function deleteRowById(tabKey, id) {
  const sheet  = getSheet(tabKey);
  const values = sheet.getDataRange().getValues();
  const idIdx  = values[0].indexOf('id');
  const rowIdx = values.findIndex((r, i) => i > 0 && String(r[idIdx]) === String(id));
  if (rowIdx > 0) sheet.deleteRow(rowIdx + 1);
}

// ══════════════════════════════════════════════
//  doGet — reads  (called with ?action=...)
// ══════════════════════════════════════════════
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'getAllData';
    let result;

    if (action === 'getAllData') {
      // Products
      const products = readAll('Products').map(p => ({
        ...p,
        cost : Number(p.cost)  || 0,
        price: Number(p.price) || 0,
        stock: Number(p.stock) || 0,
        low  : Number(p.low)   || 0,
      }));

      // Sales + saleItems joined
      const salesRaw  = readAll('Sales');
      const saleItems = readAll('SaleItems').map(i => ({
        ...i,
        qty  : Number(i.qty)   || 0,
        price: Number(i.price) || 0,
        cost : Number(i.cost)  || 0,
      }));
      const sales = salesRaw.map(s => ({
        ...s,
        total : Number(s.total)  || 0,
        cost  : Number(s.cost)   || 0,
        profit: Number(s.profit) || 0,
        items : saleItems.filter(i => String(i.saleId) === String(s.id)),
      }));

      // Expenses
      const expenses = readAll('Expenses').map(ex => ({
        ...ex, amount: Number(ex.amount) || 0,
      }));

      // Users
      const users = readAll('Users');

      result = { ok: true, products, sales, expenses, users };

    } else if (action === 'ping') {
      result = { ok: true, message: '8fifteen backend is alive ✓' };

    } else {
      result = { ok: false, error: 'Unknown action: ' + action };
    }

    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

// ══════════════════════════════════════════════
//  doPost — writes  (body: URLSearchParams)
// ══════════════════════════════════════════════
function doPost(e) {
  try {
    const params = e.parameter;
    const action = params.action;
    const data   = params.data ? JSON.parse(params.data) : {};
    let result   = { ok: true };

    // ── Products ──
    if (action === 'addProduct') {
      appendRow('Products', data);

    } else if (action === 'updateProduct') {
      updateRow('Products', data);

    } else if (action === 'deleteProduct') {
      deleteRowById('Products', data.id);

    } else if (action === 'updateStock') {
      // data: { id, stock }
      const sheet  = getSheet('Products');
      const values = sheet.getDataRange().getValues();
      const idIdx  = values[0].indexOf('id');
      const skIdx  = values[0].indexOf('stock');
      const rowIdx = values.findIndex((r, i) => i > 0 && String(r[idIdx]) === String(data.id));
      if (rowIdx > 0) sheet.getRange(rowIdx + 1, skIdx + 1).setValue(data.stock);

    // ── Sales ──
    } else if (action === 'addSale') {
      // data: sale object with .items array
      const saleRow = { ...data };
      delete saleRow.items;
      appendRow('Sales', saleRow);
      (data.items || []).forEach(item => appendRow('SaleItems', item));

    // ── Stock log ──
    } else if (action === 'logStock') {
      appendRow('StockLog', data);

    // ── Expenses ──
    } else if (action === 'addExpense') {
      appendRow('Expenses', data);

    } else if (action === 'deleteExpense') {
      deleteRowById('Expenses', data.id);

    // ── Users ──
    } else if (action === 'addUser') {
      appendRow('Users', data);

    } else if (action === 'deleteUser') {
      deleteRowById('Users', data.id);

    // ── Seed demo data ──
    } else if (action === 'seedDemo') {
      seedDemoData();

    } else {
      result = { ok: false, error: 'Unknown action: ' + action };
    }

    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

// ── Helper: return JSON response ──
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Seed demo products if sheet is empty ──
function seedDemoData() {
  const sheet = getSheet('Products');
  if (sheet.getLastRow() > 1) return; // already has data

  const demo = [
    ['p1','BTY-001','Shea Glow Body Butter','Beauty',2500,4500,18,8],
    ['p2','BTY-002','Rosewater Facial Mist','Beauty',1800,3200,5,6],
    ['p3','BTY-003','Charcoal Clay Mask','Beauty',2600,5000,12,5],
    ['p4','BTY-004','Vitamin C Serum 30ml','Beauty',5000,8500,3,5],
    ['p5','PC-001','Coconut & Aloe Body Wash','Personal Care',1600,2800,24,10],
    ['p6','PC-002','Bamboo Toothbrush Set','Personal Care',800,1500,30,10],
    ['p7','PC-003','Lavender Hand Cream','Personal Care',1200,2200,7,8],
    ['p8','PC-004','Unscented Bar Soap','Personal Care',450,900,40,15],
    ['p9','TEA-001','Hibiscus & Ginger Loose Tea','Tea',1600,3000,14,6],
    ['p10','TEA-002','Moringa Green Tea Bags (20s)','Tea',1400,2600,2,6],
    ['p11','TEA-003','Chamomile Sleep Blend','Tea',1500,2900,9,6],
    ['p12','TEA-004','Lemongrass & Mint Tea','Tea',1400,2700,16,6],
    ['p13','GEN-001','Reusable Glass Tumbler','General',2000,3500,11,5],
    ['p14','GEN-002','Tote Bag — 8fifteen Print','General',2200,4000,6,5],
    ['p15','GEN-003','Scented Soy Candle','General',2800,5500,4,6],
  ];
  demo.forEach(row => sheet.appendRow(row));
}
