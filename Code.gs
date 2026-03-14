// ============================================================
// Google Apps Script - PO Manager Backend
// ============================================================
// SETUP INSTRUCTIONS:
// 1. Go to https://script.google.com and create a new project
// 2. Paste this entire code into Code.gs
// 3. Create a Google Sheet and copy its ID from the URL
//    (the long string between /d/ and /edit in the sheet URL)
// 4. Replace the SPREADSHEET_ID below with your Sheet ID
// 5. Run the "setupSheets" function once to create the required sheets
// 6. Deploy > New deployment > Web app
//    - Execute as: Me
//    - Who has access: Anyone
// 7. Copy the Web App URL and paste it into main.js as GOOGLE_SCRIPT_URL
// ============================================================

const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';

// Sheet names
const SHEET_PO = 'PurchaseOrders';
const SHEET_VENDORS = 'Vendors';
const SHEET_ITEMS = 'Items';
const SHEET_USERS = 'Users';
const SHEET_SETTINGS = 'Settings';

// Cached spreadsheet reference — avoids repeated openById() calls per request
function getSS_() {
  if (!getSS_._ss) getSS_._ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return getSS_._ss;
}
function getSheet_(name) {
  return getSS_().getSheetByName(name);
}

// ============================================================
// SETUP - Run this once to create sheets with headers
// ============================================================
function setupSheets() {
  const ss = getSS_();

  // Purchase Orders sheet
  let poSheet = ss.getSheetByName(SHEET_PO);
  if (!poSheet) {
    poSheet = ss.insertSheet(SHEET_PO);
    poSheet.appendRow([
      'po_id', 'po_no', 'po_date', 'vendor_name', 'vendor_address', 'vendor_gstin',
      'client_name', 'event_name', 'event_location', 'event_date',
      'items', 'transport', 'cgst_percent', 'sgst_percent',
      'terms', 'total', 'created_at'
    ]);
    poSheet.getRange(1, 1, 1, 17).setFontWeight('bold');
  }

  // Vendors sheet
  let vendorSheet = ss.getSheetByName(SHEET_VENDORS);
  if (!vendorSheet) {
    vendorSheet = ss.insertSheet(SHEET_VENDORS);
    vendorSheet.appendRow(['name', 'address', 'gstin']);
    vendorSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  }

  // Items sheet
  let itemSheet = ss.getSheetByName(SHEET_ITEMS);
  if (!itemSheet) {
    itemSheet = ss.insertSheet(SHEET_ITEMS);
    itemSheet.appendRow(['desc', 'uom', 'rate', 'default_qty']);
    itemSheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  }

  SpreadsheetApp.getUi().alert('Sheets created successfully!');
}

// ============================================================
// WEB APP ENTRY POINTS
// ============================================================
function doGet(e) {
  const action = (e.parameter && e.parameter.action) || '';
  let result;

  switch (action) {
    case 'read':
      result = readPOs();
      break;
    case 'read_vendors':
      result = readVendors();
      break;
    case 'read_items':
      result = readItems();
      break;
    case 'read_users':
      result = readUsers();
      break;
    case 'read_all':
      result = { pos: readPOs(), vendors: readVendors(), items: readItems() };
      break;
    case 'next_po_number':
      result = getNextPONumber();
      break;
    case 'read_settings':
      result = readSettings();
      break;
    default:
      result = { error: 'Unknown action: ' + action };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let body;
  try {
    const raw = e.postData ? e.postData.contents : '';
    body = JSON.parse(raw);
  } catch (err) {
    return jsonResponse({ success: false, error: 'Invalid JSON: ' + err.message });
  }

  const action = body.action || '';
  let result;

  switch (action) {
    // --- PO ---
    case 'create':
      result = createPO(body);
      break;
    case 'update':
      result = updatePO(body);
      break;
    case 'delete':
      result = deletePO(body);
      break;

    // --- VENDORS ---
    case 'add_vendor':
      result = addVendor(body);
      break;
    case 'update_vendor':
      result = updateVendor(body);
      break;
    case 'delete_vendor':
      result = deleteVendor(body);
      break;
    case 'bulk_upsert_vendors':
      result = bulkUpsertVendors(body.rows || []);
      break;

    // --- ITEMS ---
    case 'add_item':
      result = addItem(body);
      break;
    case 'update_item':
      result = updateItem(body);
      break;
    case 'delete_item':
      result = deleteItem(body);
      break;
    case 'bulk_upsert_items':
      result = bulkUpsertItems(body.rows || []);
      break;

    // --- BULK POs ---
    case 'bulk_upsert_pos':
      result = bulkUpsertPOs(body.rows || []);
      break;

    // --- USERS ---
    case 'save_users':
      result = saveUsers(body.users || []);
      break;

    // --- SETTINGS ---
    case 'save_setting':
      result = saveSetting(body.key, body.value);
      break;
    case 'save_settings_bulk':
      result = saveSettingsBulk(body.settings || {});
      break;

    default:
      result = { success: false, error: 'Unknown action: ' + action };
  }

  return jsonResponse(result);
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// AUTO PO NUMBER GENERATION  —  Format: PREFIX/EO/001
// ============================================================
function getNextPONumber() {
  // Read prefix and start-seq from Settings
  let poPrefix = 'RE';
  let startSeq = 1;
  try {
    const settings = readSettings();
    if (settings.po_prefix && String(settings.po_prefix).trim()) {
      poPrefix = String(settings.po_prefix).trim();
    }
    if (settings.po_start_seq) {
      const parsed = parseInt(String(settings.po_start_seq), 10);
      if (!isNaN(parsed) && parsed >= 1) startSeq = parsed;
    }
  } catch(e) {}

  const prefix = poPrefix + '/EO/';     // e.g. RE/EO/

  // Scan ALL existing POs for the highest global sequence matching this prefix
  const sheet = getSheet_(SHEET_PO);
  let maxSeq = 0;
  if (sheet && sheet.getLastRow() > 1) {
    const poNos = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues(); // col B = po_no
    poNos.forEach(function(row) {
      const val = String(row[0] || '').trim();
      if (val.indexOf(prefix) === 0) {
        const seqStr = val.substring(prefix.length);
        const seq = parseInt(seqStr, 10);
        if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
      }
    });
  }

  // Next = max existing + 1, but at least startSeq
  const nextNum = Math.max(maxSeq + 1, startSeq);
  const nextSeq = String(nextNum).padStart(3, '0');
  return { success: true, po_no: prefix + nextSeq, seq: nextNum };
}

// ============================================================
// PURCHASE ORDERS
// ============================================================
function readPOs() {
  const sheet = getSheet_(SHEET_PO);
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];

  const headers = rows[0];
  const data = [];
  for (let i = 1; i < rows.length; i++) {
    const obj = {};
    headers.forEach((h, idx) => {
      let val = rows[i][idx];
      // Parse items JSON string back to array
      if (h === 'items' && typeof val === 'string' && val.startsWith('[')) {
        try { val = JSON.parse(val); } catch (e) { /* keep as string */ }
      }
      // Format date values
      if ((h === 'po_date' || h === 'created_at') && val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      // Ensure po_no and po_id are always strings
      if ((h === 'po_no' || h === 'po_id') && val !== '' && val != null) {
        val = String(val).trim();
      }
      obj[h] = val;
    });
    data.push(obj);
  }
  return data;
}

function createPO(body) {
  const sheet = getSheet_(SHEET_PO);
  if (!sheet) return { success: false, error: 'Sheet not found' };

  const poId = Utilities.getUuid();
  const items = body.items || [];
  const itemsJson = JSON.stringify(items);

  // Calculate grand total
  const itemsTotal = items.reduce((sum, item) => sum + (parseFloat(item.total) || 0), 0);
  const transport = parseFloat(body.transport) || 0;
  const subtotal = itemsTotal + transport;
  const cgstPct = parseFloat(body.cgst_percent) || 0;
  const sgstPct = parseFloat(body.sgst_percent) || 0;
  const cgstAmt = Math.round(subtotal * cgstPct / 100);
  const sgstAmt = Math.round(subtotal * sgstPct / 100);
  const grandTotal = subtotal + cgstAmt + sgstAmt;

  sheet.appendRow([
    poId,
    body.po_no || '',
    body.po_date || '',
    body.vendor_name || '',
    body.vendor_address || '',
    body.vendor_gstin || '',
    body.client_name || '',
    body.event_name || '',
    body.event_location || '',
    body.event_date || '',
    itemsJson,
    transport,
    cgstPct,
    sgstPct,
    body.terms || '',
    grandTotal,
    new Date()
  ]);

  return { success: true, po_id: poId };
}

function updatePO(body) {
  const sheet = getSheet_(SHEET_PO);
  if (!sheet) return { success: false, error: 'Sheet not found' };

  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const poIdCol = headers.indexOf('po_id');
  const poNoCol = headers.indexOf('po_no');

  // Find the row by po_id or po_no
  let rowIndex = -1;
  for (let i = 1; i < rows.length; i++) {
    if (
      (body.po_id && rows[i][poIdCol] === body.po_id) ||
      (body.po_no && rows[i][poNoCol] === body.po_no)
    ) {
      rowIndex = i + 1; // 1-indexed for sheet
      break;
    }
  }

  if (rowIndex === -1) return { success: false, error: 'PO not found' };

  const items = body.items || [];
  const itemsJson = JSON.stringify(items);
  const itemsTotal = items.reduce((sum, item) => sum + (parseFloat(item.total) || 0), 0);
  const transport = parseFloat(body.transport) || 0;
  const subtotal = itemsTotal + transport;
  const cgstPct = parseFloat(body.cgst_percent) || 0;
  const sgstPct = parseFloat(body.sgst_percent) || 0;
  const cgstAmt = Math.round(subtotal * cgstPct / 100);
  const sgstAmt = Math.round(subtotal * sgstPct / 100);
  const grandTotal = subtotal + cgstAmt + sgstAmt;

  // Update columns (keep po_id and created_at unchanged)
  const updatedRow = [
    rows[rowIndex - 1][poIdCol],   // po_id (unchanged)
    body.po_no || '',
    body.po_date || '',
    body.vendor_name || '',
    body.vendor_address || '',
    body.vendor_gstin || '',
    body.client_name || '',
    body.event_name || '',
    body.event_location || '',
    body.event_date || '',
    itemsJson,
    transport,
    cgstPct,
    sgstPct,
    body.terms || '',
    grandTotal,
    rows[rowIndex - 1][headers.indexOf('created_at')] // created_at (unchanged)
  ];

  sheet.getRange(rowIndex, 1, 1, updatedRow.length).setValues([updatedRow]);
  return { success: true };
}

function deletePO(body) {
  const sheet = getSheet_(SHEET_PO);
  if (!sheet) return { success: false, error: 'Sheet not found' };

  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const poIdCol = headers.indexOf('po_id');
  const poNoCol = headers.indexOf('po_no');

  for (let i = 1; i < rows.length; i++) {
    if (
      (body.po_id && rows[i][poIdCol] === body.po_id) ||
      (body.po_no && rows[i][poNoCol] === body.po_no)
    ) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'PO not found' };
}

// ============================================================
// VENDORS
// ============================================================
function readVendors() {
  const sheet = getSheet_(SHEET_VENDORS);
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];

  const headers = rows[0];
  const data = [];
  for (let i = 1; i < rows.length; i++) {
    data.push({
      name: rows[i][0] || '',
      address: rows[i][1] || '',
      gstin: rows[i][2] || ''
    });
  }
  return data;
}

function addVendor(body) {
  const sheet = getSheet_(SHEET_VENDORS);
  if (!sheet) return { success: false, error: 'Sheet not found' };

  const name = (body.name || '').trim();
  if (!name) return { success: false, error: 'Vendor name is required' };

  const address = (body.address || '').trim();
  const gstin = (body.gstin || '').trim();

  // Check for duplicate
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0].toString().toLowerCase() === name.toLowerCase()) {
      return { success: false, error: 'Vendor already exists' };
    }
  }

  sheet.appendRow([name, address, gstin]);
  return { success: true };
}

function deleteVendor(body) {
  const sheet = getSheet_(SHEET_VENDORS);
  if (!sheet) return { success: false, error: 'Sheet not found' };

  const idx = parseInt(body.idx);
  if (isNaN(idx) || idx < 0) return { success: false, error: 'Invalid index' };

  const rowToDelete = idx + 2; // +1 for header, +1 for 1-indexed
  if (rowToDelete > sheet.getLastRow()) return { success: false, error: 'Index out of range' };

  sheet.deleteRow(rowToDelete);
  return { success: true };
}

function updateVendor(body) {
  const sheet = getSheet_(SHEET_VENDORS);
  if (!sheet) return { success: false, error: 'Sheet not found' };

  const idx = parseInt(body.idx, 10);
  if (isNaN(idx) || idx < 0) return { success: false, error: 'Invalid index' };

  const name = String(body.name || '').trim();
  const address = String(body.address || '').trim();
  const gstin = String(body.gstin || '').trim();
  if (!name) return { success: false, error: 'Vendor name is required' };

  const rowNum = idx + 2;
  if (rowNum > sheet.getLastRow()) return { success: false, error: 'Index out of range' };

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (i === idx + 1) continue;
    if (String(rows[i][0] || '').trim().toLowerCase() === name.toLowerCase()) {
      return { success: false, error: 'Another vendor already uses this name' };
    }
  }

  sheet.getRange(rowNum, 1, 1, 3).setValues([[name, address, gstin]]);
  return { success: true };
}

function bulkUpsertVendors(rows) {
  if (!Array.isArray(rows)) return { success: false, error: 'Invalid rows payload' };
  const sheet = getSheet_(SHEET_VENDORS);
  if (!sheet) return { success: false, error: 'Sheet not found' };

  const existing = readVendors();
  const indexByName = {};
  existing.forEach(function(v, i) {
    const key = String(v.name || '').trim().toLowerCase();
    if (key) indexByName[key] = i + 2;
  });

  let added = 0;
  let updated = 0;
  let skipped = 0;

  rows.forEach(function(r) {
    const name = String((r && r.name) || '').trim();
    if (!name) { skipped++; return; }
    const address = String((r && r.address) || '').trim();
    const gstin = String((r && r.gstin) || '').trim();
    const key = name.toLowerCase();

    if (indexByName[key]) {
      sheet.getRange(indexByName[key], 1, 1, 3).setValues([[name, address, gstin]]);
      updated++;
    } else {
      sheet.appendRow([name, address, gstin]);
      const newRow = sheet.getLastRow();
      indexByName[key] = newRow;
      added++;
    }
  });

  return { success: true, added: added, updated: updated, skipped: skipped };
}

// ============================================================
// ITEMS
// ============================================================
function readItems() {
  const sheet = getSheet_(SHEET_ITEMS);
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];

  const data = [];
  for (let i = 1; i < rows.length; i++) {
    data.push({
      desc: rows[i][0],
      uom: rows[i][1],
      rate: rows[i][2],
      default_qty: rows[i][3]
    });
  }
  return data;
}

function addItem(body) {
  const sheet = getSheet_(SHEET_ITEMS);
  if (!sheet) return { success: false, error: 'Sheet not found' };

  const desc = (body.desc || '').trim();
  const uom = (body.uom || '').trim();
  const rate = parseFloat(body.rate) || 0;
  const default_qty = parseInt(body.default_qty) || 1;

  if (!desc || !uom) return { success: false, error: 'Description and UOM required' };

  sheet.appendRow([desc, uom, rate, default_qty]);
  return { success: true };
}

function deleteItem(body) {
  const sheet = getSheet_(SHEET_ITEMS);
  if (!sheet) return { success: false, error: 'Sheet not found' };

  const idx = parseInt(body.idx);
  if (isNaN(idx) || idx < 0) return { success: false, error: 'Invalid index' };

  const rowToDelete = idx + 2; // +1 for header, +1 for 1-indexed
  if (rowToDelete > sheet.getLastRow()) return { success: false, error: 'Index out of range' };

  sheet.deleteRow(rowToDelete);
  return { success: true };
}

function updateItem(body) {
  const sheet = getSheet_(SHEET_ITEMS);
  if (!sheet) return { success: false, error: 'Sheet not found' };

  const idx = parseInt(body.idx, 10);
  if (isNaN(idx) || idx < 0) return { success: false, error: 'Invalid index' };

  const desc = String(body.desc || '').trim();
  const uom = String(body.uom || '').trim();
  const rate = parseFloat(body.rate) || 0;
  const default_qty = parseInt(body.default_qty, 10) || 1;
  if (!desc || !uom) return { success: false, error: 'Description and UOM required' };

  const rowNum = idx + 2;
  if (rowNum > sheet.getLastRow()) return { success: false, error: 'Index out of range' };

  sheet.getRange(rowNum, 1, 1, 4).setValues([[desc, uom, rate, default_qty]]);
  return { success: true };
}

function bulkUpsertItems(rows) {
  if (!Array.isArray(rows)) return { success: false, error: 'Invalid rows payload' };
  const sheet = getSheet_(SHEET_ITEMS);
  if (!sheet) return { success: false, error: 'Sheet not found' };

  const existing = readItems();
  const indexByKey = {};
  existing.forEach(function(it, i) {
    const key = (String(it.desc || '').trim() + '|' + String(it.uom || '').trim()).toLowerCase();
    if (key !== '|') indexByKey[key] = i + 2;
  });

  let added = 0;
  let updated = 0;
  let skipped = 0;

  rows.forEach(function(r) {
    const desc = String((r && r.desc) || '').trim();
    const uom = String((r && r.uom) || '').trim() || 'NOS';
    if (!desc) { skipped++; return; }

    const rate = parseFloat((r && r.rate) || 0) || 0;
    const defaultQty = parseInt((r && r.default_qty), 10) || 1;
    const key = (desc + '|' + uom).toLowerCase();

    if (indexByKey[key]) {
      sheet.getRange(indexByKey[key], 1, 1, 4).setValues([[desc, uom, rate, defaultQty]]);
      updated++;
    } else {
      sheet.appendRow([desc, uom, rate, defaultQty]);
      const newRow = sheet.getLastRow();
      indexByKey[key] = newRow;
      added++;
    }
  });

  return { success: true, added: added, updated: updated, skipped: skipped };
}

function calcGrandTotal_(items, transport, cgstPct, sgstPct) {
  const itemTotal = (items || []).reduce(function(sum, item) {
    return sum + (parseFloat(item.total) || 0);
  }, 0);
  const subtotal = itemTotal + (parseFloat(transport) || 0);
  const cgstAmt = Math.round(subtotal * (parseFloat(cgstPct) || 0) / 100);
  const sgstAmt = Math.round(subtotal * (parseFloat(sgstPct) || 0) / 100);
  return subtotal + cgstAmt + sgstAmt;
}

function bulkUpsertPOs(rows) {
  if (!Array.isArray(rows)) return { success: false, error: 'Invalid rows payload' };
  const sheet = getSheet_(SHEET_PO);
  if (!sheet) return { success: false, error: 'Sheet not found' };

  const existingRows = sheet.getDataRange().getValues();
  if (existingRows.length === 0) return { success: false, error: 'PO sheet is not initialized' };

  const headers = existingRows[0];
  const poNoCol = headers.indexOf('po_no');
  const poIdCol = headers.indexOf('po_id');
  const createdAtCol = headers.indexOf('created_at');

  const rowByPoNo = {};
  for (let i = 1; i < existingRows.length; i++) {
    const poNo = String(existingRows[i][poNoCol] || '').trim();
    if (poNo) rowByPoNo[poNo.toLowerCase()] = i + 1;
  }

  let added = 0;
  let updated = 0;
  let skipped = 0;

  rows.forEach(function(r) {
    const poNo = String((r && r.po_no) || '').trim();
    const poDate = String((r && r.po_date) || '').trim();
    const vendorName = String((r && r.vendor_name) || '').trim();
    if (!poNo || !poDate || !vendorName) { skipped++; return; }

    const items = Array.isArray(r.items) ? r.items : [];
    const transport = parseFloat((r && r.transport) || 0) || 0;
    const cgstPct = parseFloat((r && r.cgst_percent) || 0) || 0;
    const sgstPct = parseFloat((r && r.sgst_percent) || 0) || 0;
    const total = (r && r.total !== undefined && r.total !== null && r.total !== '')
      ? (parseFloat(r.total) || 0)
      : calcGrandTotal_(items, transport, cgstPct, sgstPct);

    const key = poNo.toLowerCase();
    const targetRow = rowByPoNo[key];

    let poId = Utilities.getUuid();
    let createdAt = new Date();
    if (targetRow) {
      poId = existingRows[targetRow - 1][poIdCol] || poId;
      createdAt = existingRows[targetRow - 1][createdAtCol] || createdAt;
    }

    const out = [
      poId,
      poNo,
      poDate,
      vendorName,
      String((r && r.vendor_address) || '').trim(),
      String((r && r.vendor_gstin) || '').trim(),
      String((r && r.client_name) || '').trim(),
      String((r && r.event_name) || '').trim(),
      String((r && r.event_location) || '').trim(),
      String((r && r.event_date) || '').trim(),
      JSON.stringify(items),
      transport,
      cgstPct,
      sgstPct,
      String((r && r.terms) || '').trim(),
      total,
      createdAt
    ];

    if (targetRow) {
      sheet.getRange(targetRow, 1, 1, out.length).setValues([out]);
      updated++;
    } else {
      sheet.appendRow(out);
      const newRow = sheet.getLastRow();
      rowByPoNo[key] = newRow;
      added++;
    }
  });

  return { success: true, added: added, updated: updated, skipped: skipped };
}

// ============================================================
// USERS
// ============================================================
function getUsersSheet() {
  const ss = getSS_();
  let sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_USERS);
    sheet.appendRow(['username', 'password', 'role']);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    sheet.appendRow(['ranga', 'Ranga@7677', 'admin']); // default admin
  }
  return sheet;
}

function readUsers() {
  const sheet = getUsersSheet();
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];

  const data = [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) {
      data.push({
        username: String(rows[i][0]).trim(),
        password: String(rows[i][1]).trim(),
        role: String(rows[i][2]).trim() || 'viewer'
      });
    }
  }
  return data;
}

function saveUsers(users) {
  const sheet = getUsersSheet();
  // Clear everything except header
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).clear();
  }
  // Batch write all users at once (much faster than row-by-row)
  if (users.length > 0) {
    const rows = users.map(u => [u.username, u.password, u.role || 'viewer']);
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }
  return { success: true };
}

// ============================================================
// SETTINGS  (key-value store)
// ============================================================
function getSettingsSheet_() {
  const ss = getSS_();
  let sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SETTINGS);
    sheet.appendRow(['key', 'value']);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  }
  return sheet;
}

function readSettings() {
  const sheet = getSettingsSheet_();
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return {};

  const obj = {};
  for (let i = 1; i < rows.length; i++) {
    const key = String(rows[i][0] || '').trim();
    if (key) obj[key] = rows[i][1] || '';
  }
  return obj;
}

function saveSetting(key, value) {
  if (!key) return { success: false, error: 'Key is required' };
  const sheet = getSettingsSheet_();
  const rows = sheet.getDataRange().getValues();

  // Find existing key
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === key) {
      sheet.getRange(i + 1, 2).setValue(value !== undefined ? value : '');
      return { success: true };
    }
  }
  // Key not found — append
  sheet.appendRow([key, value !== undefined ? value : '']);
  return { success: true };
}

function saveSettingsBulk(settings) {
  if (!settings || typeof settings !== 'object') return { success: false, error: 'Invalid settings object' };
  const keys = Object.keys(settings);
  if (!keys.length) return { success: true };

  const sheet = getSettingsSheet_();
  const rows = sheet.getDataRange().getValues();

  // Build a map of existing key → row index (1-based)
  const keyRowMap = {};
  for (let i = 1; i < rows.length; i++) {
    keyRowMap[String(rows[i][0]).trim()] = i + 1;
  }

  keys.forEach(function(key) {
    const val = settings[key] !== undefined ? settings[key] : '';
    if (keyRowMap[key]) {
      sheet.getRange(keyRowMap[key], 2).setValue(val);
    } else {
      sheet.appendRow([key, val]);
    }
  });

  return { success: true };
}
