/**
 * KAS — Secure Cloud Sync (Google Apps Script) - FIXED VERSION
 * 1) Auto-healing PIN & Spreadsheet binding
 * 2) Safe Merge by id + updatedAt
 * 3) Separate Tabs in the Same Sheet:
 *    - Bill_of_Supply (Sales)
 *    - Purchase_Bill (Purchase)
 *    - Malparat_Return (Return)
 *    - Cash_Receipt (Payments In/Out)
 *    - Parties, Items, Expenses
 */
const PROP_KEY_DATA = 'KAS_CLOUD_STATE_JSON_V2';
const PROP_KEY_OWNER_PIN = 'KAS_OWNER_PIN_HASH';
const PROP_KEY_STAFF_PIN = 'KAS_STAFF_PIN_HASH';
const PROP_KEY_SHEET_ID = 'KAS_BOUND_SHEET_ID';

// જો તમારી સ્ક્રિપ્ટ અલગ (Standalone) હોય, તો નીચે ડબલ કોટ્સમાં તમારી Google Sheet ની ID નાખો (નહીંતર ખાલી રહેવા દો)
const MANUAL_SPREADSHEET_ID = ''; 

function getTargetSpreadsheet() {
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (ss) {
    try { PropertiesService.getScriptProperties().setProperty(PROP_KEY_SHEET_ID, ss.getId()); } catch (e) {}
    return ss;
  }
  const savedId = MANUAL_SPREADSHEET_ID || PropertiesService.getScriptProperties().getProperty(PROP_KEY_SHEET_ID);
  if (savedId) {
    try { return SpreadsheetApp.openById(savedId); } catch (e) {}
  }
  throw new Error('Google Sheet મળી નથી! કૃપા કરીને Google Sheet માંથી Extensions > Apps Script ખોલો અથવા MANUAL_SPREADSHEET_ID નાખો.');
}

function isAuthorized(pinHash, allowAutoReset) {
  const props = PropertiesService.getScriptProperties();
  const owner = props.getProperty(PROP_KEY_OWNER_PIN) || '';
  const staff = props.getProperty(PROP_KEY_STAFF_PIN) || '';
  
  if (!owner && pinHash) {
    props.setProperty(PROP_KEY_OWNER_PIN, String(pinHash));
    return true;
  }
  if (!pinHash) return true;
  if (String(pinHash) === owner || (staff && String(pinHash) === staff)) {
    return true;
  }
  if (allowAutoReset && pinHash) {
    props.setProperty(PROP_KEY_OWNER_PIN, String(pinHash));
    return true;
  }
  return false;
}

// PIN રિસેટ અને શીટ કનેક્ટ કરવા માટેનું ફંક્શન (Apps Script માંથી Run કરો)
function resetAllPins() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(PROP_KEY_OWNER_PIN);
  props.deleteProperty(PROP_KEY_STAFF_PIN);
  const ss = getTargetSpreadsheet();
  if (ss) {
    const raw = props.getProperty(PROP_KEY_DATA);
    const existing = raw ? JSON.parse(raw) : { parties: [], items: [], bills: [], payments: [], expenses: [] };
    syncReadableSheets(existing);
  }
}

function mergeArraysById(existingArr, incomingArr) {
  const map = {};
  (existingArr || []).forEach(function(item) {
    if (item && item.id != null) map[String(item.id)] = item;
  });
  (incomingArr || []).forEach(function(item) {
    if (!item || item.id == null) return;
    const k = String(item.id);
    const prev = map[k];
    if (!prev || Number(item.updatedAt || 0) >= Number(prev.updatedAt || 0)) {
      map[k] = item;
    }
  });
  return Object.keys(map).map(function(k) { return map[k]; });
}

function doGet(e) {
  try {
    const pin = (e && e.parameter && e.parameter.pin) ? e.parameter.pin : '';
    isAuthorized(pin, true);
    const raw = PropertiesService.getScriptProperties().getProperty(PROP_KEY_DATA);
    const data = raw ? JSON.parse(raw) : { parties: [], items: [], bills: [], payments: [], expenses: [] };
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const contents = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    const body = JSON.parse(contents);
    
    isAuthorized(body.pinHash, true);

    const props = PropertiesService.getScriptProperties();
    const raw = props.getProperty(PROP_KEY_DATA);
    const existing = raw ? JSON.parse(raw) : { parties: [], items: [], bills: [], payments: [], expenses: [] };

    const merged = {
      parties: mergeArraysById(existing.parties, body.parties),
      items: mergeArraysById(existing.items, body.items),
      bills: mergeArraysById(existing.bills, body.bills),
      payments: mergeArraysById(existing.payments, body.payments),
      expenses: mergeArraysById(existing.expenses, body.expenses)
    };
    props.setProperty(PROP_KEY_DATA, JSON.stringify(merged));
    syncReadableSheets(merged);

    return ContentService.createTextOutput(JSON.stringify({ ok: true, savedAt: new Date().toISOString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function itemsSummary(rows) {
  return (rows || []).map(function(r) {
    return (r.desc || '') + ' (' + (r.kg || 0) + ' Kg @ ' + (r.rate || 0) + ')';
  }).join(', ');
}

function syncReadableSheets(data) {
  const ss = getTargetSpreadsheet();
  if (!ss) return;

  const allBills = data.bills || [];
  const salesBills = allBills.filter(function(b) { return (b.kind || 'Sales') === 'Sales'; });
  const purchaseBills = allBills.filter(function(b) { return b.kind === 'Purchase'; });
  const returnBills = allBills.filter(function(b) { return b.kind === 'Return'; });

  writeTab(ss, 'Bill_of_Supply',
    ['ID', 'Bill No', 'Date', 'Buyer', 'Phone', 'Payment', 'Collection Date', 'Total (Rs)', 'Items', 'Deleted', 'UpdatedAt'],
    salesBills.map(function(b) {
      return [b.id, b.no || '', b.date || '', b.buyer || '', b.phone || '', b.payment || '', b.collectionDate || '', Number(b.total || 0), itemsSummary(b.rows), !!b.deleted, b.updatedAt || 0];
    })
  );

  writeTab(ss, 'Purchase_Bill',
    ['ID', 'Bill No', 'Date', 'Party / Supplier', 'Phone', 'Payment', 'Total (Rs)', 'Items', 'Deleted', 'UpdatedAt'],
    purchaseBills.map(function(b) {
      return [b.id, b.no || '', b.date || '', b.buyer || '', b.phone || '', b.payment || '', Number(b.total || 0), itemsSummary(b.rows), !!b.deleted, b.updatedAt || 0];
    })
  );

  writeTab(ss, 'Malparat_Return',
    ['ID', 'Return No', 'Orig Bill No', 'Date', 'Buyer', 'Phone', 'Reason', 'Total (Rs)', 'Items', 'Deleted', 'UpdatedAt'],
    returnBills.map(function(b) {
      return [b.id, b.no || '', b.origBillNo || '', b.date || '', b.buyer || '', b.phone || '', b.returnReason || '', Number(b.total || 0), itemsSummary(b.rows), !!b.deleted, b.updatedAt || 0];
    })
  );

  writeTab(ss, 'Cash_Receipt',
    ['ID', 'Receipt No', 'Type (In/Out)', 'Date', 'Party', 'Mode', 'Note', 'Amount (Rs)', 'Deleted', 'UpdatedAt'],
    (data.payments || []).map(function(p) {
      return [p.id, p.no || '', p.type || 'In', p.date || '', p.party || '', p.mode || '', p.note || '', Number(p.amount || 0), !!p.deleted, p.updatedAt || 0];
    })
  );

  writeTab(ss, 'Parties',
    ['ID', 'Name', 'Phone', 'City', 'Due (Rs)', 'Due Date', 'Deleted', 'UpdatedAt'],
    (data.parties || []).map(function(p) {
      return [p.id, p.name || '', p.phone || '', p.city || '', Number(p.due || 0), p.dueDate || '', !!p.deleted, p.updatedAt || 0];
    })
  );

  writeTab(ss, 'Items',
    ['ID', 'Name', 'Pack', 'Rate', 'Wholesale Rate', 'Stock', 'Batch', 'Expiry', 'Deleted', 'UpdatedAt'],
    (data.items || []).map(function(it) {
      return [it.id, it.name || '', it.pack || '', Number(it.rate || 0), Number(it.wholesaleRate || 0), Number(it.stock || 0), it.batch || '', it.expiry || '', !!it.deleted, it.updatedAt || 0];
    })
  );

  writeTab(ss, 'Expenses',
    ['ID', 'Date', 'Category / Title', 'Note', 'Amount (Rs)', 'Deleted', 'UpdatedAt'],
    (data.expenses || []).map(function(ex) {
      return [ex.id, ex.date || '', ex.category || ex.title || '', ex.note || '', Number(ex.amount || 0), !!ex.deleted, ex.updatedAt || 0];
    })
  );

  SpreadsheetApp.flush();
}

function writeTab(ss, name, headers, rows) {
  let sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clearContents();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
}
