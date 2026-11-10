/* =====================================================
   InventoryPro – Code.gs (Public Access & Fully Corrected)
   Sheets: Item_Master, Vendor_List, Stock_Transactions,
           Inventory_Balance, Purchase_Orders, Stock_Adjustments,
           Item_Daily_Summary, Vendor and Material List
===================================================== */

/* ---------- 1. ROLES & ACCESS CONTROL ---------- */
// Email lookup map (kept for reference, but access restrictions are removed)
const USER_ROLES = {
  "ea@bhtindia.com": "ADMIN",
  "mail@bhtindia.com": "ADMIN",
  "enquiry@bhtindia.com": "ADMIN",
  "anumittal81@gmail.com": "ADMIN",
  "a91884424@gmail.com": "ADMIN",
  "production@bhtindia.com": "ADMIN",
  "store@bhtindia.com": "ADMIN",
  "pc1@bhtindia.com": "ADMIN"
};

/**
 * Returns ADMIN by default so any user accessing the app has full capabilities.
 */
function getRole() {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  return USER_ROLES[email] || "ADMIN";
}

/**
 * Authorization check disabled to allow open access.
 */
function checkPermission(roles) {
  return true;
}

/* ---------- 2. SMALL HELPERS ---------- */
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function str_(v) { return String(v == null ? '' : v).trim(); }
function ms_(v) { return v instanceof Date ? v.getTime() : new Date(v).getTime(); }
function dayRange_() {
  const n = new Date();
  const s = new Date(n.getFullYear(), n.getMonth(), n.getDate()).getTime();
  return [s, s + 86400000];
}

function sheetToObjects_(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];
  const headers = data[0].map(h => String(h).trim());
  return data.slice(1).filter(r => r[0]).map(row => {
    const o = {};
    headers.forEach((h, i) => o[h] = row[i]);
    o.__row = row;
    return o;
  });
}

function txnObj_(r, rowIdx, d, tz, type, qty) {
  return {
    rowIdx: rowIdx,
    date: Utilities.formatDate(d, tz, 'yyyy-MM-dd HH:mm:ss'),
    item: r[1],
    name: r[2],
    type: type,
    qty: qty,
    method: r[5],
    remarks: r[6],
    user: r[7],
    rack: r[8],
    vendor: r[11],
    invoice: r[12],
    voucher: r[13],
    party: r[14],
    reference: r[15]
  };
}

/* ---------- 3. CACHE LAYER ---------- */
const CACHE_TTL = 120;
const CHUNK = 30000, CHUNKS = 8;
const CK = {
  MASTER: 'C_MASTER', DYN: 'C_DYN', DASH_OV: 'C_DASH_OV', POS: 'C_POS',
  ITEMS: 'C_ITEM_MASTER', VENDORS: 'C_VENDOR_LIST', BALANCE: 'C_BALANCE', DAILY: 'C_DAILY'
};
const DYN_KEYS = [CK.DYN, CK.DASH_OV, CK.POS, CK.BALANCE, CK.DAILY];
const MASTER_KEYS = [CK.MASTER, CK.ITEMS, CK.VENDORS];

function cacheGetJson_(key) {
  try {
    const c = CacheService.getScriptCache();
    const v = c.get(key);
    if (!v) return null;
    if (v.charAt(0) !== '#') return JSON.parse(v);
    const n = Number(v.slice(1)), names = [];
    for (let i = 0; i < n; i++) names.push(key + '_' + i);
    const parts = c.getAll(names);
    let s = '';
    for (let i = 0; i < n; i++) {
      if (parts[names[i]] == null) return null;
      s += parts[names[i]];
    }
    return JSON.parse(s);
  } catch (e) { return null; }
}

function cachePutJson_(key, value, ttl) {
  try {
    const s = JSON.stringify(value), c = CacheService.getScriptCache();
    if (s.length <= CHUNK) { c.put(key, s, ttl || CACHE_TTL); return; }
    const n = Math.ceil(s.length / CHUNK);
    if (n > CHUNKS) return;
    const m = {};
    for (let i = 0; i < n; i++) m[key + '_' + i] = s.substr(i * CHUNK, CHUNK);
    m[key] = '#' + n;
    c.putAll(m, ttl || CACHE_TTL);
  } catch (e) {}
}

function clearCaches_(keys) {
  try {
    const names = [];
    keys.forEach(k => {
      names.push(k);
      for (let i = 0; i < CHUNKS; i++) names.push(k + '_' + i);
    });
    CacheService.getScriptCache().removeAll(names);
  } catch (e) {}
}

/* ---------- 4. WEB APP & SPREADSHEET ROUTERS ---------- */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('InventoryPro')
    .addItem('Rebuild Daily Summary', 'rebuildDailySummary')
    .addItem('Rebuild Inventory Balance', 'rebuildInventoryBalance')
    .addToUi();
}

/**
 * Serves dashboard to any visitor without email restrictions.
 */
function doGet(e) {
  const email = (Session.getActiveUser().getEmail() || (e && e.parameter && e.parameter.email) || "").toLowerCase();
  const role = USER_ROLES[email] || "ADMIN";

  const t = HtmlService.createTemplateFromFile('Dashboard');
  t.userRole = role;
  return t.evaluate()
    .setTitle("InventoryPro Dashboard")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/* ---------- 5. BOOTSTRAP & READS ---------- */
function getMasterData_() {
  const hit = cacheGetJson_(CK.MASTER);
  if (hit) return hit;
  const ss = ss_(), items = [], vendors = [];
  const itemSh = ss.getSheetByName('Item_Master');
  if (itemSh && itemSh.getLastRow() > 1) {
    itemSh.getRange(2, 1, itemSh.getLastRow() - 1, 7).getValues().forEach(r => {
      const code = str_(r[0]), name = str_(r[1]);
      if (!code || !name) return;
      items.push({ code: code, name: name, category: str_(r[2]), subCategory: str_(r[3]), line: str_(r[4]), rack: str_(r[5]), box: str_(r[6]) });
    });
  }
  const venSh = ss.getSheetByName('Vendor_List');
  if (venSh && venSh.getLastRow() > 1) {
    venSh.getRange(2, 1, venSh.getLastRow() - 1, 6).getValues().forEach(r => {
      const code = str_(r[0]), name = str_(r[1]);
      if (!code || !name) return;
      vendors.push({ code: code, name: name, personName: str_(r[2]), address: str_(r[3]), email: str_(r[4]), phone: str_(r[5]) });
    });
  }
  const out = { items: items, vendors: vendors };
  cachePutJson_(CK.MASTER, out, 600);
  return out;
}

function readPOData_() {
  const hit = cacheGetJson_(CK.POS);
  if (hit) return hit;
  const sh = ss_().getSheetByName('Purchase_Orders');
  const rows = [], notif = { openPOs: 0, todayCreated: 0, todayExpected: 0 };
  if (sh && sh.getLastRow() > 1) {
    const tz = Session.getScriptTimeZone(), day = dayRange_(), seen = {};
    sh.getRange(2, 1, sh.getLastRow() - 1, 12).getValues().forEach((r, k) => {
      const po = str_(r[0]);
      if (!po) return;
      const status = str_(r[6]).toUpperCase();
      rows.push({
        rowIdx: k + 2,
        poNumber: po,
        date: r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd HH:mm:ss') : str_(r[1]),
        vendor: str_(r[2]),
        itemCode: str_(r[3]),
        orderedQty: Number(r[4]) || 0,
        receivedQty: Number(r[5]) || 0,
        status: status,
        createdBy: str_(r[7]),
        remarks: str_(r[8]),
        rate: Number(r[9]) || 0,
        deliveryDate: r[10] instanceof Date ? Utilities.formatDate(r[10], tz, 'yyyy-MM-dd') : str_(r[10]),
        terms: str_(r[11])
      });
      if (!seen[po]) {
        seen[po] = 1;
        const dMs = ms_(r[1]), vMs = ms_(r[10]);
        if (status === 'OPEN' || status === 'PARTIAL') notif.openPOs++;
        if (dMs >= day[0] && dMs < day[1]) notif.todayCreated++;
        if (vMs >= day[0] && vMs < day[1]) notif.todayExpected++;
      }
    });
  }
  const out = { rows: rows, notif: notif };
  cachePutJson_(CK.POS, out);
  return out;
}

function getDynamicData_() {
  const hit = cacheGetJson_(CK.DYN);
  if (hit) return hit;
  const ss = ss_(), tz = Session.getScriptTimeZone(), day = dayRange_();
  const po = readPOData_();
  const kpi = { currentStockQty: 0, todayIn: 0, todayOut: 0, lowStockCount: 0 };

  const balSh = ss.getSheetByName('Inventory_Balance');
  if (balSh && balSh.getLastRow() > 1) {
    balSh.getRange(2, 5, balSh.getLastRow() - 1, 1).getValues().forEach(r => {
      const q = Number(r[0]) || 0;
      kpi.currentStockQty += q;
      if (q < 10) kpi.lowStockCount++;
    });
  }

  const recent = [];
  const tSh = ss.getSheetByName('Stock_Transactions');
  if (tSh && tSh.getLastRow() > 1) {
    const total = tSh.getLastRow() - 1, n = Math.min(total, 150), start = total - n + 2;
    const data = tSh.getRange(start, 1, n, 16).getValues();
    for (let i = data.length - 1; i >= 0; i--) {
      const r = data[i];
      if (!r[0]) continue;
      const d = r[0] instanceof Date ? r[0] : new Date(r[0]);
      if (isNaN(d.getTime())) continue;
      const type = str_(r[3]).toUpperCase(), qty = Number(r[4]) || 0, ts = d.getTime();
      if (ts >= day[0] && ts < day[1]) {
        if (type === 'IN') kpi.todayIn += qty;
        if (type === 'OUT') kpi.todayOut += qty;
      }
      if (recent.length < 50) recent.push(txnObj_(r, start + i, d, tz, type, qty));
    }
  }

  const out = {
    openPOs: po.rows.filter(p => p.status === 'OPEN' || p.status === 'PARTIAL'),
    notifCounts: po.notif,
    kpi: kpi,
    recentTransactions: recent
  };
  cachePutJson_(CK.DYN, out);
  return out;
}

function getAppBootstrapData() {
  const email = Session.getActiveUser().getEmail() || "Guest User";
  return Object.assign({ userRole: getRole(), userEmail: email }, getMasterData_(), getDynamicData_());
}

function getDynamicData() {
  return Object.assign({ userRole: getRole() }, getDynamicData_());
}

/* ---------- 6. WRITE OPERATIONS ---------- */
function processStock(payload) {
  checkPermission(["ADMIN", "STORE"]);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const batch = Array.isArray(payload) ? payload : [payload];
    if (!batch.length) throw new Error("No items to process");

    const ss = ss_();
    const trans = ss.getSheetByName('Stock_Transactions');
    const itemSh = ss.getSheetByName('Item_Master');
    const poSh = ss.getSheetByName('Purchase_Orders');
    const balSh = ss.getSheetByName('Inventory_Balance');
    if (!trans || !itemSh || !balSh) throw new Error("Required sheet missing");
    if (itemSh.getLastRow() < 2) throw new Error("Item_Master database is empty.");

    const now = new Date(), nowMs = now.getTime();
    const email = Session.getActiveUser().getEmail() || (batch[0] && batch[0].user) || "WebUser";
    const userLc = email.toLowerCase();
    const isPO = !!(batch[0] && batch[0].poNumber);
    const poNum = isPO ? str_(batch[0].poNumber) : '';
    const poAction = str_(batch[0].poAction || (batch[0].createBackorder ? 'BACKORDER' : 'KEEP_OPEN')).toUpperCase();

    // 1. Guard against in-batch duplicate items for direct stock entries
    if (!isPO) {
      const seenBatch = {};
      for (const it of batch) {
        const c = str_(it.itemCode);
        if (seenBatch[c]) throw new Error("Duplicate item " + c + " detected in submission. Please combine quantities into one row.");
        seenBatch[c] = true;
      }
    }

    // 2. Guard against rapid accidental double-clicks (within 30 seconds)
    const tLast = trans.getLastRow();
    if (tLast > 1) {
      const n = Math.min(tLast - 1, 50);
      const recent = trans.getRange(tLast - n + 1, 1, n, Math.max(17, trans.getLastColumn())).getValues();
      for (const it of batch) {
        const code = str_(it.itemCode), qty = Number(it.qty) || 0, type = str_(it.type).toUpperCase();
        if (qty <= 0) continue; // 0-qty lines are PO status updates, not physical movements

        const inv = str_(it.invoiceNo).toUpperCase();
        const party = str_(it.partyDetails).toUpperCase();

        const dup = recent.some(r => {
          const diff = nowMs - ms_(r[0]);
          if (isNaN(diff) || diff < 0 || diff > 30000) return false;
          if (str_(r[1]) !== code) return false;

          // For PO receipts, match PO number and qty
          if (isPO) {
            const rPo16 = str_(r[15]), rPo17 = r.length > 16 ? str_(r[16]) : '';
            if ((rPo16 === poNum || rPo17 === poNum) && (Number(r[4]) || 0) === qty) return true;
          }

          const matchBase = str_(r[3]).toUpperCase() === type &&
                            (Number(r[4]) || 0) === qty &&
                            str_(r[7]).toLowerCase() === userLc;
          if (!matchBase) return false;

          const rInv = str_(r[12]).toUpperCase();
          const rParty = str_(r[14]).toUpperCase();
          if (inv && rInv && inv === rInv) return true;
          if (party && rParty && party === rParty) return true;
          if (!inv && !party) return true;
          return false;
        });

        if (dup) {
          return {
            ok: false,
            message: "Duplicate entry blocked: An identical transaction for " + code + " (" + qty + " units) was submitted within the last 30 seconds."
          };
        }
      }
    }

    const itemMap = {};
    itemSh.getRange(2, 1, itemSh.getLastRow() - 1, 11).getValues().forEach(r => { if (r[0]) itemMap[str_(r[0])] = r; });

    const balRows = balSh.getLastRow() > 1 ? balSh.getRange(2, 1, balSh.getLastRow() - 1, 8).getValues() : [];
    const balIdx = {};
    balRows.forEach((r, i) => { if (r[0]) balIdx[str_(r[0])] = i; });

    const poData = (poSh && poSh.getLastRow() > 1) ? poSh.getRange(1, 1, poSh.getLastRow(), 12).getValues() : [];
    const rowsToAppend = [];
    const poLineUpdates = {};
    const backorderLines = [];

    for (const d of batch) {
      const code = str_(d.itemCode), qty = Number(d.qty) || 0;
      if (!code) throw new Error("An item code is missing in the request.");
      if (qty < 0) throw new Error("Invalid negative quantity for item " + code);

      const item = itemMap[code];
      if (!item) throw new Error("Item not found in master list: " + code);

      // Only update balance and append transaction if physical qty > 0
      if (qty > 0) {
        let bi = balIdx[code];
        if (bi === undefined) {
          const op = (item[10] !== '' && !isNaN(item[10])) ? Number(item[10]) : 0;
          balRows.push([code, op, 0, 0, op, now, 0, '']);
          bi = balIdx[code] = balRows.length - 1;
        }
        const b = balRows[bi];
        if (d.type === "OUT") {
          const avail = Number(b[4]) || 0;
          if (avail < qty) throw new Error("Insufficient stock for " + code + ". Available: " + avail + " Required: " + qty);
          b[3] = (Number(b[3]) || 0) + qty;
          b[4] = avail - qty;
        } else {
          b[2] = (Number(b[2]) || 0) + qty;
          b[4] = (Number(b[4]) || 0) + qty;
        }
        b[5] = now;

        rowsToAppend.push([
          now, code, item[1], d.type, qty, d.method || '', d.remarks || '-', email,
          item[4] || '', item[5] || '', item[6] || '',
          d.vendorName || '', d.invoiceNo || '', d.voucherNo || '', d.partyDetails || '',
          d.referenceNo || d.poNumber || '', d.poNumber || ''
        ]);
      }

      // Process PO line status and shortfall tracking (including lines with qty = 0)
      if (isPO && poSh && d.type === 'IN') {
        let idx = -1;
        if (d.poRowIdx && d.poRowIdx > 1 && d.poRowIdx <= poData.length) {
          const r = poData[d.poRowIdx - 1];
          if (r[0] === poNum && str_(r[3]) === code) idx = d.poRowIdx - 1;
        }
        if (idx === -1) {
          for (let i = 1; i < poData.length; i++) {
            if (poData[i][0] === poNum && str_(poData[i][3]) === code) { idx = i; break; }
          }
        }
        if (idx > -1) {
          const r = poData[idx];
          const ordered = Number(r[4]) || 0;
          const prevGot = Number(r[5]) || 0;
          const maxAllowed = Math.max(0, ordered - prevGot);

          if (qty > maxAllowed) {
            throw new Error("Cannot receive " + qty + " for " + code + ". Only " + maxAllowed + " remaining on PO " + poNum + ".");
          }

          const totalReceived = prevGot + qty;
          const remaining = Math.max(0, ordered - totalReceived);
          const lineSheetRow = idx + 1;

          let newStatus = "OPEN";
          if (remaining === 0) {
            newStatus = "FULFILLED";
          } else if (poAction === 'BACKORDER' || poAction === 'CLOSE_SHORT') {
            newStatus = "CLOSED-SHORT";
            if (poAction === 'BACKORDER' && remaining > 0) {
              backorderLines.push({
                vendor: r[2],
                itemCode: code,
                orderedQty: remaining,
                rate: r[9] || 0,
                deliveryDate: r[10] || '',
                terms: r[11] || ''
              });
            }
          } else {
            newStatus = totalReceived > 0 ? "PARTIAL" : "OPEN";
          }

          poLineUpdates[lineSheetRow] = {
            row: lineSheetRow,
            received: totalReceived,
            status: newStatus
          };
        }
      }
    }

    // 1. Write stock transactions
    if (rowsToAppend.length) {
      trans.getRange(trans.getLastRow() + 1, 1, rowsToAppend.length, 17).setValues(rowsToAppend);
    }

    // 2. Update original PO lines in sheet
    Object.keys(poLineUpdates).forEach(k => {
      const u = poLineUpdates[k];
      poSh.getRange(Number(u.row), 6, 1, 2).setValues([[u.received, u.status]]);
    });

    // 3. Create consolidated Backorder PO containing all remaining items
    if (poAction === 'BACKORDER' && backorderLines.length > 0) {
      const boNum = poNum.endsWith('-BO') ? poNum + '2' : poNum + '-BO';
      const boRows = backorderLines.map(bo => [
        boNum,
        now,
        bo.vendor,
        bo.itemCode,
        bo.orderedQty,
        0,
        "OPEN",
        email,
        "Auto Backorder from " + poNum,
        bo.rate,
        bo.deliveryDate,
        bo.terms
      ]);
      poSh.getRange(poSh.getLastRow() + 1, 1, boRows.length, 12).setValues(boRows);
    }

    // 4. Update balances & daily summary
    if (balRows.length) balSh.getRange(2, 1, balRows.length, 8).setValues(balRows);
    rebuildDailySummary();
    clearCaches_(DYN_KEYS);

    return {
      ok: true,
      message: "Successfully processed " + rowsToAppend.length + " item(s)" +
               (backorderLines.length ? " and created backorder " + (poNum.endsWith('-BO') ? poNum + '2' : poNum + '-BO') + " for " + backorderLines.length + " remaining item(s)." : "."),
      dyn: Object.assign({ userRole: getRole() }, getDynamicData_())
    };
  } finally {
    lock.releaseLock();
  }
}

function deleteMultipleTransactions(rowIndices) {
  checkPermission(["ADMIN"]);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = ss_(), transSh = ss.getSheetByName('Stock_Transactions'), poSh = ss.getSheetByName('Purchase_Orders');
    if (!transSh) throw new Error("Stock_Transactions sheet not found");
    const idx = (rowIndices || []).map(Number).filter(r => !isNaN(r) && r > 1);
    if (!idx.length) return "No valid rows selected.";

    const tData = transSh.getDataRange().getValues();
    const adj = {};
    idx.forEach(ri => {
      const r = tData[ri - 1];
      if (!r) return;
      const qty = Number(r[4]) || 0;
      let po = str_(r[15]);
      if (!po && r.length > 16) po = str_(r[16]);
      if (!po) { const m = String(r[6] || '').match(/BHT\/PO\/\d+/); if (m) po = m[0]; }
      if (str_(r[3]).toUpperCase() === 'IN' && po.indexOf('BHT/PO/') === 0 && qty > 0) {
        const k = po + "|" + str_(r[1]);
        adj[k] = (adj[k] || 0) + qty;
      }
    });

    if (poSh && poSh.getLastRow() > 1 && Object.keys(adj).length) {
      const poData = poSh.getDataRange().getValues();
      for (let i = 1; i < poData.length; i++) {
        const d = adj[str_(poData[i][0]) + "|" + str_(poData[i][3])];
        if (!d) continue;
        const ordered = Number(poData[i][4]) || 0, rec = Math.max(0, (Number(poData[i][5]) || 0) - d);
        poSh.getRange(i + 1, 6, 1, 2).setValues([[rec, rec >= ordered ? "FULFILLED" : (rec > 0 ? "PARTIAL" : "OPEN")]]);
      }
    }

    idx.sort((a, b) => b - a).forEach(ri => transSh.deleteRow(ri));
    rebuildInventoryBalance();
    rebuildDailySummary();
    clearCaches_(DYN_KEYS);
    return "Successfully deleted " + idx.length + " transaction(s) and synced related Purchase Orders.";
  } finally {
    lock.releaseLock();
  }
}

function deletePurchaseOrder(poNumber) {
  checkPermission(["ADMIN"]);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (!poNumber) throw new Error("PO Number is required.");
    const ss = ss_(), poSh = ss.getSheetByName('Purchase_Orders'), transSh = ss.getSheetByName('Stock_Transactions');
    if (!poSh) throw new Error("Purchase_Orders sheet not found.");
    const po = str_(poNumber), bo = po + '-BO';

    let deletedTxn = 0;
    if (transSh && transSh.getLastRow() > 1) {
      const t = transSh.getDataRange().getValues(), del = [];
      for (let i = 1; i < t.length; i++) {
        const ref = str_(t[i][15]), c17 = t[i].length > 16 ? str_(t[i][16]) : '';
        if (ref === po || ref === bo || c17 === po || c17 === bo || String(t[i][6] || '').includes(po)) del.push(i + 1);
      }
      del.sort((a, b) => b - a).forEach(r => transSh.deleteRow(r));
      deletedTxn = del.length;
    }

    const p = poSh.getDataRange().getValues(), rows = [];
    for (let i = 1; i < p.length; i++) { const c = str_(p[i][0]); if (c === po || c === bo) rows.push(i + 1); }
    if (!rows.length) throw new Error("Purchase Order not found.");
    rows.sort((a, b) => b - a).forEach(r => poSh.deleteRow(r));

    if (deletedTxn) {
      rebuildInventoryBalance();
      rebuildDailySummary();
    }
    clearCaches_(DYN_KEYS);
    return "Successfully deleted " + poNumber + " and reversed " + deletedTxn + " associated stock transaction(s).";
  } finally {
    lock.releaseLock();
  }
}

function addVendorData(data) {
  checkPermission(["ADMIN"]);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = ss_().getSheetByName('Vendor_List');
    if (!sheet) throw new Error("Vendor_List sheet not found");
    const lastRow = sheet.getLastRow(), now = new Date();
    const code = (data.name || "VEN").substring(0, 3).toUpperCase() + "-0" + (lastRow > 0 ? lastRow : 1) + "-" + (now.getMonth() + 1) + "-" + now.getFullYear();
    sheet.appendRow([code, data.name, data.personName, data.address, data.email, data.phone, Utilities.getUuid(), "Active"]);
    clearCaches_(MASTER_KEYS);
    return "Vendor successfully added!";
  } finally {
    lock.releaseLock();
  }
}

function createPurchaseOrder(payload) {
  checkPermission(["ADMIN", "PC", "STORE"]);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (!payload || !payload.vendor || !payload.items || !payload.items.length) throw new Error("Vendor and at least one item are required.");
    const poSh = ss_().getSheetByName('Purchase_Orders');
    if (!poSh) throw new Error("Purchase_Orders sheet not found.");

    const props = PropertiesService.getScriptProperties();
    let seq = parseInt(props.getProperty('BHT_PO_COUNTER')) || 0;
    if (poSh.getLastRow() > 1) {
      poSh.getRange(2, 1, poSh.getLastRow() - 1, 1).getValues().forEach(r => {
        const m = String(r[0]).trim().match(/^BHT\/PO\/(\d+)/);
        if (m && Number(m[1]) > seq) seq = Number(m[1]);
      });
    }
    seq += 1;
    props.setProperty('BHT_PO_COUNTER', String(seq));

    const poNumber = 'BHT/PO/' + String(seq).padStart(3, '0');
    const user = Session.getActiveUser().getEmail() || payload.user || "WebUser";
    const now = new Date();
    const rows = payload.items.map(i => [
      poNumber, now, payload.vendor, i.itemCode, Number(i.qty), 0, "OPEN", user,
      payload.remarks || "", Number(i.rate), payload.deliveryDate, payload.terms || ""
    ]);
    poSh.getRange(poSh.getLastRow() + 1, 1, rows.length, 12).setValues(rows);
    clearCaches_(DYN_KEYS);
    return poNumber;
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 7. READ FILTERS & VIEWS ---------- */
function getDashboardDataFiltered(f) {
  f = f || {};
  const hasFilter = f.from || f.to || f.item || f.vendor || f.invoice;
  if (!hasFilter) { const hit = cacheGetJson_(CK.DASH_OV); if (hit) return hit; }

  const tz = Session.getScriptTimeZone(), table = [];
  const tSh = ss_().getSheetByName('Stock_Transactions');
  if (tSh && tSh.getLastRow() > 1) {
    const total = tSh.getLastRow() - 1, n = hasFilter ? total : Math.min(total, 500), start = total - n + 2;
    const raw = tSh.getRange(start, 1, n, 16).getValues();
    const fromMs = f.from ? new Date(f.from).getTime() : 0;
    const toD = f.to ? new Date(f.to) : new Date(8640000000000000);
    toD.setHours(23, 59, 59);
    const toMs = toD.getTime();
    const qItem = (f.item || '').toLowerCase(), qVen = (f.vendor || '').toLowerCase(), qInv = (f.invoice || '').toLowerCase();

    for (let i = raw.length - 1; i >= 0; i--) {
      const r = raw[i];
      if (!r[0]) continue;
      const d = r[0] instanceof Date ? r[0] : new Date(r[0]);
      const ts = d.getTime();
      if (isNaN(ts) || ts < fromMs || ts > toMs) continue;
      if (qItem && str_(r[1]).toLowerCase().indexOf(qItem) === -1) continue;
      if (qVen && str_(r[11]).toLowerCase().indexOf(qVen) === -1) continue;
      if (qInv && str_(r[12]).toLowerCase().indexOf(qInv) === -1) continue;
      table.push(txnObj_(r, start + i, d, tz, str_(r[3]).toUpperCase(), Number(r[4]) || 0));
    }
  }
  const result = { table: table };
  if (!hasFilter) cachePutJson_(CK.DASH_OV, result);
  return result;
}

function getItemMasterData() {
  const hit = cacheGetJson_(CK.ITEMS);
  if (hit) return hit;
  const out = sheetToObjects_(ss_().getSheetByName('Item_Master'));
  out.forEach(o => delete o.__row);
  cachePutJson_(CK.ITEMS, out, 600);
  return out;
}

function getVendorListData() {
  const hit = cacheGetJson_(CK.VENDORS);
  if (hit) return hit;
  const out = sheetToObjects_(ss_().getSheetByName('Vendor_List'));
  out.forEach(o => delete o.__row);
  cachePutJson_(CK.VENDORS, out, 600);
  return out;
}

function getTodaySummaryData() {
  const hit = cacheGetJson_(CK.DAILY);
  if (hit) return hit;

  let sum = ss_().getSheetByName('Item_Daily_Summary');
  if (!sum || sum.getLastRow() <= 1) {
    rebuildDailySummary();
    sum = ss_().getSheetByName('Item_Daily_Summary');
  }

  const out = sheetToObjects_(sum).map(o => {
    const r = o.__row; delete o.__row;
    o.Mapped_Date = r[0]; o.Mapped_Code = r[1]; o.Mapped_Name = r[2];
    o.Mapped_Opening = r[3]; o.Mapped_In = r[4]; o.Mapped_Out = r[5]; o.Mapped_Closing = r[6];
    return o;
  });
  cachePutJson_(CK.DAILY, out, 120);
  return out;
}

function getAvailableBalanceData() {
  const hit = cacheGetJson_(CK.BALANCE);
  if (hit) return hit;
  const ss = ss_(), nameMap = {}, itemSh = ss.getSheetByName('Item_Master');
  if (itemSh && itemSh.getLastRow() > 1) {
    itemSh.getRange(2, 1, itemSh.getLastRow() - 1, 2).getValues().forEach(r => nameMap[str_(r[0])] = str_(r[1]));
  }
  const out = sheetToObjects_(ss.getSheetByName('Inventory_Balance')).map(o => {
    const code = str_(o.__row[0]); delete o.__row;
    o['Product Name'] = nameMap[code] || '-';
    return o;
  });
  cachePutJson_(CK.BALANCE, out);
  return out;
}

function getDrillDownData(type) {
  if (type !== 'STOCK' && type !== 'LOW') return [];
  const sheet = ss_().getSheetByName('Inventory_Balance');
  if (!sheet) throw new Error("Inventory_Balance missing.");
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const h = data[0].map(x => String(x).trim().toLowerCase());
  const ix = n => h.indexOf(n);
  const low = ix('low stock'), code = ix('item code'), op = ix('opening stock'), tin = ix('total stock in'),
        tout = ix('total stock out'), cl = ix('closing stock'), up = ix('last updated');
  const out = [];
  data.slice(1).filter(r => r[0]).forEach(r => {
    const isLow = low >= 0 && r[low] ? String(r[low]).trim().toUpperCase() === 'LOW STOCK' : false;
    if (type === 'LOW' && !isLow) return;
    out.push({
      'Item Code': r[code],
      'Opening Stock': r[op] || 0,
      'In Stock': r[tin] || 0,
      'Out Stock': r[tout] || 0,
      'Closing Stock': r[cl] || 0,
      'Last Update': up >= 0 && r[up] ? new Date(r[up]).getTime() : '-',
      'isLow': isLow
    });
  });
  return out;
}

function getPOListData() {
  return readPOData_().rows;
}

/* ---------- 8. AUDIT & RECONCILIATION ---------- */
function parseAuditDate_(v) {
  if (!v) return new Date();
  if (v instanceof Date) return v;
  const s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    const p = s.split("-");
    return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 12, 0, 0);
  }
  const d = new Date(s);
  return isNaN(d.getTime()) ? new Date() : d;
}

function isFutureDate_(d) {
  return d.getTime() > Date.now() + 24 * 3600 * 1000;
}

function ensureStockAdjustmentsSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName('Stock_Adjustments');
  if (!sh) {
    sh = ss.insertSheet('Stock_Adjustments');
    sh.getRange(1, 1, 1, 13).setValues([[
      'Adjustment ID', 'Audit Date', 'Item Code', 'Item Name', 'System Stock Before', 'Physical Stock',
      'Adjustment Qty', 'Reason', 'Audit Reference', 'Remarks', 'Adjusted By', 'Recorded At', 'Status'
    ]]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 13).setFontWeight('bold');
    sh.getRange(1, 2, sh.getMaxRows(), 1).setNumberFormat('dd-mmm-yyyy');
    sh.getRange(1, 12, sh.getMaxRows(), 1).setNumberFormat('dd-mmm-yyyy hh:mm:ss');
  }
  return sh;
}

function getSystemStocksAsOf_(codes, auditDate) {
  const ss = ss_(), tz = ss.getSpreadsheetTimeZone();
  const target = Utilities.formatDate(auditDate, tz, 'yyyy-MM-dd');
  const want = {}, out = {};
  codes.forEach(c => { want[c] = true; out[c] = 0; });

  const itemSh = ss.getSheetByName('Item_Master');
  if (itemSh && itemSh.getLastRow() > 1) {
    itemSh.getRange(2, 1, itemSh.getLastRow() - 1, 11).getValues().forEach(r => {
      const c = str_(r[0]);
      if (want[c]) out[c] = (r[10] !== '' && !isNaN(r[10])) ? Number(r[10]) : 0;
    });
  }
  const tSh = ss.getSheetByName('Stock_Transactions');
  if (tSh && tSh.getLastRow() > 1) {
    tSh.getRange(2, 1, tSh.getLastRow() - 1, 5).getValues().forEach(r => {
      const c = str_(r[1]);
      if (!want[c] || !r[0]) return;
      const d = r[0] instanceof Date ? r[0] : new Date(r[0]);
      if (isNaN(d.getTime()) || Utilities.formatDate(d, tz, 'yyyy-MM-dd') > target) return;
      const type = str_(r[3]).toUpperCase(), q = Number(r[4]) || 0;
      if (type === 'IN') out[c] += q; else if (type === 'OUT') out[c] -= q;
    });
  }
  const aSh = ss.getSheetByName('Stock_Adjustments');
  if (aSh && aSh.getLastRow() > 1) {
    aSh.getRange(2, 1, aSh.getLastRow() - 1, 13).getValues().forEach(r => {
      const c = str_(r[2]);
      if (!want[c] || str_(r[12]).toUpperCase() !== 'POSTED' || !r[1]) return;
      const d = r[1] instanceof Date ? r[1] : new Date(r[1]);
      if (!isNaN(d.getTime()) && Utilities.formatDate(d, tz, 'yyyy-MM-dd') <= target) out[c] += Number(r[6]) || 0;
    });
  }
  return out;
}

function getSystemStockForAudit(itemCode, auditDate) {
  checkPermission(['ADMIN']);
  if (!itemCode) throw new Error('Item Code is required.');
  const c = str_(itemCode);
  return { itemCode: c, stock: getSystemStocksAsOf_([c], parseAuditDate_(auditDate))[c] };
}

function getSystemStockForAuditBulk(itemCodes, auditDate) {
  checkPermission(['ADMIN']);
  if (!Array.isArray(itemCodes) || !itemCodes.length) return {};
  const codes = [...new Set(itemCodes.map(str_).filter(Boolean))];
  return getSystemStocksAsOf_(codes, parseAuditDate_(auditDate));
}

function createStockAdjustments(payloads) {
  checkPermission(['ADMIN']);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (!Array.isArray(payloads) || !payloads.length) throw new Error('At least one audit adjustment is required.');
    const ss = ss_(), itemSh = ss.getSheetByName('Item_Master');
    if (!itemSh || itemSh.getLastRow() < 2) throw new Error('Item_Master database is empty.');

    const itemMap = {};
    itemSh.getRange(2, 1, itemSh.getLastRow() - 1, 2).getValues().forEach(r => {
      if (r[0]) {
        const c = str_(r[0]);
        itemMap[c] = r;
        itemMap[c.toUpperCase()] = r;
      }
    });

    const norm = [], seenItem = {}, seenRef = {};
    payloads.forEach((p, idx) => {
      const n = 'Audit row ' + (idx + 1) + ': ';
      if (!p) throw new Error(n + 'data is missing.');
      const itemCode = str_(p.itemCode), reason = str_(p.reason), ref = str_(p.auditReference).toUpperCase();
      const physical = Number(p.physicalStock);
      if (!itemCode) throw new Error(n + 'Item Code is required.');
      if (!itemMap[itemCode] && !itemMap[itemCode.toUpperCase()]) throw new Error(n + 'Item not found: ' + itemCode);
      if (!Number.isFinite(physical) || physical < 0) throw new Error(n + 'Physical stock must be 0 or greater.');
      if (!reason) throw new Error(n + 'Audit reason is required.');
      if (!ref) throw new Error(n + 'Audit Reference is required.');
      if (seenItem[itemCode]) throw new Error('Duplicate item in audit batch: ' + itemCode);
      if (seenRef[ref]) throw new Error('Duplicate Audit Reference in batch: ' + ref);
      const auditDate = parseAuditDate_(p.auditDate);
      if (isFutureDate_(auditDate)) throw new Error(n + 'Audit date cannot be in the future.');
      seenItem[itemCode] = seenRef[ref] = true;
      norm.push({ itemCode: itemCode, reason: reason, ref: ref, remarks: str_(p.remarks), physical: physical, auditDate: auditDate });
    });

    const sh = ensureStockAdjustmentsSheet_();
    const existing = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 13).getValues() : [];
    const usedRefs = {};
    existing.forEach(r => { if (str_(r[12]).toUpperCase() === 'POSTED') usedRefs[str_(r[8]).toUpperCase()] = true; });

    const byDate = {};
    norm.forEach(x => {
      const k = Utilities.formatDate(x.auditDate, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
      (byDate[k] = byDate[k] || { date: x.auditDate, codes: [] }).codes.push(x.itemCode);
    });
    const stocks = {};
    Object.keys(byDate).forEach(k => { stocks[k] = getSystemStocksAsOf_(byDate[k].codes, byDate[k].date); });

    const user = Session.getActiveUser().getEmail() || 'WebUser', now = new Date();
    const stamp = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), 'yyyyMMdd-HHmmss');
    const rows = [], results = [];
    norm.forEach(x => {
      const dk = Utilities.formatDate(x.auditDate, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
      const system = stocks[dk][x.itemCode], adj = x.physical - system;
      const ref = usedRefs[x.ref] ? x.ref + '-' + Utilities.getUuid().slice(0, 4).toUpperCase() : x.ref;
      const id = 'ADJ-' + stamp + '-' + Utilities.getUuid().slice(0, 8).toUpperCase();
      const nm = (itemMap[x.itemCode] || itemMap[x.itemCode.toUpperCase()])[1] || '';
      rows.push([id, x.auditDate, x.itemCode, nm, system, x.physical, adj, x.reason, ref, x.remarks || '-', user, now, 'POSTED']);
      results.push({ adjustmentId: id, itemCode: x.itemCode, systemStockBefore: system, physicalStock: x.physical, adjustmentQty: adj });
    });

    sh.getRange(sh.getLastRow() + 1, 1, rows.length, 13).setValues(rows);
    rebuildInventoryBalance();
    rebuildDailySummary();
    clearCaches_(DYN_KEYS);
    return { success: true, results: results, message: 'Successfully posted ' + results.length + ' audit adjustment(s).' };
  } finally {
    lock.releaseLock();
  }
}

function createStockAdjustment(payload) {
  if (!payload) throw new Error('Audit adjustment data is required.');
  const res = createStockAdjustments([payload]), r = res.results[0];
  return {
    success: true,
    adjustmentId: r.adjustmentId,
    itemCode: r.itemCode,
    systemStockBefore: r.systemStockBefore,
    physicalStock: r.physicalStock,
    adjustmentQty: r.adjustmentQty,
    message: r.adjustmentQty === 0 ? 'Audit completed. No stock variance was found.'
      : 'Audit adjustment posted: ' + (r.adjustmentQty > 0 ? '+' : '') + r.adjustmentQty + ' unit(s).'
  };
}

function getStockAuditHistory(itemCode, limit) {
  checkPermission(['ADMIN']);
  const sh = ss_().getSheetByName('Stock_Adjustments');
  if (!sh || sh.getLastRow() < 2) return [];
  const code = itemCode ? str_(itemCode) : '';
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 13).getDisplayValues().filter(r => !code || str_(r[2]) === code);
  return rows.slice(-Math.max(1, Number(limit) || 50)).reverse().map(r => ({
    id: r[0],
    auditDate: r[1],
    itemCode: r[2],
    itemName: r[3],
    systemStockBefore: Number(r[4]) || 0,
    physicalStock: Number(r[5]) || 0,
    adjustmentQty: Number(r[6]) || 0,
    reason: r[7],
    reference: r[8],
    remarks: r[9],
    adjustedBy: r[10],
    recordedAt: r[11],
    status: r[12]
  }));
}

/* ---------- 9. SUMMARY & REBUILD ENGINES ---------- */
function rebuildInventoryBalance() {
  const ss = ss_();
  const trans = ss.getSheetByName('Stock_Transactions'), bal = ss.getSheetByName('Inventory_Balance'), itemSh = ss.getSheetByName('Item_Master');
  if (!trans || !bal || !itemSh) return;

  const tz = ss.getSpreadsheetTimeZone(), todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const blank = op => ({ opening: op, totalIn: 0, totalOut: 0, auditDelta: 0, todayAudit: 0, lastUpdated: new Date(0), lastAuditDate: '' });
  const inv = {};

  const iData = itemSh.getDataRange().getValues();
  for (let i = 1; i < iData.length; i++) {
    const code = str_(iData[i][0]);
    if (!code) continue;
    const o = iData[i][10];
    inv[code] = blank((o !== '' && !isNaN(o)) ? Number(o) : 0);
  }

  const tData = trans.getDataRange().getValues();
  for (let i = 1; i < tData.length; i++) {
    const code = str_(tData[i][1]);
    if (!tData[i][0] || !code) continue;
    const d = tData[i][0] instanceof Date ? tData[i][0] : new Date(tData[i][0]);
    if (isNaN(d.getTime())) continue;
    const b = inv[code] || (inv[code] = blank(0));
    const type = str_(tData[i][3]).toUpperCase(), qty = Number(tData[i][4]) || 0;
    if (type === 'IN') b.totalIn += qty; else if (type === 'OUT') b.totalOut += qty;
    if (d > b.lastUpdated) b.lastUpdated = d;
  }

  const adjSh = ss.getSheetByName('Stock_Adjustments');
  if (adjSh && adjSh.getLastRow() > 1) {
    adjSh.getRange(2, 1, adjSh.getLastRow() - 1, 13).getValues().forEach(r => {
      const code = str_(r[2]);
      if (str_(r[12]).toUpperCase() !== 'POSTED' || !code) return;
      const ad = r[1] instanceof Date ? r[1] : new Date(r[1]);
      if (isNaN(ad.getTime())) return;
      const b = inv[code] || (inv[code] = blank(0));
      const delta = Number(r[6]) || 0, ads = Utilities.formatDate(ad, tz, 'yyyy-MM-dd');
      b.auditDelta += delta;
      if (ads === todayStr) b.todayAudit += delta;
      b.lastAuditDate = ads;
      const rec = r[11] ? (r[11] instanceof Date ? r[11] : new Date(r[11])) : ad;
      if (!isNaN(rec.getTime()) && rec > b.lastUpdated) b.lastUpdated = rec;
    });
  }

  const out = Object.keys(inv).map(code => {
    const b = inv[code];
    return [
      code,
      b.opening,
      b.totalIn,
      b.totalOut,
      b.opening + b.totalIn + b.auditDelta - b.totalOut,
      b.lastUpdated.getTime() === 0 ? new Date() : b.lastUpdated,
      b.todayAudit,
      b.lastAuditDate || ''
    ];
  });
  if (!out.length) return;
  if (bal.getLastRow() > 1) bal.getRange(2, 1, bal.getLastRow() - 1, 8).clearContent();
  bal.getRange(1, 1, 1, 8).setValues([['Item Code', 'Opening Stock', 'Total Stock In', 'Total Stock Out', 'Closing Stock', 'Last Updated', 'Today Audit Adjustment', 'Last Audit Date']]);
  bal.getRange(2, 1, out.length, 8).setValues(out);
}

function rebuildDailySummary() {
  const ss = ss_();
  let sum = ss.getSheetByName('Item_Daily_Summary');
  const trans = ss.getSheetByName('Stock_Transactions');
  const itemSh = ss.getSheetByName('Item_Master');
  const adjSh = ss.getSheetByName('Stock_Adjustments');

  if (!trans || !itemSh) return;

  // Auto-create sheet if missing
  if (!sum) sum = ss.insertSheet('Item_Daily_Summary');

  const headers = ['Date', 'Item Code', 'Item Name', 'Opening Stock', 'Total In', 'Total Out', 'Closing Stock'];
  sum.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');

  if (sum.getLastRow() > 1) {
    sum.getRange(2, 1, sum.getLastRow() - 1, headers.length).clearContent();
  }

  const tz = ss.getSpreadsheetTimeZone();
  const running = {}, nameMap = {};

  const iData = itemSh.getDataRange().getValues();
  for (let i = 1; i < iData.length; i++) {
    const code = str_(iData[i][0]);
    if (code) {
      nameMap[code] = str_(iData[i][1]);
      running[code] = (iData[i][10] !== "" && !isNaN(iData[i][10])) ? Number(iData[i][10]) : 0;
    }
  }

  const daily = {};

  // 1. Process Stock Transactions
  if (trans.getLastRow() > 1) {
    const tData = trans.getDataRange().getValues();
    for (let i = 1; i < tData.length; i++) {
      const r = tData[i];
      if (!r[0] || !r[1]) continue;
      const d = r[0] instanceof Date ? r[0] : new Date(r[0]);
      if (isNaN(d.getTime())) continue;

      const code = str_(r[1]);
      const key = Utilities.formatDate(d, tz, 'yyyy-MM-dd') + "|" + code;
      const e = daily[key] || (daily[key] = {
        show: Utilities.formatDate(d, tz, 'dd-MMM-yyyy'),
        code: code,
        name: str_(r[2]) || nameMap[code] || code,
        in: 0,
        out: 0
      });

      const type = str_(r[3]).toUpperCase(), qty = Number(r[4]) || 0;
      if (type === 'IN') e.in += qty;
      else if (type === 'OUT') e.out += qty;
    }
  }

  // 2. Process Stock Adjustments
  if (adjSh && adjSh.getLastRow() > 1) {
    const aData = adjSh.getRange(2, 1, adjSh.getLastRow() - 1, 13).getValues();
    for (let i = 0; i < aData.length; i++) {
      const r = aData[i];
      if (str_(r[12]).toUpperCase() !== 'POSTED') continue;
      const code = str_(r[2]);
      if (!code) continue;
      const d = r[1] instanceof Date ? r[1] : new Date(r[1]);
      if (isNaN(d.getTime())) continue;

      const key = Utilities.formatDate(d, tz, 'yyyy-MM-dd') + "|" + code;
      const e = daily[key] || (daily[key] = {
        show: Utilities.formatDate(d, tz, 'dd-MMM-yyyy'),
        code: code,
        name: str_(r[3]) || nameMap[code] || code,
        in: 0,
        out: 0
      });

      const adjQty = Number(r[6]) || 0;
      if (adjQty > 0) e.in += adjQty;
      else if (adjQty < 0) e.out += Math.abs(adjQty);
    }
  }

  // 3. Calculate Running Stock & Output Rows
  const rows = Object.keys(daily).sort().map(k => {
    const r = daily[k];
    const open = running[r.code] || 0;
    const net = open + r.in - r.out;
    running[r.code] = net;
    return [r.show, r.code, r.name, open, r.in, r.out, net];
  });

  if (rows.length) {
    sum.getRange(2, 1, rows.length, 7).setValues(rows);
  }
}

/* ---------- 10. LOGISTICS ROUTE PLANNER ---------- */
function getDistanceFromLatLonInKm(lat1, lon1, lat2, lon2) {
  const R = 6371, rad = Math.PI / 180;
  const dLat = (lat2 - lat1) * rad, dLon = (lon2 - lon1) * rad;
  const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) + Math.cos(lat1 * rad) * Math.cos(lat2 * rad) * Math.sin(dLon / 2) * Math.sin(dLon / 2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function getNearbyVendors(baseLocation) {
  checkPermission(["ADMIN", "STORE"]);
  const sheet = ss_().getSheetByName('Vendor and Material List');
  if (!sheet) throw new Error("Could not find the 'Vendor and Material List' sheet.");
  const data = sheet.getDataRange().getValues();

  const baseGeo = Maps.newGeocoder().geocode(baseLocation + ", Delhi NCR, India");
  if (baseGeo.status !== 'OK') throw new Error("Base location not found. Please try more specific details.");
  const base = baseGeo.results[0].geometry.location;

  const list = [];
  for (let i = 1; i < data.length; i++) {
    const location = data[i][2];
    if (!location) continue;
    let lat = data[i][4], lng = data[i][5];
    if (!lat || !lng) {
      Utilities.sleep(100);
      const geo = Maps.newGeocoder().geocode(location + ", Delhi NCR, India");
      if (geo.status === 'OK') {
        lat = geo.results[0].geometry.location.lat;
        lng = geo.results[0].geometry.location.lng;
        sheet.getRange(i + 1, 5, 1, 2).setValues([[lat, lng]]);
      }
    }
    if (lat && lng) {
      const dist = getDistanceFromLatLonInKm(base.lat, base.lng, lat, lng);
      if (dist <= 15) list.push({ vendor: data[i][1], location: location, material: data[i][3] || 'N/A', distance: dist.toFixed(1) });
    }
  }
  list.sort((a, b) => a.distance - b.distance);
  return list;
}