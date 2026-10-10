const express = require("express");
const { google } = require("googleapis");
const NodeCache = require("node-cache");
const cors = require("cors");

const app = express();
app.use(cors());
app.use(express.urlencoded({ extended: true }));
app.use(express.json({ limit: "20mb", type: ["application/json", "text/plain"] }));

const cache = new NodeCache({ stdTTL: 120, checkperiod: 30 });

const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT),
  scopes: ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"],
});
const sheets = google.sheets({ version: "v4", auth });
const SHEET_ID = process.env.SHEET_ID;
const APPS_SCRIPT_URL = process.env.APPS_SCRIPT_URL || "";

const S = {
  PRODUCTS: "Product & Stock Master", ATTENDANCE: "Attendance", DEMAND: "Demand",
  PO: "Purchase Orders", STAFF: "Staff", WHLOG: "Warehouse Entry Log",
  CYLINDER: "Cylinder Entry", SETTINGS: "Settings",
};

// =================== HELPERS ===================
function num(v) { const n = Number(v); return isNaN(n) ? 0 : n; }
function round3(v) { return Math.round(v * 1000) / 1000; }
function normH(s) { return String(s).toLowerCase().replace(/\s+/g, " ").trim(); }
function pad2(n) { return String(n).padStart(2, "0"); }
function toISODate(d) { return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()); }

// 🆕 IST helpers — Indian Standard Time = UTC + 5:30
function toISTISO(d) {
  const dt = (d instanceof Date) ? d : new Date(d);
  if (isNaN(dt.getTime())) return "";
  const ist = new Date(dt.getTime() + 5.5 * 3600000);
  return ist.toISOString().slice(0, 19); // "2026-10-09T17:30:00" (IST wall-clock)
}

function normalizeDate(v) {
  if (v === null || v === undefined || v === "") return "";
  if (v instanceof Date) return toISODate(v);
  const s = String(v).trim(); if (!s) return "";
  if (/^\d{4}-\d{1,2}-\d{1,2}/.test(s)) return s.slice(0, 10);
  const us = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (us) return `${us[3]}-${pad2(us[1])}-${pad2(us[2])}`;
  const n = Number(s);
  if (!isNaN(n) && n > 20000 && n < 60000) {
    const d = new Date(Date.UTC(1899, 11, 30) + n * 86400000);
    return toISODate(d);
  }
  const d = new Date(s);
  if (!isNaN(d.getTime())) return toISODate(d);
  return "";
}

function normalizeDateTime(v, dateHint) {
  if (v === null || v === undefined || v === "") return "";
  if (v instanceof Date) return toISTISO(v);
  const s = String(v).trim(); if (!s) return "";

  // 🆕 ISO with Z suffix (UTC) → convert to IST
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?Z$/i.test(s)) {
    const d = new Date(s);
    if (!isNaN(d.getTime())) return toISTISO(d);
  }

  // ISO without Z (naive IST) → keep as-is
  if (/^\d{4}-\d{2}-\d{2}T/.test(s)) return s.slice(0, 19);
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) { const d = new Date(s); if (!isNaN(d.getTime())) return d.toISOString().slice(0, 19); }
  const n = Number(s);
  if (!isNaN(n) && n > 20000 && n < 60000) { const ms = Date.UTC(1899, 11, 30) + n * 86400000; return toISTISO(new Date(ms)); }
  const usFull = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?/i);
  if (usFull) {
    let h = parseInt(usFull[4]); const ap = (usFull[7] || "").toUpperCase();
    if (ap === "PM" && h !== 12) h += 12; if (ap === "AM" && h === 12) h = 0;
    return `${usFull[3]}-${pad2(usFull[1])}-${pad2(usFull[2])}T${pad2(h)}:${usFull[5]}:${usFull[6] || "00"}`;
  }
  const tOnly = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?/i);
  if (tOnly && dateHint) {
    let h = parseInt(tOnly[1]); const ap = (tOnly[4] || "").toUpperCase();
    if (ap === "PM" && h !== 12) h += 12; if (ap === "AM" && h === 12) h = 0;
    return `${dateHint}T${pad2(h)}:${tOnly[2]}:${tOnly[3] || "00"}`;
  }
  const d = new Date(s);
  if (!isNaN(d.getTime())) return toISTISO(d);
  return "";
}

function colLetter(n) {
  let s = "";
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

async function readRange(sheetName, range = "A1:ZZ5000", renderOption = "UNFORMATTED_VALUE") {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: `'${sheetName}'!${range}`,
    valueRenderOption: renderOption,
  });
  return res.data.values || [];
}

async function writeCell(sheetName, row, col, value) {
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID, range: `'${sheetName}'!${colLetter(col)}${row}`,
    valueInputOption: "USER_ENTERED", requestBody: { values: [[value]] },
  });
}

async function appendRow(sheetName, values) {
  const resp = await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID, range: `'${sheetName}'!A1`,
    valueInputOption: "USER_ENTERED", insertDataOption: "INSERT_ROWS",
    requestBody: { values: [values] },
  });
  const updatedRange = (resp.data.updates && resp.data.updates.updatedRange) || "";
  const m = updatedRange.match(/![A-Z]+(\d+)/);
  return m ? parseInt(m[1], 10) : 0;
}

function rowsToObjects(rows) {
  if (!rows.length) return [];
  const headers = rows[0].map((h) => String(h).trim());
  return rows.slice(1).map((r, i) => {
    const o = { _row: i + 2 };
    headers.forEach((h, idx) => { if (h) o[h] = r[idx]; });
    return o;
  });
}

async function cached(key, builder) {
  const hit = cache.get(key); if (hit !== undefined) return hit;
  const data = await builder(); cache.set(key, data); return data;
}
function invalidate(...keys) { keys.forEach(k => cache.del(k)); }

// =================== 🆕 PHOTO UPLOAD (Apps Script → Google Drive) ===================
// Photo Google Drive mein Apps Script ke through save hoti hai. Agar upload fail ho
// to error Render Logs mein dikhega (pehle silently chhup jata tha).
async function uploadPhotoToDrive(photoBase64, userId, name) {
  if (!photoBase64) return "";
  if (!APPS_SCRIPT_URL) { console.error("Photo upload skipped: APPS_SCRIPT_URL env variable set nahi hai"); return ""; }
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 60000);
  try {
    const r = await fetch(APPS_SCRIPT_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({ action: "uploadPhoto", photoBase64, userId: userId || "photo", name: name || "photo" }),
      signal: controller.signal,
    });
    const txt = await r.text();
    let j;
    try { j = JSON.parse(txt); } catch (e) { console.error("Photo upload: JSON nahi mila. Response:", txt.slice(0, 200)); return ""; }
    const url = (j && (j.url || (j.data && j.data.url))) || "";
    if (!url) console.error("Photo upload failed:", JSON.stringify(j).slice(0, 300));
    return url;
  } catch (e) {
    console.error("Photo upload error:", e.message);
    return "";
  } finally { clearTimeout(timer); }
}

// =================== 🆕 SHIFT STATUS (On Time / Late / Early Leave / Overtime) ===================
function parseClock(str) {
  const m = String(str).match(/(\d+):(\d+)\s*(AM|PM)/i);
  if (!m) return null;
  let h = Number(m[1]) % 12;
  if (m[3].toUpperCase() === "PM") h += 12;
  return h * 60 + Number(m[2]);
}
function istMinutes(d) {
  const t = new Date(d.getTime() + 5.5 * 3600000);
  return t.getUTCHours() * 60 + t.getUTCMinutes();
}
function computeInStatusAt(shift, minutes) {
  try {
    const start = parseClock(String(shift).split("-")[0].trim());
    if (start === null) return "Marked";
    return minutes > start + 15 ? "Late" : "On Time";
  } catch (e) { return "Marked"; }
}
function computeOutStatusAt(shift, minutes) {
  try {
    const parts = String(shift).split("-");
    const start = parseClock(parts[0].trim());
    let end = parseClock((parts[1] || "").trim());
    if (start === null || end === null) return "Marked";
    let n = minutes;
    if (end <= start) { end += 1440; if (n < start) n += 1440; } // overnight shift (jaise 6 PM - 2 AM)
    if (n < end - 60) return "Early Leave";
    if (n > end + 30) return "Overtime";
    return "On Time";
  } catch (e) { return "Marked"; }
}
function computeInStatus(shift, now) { return computeInStatusAt(shift, istMinutes(now)); }
function computeOutStatus(shift, now) { return computeOutStatusAt(shift, istMinutes(now)); }
function hmToMinutes(hm) {
  const m = String(hm || "").match(/^(\d{1,2}):(\d{2})/);
  return m ? Number(m[1]) * 60 + Number(m[2]) : null;
}

// =================== READERS ===================
async function readStaff() { return rowsToObjects(await readRange(S.STAFF)); }

async function readDemands() {
  const rows = await readRange(S.DEMAND, "A1:N500", "FORMATTED_VALUE");
  return rows.slice(1).reverse().map((r, i) => ({
    _row: rows.length - i, "Demand ID": r[0] || "",
    Timestamp: normalizeDateTime(r[1]),
    "Raised By": r[2] || "", "Item Name": r[3] || "", Category: r[4] || "",
    Unit: r[5] || "", "Qty Requested": r[6] || 0, Status: r[7] || "Pending",
    "Warehouse Stock After": r[8] || "", Notes: r[9] || "",
    "Approved Qty": r[10] || "", "Approved At": normalizeDateTime(r[11]),
    Feedback: r[12] || "", "Feedback At": normalizeDateTime(r[13]),
  }));
}

async function readProducts() {
  const rows = await readRange(S.PRODUCTS, "A1:ZZ5000", "FORMATTED_VALUE");
  if (rows.length < 2) return [];
  const H = rows[0].map(normH);
  const find = (pred, def) => { for (let i = 0; i < H.length; i++) if (H[i] && pred(H[i])) return i; return def; };
  const C = {
    NAME: find(x => x.indexOf("item name") === 0, 1), CAT: find(x => x === "category", 2),
    PACK: find(x => x.indexOf("packaging") !== -1, 3), UNIT: find(x => x === "unit", 4),
    OPENFULL: find(x => x.indexOf("full pack") !== -1, 5), OPENLOOSE: find(x => x.indexOf("loose") !== -1, 6),
    DEMAND: find(x => x.indexOf("deman") === 0 && x.indexOf("approved") === -1 && x.indexOf("unit") === -1, 7),
    DEMANDUNIT: find(x => x.indexOf("demand unit") !== -1, 8),
    APPROVED: find(x => x.indexOf("approved") !== -1, 9), PURCHASE: find(x => x.indexOf("purchase") !== -1, 10),
    MIN: find(x => x === "min", 13), MAX: find(x => x === "max", 14), VENDOR: find(x => x === "vendor", 15),
    RATE: find(x => x === "rate", 16), RATEUNIT: find(x => x.indexOf("unit of rate") !== -1, 17),
    EXPIRY: find(x => x.indexOf("expiry") === 0, 18), BASEUNIT: find(x => x.indexOf("base unit") !== -1, 19),
    STOCKUNITQTY: find(x => x.indexOf("stock unit") !== -1 || x.indexOf("base qty") !== -1, 20),
    STOCKTYPE: find(x => x.indexOf("stock type") !== -1, 21),
  };
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i]; if (!r || !String(r[C.NAME] || "").trim()) continue;
    const openFull = num(r[C.OPENFULL]), openLoose = num(r[C.OPENLOOSE]);
    const purchase = num(r[C.PURCHASE]), approved = num(r[C.APPROVED]);
    const suq = num(r[C.STOCKUNITQTY]) || 1;
    const opening = round3(openFull * suq + openLoose);
    const currentPcs = round3(opening + purchase - approved);
    out.push({
      _row: i + 1, "Item Name (Standardized)": String(r[C.NAME]).trim(),
      Category: r[C.CAT] || "", Unit: r[C.UNIT] || "", Packaging: r[C.PACK] || "",
      "W/H Opening Full Pack": openFull, "W/H Opening Loose": openLoose,
      "Warehouse Stock Qty Opening STock": opening,
      "Deman for canteen": num(r[C.DEMAND]), "Demand Unit": r[C.DEMANDUNIT] || "",
      "Approved Demand Qty by supervisor": approved, "Purchase Stock  QTy": purchase,
      "Current Stock": currentPcs, "Current Stock In Carton": round3(currentPcs / suq),
      Min: r[C.MIN], Max: r[C.MAX], Vendor: r[C.VENDOR] || "",
      Rate: r[C.RATE], "Unit of rate": r[C.RATEUNIT] || "",
      "Expiry Date": normalizeDate(r[C.EXPIRY]),
      "Base Unit": r[C.BASEUNIT] || r[C.UNIT] || "", "1 Stock Unit = (Base Qty)": suq,
      "Stock Type": r[C.STOCKTYPE] || "Stocked",
    });
  }
  return out;
}

async function readWarehouseEntries() {
  const objs = rowsToObjects(await readRange(S.WHLOG, "A1:ZZ500", "FORMATTED_VALUE"));
  objs.forEach(o => { o['Timestamp'] = normalizeDateTime(o['Timestamp']); o['Expiry Date'] = normalizeDate(o['Expiry Date']); o['Last Edited'] = normalizeDateTime(o['Last Edited']); });
  return objs.slice(-200).reverse();
}
async function readCylinderEntries() {
  const objs = rowsToObjects(await readRange(S.CYLINDER, "A1:ZZ500", "FORMATTED_VALUE"));
  objs.forEach(o => { o['Timestamp'] = normalizeDateTime(o['Timestamp']); });
  return objs.slice(-200).reverse();
}
async function readPO() {
  const objs = rowsToObjects(await readRange(S.PO, "A1:ZZ2000", "FORMATTED_VALUE"));
  objs.forEach(o => { o['Timestamp'] = normalizeDateTime(o['Timestamp']); o['Expiry Date'] = normalizeDate(o['Expiry Date']); });
  return objs;
}
async function readManualPurchases() {
  const objs = await readPO();
  return objs.filter(r => String(r["PO ID"]).indexOf("MPO-") === 0).slice(-400).reverse();
}
async function readAttendance() {
  const rows = await readRange(S.ATTENDANCE, "A1:R2000", "FORMATTED_VALUE");
  if (rows.length < 2) return [];
  const headers = rows[0].map(h => String(h).trim());
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i]; const obj = { _row: i + 2 };
    headers.forEach((h, idx) => { if (h) obj[h] = r[idx]; });
    obj['Date'] = normalizeDate(obj['Date']);
    obj['IN Time'] = normalizeDateTime(obj['IN Time'], obj['Date']);
    obj['OUT Time'] = normalizeDateTime(obj['OUT Time'], obj['Date']);
    if (!obj['User ID'] && !obj['Date']) continue;
    out.push(obj);
  }
  return out;
}
async function readSettings() {
  const objs = rowsToObjects(await readRange(S.SETTINGS));
  const map = {}; objs.forEach(r => { map[String(r.Key).trim()] = r.Value; });
  return map;
}
async function getSetting(key, def) {
  const map = await cached("settings", readSettings);
  const v = map[String(key).trim()];
  return v === undefined ? def : v;
}
async function setSetting(key, val) {
  const objs = rowsToObjects(await readRange(S.SETTINGS));
  const match = objs.find(r => String(r.Key).trim() === key);
  if (match) await writeCell(S.SETTINGS, match._row, 2, val);
  else await appendRow(S.SETTINGS, [key, val]);
  cache.del("settings");
}
async function readVendors() {
  const products = await cached("products", readProducts);
  const names = new Set();
  products.forEach(p => { const v = String(p.Vendor || "").trim(); if (v) names.add(v); });
  const extra = await getSetting("ExtraVendors", "");
  if (extra) { try { JSON.parse(extra).forEach(v => { if (v) names.add(v); }); } catch (e) {} }
  return [...names].sort();
}

const VENDOR_ITEMS = {
  "Mauranipur (Shivam Sahu)": ["tel tin","nariyal","chabal","dal","haldi","lal mirch powder","dhaniya powder","namak","nirma","khatta meetha","aalu bhujia","navratan","bhel","kaju","magaj","meet masala","kichin king","chat masala","jeera powder masala","papad","jeera","nisree","green chilli","vinegar","red chilli","heeng","hani","gulab jal","methi","aajban","ararot","idli powder","ajinomoto","kali mirch powder","peanut","chole","long","sarso tel botal","imli","sambhar masala","white pepper","kastoori maithi","red kalar","green kalar","lal mirch khadi","khada dhaniya"],
  "Bada Bazar": ["kale dibba","combo","nepkin","spoon","clean wrap","silver rol","french fries lifafa","stro","chatni dibbi badi","lakdi chammach","chay gilas chote","chay gilas bade","pani gilas"],
  "Venktesh": ["french fries","dicheese","corn","cheese slices","cream","thaujan","myoonis","meggi","tameto ketchup","tameto sese","oregano","chilli flakes","white sauces","bread crumbs"],
};
function matchVendorForItem(name) {
  const norm = String(name).toLowerCase().replace(/\([^)]*\)/g, "").replace(/\s+/g, " ").trim();
  for (const v in VENDOR_ITEMS) if (VENDOR_ITEMS[v].indexOf(norm) !== -1) return v;
  return "";
}

// =================== 🆕 AUTO-PO ENGINE ===================
async function checkAndCreateAutoPO(itemName) {
  if (!itemName) return null;
  const products = await cached("products", readProducts);
  const item = products.find(x => normH(x["Item Name (Standardized)"]) === normH(itemName));
  if (!item) return null;

  const min = num(item.Min);
  const max = num(item.Max);
  const cur = num(item["Current Stock"]);

  // Only trigger if Min set and stock at/below Min
  if (min <= 0) return null;
  if (cur > min) return null;

  // Formula: qty = max > min ? (max - current) : (min * 2 - current)
  let orderQty;
  if (max > min) orderQty = Math.max(1, round3(max - cur));
  else orderQty = Math.max(1, round3(min * 2 - cur));

  // Duplicate safety: agar Pending PO pehle se hai us item ka
  const pos = await cached("po", readPO);
  const existingPending = pos.find(p =>
    normH(String(p["Item Name"] || "")) === normH(itemName) &&
    String(p.Status || "").trim() === "Pending"
  );
  if (existingPending) {
    return { skipped: true, reason: "Pending PO already exists", row: existingPending._row, existingQty: existingPending["Qty Ordered"] };
  }

  const poId = "APO-" + Date.now() + "-" + Math.floor(Math.random() * 90 + 10);
  const vendor = item.Vendor || "Unassigned";
  const ts = toISTISO(new Date());
  // Columns: PO ID, Timestamp, Item Name, Category, Unit, Qty, Status, Requested By, Notes, Vendor
  await appendRow(S.PO, [
    poId, ts, item["Item Name (Standardized)"], item.Category || "",
    item.Unit || "", orderQty, "Pending", "AUTO-PO",
    `Auto: Stock ${cur} <= Min ${min}`, vendor
  ]);
  invalidate("po", "po_batches", "mpo", "dash");
  return { created: true, poId, itemName: item["Item Name (Standardized)"], qty: orderQty, current: cur, min, max, vendor };
}

async function runAutoPOForAll() {
  const products = await cached("products", readProducts);
  const results = { created: [], skipped: [], failed: [] };
  for (const p of products) {
    const min = num(p.Min);
    const cur = num(p["Current Stock"]);
    if (min > 0 && cur <= min) {
      try {
        const r = await checkAndCreateAutoPO(p["Item Name (Standardized)"]);
        if (r && r.created) results.created.push(r);
        else if (r && r.skipped) results.skipped.push({ name: p["Item Name (Standardized)"], reason: r.reason });
      } catch (e) {
        results.failed.push({ name: p["Item Name (Standardized)"], error: e.message });
      }
    }
  }
  return results;
}

// =================== HANDLERS ===================
async function handleLogin(p) {
  const staff = await cached("staff", readStaff);
  const match = staff.find(r => String(r.User || "").toLowerCase().trim() === String(p.userId).toLowerCase().trim() && String(r.Password || "").trim() === String(p.password).trim());
  if (!match) throw new Error("Invalid User ID or Password");
  const result = { userId: match.User, name: match.Name, designation: match.Designation, shift: match["Shift Timing"], permission: match.Permission };
  try {
    const want = []; const perm = result.permission;
    if (perm === "SuperAdmin") want.push("staff", "allAttendance");
    else {
      want.push("myAttendance");
      if (perm === "Admin") want.push("stats", "products", "demands");
      else if (perm === "Supervisor") want.push("demands", "products", "vendors");
      else if (perm === "Cashier") want.push("products", "demands");
    }
    result.bootstrap = await handleBootstrap({ want, userId: result.userId });
  } catch (e) { console.error("Bootstrap:", e.message); result.bootstrap = {}; }
  return result;
}

async function handleBootstrap(p) {
  const want = p.want || [];
  const out = { products: [], demands: [], staff: [], myAttendance: [], allAttendance: [], vendors: [], stats: null, poll: null };
  const tasks = [];
  if (want.includes("products")) tasks.push(cached("products", readProducts).then(d => out.products = d).catch(() => {}));
  if (want.includes("demands")) tasks.push(cached("demands", readDemands).then(d => out.demands = d).catch(() => {}));
  if (want.includes("staff")) tasks.push(cached("staff", readStaff).then(d => out.staff = d).catch(() => {}));
  if (want.includes("vendors")) tasks.push(cached("vendors", readVendors).then(d => out.vendors = d).catch(() => {}));
  if (want.includes("stats")) tasks.push(buildDashboardStats().then(d => out.stats = d).catch(() => {}));
  if (want.includes("myAttendance") || want.includes("allAttendance")) {
    tasks.push(cached("attendance", readAttendance).then(d => {
      if (want.includes("myAttendance")) out.myAttendance = d.filter(r => String(r["User ID"]).toLowerCase().trim() === String(p.userId).toLowerCase().trim());
      if (want.includes("allAttendance")) out.allAttendance = d;
    }).catch(() => {}));
  }
  await Promise.all(tasks);
  try { out.poll = await buildPoll(); } catch (e) {}
  return out;
}

async function buildPoll() {
  const demands = await cached("demands", readDemands);
  let pending = 0, newest = null;
  demands.forEach(d => { if (d.Status === "Pending") { pending++; if (!newest || new Date(d.Timestamp) > new Date(newest.Timestamp)) newest = d; } });
  return { ver: Date.now(), pending, newest: newest ? { item: newest["Item Name"], qty: newest["Qty Requested"], unit: newest.Unit, by: newest["Raised By"] } : null };
}

async function buildDashboardStats() {
  const products = (await cached("products", readProducts)).filter(p => String(p["Item Name (Standardized)"] || "").trim() !== "");
  const attendance = (await cached("attendance", readAttendance)).slice(-700);
  const demands = await cached("demands", readDemands);
  const pos = await cached("po", readPO);
  const staff = await cached("staff", readStaff);
  const istNow = new Date(Date.now() + 5.5 * 3600000);
  const today = istNow.toISOString().slice(0, 10);
  const presentIds = new Set(attendance.filter(a => String(a.Date || "").slice(0, 10) === today && a["IN Time"]).map(a => a["User ID"]));
  const totalStaff = staff.filter(s => s.Permission !== "Admin" && s.Permission !== "SuperAdmin").length;
  const health = { healthy: 0, low: 0, overstock: 0 };
  const categoryStock = {};
  products.forEach(p => {
    const min = num(p.Min), max = num(p.Max);
    // 🔧 FIX: Compare against PCS (Current Stock), not carton
    const cur = num(p["Current Stock"]);
    const cat = p.Category || "Other"; categoryStock[cat] = (categoryStock[cat] || 0) + num(p["Current Stock"]);
    if (min > 0 && cur <= min) health.low++; else if (max > 0 && cur >= max) health.overstock++; else health.healthy++;
  });
  const trend = [];
  for (let i = 6; i >= 0; i--) {
    const d = new Date(istNow); d.setDate(d.getDate() - i);
    const key = d.toISOString().slice(0, 10);
    const ids = new Set(attendance.filter(a => String(a.Date || "").slice(0, 10) === key && a["IN Time"]).map(a => a["User ID"]));
    trend.push({ date: key, present: ids.size });
  }
  const poStatus = { Pending: 0, Ordered: 0, Received: 0, Cancelled: 0 };
  pos.forEach(p => { const s = String(p.Status || "").trim(); if (poStatus.hasOwnProperty(s)) poStatus[s]++; else if (s.indexOf("Received") === 0) poStatus.Received++; });
  const topDemandItems = products.map(p => ({ name: p["Item Name (Standardized)"], demand: num(p["Deman for canteen"]) })).filter(p => p.demand > 0).sort((a, b) => b.demand - a.demand).slice(0, 5);
  return { totalStaff, presentToday: presentIds.size, absentToday: Math.max(totalStaff - presentIds.size, 0), totalProducts: products.length, lowStockCount: health.low, pendingDemands: demands.filter(d => d.Status === "Pending").length, pendingPOs: pos.filter(p => p.Status === "Pending").length, attendanceTrend: trend, categoryStock, lateToday: attendance.filter(a => String(a.Date || "").slice(0, 10) === today && a["IN Status"] === "Late").length, stockHealth: health, poStatus, topDemandItems };
}

// =================== DEMAND HANDLERS ===================
async function handleAddDemand(p) {
  const products = await cached("products", readProducts);
  const item = products.find(x => normH(x["Item Name (Standardized)"]) === normH(p.itemName));
  if (!item) throw new Error("Item not found: " + p.itemName);
  const id = "DMD-" + Date.now() + "-" + Math.floor(Math.random() * 90 + 10);
  const ts = toISTISO(new Date()); // 🆕 IST
  const rowNum = await appendRow(S.DEMAND, [id, ts, p.raisedBy || "", p.itemName, p.category || item.Category || "", p.unit || item.Unit || "", Number(p.qty) || 0, "Pending", "", p.notes || "", "", "", "", ""]);
  if (item._row) {
    const newDemand = num(item["Deman for canteen"]) + num(p.qty);
    await writeCell(S.PRODUCTS, item._row, 8, newDemand);
  }
  invalidate("demands", "products", "dash");
  return { id, row: rowNum };
}

async function handleApproveDemands(p) {
  const items = Array.isArray(p.items) ? p.items : [];
  if (!items.length) throw new Error("Koi demand select nahi hui.");
  const results = [];
  const autoPOs = [];
  for (const it of items) {
    try {
      const row = Number(it.row); if (!row || row < 2) throw new Error("Invalid row");
      const approved = Number(it.approvedQty);
      const feedback = (it.feedback || "").trim();
      const nowIso = toISTISO(new Date()); // 🆕 IST

      // Read item name before approving (for auto-PO check)
      const demRows = await readRange(S.DEMAND, `A${row}:N${row}`, "FORMATTED_VALUE");
      const demR = demRows[0] || [];
      const itemName = String(demR[3] || "").trim();

      await writeCell(S.DEMAND, row, 8, "Approved");
      await writeCell(S.DEMAND, row, 11, isNaN(approved) ? "" : approved);
      await writeCell(S.DEMAND, row, 12, nowIso);
      if (feedback) { await writeCell(S.DEMAND, row, 13, feedback); await writeCell(S.DEMAND, row, 14, nowIso); }

      // 🆕 AUTO-PO check after approval
      if (itemName) {
        try {
          // Bust products cache first so we read fresh stock
          invalidate("products");
          const poRes = await checkAndCreateAutoPO(itemName);
          if (poRes && poRes.created) autoPOs.push(poRes);
        } catch (e) { console.error("AutoPO on approve:", e.message); }
      }

      results.push({ row, ok: true, approvedQty: approved });
    } catch (e) { results.push({ row: it.row, ok: false, error: e.message }); }
  }
  invalidate("demands", "products", "dash");
  return { results, autoPOs };
}

async function handleRejectDemands(p) {
  const rows = Array.isArray(p.rows) ? p.rows.map(Number) : [];
  if (!rows.length) throw new Error("Koi demand select nahi hui.");
  const results = [];
  for (const row of rows) {
    try {
      if (!row || row < 2) throw new Error("Invalid row");
      const demRows = await readRange(S.DEMAND, `A${row}:N${row}`, "FORMATTED_VALUE");
      const r = demRows[0] || [];
      const status = r[7] || "Pending";
      if (status !== "Pending") throw new Error("Ye demand pehle hi " + status + " ho chuki hai.");
      const itemName = String(r[3] || "").trim();
      const qty = num(r[6]);
      if (itemName && qty > 0) {
        const products = await cached("products", readProducts);
        const item = products.find(x => normH(x["Item Name (Standardized)"]) === normH(itemName));
        if (item && item._row) {
          const currentDemand = num(item["Deman for canteen"]);
          const newDemand = Math.max(0, round3(currentDemand - qty));
          await writeCell(S.PRODUCTS, item._row, 8, newDemand);
        }
      }
      await writeCell(S.DEMAND, row, 8, "Rejected");
      results.push({ row, ok: true, subtracted: qty });
    } catch (e) { results.push({ row, ok: false, error: e.message }); }
  }
  invalidate("demands", "products", "dash");
  return { results };
}

async function handleEditDemand(p) {
  const row = Number(p.row);
  if (!row || row < 2) throw new Error("Invalid row");
  const newQty = Number(p.newQty);
  if (isNaN(newQty) || newQty <= 0) throw new Error("Invalid quantity");

  const demRows = await readRange(S.DEMAND, `A${row}:N${row}`, "FORMATTED_VALUE");
  const r = demRows[0] || [];
  const status = r[7] || "Pending";
  if (status !== "Pending") throw new Error("Ye demand pehle hi " + status + " ho chuki hai.");
  const itemName = String(r[3] || "").trim();
  const oldQty = num(r[6]);
  const diff = round3(newQty - oldQty);

  if (itemName && diff !== 0) {
    const products = await cached("products", readProducts);
    const item = products.find(x => normH(x["Item Name (Standardized)"]) === normH(itemName));
    if (item && item._row) {
      const currentDemand = num(item["Deman for canteen"]);
      const newDemand = Math.max(0, round3(currentDemand + diff));
      await writeCell(S.PRODUCTS, item._row, 8, newDemand);
    }
  }

  await writeCell(S.DEMAND, row, 7, newQty);
  invalidate("demands", "products", "dash");
  return { row, oldQty, newQty, diff };
}

async function handleDeleteDemand(p) {
  const row = Number(p.row);
  if (!row || row < 2) throw new Error("Invalid row");

  const demRows = await readRange(S.DEMAND, `A${row}:N${row}`, "FORMATTED_VALUE");
  const r = demRows[0] || [];
  const status = r[7] || "Pending";
  if (status !== "Pending") throw new Error("Ye demand pehle hi " + status + " ho chuki hai.");
  const itemName = String(r[3] || "").trim();
  const qty = num(r[6]);

  if (itemName && qty > 0) {
    const products = await cached("products", readProducts);
    const item = products.find(x => normH(x["Item Name (Standardized)"]) === normH(itemName));
    if (item && item._row) {
      const currentDemand = num(item["Deman for canteen"]);
      const newDemand = Math.max(0, round3(currentDemand - qty));
      await writeCell(S.PRODUCTS, item._row, 8, newDemand);
    }
  }

  await writeCell(S.DEMAND, row, 8, "Cancelled");
  invalidate("demands", "products", "dash");
  return { row, subtracted: qty };
}

async function handleSaveDemandFeedback(p) {
  const row = Number(p.row); if (!row || row < 2) throw new Error("Invalid row");
  const feedback = String(p.feedback || "").trim().slice(0, 500);
  if (!feedback) throw new Error("Feedback khali hai");
  const nowIso = toISTISO(new Date()); // 🆕 IST
  await writeCell(S.DEMAND, row, 13, feedback);
  await writeCell(S.DEMAND, row, 14, nowIso);
  invalidate("demands");
  return { row, feedback, feedbackAt: nowIso };
}

// =================== ATTENDANCE HANDLERS ===================
function businessDateKey(d) {
  const ist = new Date(new Date(d).getTime() + 5.5 * 3600000);
  const hour = ist.getUTCHours();
  if (hour < 3) ist.setUTCDate(ist.getUTCDate() - 1);
  return ist.toISOString().slice(0, 10);
}

async function handleMarkAttendance(p) {
  const now = new Date();
  const dateStr = businessDateKey(now);
  const type = p.type === "OUT" ? "OUT" : "IN";

  const rows = await readRange(S.ATTENDANCE, "A1:R2000", "FORMATTED_VALUE");
  if (rows.length < 1) throw new Error("Attendance sheet khaali");
  const headers = rows[0].map(h => String(h).trim());
  const ci = {}; headers.forEach((h, i) => { ci[h] = i; });
  let todayRow = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][ci["User ID"]] || "").trim() === String(p.userId).trim() && normalizeDate(rows[i][ci["Date"]]) === dateStr) { todayRow = i + 1; break; }
  }

  // 🆕 Pehle validate karo — rejected punch par photo upload nahi hogi
  if (type === "IN" && todayRow > 0 && rows[todayRow - 1][ci["IN Time"]]) throw new Error("Aap already IN punch kar chuke ho.");
  if (type === "OUT") {
    if (todayRow <= 0) throw new Error("Pehle IN punch karo.");
    if (rows[todayRow - 1][ci["OUT Time"]]) throw new Error("Aap already OUT punch kar chuke ho.");
  }

  // 🆕 Photo upload (Apps Script → Drive). Fail hone par error Render Logs mein dikhega.
  const photoUrl = await uploadPhotoToDrive(p.photoBase64, p.userId, p.name || p.userId);

  const mapsLink = (p.lat && p.lng) ? ("https://www.google.com/maps?q=" + p.lat + "," + p.lng) : "";
  // 🔧 FIX: Store IST wall-clock time
  const nowIso = toISTISO(now);
  let status;
  if (type === "IN") {
    status = computeInStatus(p.shift, now); // On Time / Late
    await appendRow(S.ATTENDANCE, [dateStr, p.userId, p.name || "", p.designation || "", p.shift || "", nowIso, photoUrl, mapsLink, p.lat || "", p.lng || "", status, "", "", "", "", "", "", "Present"]);
  } else {
    status = computeOutStatus(p.shift, now); // On Time / Early Leave / Overtime
    await writeCell(S.ATTENDANCE, todayRow, ci["OUT Time"] + 1, nowIso);
    await writeCell(S.ATTENDANCE, todayRow, ci["OUT Photo URL"] + 1, photoUrl);
    await writeCell(S.ATTENDANCE, todayRow, ci["OUT Maps Link"] + 1, mapsLink);
    await writeCell(S.ATTENDANCE, todayRow, ci["OUT Latitude"] + 1, p.lat || "");
    await writeCell(S.ATTENDANCE, todayRow, ci["OUT Longitude"] + 1, p.lng || "");
    await writeCell(S.ATTENDANCE, todayRow, ci["OUT Status"] + 1, status);
  }
  invalidate("attendance", "dash");
  return { status, type, photoUrl, istTime: nowIso };
}

async function handleEditAttendancePunch(p) {
  const rows = await readRange(S.ATTENDANCE, "A1:R2000", "FORMATTED_VALUE");
  const headers = rows[0].map(h => String(h).trim());
  const ci = {}; headers.forEach((h, i) => { ci[h] = i; });
  const note = "Manually corrected by " + (p.editedBy || "SuperAdmin");
  let ex = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][ci["User ID"]] || "").trim() === String(p.userId).trim() && normalizeDate(rows[i][ci["Date"]]) === p.dateISO) { ex = i + 1; break; }
  }
  if (p.inTime) {
    const ts = p.dateISO + "T" + p.inTime + ":00"; // already IST since typed manually
    const inMin = hmToMinutes(p.inTime);
    const inStatus = inMin === null ? "Marked" : computeInStatusAt(p.shift, inMin);
    if (ex > 0) {
      await writeCell(S.ATTENDANCE, ex, ci["IN Time"] + 1, ts);
      await writeCell(S.ATTENDANCE, ex, ci["IN Photo URL"] + 1, note);
      if (ci["IN Status"] !== undefined) await writeCell(S.ATTENDANCE, ex, ci["IN Status"] + 1, inStatus);
    } else {
      await appendRow(S.ATTENDANCE, [p.dateISO, p.userId, p.name, p.designation, p.shift, ts, note, "", "", "", inStatus, "", "", "", "", "", "", "Present"]);
    }
  }
  if (p.outTime) {
    let ex2 = ex;
    if (ex2 <= 0) {
      const r2 = await readRange(S.ATTENDANCE, "A1:R2000", "FORMATTED_VALUE");
      for (let i = 1; i < r2.length; i++) {
        if (String(r2[i][ci["User ID"]] || "").trim() === String(p.userId).trim() && normalizeDate(r2[i][ci["Date"]]) === p.dateISO) { ex2 = i + 1; break; }
      }
    }
    if (ex2 <= 0) throw new Error("Pehle IN set karo");
    const ts = p.dateISO + "T" + p.outTime + ":00";
    const outMin = hmToMinutes(p.outTime);
    const outStatus = outMin === null ? "Marked" : computeOutStatusAt(p.shift, outMin);
    await writeCell(S.ATTENDANCE, ex2, ci["OUT Time"] + 1, ts);
    await writeCell(S.ATTENDANCE, ex2, ci["OUT Photo URL"] + 1, note);
    if (ci["OUT Status"] !== undefined) await writeCell(S.ATTENDANCE, ex2, ci["OUT Status"] + 1, outStatus);
  }
  invalidate("attendance", "dash");
  return { ok: true };
}

// =================== WAREHOUSE HANDLERS ===================
async function handleAddWarehouseEntry(p) {
  const products = await cached("products", readProducts);
  const item = products.find(x => normH(x["Item Name (Standardized)"]) === normH(p.itemName || p.newName));
  if (!item) throw new Error("Item not found");
  const qty = Number(p.qty) || 0; if (qty <= 0) throw new Error("Invalid quantity");
  const newLoose = round3(num(item["W/H Opening Loose"]) + qty);
  await writeCell(S.PRODUCTS, item._row, 7, newLoose);
  if (p.rate) await writeCell(S.PRODUCTS, item._row, 17, Number(p.rate));
  if (p.expiryDate) await writeCell(S.PRODUCTS, item._row, 19, p.expiryDate);
  const logRow = await appendRow(S.WHLOG, ["WH-" + Date.now(), toISTISO(new Date()), item["Item Name (Standardized)"], "add", qty, item["Current Stock"], round3(num(item["Current Stock"]) + qty), p.enteredBy || "", ""]);
  invalidate("products", "whlog", "dash");
  return { itemName: item["Item Name (Standardized)"], newValue: round3(num(item["Current Stock"]) + qty), row: logRow };
}
async function handleEditWarehouseEntry(p) {
  const row = Number(p.logRow); if (!row || row < 2) throw new Error("Invalid row");
  await writeCell(S.WHLOG, row, 5, Number(p.newQty) || 0);
  await writeCell(S.WHLOG, row, 9, toISTISO(new Date()));
  invalidate("whlog", "products", "dash");
  return { row };
}
async function handleDeleteWarehouseEntry(p) {
  const row = Number(p.logRow); if (!row || row < 2) throw new Error("Invalid row");
  await writeCell(S.WHLOG, row, 8, "Cancelled");
  invalidate("whlog", "products", "dash");
  return { row };
}

// =================== CYLINDER HANDLERS ===================
async function handleAddCylinderEntry(p) {
  const qty = Number(p.qty) || 0; if (qty <= 0) throw new Error("Invalid quantity");
  const price = Number(await getSetting("CylinderPrice", "2900")) || 2900;
  const total = qty * price;
  // userId "cyl_..." se Apps Script photo ko "Peetambra Cylinder Photos" folder mein rakhta hai
  const photoUrl = await uploadPhotoToDrive(p.photoBase64, "cyl_" + Date.now(), "cyl");
  await appendRow(S.CYLINDER, ["CYL-" + Date.now(), toISTISO(new Date()), p.receivedBy || "", qty, price, total, photoUrl]);
  invalidate("cyl", "dash");
  return { qty, price, total, photoUrl };
}

// =================== PO HANDLERS ===================
async function handleAddPurchaseOrder(p) {
  await appendRow(S.PO, ["PO-" + Date.now(), toISTISO(new Date()), p.itemName, p.category || "", p.unit || "", Number(p.qty) || 0, "Pending", p.requestedBy || "", p.notes || ""]);
  invalidate("po", "po_batches");
  return { id: "PO-" + Date.now() };
}
async function handleCancelPO(p) {
  const row = Number(p.row); if (!row || row < 2) throw new Error("Invalid row");
  await writeCell(S.PO, row, 7, "Cancelled");
  invalidate("po", "po_batches");
  return { row };
}
async function handleClearPOBatch(p) {
  const rows = await readRange(S.PO, "A1:ZZ2000", "FORMATTED_VALUE");
  const headers = rows[0].map(h => String(h).trim());
  const ci = {}; headers.forEach((h, i) => { ci[h] = i; });
  let cleared = 0;
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if ((r[ci["Vendor"]] || "Unassigned") === p.vendor && (r[ci["Batch ID"]] || "Legacy") === p.batchId && r[ci["Status"]] === "Generated") {
      await writeCell(S.PO, i + 1, ci["Status"] + 1, "Cleared"); cleared++;
    }
  }
  invalidate("po", "po_batches");
  return { cleared };
}
async function handleGeneratePOBatch(p) {
  const rows = await readRange(S.PO, "A1:ZZ2000", "FORMATTED_VALUE");
  const headers = rows[0].map(h => String(h).trim());
  const ci = {}; headers.forEach((h, i) => { ci[h] = i; });
  const selected = Array.isArray(p.rows) ? p.rows.map(Number) : [];
  let itemCount = 0, remainingPending = 0;
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const matchVendor = (r[ci["Vendor"]] || "Unassigned") === p.vendor;
    const matchBatch = (r[ci["Batch ID"]] || "Legacy") === p.batchId;
    const isPending = (r[ci["Status"]] === "Pending");
    if (matchVendor && matchBatch && isPending) {
      if (selected.length === 0 || selected.indexOf(i + 1) !== -1) {
        await writeCell(S.PO, i + 1, ci["Status"] + 1, "Generated"); itemCount++;
      } else remainingPending++;
    }
  }
  invalidate("po", "po_batches");
  return { itemCount, remainingPending, pdfUrl: "" };
}
async function handleAutoTagVendors() {
  const products = await cached("products", readProducts);
  let tagged = 0;
  for (const p of products) {
    if (String(p.Vendor || "").trim()) continue;
    const m = matchVendorForItem(p["Item Name (Standardized)"]);
    if (m) { await writeCell(S.PRODUCTS, p._row, 16, m); tagged++; }
  }
  invalidate("products");
  return { tagged, total: products.length };
}

// =================== MANUAL PURCHASE HANDLERS ===================
async function handleAddManualPurchase(p) {
  const products = await cached("products", readProducts);
  const item = products.find(x => normH(x["Item Name (Standardized)"]) === normH(p.itemName || p.newName));
  if (!item) throw new Error("Item not found");
  const qty = Number(p.qty) || 0; if (qty <= 0) throw new Error("Invalid quantity");
  const newPurchase = round3(num(item["Purchase Stock  QTy"]) + qty);
  await writeCell(S.PRODUCTS, item._row, 11, newPurchase);
  if (p.rate) await writeCell(S.PRODUCTS, item._row, 17, Number(p.rate));
  if (p.expiryDate) await writeCell(S.PRODUCTS, item._row, 19, p.expiryDate);
  await appendRow(S.PO, ["MPO-" + Date.now(), toISTISO(new Date()), item["Item Name (Standardized)"], p.category || item.Category || "", p.unit || item.Unit || "", qty, "Received (Manual)", p.enteredBy || "", "Manual purchase entry"]);
  invalidate("products", "po", "dash", "mpo");
  return { itemName: item["Item Name (Standardized)"], newValue: round3(num(item["Current Stock"]) + qty) };
}
async function handleEditManualPurchase(p) {
  const row = Number(p.row); if (!row || row < 2) throw new Error("Invalid row");
  await writeCell(S.PO, row, 6, Number(p.newQty) || 0);
  invalidate("po", "mpo", "products", "dash");
  return { row };
}
async function handleDeleteManualPurchase(p) {
  const row = Number(p.row); if (!row || row < 2) throw new Error("Invalid row");
  await writeCell(S.PO, row, 7, "Cancelled");
  invalidate("po", "mpo", "products", "dash");
  return { row };
}

// =================== STAFF HANDLERS ===================
async function handleAddStaff(p) {
  const staff = await cached("staff", readStaff);
  const userId = p.userId || (String(p.name).replace(/\s+/g, "") + "-" + String(staff.length + 1).padStart(3, "0"));
  await appendRow(S.STAFF, [p.name, p.designation, p.shift, userId, p.permission, p.password || "123"]);
  invalidate("staff", "dash");
  return { userId };
}
async function handleUpdateStaff(p) {
  const row = Number(p.row); if (!row || row < 2) throw new Error("Invalid row");
  await writeCell(S.STAFF, row, 1, p.name);
  await writeCell(S.STAFF, row, 2, p.designation);
  await writeCell(S.STAFF, row, 3, p.shift);
  await writeCell(S.STAFF, row, 4, p.userId);
  await writeCell(S.STAFF, row, 5, p.permission);
  await writeCell(S.STAFF, row, 6, p.password);
  invalidate("staff", "dash");
  return { row };
}
async function handleDeleteStaff(p) {
  const row = Number(p.row); if (!row || row < 2) throw new Error("Invalid row");
  try {
    const sh = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
    const sheetId = sh.data.sheets.find(s => s.properties.title === S.STAFF)?.properties.sheetId;
    if (sheetId !== undefined) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SHEET_ID,
        requestBody: { requests: [{ deleteDimension: { range: { sheetId, dimension: "ROWS", startIndex: row - 1, endIndex: row } } }] },
      });
    }
  } catch (e) {}
  invalidate("staff", "dash");
  return { deleted: row };
}
async function handleAddProduct(p) {
  const products = await cached("products", readProducts);
  if (products.find(x => normH(x["Item Name (Standardized)"]) === normH(p.name))) throw new Error("Item already exists");
  const suq = Number(p.stockUnitQty) || 1;
  const rowNum = await appendRow(S.PRODUCTS, [products.length + 1, p.name, p.category || "", p.packaging || "", p.unit || "", 0, 0, 0, p.demandUnit || p.unit || "", 0, 0, "", "", Number(p.minStock) || 0, Number(p.maxStock) || 0, "", "", p.unit || "", "", p.baseUnit || p.unit || "", suq, p.stockType || "Stocked"]);
  invalidate("products", "dash");
  return { added: p.name, row: rowNum };
}
async function handleAddVendor(p) {
  const name = String(p.name || "").trim(); if (!name) throw new Error("Vendor name khali hai");
  const extra = await getSetting("ExtraVendors", "");
  let list = []; if (extra) { try { list = JSON.parse(extra); } catch (e) {} }
  if (!list.some(v => String(v).toLowerCase() === name.toLowerCase())) { list.push(name); await setSetting("ExtraVendors", JSON.stringify(list)); }
  invalidate("vendors");
  return { name, added: true };
}
async function handleSetProductVendor(p) {
  const products = await cached("products", readProducts);
  const item = products.find(x => normH(x["Item Name (Standardized)"]) === normH(p.itemName));
  if (!item) throw new Error("Item not found");
  await writeCell(S.PRODUCTS, item._row, 16, p.vendor || "");
  invalidate("products");
  return { itemName: p.itemName, vendor: p.vendor };
}

// =================== PO BATCHES ===================
async function readPOBatches() {
  const pos = (await cached("po", readPO)).filter(po => po.Status !== "Cancelled");
  const batches = {};
  pos.forEach(po => {
    const vendor = po.Vendor || "Unassigned"; const batchId = po["Batch ID"] || "Legacy";
    const key = vendor + "|" + batchId;
    if (!batches[key]) batches[key] = { vendor, batchId, items: [], pdfUrl: "" };
    batches[key].items.push({ row: po._row, name: po["Item Name"], unit: po.Unit, qty: po["Qty Ordered"], status: po.Status, timestamp: po.Timestamp });
    if (po["PDF URL"]) batches[key].pdfUrl = po["PDF URL"];
  });
  const list = Object.values(batches).filter(b => b.items.length);
  list.forEach(b => {
    if (b.items.every(i => i.status === "Cleared")) b.status = "Cleared";
    else if (b.items.some(i => i.status === "Generated")) b.status = "Generated";
    else b.status = "Pending";
  });
  const order = { Pending: 0, Generated: 1, Cleared: 2 };
  return list.sort((a, b) => (order[a.status] - order[b.status]));
}

// =================== ROUTER ===================
app.all("/", async (req, res) => {
  const payload = { ...req.query, ...req.body };
  const action = payload.action;
  if (!action) return res.json({ ok: true, service: "Peetambra Bridge Running" });
  try {
    let data;
    if (action === "login") data = await handleLogin(payload);
    else if (action === "bootstrap") data = await handleBootstrap(payload);
    else if (action === "poll") data = await buildPoll();
    else if (action === "getProducts") data = await cached("products", readProducts);
    else if (action === "getCategories") data = [...new Set((await cached("products", readProducts)).map(p => p.Category).filter(Boolean))].sort();
    else if (action === "getStaff") data = await cached("staff", readStaff);
    else if (action === "getVendors") data = await cached("vendors", readVendors);
    else if (action === "getDemands") data = await cached("demands", readDemands);
    else if (action === "getDashboardStats") data = await buildDashboardStats();
    else if (action === "getLowStockAlerts") {
      const products = await cached("products", readProducts);
      // 🔧 FIX: Compare against Current Stock (pcs), not carton
      data = products.filter(p => {
        const min = num(p.Min), max = num(p.Max), cur = num(p["Current Stock"]);
        return (min > 0 && cur <= min) || (max > 0 && cur >= max);
      }).map(p => ({
        "Item Name": p["Item Name (Standardized)"],
        Category: p.Category,
        Unit: p.Unit,
        "Current Stock": p["Current Stock"],
        "Min Stock": p.Min,
        "Max Stock": p.Max,
        "Status": (num(p.Min) > 0 && num(p["Current Stock"]) <= num(p.Min)) ? "LOW" : "OVERSTOCK"
      }));
    }
    else if (action === "getWarehouseEntries") data = await cached("whlog", readWarehouseEntries);
    else if (action === "getCylinderEntries") data = await cached("cyl", readCylinderEntries);
    else if (action === "getCylinderPrice") data = Number(await getSetting("CylinderPrice", "2900")) || 2900;
    else if (action === "getPOBatches") data = await cached("po_batches", readPOBatches);
    else if (action === "getPurchaseOrders") data = await cached("po", readPO);
    else if (action === "getManualPurchases") data = await cached("mpo", readManualPurchases);
    else if (action === "getAttendance") {
      let att = await cached("attendance", readAttendance);
      if (payload.userId) att = att.filter(r => String(r["User ID"]).toLowerCase().trim() === String(payload.userId).toLowerCase().trim());
      data = att;
    }
    else if (action === "addProduct") data = await handleAddProduct(payload);
    else if (action === "addWarehouseEntry") data = await handleAddWarehouseEntry(payload);
    else if (action === "editWarehouseEntry") data = await handleEditWarehouseEntry(payload);
    else if (action === "deleteWarehouseEntry") data = await handleDeleteWarehouseEntry(payload);
    else if (action === "addDemand") data = await handleAddDemand(payload);
    else if (action === "approveDemands") data = await handleApproveDemands(payload);
    else if (action === "approveDemand") data = await handleApproveDemands({ items: [{ row: payload.row, approvedQty: payload.approvedQty, feedback: payload.feedback }] });
    else if (action === "rejectDemands") data = await handleRejectDemands(payload);
    else if (action === "rejectDemand") data = await handleRejectDemands({ rows: [payload.row] });
    else if (action === "editDemand") data = await handleEditDemand(payload);
    else if (action === "deleteDemand") data = await handleDeleteDemand(payload);
    else if (action === "saveDemandFeedback") data = await handleSaveDemandFeedback(payload);
    else if (action === "addPurchaseOrder") data = await handleAddPurchaseOrder(payload);
    else if (action === "cancelPO") data = await handleCancelPO(payload);
    else if (action === "clearPOBatch") data = await handleClearPOBatch(payload);
    else if (action === "generatePOBatch") data = await handleGeneratePOBatch(payload);
    else if (action === "autoTagVendors") data = await handleAutoTagVendors();
    else if (action === "addManualPurchase") data = await handleAddManualPurchase(payload);
    else if (action === "editManualPurchase") data = await handleEditManualPurchase(payload);
    else if (action === "deleteManualPurchase") data = await handleDeleteManualPurchase(payload);
    else if (action === "addVendor") data = await handleAddVendor(payload);
    else if (action === "setProductVendor") data = await handleSetProductVendor(payload);
    else if (action === "markAttendance") data = await handleMarkAttendance(payload);
    else if (action === "editAttendancePunch") data = await handleEditAttendancePunch(payload);
    else if (action === "addCylinderEntry") data = await handleAddCylinderEntry(payload);
    else if (action === "updateCylinderPrice") { const price = Number(payload.price) || 0; if (price <= 0) throw new Error("Invalid price"); await setSetting("CylinderPrice", price); invalidate("cyl"); data = { price }; }
    else if (action === "addStaff") data = await handleAddStaff(payload);
    else if (action === "updateStaff") data = await handleUpdateStaff(payload);
    else if (action === "deleteStaff") data = await handleDeleteStaff(payload);
    // 🆕 AUTO-PO actions
    else if (action === "runAutoPO") data = await runAutoPOForAll();
    else if (action === "checkAutoPOForItem") data = await checkAndCreateAutoPO(payload.itemName);
    else if (action === "uploadPhoto") {
      const url = await uploadPhotoToDrive(payload.photoBase64, payload.userId || "photo", payload.name || "photo");
      data = { url };
    }
    else data = { note: "Action not implemented yet: " + action };
    res.json({ ok: true, data });
  } catch (e) {
    console.error("Router error on", action, ":", e.message);
    res.json({ ok: false, error: e.message });
  }
});

app.get("/", (req, res) => res.json({ ok: true, service: "Peetambra Bridge Running" }));

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log("🚀 Bridge running on port " + PORT));
