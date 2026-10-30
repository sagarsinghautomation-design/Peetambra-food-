const express = require("express");
const { google } = require("googleapis");
const NodeCache = require("node-cache");
const cors = require("cors");

const app = express();
app.use(cors());
app.use(express.urlencoded({ extended: true }));
app.use(express.json({ limit: "20mb", type: ["application/json", "text/plain"] }));

// useClones:false — cache se data turant milta hai (copy banane ka time bachta hai)
const cache = new NodeCache({ stdTTL: 120, checkperiod: 30, useClones: false });

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
  CYLINDER: "Cylinder Entry", SETTINGS: "Settings", RATEHIST: "Rate History", MSG: "Messages", COMPLAINT: "Complaints",
};

// =================== HELPERS ===================
// num(): "₹1,246" / "1,200" / " 45 " sab ko sahi number banata hai; "NA"/khali → 0
function num(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return isNaN(v) ? 0 : v;
  const n = Number(String(v).replace(/[₹,\s]/g, ""));
  return isNaN(n) ? 0 : n;
}
// money(): rate ke liye — "NA"/khali → "" (blank), warna number
function money(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "number") return isNaN(v) ? "" : v;
  const s = String(v).replace(/[₹,\s]/g, "");
  if (!s) return "";
  const n = Number(s);
  return isNaN(n) ? "" : n;
}
function round3(v) { return Math.round(v * 1000) / 1000; }
function round2(v) { return Math.round(v * 100) / 100; }
function normH(s) { return String(s === undefined || s === null ? "" : s).toLowerCase().replace(/\s+/g, " ").trim(); }
function pad2(n) { return String(n).padStart(2, "0"); }
function toISODate(d) { return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate()); }
function isDaily(v) { return /daily/i.test(String(v || "")); }

// IST helpers — Indian Standard Time = UTC + 5:30
function toISTISO(d) {
  const dt = (d instanceof Date) ? d : new Date(d);
  if (isNaN(dt.getTime())) return "";
  const ist = new Date(dt.getTime() + 5.5 * 3600000);
  return ist.toISOString().slice(0, 19);
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
    return d.getUTCFullYear() + "-" + pad2(d.getUTCMonth() + 1) + "-" + pad2(d.getUTCDate());
  }
  if (!isNaN(n)) return "";
  const d = new Date(s);
  if (!isNaN(d.getTime())) return toISODate(d);
  return "";
}

function normalizeDateTime(v, dateHint) {
  if (v === null || v === undefined || v === "") return "";
  if (v instanceof Date) return toISTISO(v);
  const s = String(v).trim(); if (!s) return "";
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?Z$/i.test(s)) {
    const d = new Date(s);
    if (!isNaN(d.getTime())) return toISTISO(d);
  }
  if (/^\d{4}-\d{2}-\d{2}T/.test(s)) return s.slice(0, 19);
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) { const d = new Date(s); if (!isNaN(d.getTime())) return d.toISOString().slice(0, 19); }
  const n = Number(s);
  if (!isNaN(n) && n > 20000 && n < 60000) { const ms = Date.UTC(1899, 11, 30) + Math.round(n * 86400000); return new Date(ms).toISOString().slice(0, 19); }
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

async function readRange(sheetName, range = "A1:ZZ", renderOption = "UNFORMATTED_VALUE") {
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

// 🆕 Ek hi API call mein kai cells likhna — sequential writeCell se 5-10x tez
async function batchWrite(sheetName, cells) {
  const list = (cells || []).filter(c => c && c.row && c.col);
  if (!list.length) return;
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: {
      valueInputOption: "USER_ENTERED",
      data: list.map(c => ({ range: `'${sheetName}'!${colLetter(c.col)}${c.row}`, values: [[c.value === undefined || c.value === null ? "" : c.value]] })),
    },
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

// =================== SHEET META / HEADERS ===================
let META = null;
async function sheetMeta(force) {
  if (META && !force) return META;
  const r = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID, fields: "sheets.properties(sheetId,title,gridProperties)" });
  META = {};
  (r.data.sheets || []).forEach(s => {
    META[s.properties.title] = { sheetId: s.properties.sheetId, cols: (s.properties.gridProperties || {}).columnCount || 26 };
  });
  return META;
}
async function ensureSheet(name, headers) {
  const m = await sheetMeta();
  if (m[name]) return;
  await sheets.spreadsheets.batchUpdate({ spreadsheetId: SHEET_ID, requestBody: { requests: [{ addSheet: { properties: { title: name } } }] } });
  await sheetMeta(true);
  if (headers && headers.length) await batchWrite(name, headers.map((h, i) => ({ row: 1, col: i + 1, value: h })));
}
async function ensureCols(name, n) {
  const m = await sheetMeta();
  const s = m[name];
  if (!s || s.cols >= n) return;
  await sheets.spreadsheets.batchUpdate({ spreadsheetId: SHEET_ID, requestBody: { requests: [{ appendDimension: { sheetId: s.sheetId, dimension: "COLUMNS", length: n - s.cols } }] } });
  s.cols = n;
}
// Header naam se column dhoondhta hai; na mile to sheet ke end mein header bana deta hai.
// Isse kisi bhi sheet mein column aage-peeche ho jaye to bhi sahi column mein hi likha jayega.
const HDR_CACHE = {};
async function ensureHeaders(sheetName, wanted, aliases) {
  aliases = aliases || {};
  let hdr = HDR_CACHE[sheetName];
  if (!hdr) hdr = ((await readRange(sheetName, "1:1", "FORMATTED_VALUE"))[0] || []).map(h => String(h).trim());
  const map = {}, add = [];
  wanted.forEach(w => {
    let idx = hdr.findIndex(h => normH(h) === normH(w));
    if (idx < 0 && aliases[w]) idx = hdr.findIndex(h => aliases[w].test(h));
    if (idx < 0) { const j = add.indexOf(w); idx = j >= 0 ? hdr.length + j : hdr.length + add.push(w) - 1; }
    map[w] = idx + 1;
  });
  if (add.length) {
    await ensureCols(sheetName, hdr.length + add.length);
    await batchWrite(sheetName, add.map((w, i) => ({ row: 1, col: hdr.length + i + 1, value: w })));
    hdr = hdr.concat(add);
  }
  HDR_CACHE[sheetName] = hdr;
  return map;
}
async function appendByHeaders(sheetName, obj, aliases) {
  const map = await ensureHeaders(sheetName, Object.keys(obj), aliases);
  const len = Math.max(...Object.values(map));
  const arr = new Array(len).fill("");
  Object.keys(obj).forEach(k => { arr[map[k] - 1] = obj[k] === undefined || obj[k] === null ? "" : obj[k]; });
  return appendRow(sheetName, arr);
}
async function writeByHeaders(sheetName, row, obj, aliases) {
  const map = await ensureHeaders(sheetName, Object.keys(obj), aliases);
  await batchWrite(sheetName, Object.keys(obj).map(k => ({ row, col: map[k], value: obj[k] })));
}
async function readRowByHeaders(sheetName, row, wanted, aliases) {
  const map = await ensureHeaders(sheetName, wanted, aliases);
  const r = (await readRange(sheetName, `A${row}:ZZ${row}`, "UNFORMATTED_VALUE"))[0] || [];
  const o = {};
  wanted.forEach(k => { o[k] = r[map[k] - 1]; });
  return o;
}

// =================== CACHE ===================
const TTL = { complaints: 10, messages: 10, complaints: 10, demands: 15, attendance: 20, products: 20, po: 20, po_batches: 20, mpo: 20, whlog: 20, cyl: 20, staff: 60, vendors: 60, settings: 60, ratehist: 60 };
let DATA_VER = 1; // har write par badhta hai — frontend isse turant refresh karta hai
async function cached(key, builder) {
  const hit = cache.get(key); if (hit !== undefined) return hit;
  const data = await builder(); cache.set(key, data, TTL[key] || 30); return data;
}
function invalidate(...keys) { DATA_VER++; keys.forEach(k => cache.del(k)); }
function hashStr(str) { let h = 0; for (let i = 0; i < str.length; i++) { h = (h * 31 + str.charCodeAt(i)) | 0; } return (h >>> 0).toString(36); }
function parseIST(v) {
  const s = String(v || "").trim(); if (!s) return NaN;
  if (/Z$|[+-]\d{2}:\d{2}$/.test(s)) return Date.parse(s);
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(s)) return Date.parse(s.slice(0, 19) + "+05:30");
  return Date.parse(s);
}
const DEMAND_EDIT_WINDOW_MS = 60 * 60 * 1000;          // edit/delete: 1 ghanta
const DEMAND_FEEDBACK_WINDOW_MS = 3 * 24 * 3600 * 1000; // feedback: 3 din
const ENTRY_EDIT_WINDOW_MS = 60 * 60 * 1000;           // warehouse/purchase entry edit: 1 ghanta
const GRACE_MS = 2 * 60 * 1000;

// =================== 🆕 LOCK — ek saath do request ek hi item ko na bigaadein ===================
// Product Master ke saare read→calculate→write isi lock mein chalte hain, isliye do log ek saath
// approve/entry karein to bhi stock ka hisaab kabhi galat nahi hota.
const LOCKS = new Map();
function withLock(key, fn) {
  const prev = LOCKS.get(key) || Promise.resolve();
  const run = prev.then(() => fn(), () => fn());
  const tail = run.then(() => {}, () => {});
  LOCKS.set(key, tail);
  tail.then(() => { if (LOCKS.get(key) === tail) LOCKS.delete(key); });
  return run;
}

// =================== DUPLICATE-REQUEST GUARD (_rid) ===================
const RID_TTL_MS = 10 * 60 * 1000;
const ridStore = new Map();
setInterval(() => {
  const now = Date.now();
  for (const [k, v] of ridStore) if (now - v.t > RID_TTL_MS) ridStore.delete(k);
}, 60 * 1000).unref();

// =================== PHOTO UPLOAD (Apps Script → Google Drive) ===================
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

// =================== SHIFT STATUS ===================
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
    if (end <= start) end += 1440;      // raat wali shift (jaise 6 PM - 2 AM)
    if (n < start) n += 1440;           // shift shuru hone se pehle ka OUT = aadhi raat ke baad (agla din)
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

// =================== ROLES ===================
async function staffOf(p) {
  const uid = String(p._uid || "").toLowerCase().trim();
  if (!uid) return null;
  const staff = await cached("staff", readStaff);
  return staff.find(r => String(r.User || "").toLowerCase().trim() === uid) || null;
}
async function roleOf(p) { const s = await staffOf(p); return s ? String(s.Permission || "").trim() : ""; }
async function requireRole(p, roles, what) {
  const r = await roleOf(p);
  if (!roles.includes(r)) throw new Error((what || "Ye kaam") + " sirf " + roles.join(" / ") + " kar sakta hai.");
  return r;
}

// =================== PRODUCT MASTER — COLUMN MAP (header se) ===================
// Sheet A:X — S no | Item Names | Category | Packaging | Unit | W/H Opening Full Pack | W/H Opening Loose |
// Deman for canteen QTY | Demand Unit | Approved Demand Qty by supervisor | Approved QTY (Unit) |
// Purchase Stock QTy | Purchase Stock QTy (Unit) | Current Stock In Carton (FORMULA) | Current Stock in pcs (FORMULA) |
// Min | Max | Vendor | Rate | Unit of rate | Expiry Date | Base Unit | 1 Stock Unit = (Base Qty) | Stock Type
// Column ki position header ke naam se nikalti hai — aage koi column badhe to bhi galat column mein nahi likhega.
let PCOL = null;
function buildProdCols(hdrRow) {
  const H = (hdrRow || []).map(normH);
  const used = new Set();
  const find = (pred, def) => {
    for (let i = 0; i < H.length; i++) if (H[i] && !used.has(i) && pred(H[i])) { used.add(i); return i; }
    return def;
  };
  const C = {};
  // specific pehle, general baad mein (order zaroori hai)
  C.CURCARTON = find(x => x.includes("current stock") && x.includes("carton"), 13);
  C.CURPCS = find(x => x.includes("current stock"), 14);
  C.APPROVEDUNIT = find(x => x.includes("approved") && x.includes("(unit)"), 10);
  C.PURCHASEUNIT = find(x => x.includes("purchase") && x.includes("(unit)"), 12);
  C.APPROVED = find(x => x.includes("approved"), 9);
  C.PURCHASE = find(x => x.includes("purchase"), 11);
  C.DEMANDUNIT = find(x => x.startsWith("demand unit"), 8);
  C.DEMAND = find(x => x.startsWith("deman"), 7);
  C.RATEUNIT = find(x => x.startsWith("unit of rate"), 19);
  C.RATE = find(x => x === "rate" || x.startsWith("rate "), 18);
  C.SUQ = find(x => x.startsWith("1 stock unit") || x.includes("base qty"), 22);
  C.BASEUNIT = find(x => x.startsWith("base unit"), 21);
  C.STOCKTYPE = find(x => x.startsWith("stock type"), 23);
  C.UNIT = find(x => x === "unit", 4);
  C.NAME = find(x => x.startsWith("item name"), 1);
  C.SNO = find(x => /^s\.?\s*no\b/.test(x), 0);
  C.CAT = find(x => x.startsWith("category"), 2);
  C.PACK = find(x => x.startsWith("packaging"), 3);
  C.OPENFULL = find(x => x.includes("full pack"), 5);
  C.OPENLOOSE = find(x => x.includes("loose"), 6);
  C.MIN = find(x => x === "min" || x.startsWith("min "), 15);
  C.MAX = find(x => x === "max" || x.startsWith("max "), 16);
  C.VENDOR = find(x => x.startsWith("vendor"), 17);
  C.EXPIRY = find(x => x.startsWith("expiry"), 20);
  return C;
}
// 117 pcs, 1 pkt = 20 → "5 pkt + 17 pcs" (sheet mein N column "5.17")
function packText(cur, suq, unit, baseUnit) {
  if (suq <= 1 || normH(unit) === normH(baseUnit)) return `${cur} ${unit}`;
  const neg = cur < 0, a = Math.abs(cur), full = Math.floor(a / suq + 1e-9), loose = round3(a - full * suq);
  return `${neg ? "-" : ""}${full} ${baseUnit} + ${loose} ${unit}`;
}
function nb(v) { return (v === "" || v === undefined || v === null) ? "" : num(v); }

// Ek row se product object — Current Stock hamesha yahin calculate hota hai:
// Current (Unit/pcs) = Full Pack × (1 Base = qty) + Loose + Purchase − Approved   (Daily Market: Approved minus nahi hota)
function rowToProduct(r, rowNum, C) {
  const g = (i) => (r && i !== undefined && i >= 0) ? r[i] : undefined;
  const name = String(g(C.NAME) === undefined ? "" : g(C.NAME)).trim();
  if (!name) return null;
  const unit = String(g(C.UNIT) || "").trim();
  const baseUnit = String(g(C.BASEUNIT) || "").trim() || unit;
  const suq = num(g(C.SUQ)) || 1;
  const openFull = num(g(C.OPENFULL)), openLoose = num(g(C.OPENLOOSE));
  const purchase = num(g(C.PURCHASE)), approved = num(g(C.APPROVED));
  const stockType = String(g(C.STOCKTYPE) || "").trim() || "Stocked";
  const daily = isDaily(stockType);
  const opening = round3(openFull * suq + openLoose);
  const cur = round3(opening + purchase - (daily ? 0 : approved));
  return {
    _row: rowNum, "S no": g(C.SNO) === undefined ? "" : g(C.SNO),
    "Item Name (Standardized)": name,
    Category: String(g(C.CAT) || "").trim(), Unit: unit, Packaging: g(C.PACK) === undefined ? "" : g(C.PACK),
    "W/H Opening Full Pack": openFull, "W/H Opening Loose": openLoose,
    "Warehouse Stock Qty Opening STock": opening,
    "Deman for canteen": num(g(C.DEMAND)), "Demand Unit": String(g(C.DEMANDUNIT) || "").trim(),
    "Approved Demand Qty by supervisor": approved, "Approved QTY (Unit)": String(g(C.APPROVEDUNIT) || "").trim(),
    "Purchase Stock  QTy": purchase, "Purchase Stock QTy (Unit)": String(g(C.PURCHASEUNIT) || "").trim(),
    "Current Stock": cur, "Current Stock In Carton": round3(cur / suq), "Current Stock Pack": packText(cur, suq, unit, baseUnit),
    Min: nb(g(C.MIN)), Max: nb(g(C.MAX)), Vendor: String(g(C.VENDOR) || "").trim(),
    Rate: money(g(C.RATE)), "Unit of rate": String(g(C.RATEUNIT) || "").trim(),
    "Expiry Date": normalizeDate(g(C.EXPIRY)),
    "Base Unit": baseUnit, "1 Stock Unit = (Base Qty)": suq,
    "Stock Type": daily ? "Daily Market" : "Stocked",
  };
}

// =================== READERS ===================
async function readStaff() { return rowsToObjects(await readRange(S.STAFF)); }

async function readDemands() {
  const rows = await readRange(S.DEMAND, "A1:Q", "UNFORMATTED_VALUE");
  return rows.slice(1).reverse().map((r, i) => ({
    _row: rows.length - i, "Demand ID": r[0] || "",
    Timestamp: normalizeDateTime(r[1]),
    "Raised By": r[2] || "", "Item Name": r[3] || "", Category: r[4] || "",
    Unit: r[5] || "", "Qty Requested": r[6] || 0, Status: r[7] || "Pending",
    "Warehouse Stock After": r[8] === undefined ? "" : r[8], Notes: r[9] || "",
    "Approved Qty": r[10] === undefined ? "" : r[10], "Approved At": normalizeDateTime(r[11]),
    Feedback: r[12] || "", "Feedback At": normalizeDateTime(r[13]),
    "Approved By": r[14] || "",
    Rate: money(r[15]), Amount: money(r[16]),
  })).filter(d => d["Demand ID"] || d["Item Name"]);
}

async function readProducts() {
  const rows = await readRange(S.PRODUCTS, "A1:ZZ", "UNFORMATTED_VALUE");
  if (!rows.length) return [];
  PCOL = buildProdCols(rows[0]);
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const p = rowToProduct(rows[i], i + 1, PCOL);
    if (p) out.push(p);
  }
  return out;
}
async function getProdCols() {
  if (!PCOL) { cache.del("products"); await cached("products", readProducts); }
  return PCOL;
}
async function readProductRowRaw(row) {
  return (await readRange(S.PRODUCTS, `A${row}:ZZ${row}`, "UNFORMATTED_VALUE"))[0] || [];
}
async function findProduct(name) {
  const n = normH(name); if (!n) return null;
  let list = await cached("products", readProducts);
  let it = list.find(x => normH(x["Item Name (Standardized)"]) === n);
  if (!it) { cache.del("products"); list = await cached("products", readProducts); it = list.find(x => normH(x["Item Name (Standardized)"]) === n); }
  return it || null;
}
function patchProductCache(prod) {
  const list = cache.get("products");
  if (!list || !prod) return;
  let found = false;
  const next = list.map(x => { if (x._row === prod._row) { found = true; return prod; } return x; });
  if (!found) next.push(prod);
  cache.set("products", next, TTL.products);
}

// 🆕 Product Master ki kisi bhi row ko sahi tarike se badalna:
// lock → us row ko sheet se FRESH padho → naya value nikaalo → ek hi call mein likho → cache update.
// Current Stock (carton/pcs) ke formula columns mein kabhi nahi likhta.
async function mutateProduct(ref, fn) {
  return withLock("products", async () => {
    const C = await getProdCols();
    let item = null;
    if (ref.row) {
      item = (await cached("products", readProducts)).find(x => x._row === Number(ref.row));
      if (!item) { cache.del("products"); item = (await cached("products", readProducts)).find(x => x._row === Number(ref.row)); }
    } else item = await findProduct(ref.name);
    if (!item) throw new Error("Item nahi mila: " + (ref.name || ("row " + ref.row)));

    let raw = await readProductRowRaw(item._row);
    if (normH(raw[C.NAME]) !== normH(item["Item Name (Standardized)"])) {
      // Sheet mein kisi ne row upar-neeche kar di — fresh list se dobara dhoondo
      cache.del("products");
      const again = await findProduct(item["Item Name (Standardized)"]);
      if (!again) throw new Error("Sheet badal gayi hai — page refresh karke dobara try karo.");
      item = again;
      raw = await readProductRowRaw(item._row);
      if (normH(raw[C.NAME]) !== normH(item["Item Name (Standardized)"])) throw new Error("Sheet badal gayi hai — page refresh karke dobara try karo.");
    }
    const before = rowToProduct(raw, item._row, C);
    const changes = (await fn(before, C)) || {};
    const cells = [];
    for (const k of Object.keys(changes)) {
      const idx = Number(k);
      if (idx === C.CURCARTON || idx === C.CURPCS) throw new Error("Current Stock formula column mein likhna allowed nahi hai.");
      cells.push({ row: item._row, col: idx + 1, value: changes[k] });
      raw[idx] = changes[k];
    }
    await batchWrite(S.PRODUCTS, cells);
    const after = rowToProduct(raw, item._row, C);
    patchProductCache(after);
    DATA_VER++; cache.del("dash"); cache.del("vendors");
    return { before, after };
  });
}

// =================== 🆕 UNIT CONVERSION ===================
// Har item ke 2 hi unit allowed hain: "Unit" (chhota, jaise pcs) aur "Base Unit" (bada, jaise pkt/box/vori).
// Sheet ke saare totals (Opening Loose, Demand, Approved, Purchase) hamesha "Unit" mein rakhe jaate hain,
// isliye Base Unit mein di gayi qty ko × "1 Stock Unit = (Base Qty)" karke likha jata hai.
// Standard units: gm↔kg aur ml↔ltr apne aap convert (jaise 250 gm pudina = 0.25 kg)
const STD_UNIT = { g: ["mass", 1], gm: ["mass", 1], gms: ["mass", 1], gram: ["mass", 1], grams: ["mass", 1], kg: ["mass", 1000], kgs: ["mass", 1000], kilo: ["mass", 1000],
  ml: ["vol", 1], l: ["vol", 1000], ltr: ["vol", 1000], litre: ["vol", 1000], liter: ["vol", 1000] };
function stdOf(u) { return STD_UNIT[normH(u).replace(/\./g, "")] || null; }
function stdSibling(u) { const s = stdOf(u); if (!s) return ""; return s[0] === "mass" ? (s[1] === 1 ? "kg" : "gm") : (s[1] === 1 ? "ltr" : "ml"); }
function allowedUnits(item) {
  const s = [];
  [item.Unit, item["Base Unit"], stdSibling(item.Unit)].forEach(u => { u = String(u || "").trim(); if (u && !s.some(x => normH(x) === normH(u) || (stdOf(x) && stdOf(u) && stdOf(x)[0] === stdOf(u)[0] && stdOf(x)[1] === stdOf(u)[1]))) s.push(u); });
  return s;
}
// Kisi bhi unit ki qty ko item ke "Unit" mein badlo
function unitFactor(item, unit) {
  const u = normH(unit || item.Unit), nat = normH(item.Unit), base = normH(item["Base Unit"]);
  const suq = num(item["1 Stock Unit = (Base Qty)"]) || 1;
  if (!u || u === nat) return 1;
  if (u === base && base !== nat) return suq;
  const su = stdOf(u), sn = stdOf(nat), sb = stdOf(base);
  if (su && sn && su[0] === sn[0]) return su[1] / sn[1];
  if (su && sb && su[0] === sb[0] && base !== nat) return (su[1] / sb[1]) * suq;
  return 1;
}
function toNative(item, qty, unit) { return round3(qty * unitFactor(item, unit)); }
function checkUnit(item, unit, role) {
  const allowed = allowedUnits(item);
  if (!unit) return item.Unit;
  const hit = allowed.find(u => normH(u) === normH(unit));
  if (hit) return hit;
  if (role === "Hero") return String(unit).trim();
  throw new Error(`"${unit}" unit "${item["Item Name (Standardized)"]}" ke liye allowed nahi — sirf ${allowed.join(" / ")}. Naya unit sirf Hero bana sakta hai.`);
}
// Full Pack / Loose ko total ke hisaab se adjust karna (total = Full × 1Base + Loose)
function openingCells(b, total, C) {
  const suq = num(b["1 Stock Unit = (Base Qty)"]) || 1;
  total = round3(total);
  if (suq <= 1) return { [C.OPENLOOSE]: round3(total - b["W/H Opening Full Pack"]) };
  if (total < 0) return { [C.OPENFULL]: 0, [C.OPENLOOSE]: total };
  const full = Math.floor(total / suq + 1e-9);
  return { [C.OPENFULL]: full, [C.OPENLOOSE]: round3(total - full * suq) };
}

// =================== 🆕 RATE HISTORY (purana rate → naya rate) ===================
const RATE_HDR = ["Timestamp", "Item Name", "Old Rate", "New Rate", "Unit of rate", "Source", "By", "Ref"];
function logRateChange(itemName, oldRate, newRate, unit, source, by, ref) {
  if (newRate === "" || newRate === null || newRate === undefined || isNaN(Number(newRate))) return;
  if (oldRate !== "" && oldRate !== undefined && Number(oldRate) === Number(newRate)) return;
  (async () => {
    try {
      await ensureSheet(S.RATEHIST, RATE_HDR);
      await appendRow(S.RATEHIST, [toISTISO(new Date()), itemName, (oldRate === "" || oldRate === undefined) ? "" : oldRate, Number(newRate), unit || "", source || "", by || "", ref || ""]);
      cache.del("ratehist");
    } catch (e) { console.error("Rate history:", e.message); }
  })();
}
async function readRateHistory() {
  const m = await sheetMeta();
  if (!m[S.RATEHIST]) return [];
  const objs = rowsToObjects(await readRange(S.RATEHIST, "A1:H", "UNFORMATTED_VALUE"));
  objs.forEach(o => { o.Timestamp = normalizeDateTime(o.Timestamp); o["Old Rate"] = money(o["Old Rate"]); o["New Rate"] = money(o["New Rate"]); });
  return objs.slice(-1500).reverse();
}

// =================== OTHER READERS ===================
// Sheet ke asli headers: Log ID | Timestamp | Item Name | Mode | Qty | Previous Stock | Resulting Stock | Entered By | Last Edited | Rate | Expiry Date | Status
// (Unit, Qty (pcs), Unit of rate, Vendor naye columns end mein apne aap ban jate hain)
const WH_HDR = ["Log ID", "Timestamp", "Item Name", "Mode", "Qty", "Unit", "Qty (pcs)", "Previous Stock", "Resulting Stock", "Entered By", "Rate", "Unit of rate", "Expiry Date", "Vendor", "Status", "Last Edited"];
const WH_ALIAS = { "Log ID": /\bid\b/i };
const PO_ALIAS = { "PO ID": /^po\s*id$/i };

async function readWarehouseEntries() {
  const objs = rowsToObjects(await readRange(S.WHLOG, "A1:ZZ", "FORMATTED_VALUE"));
  objs.forEach(o => { o["Timestamp"] = normalizeDateTime(o["Timestamp"]); o["Expiry Date"] = normalizeDate(o["Expiry Date"]); o["Last Edited"] = normalizeDateTime(o["Last Edited"]); o["Rate"] = money(o["Rate"]); });
  return objs.filter(o => o["Item Name"]).slice(-3000).reverse();
}
async function readCylinderEntries() {
  const objs = rowsToObjects(await readRange(S.CYLINDER, "A1:ZZ", "FORMATTED_VALUE"));
  objs.forEach(o => { o["Timestamp"] = normalizeDateTime(o["Timestamp"]); });
  return objs.filter(o => o["Timestamp"]).slice(-3000).reverse();
}
async function readPO() {
  const objs = rowsToObjects(await readRange(S.PO, "A1:ZZ", "FORMATTED_VALUE"));
  objs.forEach(o => { o["Timestamp"] = normalizeDateTime(o["Timestamp"]); o["Expiry Date"] = normalizeDate(o["Expiry Date"]); o["Last Edited"] = normalizeDateTime(o["Last Edited"]); if (o["Rate"] !== undefined) o["Rate"] = money(o["Rate"]); });
  return objs;
}
async function readManualPurchases() {
  const objs = await readPO();
  return objs.filter(r => String(r["PO ID"]).indexOf("MPO-") === 0).slice(-3000).reverse();
}
async function readAttendance() {
  const rows = await readRange(S.ATTENDANCE, "A1:R", "FORMATTED_VALUE");
  if (rows.length < 2) return [];
  const headers = rows[0].map(h => String(h).trim());
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i]; const obj = { _row: i + 2 };
    headers.forEach((h, idx) => { if (h) obj[h] = r[idx]; });
    obj["Date"] = normalizeDate(obj["Date"]);
    obj["IN Time"] = normalizeDateTime(obj["IN Time"], obj["Date"]);
    obj["OUT Time"] = normalizeDateTime(obj["OUT Time"], obj["Date"]);
    if (!obj["User ID"] && !obj["Date"]) continue;
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

// =================== AUTO-PO ENGINE ===================
async function checkAndCreateAutoPO(itemName) {
  if (!itemName) return null;
  return withLock("autopo", async () => {
    const item = await findProduct(itemName);
    if (!item || isDaily(item["Stock Type"])) return null;
    const min = num(item.Min), max = num(item.Max), cur = num(item["Current Stock"]);
    if (min <= 0 || cur > min) return null;
    // Zaroorat "Unit" (chhota, jaise pcs) mein nikalti hai, phir jitne poore carton/pkt (Base Unit) banein
    // utne hi auto PO mein jaate hain. Baaki (loose) hissa supervisor PO mein "Add-on" se jodta hai.
    const needNative = max > min ? Math.max(1, round3(max - cur)) : Math.max(1, round3(min * 2 - cur));
    const nativeUnit = String(item.Unit || "").trim();
    const baseUnit = String(item["Base Unit"] || "").trim();
    const suq = num(item["1 Stock Unit = (Base Qty)"]) || 1;
    let orderQty = needNative, orderUnit = nativeUnit;
    if (baseUnit && normH(baseUnit) !== normH(nativeUnit) && suq > 1) {
      orderQty = Math.max(1, Math.floor(needNative / suq + 1e-9)); // kam se kam 1 carton
      orderUnit = baseUnit;
    }
    cache.del("po");
    const pos = await cached("po", readPO);
    const existing = pos.find(p => normH(p["Item Name"]) === normH(itemName) && String(p.Status || "").trim() === "Pending");
    if (existing) return { skipped: true, reason: "Pending PO already exists", row: existing._row };
    const poId = "APO-" + Date.now() + "-" + Math.floor(Math.random() * 90 + 10);
    const vendor = item.Vendor || "Unassigned";
    await appendByHeaders(S.PO, {
      "PO ID": poId, "Timestamp": toISTISO(new Date()), "Item Name": item["Item Name (Standardized)"], "Category": item.Category || "",
      "Unit": orderUnit, "Qty Ordered": orderQty, "Auto Qty": orderQty, "Status": "Pending", "Requested By": "AUTO-PO",
      "Notes": `Auto: Stock ${cur} <= Min ${min} (need ${needNative} ${nativeUnit})`, "Vendor": vendor,
    }, PO_ALIAS);
    invalidate("po", "po_batches", "mpo", "dash");
    return { created: true, poId, itemName: item["Item Name (Standardized)"], qty: orderQty, unit: orderUnit, current: cur, min, max, vendor };
  });
}
async function runAutoPOForAll() {
  const products = await cached("products", readProducts);
  const results = { created: [], skipped: [], failed: [] };
  for (const p of products) {
    if (num(p.Min) > 0 && num(p["Current Stock"]) <= num(p.Min) && !isDaily(p["Stock Type"])) {
      try {
        const r = await checkAndCreateAutoPO(p["Item Name (Standardized)"]);
        if (r && r.created) results.created.push(r);
        else if (r && r.skipped) results.skipped.push({ name: p["Item Name (Standardized)"], reason: r.reason });
      } catch (e) { results.failed.push({ name: p["Item Name (Standardized)"], error: e.message }); }
    }
  }
  return results;
}

// =================== LOGIN / BOOTSTRAP / POLL / STATS ===================
function wantForRole(perm) {
  const want = [];
  if (perm === "SuperAdmin") want.push("staff", "allAttendance");
  else {
    want.push("myAttendance");
    if (perm === "Admin") want.push("stats", "products", "demands");
    else if (perm === "Hero") want.push("stats", "products", "demands", "vendors");
    else if (perm === "Supervisor") want.push("demands", "products", "vendors");
    else if (perm === "Cashier") want.push("products", "demands");
  }
  return want;
}
async function handleLogin(p) {
  const staff = await cached("staff", readStaff);
  const match = staff.find(r => String(r.User || "").toLowerCase().trim() === String(p.userId).toLowerCase().trim() && String(r.Password || "").trim() === String(p.password).trim());
  if (!match) throw new Error("Invalid User ID or Password");
  const result = { userId: match.User, name: match.Name, designation: match.Designation, shift: match["Shift Timing"], permission: String(match.Permission || "").trim() };
  try { result.bootstrap = await handleBootstrap({ want: wantForRole(result.permission), userId: result.userId }); }
  catch (e) { console.error("Bootstrap:", e.message); result.bootstrap = {}; }
  return result;
}

async function handleBootstrap(p) {
  const want = p.want || [];
  const out = { poll: null }; // sirf maangi hui lists bhejo — khali [] bhejne se app galat cache kar leta tha
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
  const products = await cached("products", readProducts).catch(() => []);
  const dailyNames = new Set(products.filter(p => isDaily(p["Stock Type"])).map(p => normH(p["Item Name (Standardized)"])));
  let pending = 0, newest = null, pendingStocked = 0, newestStocked = null;
  demands.forEach(d => {
    if (d.Status !== "Pending") return;
    pending++; if (!newest || new Date(d.Timestamp) > new Date(newest.Timestamp)) newest = d;
    if (!dailyNames.has(normH(d["Item Name"]))) { pendingStocked++; if (!newestStocked || new Date(d.Timestamp) > new Date(newestStocked.Timestamp)) newestStocked = d; }
  });
  const brief = (d) => d ? { item: d["Item Name"], qty: d["Qty Requested"], unit: d.Unit, by: d["Raised By"] } : null;
  const sig = demands.length + "|" + demands.slice(0, 300).map(d => d["Demand ID"] + d.Status + d["Qty Requested"] + d["Approved Qty"] + d.Feedback + d.Rate).join(",")
    + "|" + products.map(p => p["Item Name (Standardized)"] + p["Current Stock"] + p.Rate).join(",");
  // pendingStocked/newestStocked = Supervisor ke liye (Daily Market wali demand Supervisor ko nahi dikhti)
  return { ver: DATA_VER + "-" + hashStr(sig), pending, newest: brief(newest), pendingStocked, newestStocked: brief(newestStocked) };
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
  const totalStaff = staff.filter(s => s.Permission !== "Admin" && s.Permission !== "SuperAdmin" && s.Permission !== "Hero").length;
  const health = { healthy: 0, low: 0, overstock: 0 };
  const categoryStock = {};
  products.forEach(p => {
    if (isDaily(p["Stock Type"])) return;
    const min = num(p.Min), max = num(p.Max), cur = num(p["Current Stock"]);
    const cat = p.Category || "Other"; categoryStock[cat] = round3((categoryStock[cat] || 0) + cur);
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
// Demand sheet: A ID | B Timestamp | C Raised By | D Item | E Category | F Unit | G Qty | H Status |
// I Warehouse Stock After | J Notes | K Approved Qty | L Approved At | M Feedback | N Feedback At |
// O Approved By | P Rate (₹/unit) | Q Amount (₹)
let DEMAND_HDR_OK = false;
async function ensureDemandHeaders() {
  if (DEMAND_HDR_OK) return;
  try {
    await ensureCols(S.DEMAND, 17);
    const h = (await readRange(S.DEMAND, "O1:Q1", "FORMATTED_VALUE"))[0] || [];
    const want = ["Approved By", "Rate (₹/unit)", "Amount (₹)"];
    const cells = [];
    want.forEach((w, i) => { if (!String(h[i] || "").trim()) cells.push({ row: 1, col: 15 + i, value: w }); });
    await batchWrite(S.DEMAND, cells);
    DEMAND_HDR_OK = true;
  } catch (e) { console.error("Demand headers:", e.message); }
}
async function readDemandRow(row) {
  return (await readRange(S.DEMAND, `A${row}:Q${row}`, "UNFORMATTED_VALUE"))[0] || [];
}
function whoLabel(name, role) {
  const n = String(name || "").trim(); const r = String(role || "").trim();
  return n ? (r ? `${n} (${r})` : n) : "";
}

async function handleAddDemand(p) {
  // Cashier sirf list wale item ki demand daal sakta hai — naya item sirf Hero banata hai
  if (p.mode === "new" || p.newName) throw new Error("Naya item sirf Hero bana sakta hai — dropdown se item chuno.");
  const role = await roleOf(p);
  const item = await findProduct(p.itemName);
  if (!item) throw new Error(`"${p.itemName}" list mein nahi hai — Hero se item banwao.`);
  const qty = num(p.qty); if (!(qty > 0)) throw new Error("Valid quantity daalo");
  const unit = checkUnit(item, p.unit, role);
  await ensureDemandHeaders();
  const itemName = item["Item Name (Standardized)"];
  const id = "DMD-" + Date.now() + "-" + Math.floor(Math.random() * 90 + 10);
  const rowNum = await appendRow(S.DEMAND, [id, toISTISO(new Date()), p.raisedBy || "", itemName, item.Category || "", unit, qty, "Pending", "", p.notes || "", "", "", "", "", "", "", ""]);
  const native = toNative(item, qty, unit);
  const { after } = await mutateProduct({ name: itemName }, (b, C) => ({ [C.DEMAND]: round3(b["Deman for canteen"] + native), [C.DEMANDUNIT]: b.Unit }));
  invalidate("demands", "dash");
  return { id, row: rowNum, itemName, unit, product: after };
}

async function approveOne(it, role, by) {
  const row = Number(it.row); if (!row || row < 2) throw new Error("Invalid row");
  const approved = Number(it.approvedQty);
  if (it.approvedQty === "" || it.approvedQty === undefined || isNaN(approved) || approved < 0) throw new Error("Approve qty sahi nahi hai");
  const d = await readDemandRow(row);
  const status = String(d[7] || "Pending").trim();
  if (status !== "Pending") throw new Error("Ye demand pehle hi " + status + " ho chuki hai.");
  const itemName = String(d[3] || "").trim();
  const item = await findProduct(itemName);
  if (!item) throw new Error(`"${itemName}" Product Master mein nahi mila — Hero se item check karwao.`);
  const daily = isDaily(item["Stock Type"]);
  if (role) {
    if (daily && !["Cashier", "Hero"].includes(role)) throw new Error("Daily-Market demand Cashier (ya Hero) approve karta hai.");
    if (!daily && !["Supervisor", "Hero"].includes(role)) throw new Error("Stocked demand sirf Supervisor (ya Hero) approve kar sakta hai.");
  }
  let unit;
  try { unit = checkUnit(item, it.unit || d[5] || item.Unit, role); }
  catch (e) { if (it.unit) throw e; unit = item.Unit; } // purani demand ka galat unit → item ka Unit

  // Daily Market: rate ya kul amount zaroori (kitne ka pada)
  let rate = "", amount = "";
  if (daily && approved > 0) {
    rate = money(it.rate); amount = money(it.amount);
    const okRate = rate !== "" && rate > 0, okAmt = amount !== "" && amount > 0;
    if (!okRate && !okAmt) throw new Error(`${itemName}: Daily Market approve karte time rate ya kul amount daalna zaroori hai.`);
    if (!okRate) rate = round2(amount / approved);
    if (!okAmt) amount = round2(rate * approved);
  }
  const nowIso = toISTISO(new Date());
  const native = toNative(item, approved, unit);
  let product = item, warning = "", productDelta = 0;
  if (!daily && native > 0) {
    // Stocked: Approved Demand Qty (J) mein Unit (pcs) ke hisaab se jodo, K = Unit
    const res = await mutateProduct({ name: itemName }, (b, C) => ({ [C.APPROVED]: round3(b["Approved Demand Qty by supervisor"] + native), [C.APPROVEDUNIT]: b.Unit }));
    product = res.after; productDelta = native;
    if (product["Current Stock"] < 0) warning = `${itemName}: approve ke baad stock ${product["Current Stock"]} ${product.Unit} (minus) ho gaya — Warehouse/Purchase entry check karo.`;
  } else if (daily && rate !== "") {
    // Daily Market: Product Master ka Rate + Unit of rate update, purana→naya Rate History mein
    const res = await mutateProduct({ name: itemName }, (b, C) => {
      const ch = {};
      if (String(b.Rate) !== String(rate)) ch[C.RATE] = rate;
      if (normH(b["Unit of rate"]) !== normH(unit)) ch[C.RATEUNIT] = unit;
      return ch;
    });
    product = res.after;
    if (String(res.before.Rate) !== String(rate)) logRateChange(itemName, res.before.Rate, rate, unit, "Daily Market approve", by, d[0]);
  }
  const fb = String(it.feedback || "").trim().slice(0, 500);
  const cells = [
    { row, col: 8, value: "Approved" }, { row, col: 6, value: unit }, { row, col: 11, value: approved },
    { row, col: 12, value: nowIso }, { row, col: 9, value: daily ? "" : product["Current Stock"] },
  ];
  if (by) cells.push({ row, col: 15, value: by });
  if (rate !== "") cells.push({ row, col: 16, value: rate }, { row, col: 17, value: amount });
  if (fb) cells.push({ row, col: 13, value: fb }, { row, col: 14, value: nowIso });
  try { await batchWrite(S.DEMAND, cells); }
  catch (e) {
    if (productDelta) { try { await mutateProduct({ name: itemName }, (b, C) => ({ [C.APPROVED]: round3(b["Approved Demand Qty by supervisor"] - productDelta) })); } catch (_) {} }
    throw e;
  }
  let autoPO = null;
  if (!daily) { try { const r = await checkAndCreateAutoPO(itemName); if (r && r.created) autoPO = r; } catch (e) { console.error("AutoPO:", e.message); } }
  return { autoPO, result: { row, ok: true, approvedQty: approved, unit, rate, amount, approvedBy: by, approvedAt: nowIso, warning, product } };
}

async function handleApproveDemands(p) {
  const items = Array.isArray(p.items) ? p.items : [];
  if (!items.length) throw new Error("Koi demand select nahi hui.");
  await ensureDemandHeaders();
  const role = await roleOf(p);
  const by = whoLabel(p.approvedBy || p.by, role || p.role);
  const results = [], autoPOs = [];
  for (const it of items) {
    try {
      const r = await withLock("demand:" + Number(it.row), () => approveOne(it, role, by));
      if (r.autoPO) autoPOs.push(r.autoPO);
      results.push(r.result);
    } catch (e) { results.push({ row: it.row, ok: false, error: e.message }); }
  }
  invalidate("demands", "dash");
  return { results, autoPOs };
}

async function reduceProductDemand(itemName, qty, unit) {
  if (!itemName || !qty) return null;
  const item = await findProduct(itemName);
  if (!item) return null;
  const native = toNative(item, qty, unit || item.Unit);
  const { after } = await mutateProduct({ name: itemName }, (b, C) => ({ [C.DEMAND]: Math.max(0, round3(b["Deman for canteen"] - native)) }));
  return after;
}

async function handleRejectDemands(p) {
  const rows = Array.isArray(p.rows) ? p.rows.map(Number) : [];
  if (!rows.length) throw new Error("Koi demand select nahi hui.");
  await ensureDemandHeaders();
  const role = await roleOf(p);
  const by = whoLabel(p.approvedBy || p.by, role || p.role);
  const results = [];
  for (const row of rows) {
    try {
      const product = await withLock("demand:" + row, async () => {
        if (!row || row < 2) throw new Error("Invalid row");
        const d = await readDemandRow(row);
        const status = String(d[7] || "Pending").trim();
        if (status !== "Pending") throw new Error("Ye demand pehle hi " + status + " ho chuki hai.");
        const after = await reduceProductDemand(String(d[3] || "").trim(), num(d[6]), d[5]);
        const cells = [{ row, col: 8, value: "Rejected" }, { row, col: 12, value: toISTISO(new Date()) }];
        if (by) cells.push({ row, col: 15, value: by });
        await batchWrite(S.DEMAND, cells);
        return after;
      });
      results.push({ row, ok: true, product });
    } catch (e) { results.push({ row, ok: false, error: e.message }); }
  }
  invalidate("demands", "dash");
  return { results };
}

function assertEditable(r, p) {
  const status = String(r[7] || "Pending").trim();
  if (status !== "Pending") throw new Error("Ye demand pehle hi " + status + " ho chuki hai.");
  const t = parseIST(normalizeDateTime(r[1]));
  if (!isNaN(t) && Date.now() - t > DEMAND_EDIT_WINDOW_MS + GRACE_MS) throw new Error("1 ghanta ho gaya — ab ye demand edit/delete nahi ho sakti (lock).");
  if (p.by && String(r[2] || "").trim() && normH(r[2]) !== normH(p.by)) throw new Error("Sirf demand banane wala hi edit/delete kar sakta hai.");
}

async function handleEditDemand(p) {
  const row = Number(p.row);
  if (!row || row < 2) throw new Error("Invalid row");
  const newQty = Number(p.newQty);
  if (isNaN(newQty) || newQty <= 0) throw new Error("Invalid quantity");
  return withLock("demand:" + row, async () => {
    const r = await readDemandRow(row);
    assertEditable(r, p);
    const itemName = String(r[3] || "").trim();
    const oldQty = num(r[6]);
    const item = await findProduct(itemName);
    let product = null;
    if (item) {
      const diff = round3(toNative(item, newQty, r[5]) - toNative(item, oldQty, r[5]));
      if (diff !== 0) product = (await mutateProduct({ name: itemName }, (b, C) => ({ [C.DEMAND]: Math.max(0, round3(b["Deman for canteen"] + diff)) }))).after;
    }
    await batchWrite(S.DEMAND, [{ row, col: 7, value: newQty }]);
    invalidate("demands", "dash");
    return { row, oldQty, newQty, product };
  });
}

async function handleDeleteDemand(p) {
  const row = Number(p.row);
  if (!row || row < 2) throw new Error("Invalid row");
  return withLock("demand:" + row, async () => {
    const r = await readDemandRow(row);
    assertEditable(r, p);
    const product = await reduceProductDemand(String(r[3] || "").trim(), num(r[6]), r[5]);
    await batchWrite(S.DEMAND, [{ row, col: 8, value: "Cancelled" }]);
    invalidate("demands", "dash");
    return { row, product };
  });
}

async function handleSaveDemandFeedback(p) {
  const row = Number(p.row); if (!row || row < 2) throw new Error("Invalid row");
  const feedback = String(p.feedback || "").trim().slice(0, 500);
  if (!feedback) throw new Error("Feedback khali hai");
  const r = await readDemandRow(row);
  if (String(r[7] || "").trim() !== "Approved") throw new Error("Feedback sirf approved demand par diya ja sakta hai.");
  const base = parseIST(normalizeDateTime(r[11])) || parseIST(normalizeDateTime(r[1]));
  if (!isNaN(base) && Date.now() - base > DEMAND_FEEDBACK_WINDOW_MS + GRACE_MS) throw new Error("3 din ho gaye — feedback ab lock ho chuka hai.");
  const nowIso = toISTISO(new Date());
  await batchWrite(S.DEMAND, [{ row, col: 13, value: feedback }, { row, col: 14, value: nowIso }]);
  invalidate("demands");
  return { row, feedback, feedbackAt: nowIso };
}

// =================== ATTENDANCE HANDLERS ===================
function businessDateKey(d) {
  const ist = new Date(new Date(d).getTime() + 5.5 * 3600000);
  if (ist.getUTCHours() < 3) ist.setUTCDate(ist.getUTCDate() - 1);
  return ist.toISOString().slice(0, 10);
}

async function handleMarkAttendance(p) {
  const now = new Date();
  const dateStr = businessDateKey(now);
  const type = p.type === "OUT" ? "OUT" : "IN";
  const rows = await readRange(S.ATTENDANCE, "A1:R", "FORMATTED_VALUE");
  if (rows.length < 1) throw new Error("Attendance sheet khaali");
  const headers = rows[0].map(h => String(h).trim());
  const ci = {}; headers.forEach((h, i) => { ci[h] = i; });
  let todayRow = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][ci["User ID"]] || "").trim() === String(p.userId).trim() && normalizeDate(rows[i][ci["Date"]]) === dateStr) { todayRow = i + 1; break; }
  }
  if (type === "IN" && todayRow > 0 && rows[todayRow - 1][ci["IN Time"]]) throw new Error("Aap already IN punch kar chuke ho.");
  if (type === "OUT") {
    if (todayRow <= 0) throw new Error("Pehle IN punch karo.");
    if (rows[todayRow - 1][ci["OUT Time"]]) throw new Error("Aap already OUT punch kar chuke ho.");
  }
  const photoUrl = await uploadPhotoToDrive(p.photoBase64, p.userId, p.name || p.userId);
  const mapsLink = (p.lat && p.lng) ? ("https://www.google.com/maps?q=" + p.lat + "," + p.lng) : "";
  const nowIso = toISTISO(now);
  let status;
  if (type === "IN") {
    status = computeInStatus(p.shift, now);
    await appendRow(S.ATTENDANCE, [dateStr, p.userId, p.name || "", p.designation || "", p.shift || "", nowIso, photoUrl, mapsLink, p.lat || "", p.lng || "", status, "", "", "", "", "", "", "Present"]);
  } else {
    status = computeOutStatus(p.shift, now);
    const cells = [];
    const put = (h, v) => { if (ci[h] !== undefined) cells.push({ row: todayRow, col: ci[h] + 1, value: v }); };
    put("OUT Time", nowIso); put("OUT Photo URL", photoUrl); put("OUT Maps Link", mapsLink);
    put("OUT Latitude", p.lat || ""); put("OUT Longitude", p.lng || ""); put("OUT Status", status);
    await batchWrite(S.ATTENDANCE, cells);
  }
  invalidate("attendance", "dash");
  return { status, type, photoUrl, istTime: nowIso };
}

async function handleEditAttendancePunch(p) {
  const rows = await readRange(S.ATTENDANCE, "A1:R", "FORMATTED_VALUE");
  const headers = rows[0].map(h => String(h).trim());
  const ci = {}; headers.forEach((h, i) => { ci[h] = i; });
  const note = "Manually corrected by " + (p.editedBy || "SuperAdmin");
  let ex = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][ci["User ID"]] || "").trim() === String(p.userId).trim() && normalizeDate(rows[i][ci["Date"]]) === p.dateISO) { ex = i + 1; break; }
  }
  if (p.inTime) {
    const ts = p.dateISO + "T" + p.inTime + ":00";
    const inMin = hmToMinutes(p.inTime);
    const inStatus = inMin === null ? "Marked" : computeInStatusAt(p.shift, inMin);
    if (ex > 0) {
      const cells = [{ row: ex, col: ci["IN Time"] + 1, value: ts }, { row: ex, col: ci["IN Photo URL"] + 1, value: note }];
      if (ci["IN Status"] !== undefined) cells.push({ row: ex, col: ci["IN Status"] + 1, value: inStatus });
      await batchWrite(S.ATTENDANCE, cells);
    } else {
      ex = await appendRow(S.ATTENDANCE, [p.dateISO, p.userId, p.name, p.designation, p.shift, ts, note, "", "", "", inStatus, "", "", "", "", "", "", "Present"]);
    }
  }
  if (p.outTime) {
    if (ex <= 0) throw new Error("Pehle IN set karo");
    const ts = p.dateISO + "T" + p.outTime + ":00";
    const outMin = hmToMinutes(p.outTime);
    const outStatus = outMin === null ? "Marked" : computeOutStatusAt(p.shift, outMin);
    const cells = [{ row: ex, col: ci["OUT Time"] + 1, value: ts }, { row: ex, col: ci["OUT Photo URL"] + 1, value: note }];
    if (ci["OUT Status"] !== undefined) cells.push({ row: ex, col: ci["OUT Status"] + 1, value: outStatus });
    await batchWrite(S.ATTENDANCE, cells);
  }
  invalidate("attendance", "dash");
  return { ok: true };
}

// =================== 🆕 STOCK ENTRY (Warehouse + Manual Purchase) — common logic ===================
// kind "wh"  → W/H Opening (Full Pack / Loose) mein judta hai, log "Warehouse Entry Log" mein
// kind "mp"  → Purchase Stock QTy (L) mein judta hai, M = Unit, log "Purchase Orders" (MPO-...) mein
function readEntryInputs(p, item, role) {
  const qty = num(p.qty !== undefined ? p.qty : p.newQty);
  if (!(qty > 0)) throw new Error("Valid quantity daalo");
  const unit = checkUnit(item, p.unit || item.Unit, role);
  const rate = money(p.rate);
  let rateUnit = String(p.rateUnit || "").trim();
  if (rateUnit) rateUnit = (allowedUnits(item).find(u => normH(u) === normH(rateUnit))) || (role === "Hero" ? rateUnit : (normH(rateUnit) === normH(item["Unit of rate"]) ? item["Unit of rate"] : unit));
  const expiry = p.expiryDate ? (normalizeDate(p.expiryDate) || "") : "";
  const vendor = (p.vendor === undefined || p.vendor === null) ? null : String(p.vendor).trim();
  return { qty, unit, native: toNative(item, qty, unit), rate, rateUnit, expiry, vendor };
}
function extraProductCells(b, C, inp) {
  const ch = {};
  if (inp.rate !== "") { if (String(b.Rate) !== String(inp.rate)) ch[C.RATE] = inp.rate; ch[C.RATEUNIT] = inp.rateUnit || b["Unit of rate"] || inp.unit; }
  if (inp.expiry) ch[C.EXPIRY] = inp.expiry;
  if (inp.vendor !== null && inp.vendor !== b.Vendor) ch[C.VENDOR] = inp.vendor;
  return ch;
}
function stockDeltaCells(kind, b, C, delta) {
  if (kind === "wh") return openingCells(b, b["Warehouse Stock Qty Opening STock"] + delta, C);
  return { [C.PURCHASE]: round3(b["Purchase Stock  QTy"] + delta), [C.PURCHASEUNIT]: b.Unit };
}
const MP_HDR = ["PO ID", "Timestamp", "Item Name", "Category", "Unit", "Qty Ordered", "Status", "Requested By", "Notes", "Vendor", "Rate", "Unit of rate", "Expiry Date", "Qty (pcs)", "Last Edited"];

async function addStockEntry(kind, p) {
  const role = await requireRole(p, ["Supervisor", "Hero"], kind === "wh" ? "Warehouse Stock Entry" : "Purchase entry");
  if (p.mode === "new" || p.newName) throw new Error("Naya item sirf Hero bana sakta hai (Item Master se).");
  const item = await findProduct(p.itemName);
  if (!item) throw new Error(`"${p.itemName}" list mein nahi hai — Hero se item banwao.`);
  if (isDaily(item["Stock Type"])) throw new Error("Daily-Market item mein stock entry nahi hoti.");
  const inp = readEntryInputs(p, item, role);
  const name = item["Item Name (Standardized)"];
  const { before, after } = await mutateProduct({ name }, (b, C) => Object.assign({}, stockDeltaCells(kind, b, C, inp.native), extraProductCells(b, C, inp)));
  const ts = toISTISO(new Date());
  let logRow;
  if (kind === "wh") {
    const id = "WH-" + Date.now();
    logRow = await appendByHeaders(S.WHLOG, {
      "Log ID": id, "Timestamp": ts, "Item Name": name, "Mode": "add", "Qty": inp.qty, "Unit": inp.unit, "Qty (pcs)": inp.native,
      "Previous Stock": before["Current Stock"], "Resulting Stock": after["Current Stock"], "Entered By": p.enteredBy || "",
      "Rate": inp.rate, "Unit of rate": inp.rate !== "" ? (inp.rateUnit || after["Unit of rate"]) : "", "Expiry Date": inp.expiry,
      "Vendor": after.Vendor, "Status": "Active",
    }, WH_ALIAS);
    if (inp.rate !== "") logRateChange(name, before.Rate, inp.rate, inp.rateUnit, "Warehouse entry", p.enteredBy, id);
    cache.del("whlog");
  } else {
    const id = "MPO-" + Date.now();
    logRow = await appendByHeaders(S.PO, {
      "PO ID": id, "Timestamp": ts, "Item Name": name, "Category": item.Category || "", "Unit": inp.unit, "Qty Ordered": inp.qty,
      "Status": "Received (Manual)", "Requested By": p.enteredBy || "", "Notes": "Manual purchase entry", "Vendor": after.Vendor,
      "Rate": inp.rate, "Unit of rate": inp.rate !== "" ? (inp.rateUnit || after["Unit of rate"]) : "", "Expiry Date": inp.expiry, "Qty (pcs)": inp.native,
    }, PO_ALIAS);
    if (inp.rate !== "") logRateChange(name, before.Rate, inp.rate, inp.rateUnit, "Manual purchase", p.enteredBy, id);
    cache.del("po"); cache.del("mpo"); cache.del("po_batches");
  }
  DATA_VER++;
  return { itemName: name, newValue: after["Current Stock"], unit: after.Unit, row: logRow, product: after };
}

async function editStockEntry(kind, p, del) {
  const role = await requireRole(p, ["Supervisor", "Hero"], "Entry edit/delete");
  const row = Number(kind === "wh" ? (p.logRow || p.row) : (p.row || p.logRow));
  if (!row || row < 2) throw new Error("Invalid row");
  const sheet = kind === "wh" ? S.WHLOG : S.PO;
  const H = kind === "wh" ? WH_HDR : MP_HDR;
  const qtyKey = kind === "wh" ? "Qty" : "Qty Ordered";
  return withLock("entry:" + kind + ":" + row, async () => {
    const e = await readRowByHeaders(sheet, row, H, kind === "wh" ? WH_ALIAS : PO_ALIAS);
    if (String(e["Status"] || "").trim() === "Cancelled" || String(e["Entered By"] || "").trim() === "Cancelled") throw new Error("Ye entry pehle hi delete ho chuki hai.");
    if (kind === "mp" && String(e["PO ID"] || "").indexOf("MPO-") !== 0) throw new Error("Ye manual purchase entry nahi hai.");
    const t = parseIST(normalizeDateTime(e["Timestamp"]));
    if (role !== "Hero" && !isNaN(t) && Date.now() - t > ENTRY_EDIT_WINDOW_MS + GRACE_MS) throw new Error("1 ghanta ho gaya — entry lock ho chuki hai (sirf Hero badal sakta hai).");
    const item = await findProduct(e["Item Name"]);
    if (!item) throw new Error("Item Product Master mein nahi mila: " + e["Item Name"]);
    const oldNative = (e["Qty (pcs)"] !== undefined && e["Qty (pcs)"] !== "") ? num(e["Qty (pcs)"]) : toNative(item, num(e[qtyKey]), e["Unit"] || item.Unit);
    const name = item["Item Name (Standardized)"];
    const nowIso = toISTISO(new Date());
    if (del) {
      const { after } = await mutateProduct({ name }, (b, C) => stockDeltaCells(kind, b, C, -oldNative));
      await writeByHeaders(sheet, row, { "Status": "Cancelled", "Last Edited": nowIso }, kind === "wh" ? WH_ALIAS : PO_ALIAS);
      invalidate(kind === "wh" ? "whlog" : "mpo", "po", "po_batches");
      return { row, product: after };
    }
    const inp = readEntryInputs(Object.assign({}, p, { unit: p.unit || e["Unit"] || item.Unit }), item, role);
    const diff = round3(inp.native - oldNative);
    const { before, after } = await mutateProduct({ name }, (b, C) => Object.assign({}, diff ? stockDeltaCells(kind, b, C, diff) : {}, extraProductCells(b, C, inp)));
    const upd = { [qtyKey]: inp.qty, "Unit": inp.unit, "Qty (pcs)": inp.native, "Last Edited": nowIso, "Expiry Date": inp.expiry };
    if (inp.rate !== "") { upd["Rate"] = inp.rate; upd["Unit of rate"] = inp.rateUnit || after["Unit of rate"]; }
    if (kind === "wh") upd["Resulting Stock"] = after["Current Stock"];
    if (inp.vendor !== null) upd["Vendor"] = inp.vendor;
    await writeByHeaders(sheet, row, upd, kind === "wh" ? WH_ALIAS : PO_ALIAS);
    if (inp.rate !== "") logRateChange(name, before.Rate, inp.rate, inp.rateUnit, kind === "wh" ? "Warehouse entry edit" : "Purchase edit", p.enteredBy || "", e["Log ID"] || e["PO ID"] || "");
    invalidate(kind === "wh" ? "whlog" : "mpo", "po", "po_batches");
    return { row, product: after, diff };
  });
}

// =================== CYLINDER ===================
async function handleAddCylinderEntry(p) {
  const qty = Number(p.qty) || 0; if (qty <= 0) throw new Error("Invalid quantity");
  const price = Number(await getSetting("CylinderPrice", "2900")) || 2900;
  const total = qty * price;
  const photoUrl = await uploadPhotoToDrive(p.photoBase64, "cyl_" + Date.now(), "cyl");
  await appendRow(S.CYLINDER, ["CYL-" + Date.now(), toISTISO(new Date()), p.receivedBy || "", qty, price, total, photoUrl]);
  invalidate("cyl", "dash");
  return { qty, price, total, photoUrl };
}

// =================== PO HANDLERS ===================
async function handleAddPurchaseOrder(p) {
  const item = await findProduct(p.itemName);
  const qty = Number(p.qty) || 0; if (qty <= 0) throw new Error("Quantity 0 se zyada honi chahiye");
  // Unit sirf item ke apne unit (Unit / Base Unit) mein se hi — naya unit yahan se nahi ban sakta (sirf Hero item master se)
  const unit = item ? checkUnit(item, p.unit, "") : (p.unit || "");
  const id = "PO-" + Date.now();
  await appendByHeaders(S.PO, {
    "PO ID": id, "Timestamp": toISTISO(new Date()), "Item Name": item ? item["Item Name (Standardized)"] : p.itemName,
    "Category": p.category || (item && item.Category) || "", "Unit": unit, "Qty Ordered": qty,
    "Status": "Pending", "Requested By": p.requestedBy || "", "Notes": p.notes || "", "Vendor": (item && item.Vendor) || "",
    "Item Note": String(p.itemNote || "").trim().slice(0, 200),
  }, PO_ALIAS);
  invalidate("po", "po_batches");
  return { id };
}
// PO item par Add-on qty (auto qty ke upar extra) aur Note — Pending / Generated dono par chalta hai
async function handleUpdatePOItem(p) {
  const row = Number(p.row); if (!row || row < 2) throw new Error("Invalid row");
  const cur = await readRowByHeaders(S.PO, row, ["PO ID", "Item Name", "Status", "Unit"], PO_ALIAS);
  if (!cur["PO ID"]) throw new Error("PO row nahi mili");
  if (String(cur["PO ID"]).indexOf("MPO-") === 0) throw new Error("Manual purchase yahan edit nahi hota");
  const st = String(cur.Status || "").trim();
  if (st !== "Pending" && st !== "Generated") throw new Error((st || "Ye") + " PO item edit nahi ho sakta");
  const upd = {};
  if (p.addonQty !== undefined) {
    const q = Number(p.addonQty) || 0;
    if (q < 0) throw new Error("Add-on qty negative nahi ho sakti");
    let u = "";
    if (q > 0) {
      u = String(p.addonUnit || "").trim() || String(cur.Unit || "").trim();
      const item = await findProduct(cur["Item Name"]);
      if (item) u = checkUnit(item, u, "");
    }
    upd["Add-on Qty"] = q > 0 ? round3(q) : "";
    upd["Add-on Unit"] = u;
  }
  if (p.itemNote !== undefined) upd["Item Note"] = String(p.itemNote || "").trim().slice(0, 200);
  if (!Object.keys(upd).length) throw new Error("Kuch badla nahi");
  await writeByHeaders(S.PO, row, upd, PO_ALIAS);
  invalidate("po", "po_batches");
  return { row, addonQty: upd["Add-on Qty"] === undefined ? undefined : upd["Add-on Qty"], addonUnit: upd["Add-on Unit"], itemNote: upd["Item Note"] };
}
async function handleCancelPO(p) {
  const row = Number(p.row); if (!row || row < 2) throw new Error("Invalid row");
  await writeByHeaders(S.PO, row, { "Status": "Cancelled" }, PO_ALIAS);
  invalidate("po", "po_batches");
  return { row };
}
async function handleClearPOBatch(p) {
  const rows = await readRange(S.PO, "A1:ZZ", "FORMATTED_VALUE");
  const headers = rows[0].map(h => String(h).trim());
  const ci = {}; headers.forEach((h, i) => { ci[h] = i; });
  const cells = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if ((r[ci["Vendor"]] || "Unassigned") === p.vendor && (r[ci["Batch ID"]] || "Legacy") === p.batchId && r[ci["Status"]] === "Generated") cells.push({ row: i + 1, col: ci["Status"] + 1, value: "Cleared" });
  }
  await batchWrite(S.PO, cells);
  invalidate("po", "po_batches");
  return { cleared: cells.length };
}
async function handleGeneratePOBatch(p) {
  const rows = await readRange(S.PO, "A1:ZZ", "FORMATTED_VALUE");
  const headers = rows[0].map(h => String(h).trim());
  const ci = {}; headers.forEach((h, i) => { ci[h] = i; });
  const selected = Array.isArray(p.rows) ? p.rows.map(Number) : [];
  const cells = []; let remainingPending = 0;
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (String(r[ci["PO ID"]] || "").indexOf("MPO-") === 0) continue;
    if ((r[ci["Vendor"]] || "Unassigned") === p.vendor && (r[ci["Batch ID"]] || "Legacy") === p.batchId && r[ci["Status"]] === "Pending") {
      if (selected.length === 0 || selected.indexOf(i + 1) !== -1) cells.push({ row: i + 1, col: ci["Status"] + 1, value: "Generated" });
      else remainingPending++;
    }
  }
  await batchWrite(S.PO, cells);
  invalidate("po", "po_batches");
  return { itemCount: cells.length, remainingPending, pdfUrl: "" };
}
async function handleAutoTagVendors() {
  return withLock("products", async () => {
    const C = await getProdCols();
    cache.del("products");
    const products = await cached("products", readProducts);
    const cells = [];
    products.forEach(p => {
      if (String(p.Vendor || "").trim()) return;
      const m = matchVendorForItem(p["Item Name (Standardized)"]);
      if (m) cells.push({ row: p._row, col: C.VENDOR + 1, value: m });
    });
    await batchWrite(S.PRODUCTS, cells);
    invalidate("products", "vendors");
    return { tagged: cells.length, total: products.length };
  });
}

// =================== STAFF ===================
async function handleAddStaff(p) {
  const staff = await cached("staff", readStaff);
  const userId = p.userId || (String(p.name).replace(/\s+/g, "") + "-" + String(staff.length + 1).padStart(3, "0"));
  await appendRow(S.STAFF, [p.name, p.designation, p.shift, userId, p.permission, p.password || "123"]);
  invalidate("staff", "dash");
  return { userId };
}
async function handleUpdateStaff(p) {
  const row = Number(p.row); if (!row || row < 2) throw new Error("Invalid row");
  await batchWrite(S.STAFF, [
    { row, col: 1, value: p.name }, { row, col: 2, value: p.designation }, { row, col: 3, value: p.shift },
    { row, col: 4, value: p.userId }, { row, col: 5, value: p.permission }, { row, col: 6, value: p.password },
  ]);
  invalidate("staff", "dash");
  return { row };
}
async function handleDeleteStaff(p) {
  const row = Number(p.row); if (!row || row < 2) throw new Error("Invalid row");
  try {
    const m = await sheetMeta();
    const sheetId = m[S.STAFF] && m[S.STAFF].sheetId;
    if (sheetId !== undefined) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SHEET_ID,
        requestBody: { requests: [{ deleteDimension: { range: { sheetId, dimension: "ROWS", startIndex: row - 1, endIndex: row } } }] },
      });
    }
  } catch (e) { console.error("Delete staff:", e.message); }
  invalidate("staff", "dash");
  return { deleted: row };
}

// =================== 🆕 HERO — ITEM CREATE / UPDATE (Product & Stock Master ka har column) ===================
async function copyStockFormulas(rowNum, C) {
  try {
    if (rowNum < 3) return;
    const src = rowNum - 1;
    const lo = Math.min(C.CURCARTON, C.CURPCS), hi = Math.max(C.CURCARTON, C.CURPCS);
    const res = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `'${S.PRODUCTS}'!${colLetter(lo + 1)}${src}:${colLetter(hi + 1)}${src}`, valueRenderOption: "FORMULA" });
    const v = (res.data.values || [])[0] || [];
    if (!v.some(x => typeof x === "string" && x.trim().startsWith("="))) return; // ARRAYFORMULA ya value — copy ki zaroorat nahi
    const m = await sheetMeta();
    const sid = m[S.PRODUCTS].sheetId;
    await sheets.spreadsheets.batchUpdate({ spreadsheetId: SHEET_ID, requestBody: { requests: [{ copyPaste: {
      source: { sheetId: sid, startRowIndex: src - 1, endRowIndex: src, startColumnIndex: lo, endColumnIndex: hi + 1 },
      destination: { sheetId: sid, startRowIndex: rowNum - 1, endRowIndex: rowNum, startColumnIndex: lo, endColumnIndex: hi + 1 },
      pasteType: "PASTE_FORMULA" } }] } });
  } catch (e) { console.error("Formula copy:", e.message); }
}
function heroVals(f, unit) {
  return {
    CAT: String(f.category || "").trim(), PACK: f.packaging === undefined ? undefined : String(f.packaging).trim(),
    OPENFULL: f.openFull === undefined ? undefined : num(f.openFull), OPENLOOSE: f.openLoose === undefined ? undefined : num(f.openLoose),
    DEMAND: f.demand === undefined ? undefined : num(f.demand), APPROVED: f.approved === undefined ? undefined : num(f.approved),
    PURCHASE: f.purchase === undefined ? undefined : num(f.purchase),
    DEMANDUNIT: unit, APPROVEDUNIT: unit, PURCHASEUNIT: unit, UNIT: unit,
    MIN: f.min === undefined ? undefined : (String(f.min).trim() === "" ? "" : num(f.min)),
    MAX: f.max === undefined ? undefined : (String(f.max).trim() === "" ? "" : num(f.max)),
    VENDOR: f.vendor === undefined ? undefined : String(f.vendor).trim(),
    RATE: f.rate === undefined ? undefined : money(f.rate),
    RATEUNIT: f.rateUnit === undefined ? undefined : (String(f.rateUnit).trim() || unit),
    EXPIRY: f.expiry === undefined ? undefined : (f.expiry ? (normalizeDate(f.expiry) || String(f.expiry)) : ""),
    BASEUNIT: f.baseUnit === undefined ? undefined : (String(f.baseUnit).trim() || unit),
    SUQ: f.suq === undefined ? undefined : (num(f.suq) || 1),
    STOCKTYPE: f.stockType === undefined ? undefined : (isDaily(f.stockType) ? "Daily Market" : "Stocked"),
  };
}
async function heroCreateProduct(f, by) {
  const out = await withLock("products", async () => {
    const C = await getProdCols();
    cache.del("products");
    const list = await cached("products", readProducts);
    if (list.find(x => normH(x["Item Name (Standardized)"]) === normH(f.name))) throw new Error("Ye item pehle se hai: " + f.name);
    const unit = f.unit;
    const vals = heroVals(Object.assign({ openFull: 0, openLoose: 0, demand: 0, approved: 0, purchase: 0, baseUnit: unit, suq: 1, stockType: "Stocked", rateUnit: unit }, f), unit);
    const maxIdx = Math.max(...Object.values(C));
    const arr = new Array(maxIdx + 1).fill("");
    arr[C.SNO] = "=ROW()-1"; // sheet jaisa hi S no formula
    arr[C.NAME] = f.name;
    Object.keys(vals).forEach(k => { if (vals[k] !== undefined && C[k] !== undefined) arr[C[k]] = vals[k]; });
    arr[C.CURCARTON] = null; arr[C.CURPCS] = null; // formula columns — chhod do
    const rowNum = await appendRow(S.PRODUCTS, arr);
    await copyStockFormulas(rowNum, C);
    cache.del("products"); DATA_VER++; cache.del("dash"); cache.del("vendors");
    const fresh = (await cached("products", readProducts)).find(x => x._row === rowNum) || null;
    return { fresh, rate: vals.RATE, rateUnit: vals.RATEUNIT };
  });
  if (out.rate !== "" && out.rate !== undefined) logRateChange(f.name, "", out.rate, out.rateUnit, "Hero: naya item", by, "");
  return out.fresh || await findProduct(f.name);
}
async function handleHeroSaveProduct(p) {
  await requireRole(p, ["Hero"], "Item create/update");
  const st = await staffOf(p);
  const by = st ? st.Name : "Hero";
  const f = p.fields || {};
  const name = String(f.name || "").trim(), category = String(f.category || "").trim(), unit = String(f.unit || "").trim();
  if (!name) throw new Error("Item ka naam likho");
  if (!category || !unit) throw new Error("Category aur Unit dono zaroori hain");
  if (f.suq !== undefined && !(num(f.suq) > 0)) throw new Error("1 Base Unit = qty 0 se zyada honi chahiye");
  const ff = Object.assign({}, f, { name, category, unit });
  if (!p.row) return { product: await heroCreateProduct(ff, by), created: true };

  const list = await cached("products", readProducts);
  const clash = list.find(x => normH(x["Item Name (Standardized)"]) === normH(name) && x._row !== Number(p.row));
  if (clash) throw new Error("Is naam ka item pehle se hai (row " + clash._row + ")");
  const res = await mutateProduct({ row: Number(p.row) }, (b, C) => {
    if (p.originalName && normH(b["Item Name (Standardized)"]) !== normH(p.originalName)) throw new Error("Sheet mein ye row badal gayi hai — page refresh karke dobara try karo.");
    const ch = { [C.NAME]: name };
    const vals = heroVals(ff, unit);
    Object.keys(vals).forEach(k => { if (vals[k] !== undefined && C[k] !== undefined) ch[C[k]] = vals[k]; });
    return ch;
  });
  const newRate = money(f.rate);
  if (f.rate !== undefined && newRate !== "" && String(res.before.Rate) !== String(newRate)) logRateChange(name, res.before.Rate, newRate, res.after["Unit of rate"], "Hero update", by, "");
  cache.del("vendors");
  return { product: res.after };
}
async function handleAddProduct(p) {
  // purana action — ab sirf Hero
  return handleHeroSaveProduct(Object.assign({}, p, { row: null, fields: {
    name: p.name, category: p.category, unit: p.unit, packaging: p.packaging, min: p.minStock, max: p.maxStock,
    baseUnit: p.baseUnit, suq: p.stockUnitQty, stockType: p.stockType,
  } }));
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
  const { after } = await mutateProduct({ name: p.itemName }, (b, C) => ({ [C.VENDOR]: String(p.vendor || "").trim() }));
  return { itemName: p.itemName, vendor: p.vendor, product: after };
}

// =================== PO BATCHES ===================
async function readPOBatches() {
  const pos = (await cached("po", readPO)).filter(po => po.Status !== "Cancelled" && String(po["PO ID"] || "").indexOf("MPO-") !== 0 && String(po.Status || "").indexOf("Received") !== 0);
  const batches = {};
  pos.forEach(po => {
    const vendor = po.Vendor || "Unassigned"; const batchId = po["Batch ID"] || "Legacy";
    const key = vendor + "|" + batchId;
    if (!batches[key]) batches[key] = { vendor, batchId, items: [], pdfUrl: "" };
    batches[key].items.push({ row: po._row, name: po["Item Name"], unit: po.Unit, qty: po["Qty Ordered"], autoQty: po["Auto Qty"] === undefined ? "" : po["Auto Qty"], addonQty: po["Add-on Qty"] === undefined ? "" : po["Add-on Qty"], addonUnit: po["Add-on Unit"] || "", note: po["Item Note"] || "", status: po.Status, timestamp: po.Timestamp });
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

// =================== 🆕 MESSAGES (Admin → Staff, staff ko reply karna zaroori) ===================
// "Messages" sheet apne aap banti hai: har recipient ki ek row.
const MSG_HDR = ["Msg ID", "Timestamp", "From", "From ID", "To ID", "To Name", "Message", "Status", "Reply", "Replied At", "Seen At"];
const MSG_SENDERS = ["Admin", "Hero", "SuperAdmin"];
async function appendRows(sheetName, rows) {
  if (!rows.length) return;
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID, range: `'${sheetName}'!A1`, valueInputOption: "RAW", insertDataOption: "INSERT_ROWS",
    requestBody: { values: rows },
  });
}
async function readMessages() {
  const m = await sheetMeta();
  if (!m[S.MSG]) return [];
  const objs = rowsToObjects(await readRange(S.MSG, "A1:K", "FORMATTED_VALUE"));
  objs.forEach(o => { o.Timestamp = normalizeDateTime(o.Timestamp); o["Replied At"] = normalizeDateTime(o["Replied At"]); o["Seen At"] = normalizeDateTime(o["Seen At"]); o.Status = o.Status || "Pending"; });
  return objs.filter(o => o["Msg ID"]).reverse();
}
async function handleSendMessage(p) {
  await requireRole(p, MSG_SENDERS, "Message bhejna");
  const me = await staffOf(p);
  const text = String(p.message || "").trim().slice(0, 1000);
  if (!text) throw new Error("Message khali hai");
  const staff = await cached("staff", readStaff);
  let to = Array.isArray(p.to) ? p.to.map(x => String(x).toLowerCase().trim()) : [];
  const list = staff.filter(s => String(s.User || "").trim() && String(s.User).toLowerCase().trim() !== String(me.User).toLowerCase().trim() && (p.toAll || to.includes(String(s.User).toLowerCase().trim())));
  if (!list.length) throw new Error("Kam se kam ek staff chuno");
  await ensureSheet(S.MSG, MSG_HDR);
  const ts = toISTISO(new Date());
  const base = "MSG-" + Date.now();
  await appendRows(S.MSG, list.map((s, i) => [base + "-" + i, ts, me.Name, me.User, s.User, s.Name, text, "Pending", "", "", ""]));
  invalidate("messages");
  return { sent: list.length, to: list.map(s => s.Name) };
}
async function handleGetMessages(p) {
  const role = await roleOf(p);
  const all = await cached("messages", readMessages);
  if (MSG_SENDERS.includes(role) && !p.mine) return all.slice(0, 1000);
  const uid = String(p._uid || p.userId || "").toLowerCase().trim();
  if (!uid) return [];
  return all.filter(m => String(m["To ID"] || "").toLowerCase().trim() === uid).slice(0, 200);
}
async function msgRowFor(p) {
  const row = Number(p.row); if (!row || row < 2) throw new Error("Invalid message");
  const r = (await readRange(S.MSG, `A${row}:K${row}`, "FORMATTED_VALUE"))[0] || [];
  const uid = String(p._uid || "").toLowerCase().trim();
  if (!uid || String(r[4] || "").toLowerCase().trim() !== uid) throw new Error("Ye message aapke liye nahi hai");
  return { row, r };
}
async function handleReplyMessage(p) {
  const reply = String(p.reply || "").trim().slice(0, 1000);
  if (!reply) throw new Error("Reply likho");
  const { row, r } = await msgRowFor(p);
  const now = toISTISO(new Date());
  await batchWrite(S.MSG, [{ row, col: 8, value: "Replied" }, { row, col: 9, value: reply }, { row, col: 10, value: now }].concat(r[10] ? [] : [{ row, col: 11, value: now }]));
  invalidate("messages");
  return { row, reply, repliedAt: now };
}
async function handleMarkMessageSeen(p) {
  const { row, r } = await msgRowFor(p);
  if (!r[10]) { await batchWrite(S.MSG, [{ row, col: 11, value: toISTISO(new Date()) }]); cache.del("messages"); }
  return { row };
}

// =================== 🆕 COMPLAINTS (Staff → Admin/Hero). Staff sirf apne dekh sakta hai ===================
const COMPLAINT_HDR = ["Ticket ID", "Timestamp", "From", "From ID", "Role", "Type", "Message", "Status", "Admin Reply", "Resolved By", "Resolved At"];
const COMPLAINT_VIEWERS = ["Admin", "Hero"];
const COMPLAINT_TYPES = ["Shikayat", "Request", "Sujhav", "Maal / Stock", "Attendance", "Aur kuch"];
async function readComplaints() {
  const m = await sheetMeta();
  if (!m[S.COMPLAINT]) return [];
  const objs = rowsToObjects(await readRange(S.COMPLAINT, "A1:K", "FORMATTED_VALUE"));
  objs.forEach(o => { o.Timestamp = normalizeDateTime(o.Timestamp); o["Resolved At"] = normalizeDateTime(o["Resolved At"]); o.Status = o.Status || "Open"; });
  return objs.filter(o => o["Ticket ID"]).reverse();
}
async function handleSendComplaint(p) {
  const me = await staffOf(p);
  if (!me) throw new Error("Pehle login karo");
  const text = String(p.message || "").trim().slice(0, 1500);
  if (!text) throw new Error("Message likho");
  const type = COMPLAINT_TYPES.includes(p.type) ? p.type : "Aur kuch";
  await ensureSheet(S.COMPLAINT, COMPLAINT_HDR);
  const id = "TKT-" + Date.now();
  const row = await appendRow(S.COMPLAINT, [id, toISTISO(new Date()), me.Name, me.User, me.Permission, type, text, "Open", "", "", ""]);
  invalidate("complaints");
  return { id, row };
}
async function handleGetComplaints(p) {
  const role = await roleOf(p);
  const all = await cached("complaints", readComplaints);
  if (COMPLAINT_VIEWERS.includes(role)) return all.slice(0, 1500);
  const uid = String(p._uid || "").toLowerCase().trim();
  if (!uid) return [];
  // dusre staff ke messages kabhi nahi
  return all.filter(c => String(c["From ID"] || "").toLowerCase().trim() === uid).slice(0, 300);
}
async function handleUpdateComplaint(p) {
  await requireRole(p, COMPLAINT_VIEWERS, "Complaint resolve");
  const me = await staffOf(p);
  const row = Number(p.row); if (!row || row < 2) throw new Error("Invalid complaint");
  const r = (await readRange(S.COMPLAINT, `A${row}:K${row}`, "FORMATTED_VALUE"))[0] || [];
  if (!r[0]) throw new Error("Complaint nahi mili");
  const reply = String(p.reply === undefined ? "" : p.reply).trim().slice(0, 1500);
  const now = toISTISO(new Date());
  const cells = [];
  if (reply) cells.push({ row, col: 9, value: reply });
  if (p.status === "Resolved") cells.push({ row, col: 8, value: "Resolved" }, { row, col: 10, value: me.Name }, { row, col: 11, value: now });
  else if (p.status === "Open") cells.push({ row, col: 8, value: "Open" }, { row, col: 10, value: "" }, { row, col: 11, value: "" });
  if (!cells.length) throw new Error("Kuch badla nahi");
  await batchWrite(S.COMPLAINT, cells);
  invalidate("complaints");
  return { row, status: p.status || r[7], reply };
}

// =================== ACTION DISPATCHER ===================
async function dispatchAction(action, payload) {
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
  else if (action === "getRateHistory") data = await cached("ratehist", readRateHistory);
  else if (action === "getMessages") data = await handleGetMessages(payload);
  else if (action === "getComplaints") data = await handleGetComplaints(payload);
  else if (action === "sendComplaint") data = await handleSendComplaint(payload);
  else if (action === "updateComplaint") data = await handleUpdateComplaint(payload);
  else if (action === "sendMessage") data = await handleSendMessage(payload);
  else if (action === "replyMessage") data = await handleReplyMessage(payload);
  else if (action === "markMessageSeen") data = await handleMarkMessageSeen(payload);
  else if (action === "getLowStockAlerts") {
    const products = await cached("products", readProducts);
    data = products.filter(p => {
      if (isDaily(p["Stock Type"])) return false;
      const min = num(p.Min), max = num(p.Max), cur = num(p["Current Stock"]);
      return (min > 0 && cur <= min) || (max > 0 && cur >= max);
    }).map(p => ({
      "Item Name": p["Item Name (Standardized)"], Category: p.Category, Unit: p.Unit,
      "Current Stock": p["Current Stock"], "Min Stock": p.Min, "Max Stock": p.Max,
      "Status": (num(p.Min) > 0 && num(p["Current Stock"]) <= num(p.Min)) ? "LOW" : "OVERSTOCK",
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
  else if (action === "heroSaveProduct") data = await handleHeroSaveProduct(payload);
  else if (action === "addProduct") data = await handleAddProduct(payload);
  else if (action === "addWarehouseEntry") data = await addStockEntry("wh", payload);
  else if (action === "editWarehouseEntry") data = await editStockEntry("wh", payload, false);
  else if (action === "deleteWarehouseEntry") data = await editStockEntry("wh", payload, true);
  else if (action === "addManualPurchase") data = await addStockEntry("mp", payload);
  else if (action === "editManualPurchase") data = await editStockEntry("mp", payload, false);
  else if (action === "deleteManualPurchase") data = await editStockEntry("mp", payload, true);
  else if (action === "addDemand") data = await handleAddDemand(payload);
  else if (action === "approveDemands") data = await handleApproveDemands(payload);
  else if (action === "approveDemand") data = await handleApproveDemands({ ...payload, items: [{ row: payload.row, approvedQty: payload.approvedQty, feedback: payload.feedback, unit: payload.unit, rate: payload.rate, amount: payload.amount }] });
  else if (action === "rejectDemands") data = await handleRejectDemands(payload);
  else if (action === "rejectDemand") data = await handleRejectDemands({ ...payload, rows: [payload.row] });
  else if (action === "editDemand") data = await handleEditDemand(payload);
  else if (action === "deleteDemand") data = await handleDeleteDemand(payload);
  else if (action === "saveDemandFeedback") data = await handleSaveDemandFeedback(payload);
  else if (action === "addPurchaseOrder") data = await handleAddPurchaseOrder(payload);
  else if (action === "updatePOItem") data = await handleUpdatePOItem(payload);
  else if (action === "cancelPO") data = await handleCancelPO(payload);
  else if (action === "clearPOBatch") data = await handleClearPOBatch(payload);
  else if (action === "generatePOBatch") data = await handleGeneratePOBatch(payload);
  else if (action === "autoTagVendors") data = await handleAutoTagVendors();
  else if (action === "addVendor") data = await handleAddVendor(payload);
  else if (action === "setProductVendor") data = await handleSetProductVendor(payload);
  else if (action === "markAttendance") data = await handleMarkAttendance(payload);
  else if (action === "editAttendancePunch") data = await handleEditAttendancePunch(payload);
  else if (action === "addCylinderEntry") data = await handleAddCylinderEntry(payload);
  else if (action === "updateCylinderPrice") { const price = Number(payload.price) || 0; if (price <= 0) throw new Error("Invalid price"); await setSetting("CylinderPrice", price); invalidate("cyl"); data = { price }; }
  else if (action === "addStaff") data = await handleAddStaff(payload);
  else if (action === "updateStaff") data = await handleUpdateStaff(payload);
  else if (action === "deleteStaff") data = await handleDeleteStaff(payload);
  else if (action === "runAutoPO") data = await runAutoPOForAll();
  else if (action === "checkAutoPOForItem") data = await checkAndCreateAutoPO(payload.itemName);
  else if (action === "uploadPhoto") data = { url: await uploadPhotoToDrive(payload.photoBase64, payload.userId || "photo", payload.name || "photo") };
  else data = { note: "Action not implemented yet: " + action };
  return data;
}

// =================== ROUTER ===================
app.all("/", async (req, res) => {
  const payload = { ...req.query, ...req.body };
  const action = payload.action;
  if (!action) return res.json({ ok: true, service: "Peetambra Bridge Running" });
  try {
    const isRead = /^get/.test(action) || action === "poll" || action === "login" || action === "bootstrap";
    const rid = (!isRead && payload._rid) ? (action + ":" + String(payload._rid)) : "";
    let data;
    if (rid) {
      let entry = ridStore.get(rid);
      if (!entry) {
        entry = { t: Date.now(), promise: dispatchAction(action, payload) };
        ridStore.set(rid, entry);
        entry.promise.catch(() => { if (ridStore.get(rid) === entry) ridStore.delete(rid); });
      } else console.log("Duplicate request blocked:", action, payload._rid);
      data = await entry.promise;
    } else data = await dispatchAction(action, payload);
    res.json({ ok: true, data });
  } catch (e) {
    console.error("Router error on", action, ":", e.message);
    res.json({ ok: false, error: e.message });
  }
});

app.get("/", (req, res) => res.json({ ok: true, service: "Peetambra Bridge Running" }));

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => {
  console.log("🚀 Bridge running on port " + PORT);
  // Server start hote hi Product Master + Demand headers ready
  cached("products", readProducts).then(() => console.log("Product columns:", JSON.stringify(PCOL))).catch(e => console.error("Warmup:", e.message));
  ensureDemandHeaders();
});
