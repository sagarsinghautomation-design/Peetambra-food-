const express = require("express");
const { google } = require("googleapis");
const NodeCache = require("node-cache");
const cors = require("cors");

const app = express();
app.use(cors());
app.use(express.urlencoded({ extended: true }));
app.use(express.json({ limit: "20mb", type: ['application/json', 'text/plain'] }));

const cache = new NodeCache({ stdTTL: 120, checkperiod: 30 });

const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT),
  scopes: ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"],
});
const sheets = google.sheets({ version: "v4", auth });
const SHEET_ID = process.env.SHEET_ID;

const S = {
  PRODUCTS: "Product & Stock Master", ATTENDANCE: "Attendance", DEMAND: "Demand",
  PO: "Purchase Orders", STAFF: "Staff", WHLOG: "Warehouse Entry Log",
  CYLINDER: "Cylinder Entry", SETTINGS: "Settings",
};

function num(v) { const n = Number(v); return isNaN(n) ? 0 : n; }
function round3(v) { return Math.round(v * 1000) / 1000; }
function normH(s) { return String(s).toLowerCase().replace(/\s+/g, " ").trim(); }

async function readRange(sheetName, range = "A1:ZZ5000") {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: `'${sheetName}'!${range}`,
    valueRenderOption: "UNFORMATTED_VALUE",
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
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID, range: `'${sheetName}'!A1`,
    valueInputOption: "USER_ENTERED", insertDataOption: "INSERT_ROWS",
    requestBody: { values: [values] },
  });
}
function colLetter(n) {
  let s = "";
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
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

// =================== READERS ===================
async function readStaff() { return rowsToObjects(await readRange(S.STAFF)); }
async function readDemands() {
  const rows = await readRange(S.DEMAND, "A1:N500");
  return rows.slice(1).reverse().map((r, i) => ({
    _row: rows.length - i, "Demand ID": r[0], Timestamp: r[1], "Raised By": r[2], "Item Name": r[3],
    Category: r[4], Unit: r[5], "Qty Requested": r[6], Status: r[7] || "Pending",
    "Warehouse Stock After": r[8], Notes: r[9], "Approved Qty": r[10], "Approved At": r[11] || "",
    Feedback: r[12] || "", "Feedback At": r[13] || "",
  }));
}
async function readProducts() {
  const rows = await readRange(S.PRODUCTS);
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
      Rate: r[C.RATE], "Unit of rate": r[C.RATEUNIT] || "", "Expiry Date": r[C.EXPIRY] || "",
      "Base Unit": r[C.BASEUNIT] || r[C.UNIT] || "", "1 Stock Unit = (Base Qty)": suq,
      "Stock Type": r[C.STOCKTYPE] || "Stocked",
    });
  }
  return out;
}
async function readWarehouseEntries() {
  const rows = await readRange(S.WHLOG);
  return rowsToObjects(rows).slice(-200).reverse();
}
async function readCylinderEntries() {
  const rows = await readRange(S.CYLINDER);
  return rowsToObjects(rows).slice(-200).reverse();
}
async function readPO() { return rowsToObjects(await readRange(S.PO)); }
async function readManualPurchases() {
  const rows = await readRange(S.PO);
  return rowsToObjects(rows).filter(r => String(r["PO ID"]).indexOf("MPO-") === 0).slice(-400).reverse();
}
async function readAttendance() { return rowsToObjects(await readRange(S.ATTENDANCE)); }
async function readVendors() {
  const products = await cached("products", readProducts);
  const names = new Set(Object.keys(VENDOR_ITEMS));
  products.forEach((p) => { const v = String(p.Vendor || "").trim(); if (v) names.add(v); });
  const extra = await getSetting("ExtraVendors", "");
  if (extra) { try { JSON.parse(extra).forEach(v => { if (v) names.add(v); }); } catch (e) {} }
  return [...names].sort();
}
async function readSettings() {
  const rows = await readRange(S.SETTINGS);
  const map = {};
  rowsToObjects(rows).forEach(r => { map[String(r.Key).trim()] = r.Value; });
  return map;
}
async function getSetting(key, def) {
  const map = await cached("settings", readSettings);
  const v = map[String(key).trim()];
  return v === undefined ? def : v;
}
async function setSetting(key, val) {
  const sh = S.SETTINGS;
  const rows = await readRange(sh);
  const objs = rowsToObjects(rows);
  const match = objs.find(r => String(r.Key).trim() === key);
  if (match) await writeCell(sh, match._row, 2, val);
  else await appendRow(sh, [key, val]);
  cache.del("settings");
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

async function cached(key, builder) {
  const hit = cache.get(key); if (hit) return hit;
  const data = await builder(); cache.set(key, data); return data;
}
function invalidate(...keys) { keys.forEach(k => cache.del(k)); }

// =================== HANDLERS ===================
async function handleLogin(p) {
  const staff = await cached("staff", readStaff);
  const match = staff.find(r => String(r.User || "").toLowerCase().trim() === String(p.userId).toLowerCase().trim() && String(r.Password || "").trim() === String(p.password).trim());
  if (!match) throw new Error("Invalid User ID or Password");
  const result = {
    userId: match.User, name: match.Name, designation: match.Designation,
    shift: match["Shift Timing"], permission: match.Permission,
  };
  try {
    const want = [];
    const perm = result.permission;
    if (perm === "SuperAdmin") want.push("staff", "allAttendance");
    else {
      want.push("myAttendance");
      if (perm === "Admin") want.push("stats", "products", "demands");
      else if (perm === "Supervisor") want.push("demands", "products", "vendors");
      else if (perm === "Cashier") want.push("products", "demands");
    }
    result.bootstrap = await handleBootstrap({ want, userId: result.userId });
  } catch (e) { console.error("Bootstrap error:", e.message); result.bootstrap = {}; }
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
      if (want.includes("myAttendance")) out.myAttendance = d.filter(r => String(r["User ID"]) === String(p.userId));
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
  demands.forEach(d => {
    if (d.Status === "Pending") {
      pending++;
      if (!newest || new Date(d.Timestamp) > new Date(newest.Timestamp)) newest = d;
    }
  });
  return { ver: Date.now(), pending, newest: newest ? { item: newest["Item Name"], qty: newest["Qty Requested"], unit: newest.Unit, by: newest["Raised By"] } : null };
}
async function buildDashboardStats() {
  const products = (await cached("products", readProducts)).filter(p => String(p["Item Name (Standardized)"] || "").trim() !== "");
  const attendance = (await cached("attendance", readAttendance)).slice(-700);
  const demands = await cached("demands", readDemands);
  const pos = await cached("po", readPO);
  const staff = await cached("staff", readStaff);
  const tz = "Asia/Kolkata";
  const today = new Date().toISOString().slice(0,10);
  const presentIds = new Set(attendance.filter(a => String(a.Date || "").slice(0,10) === today && a["IN Time"]).map(a => a["User ID"]));
  const totalStaff = staff.filter(s => s.Permission !== "Admin" && s.Permission !== "SuperAdmin").length;
  const health = { healthy: 0, low: 0, overstock: 0 };
  const categoryStock = {};
  products.forEach(p => {
    const min = num(p.Min), max = num(p.Max), cur = num(p["Current Stock In Carton"]);
    const cat = p.Category || "Other";
    categoryStock[cat] = (categoryStock[cat] || 0) + num(p["Current Stock"]);
    if (min > 0 && cur <= min) health.low++;
    else if (max > 0 && cur >= max) health.overstock++;
    else health.healthy++;
  });
  const trend = [];
  for (let i = 6; i >= 0; i--) {
    const d = new Date(); d.setDate(d.getDate() - i);
    const key = d.toISOString().slice(0,10);
    const ids = new Set(attendance.filter(a => String(a.Date || "").slice(0,10) === key && a["IN Time"]).map(a => a["User ID"]));
    trend.push({ date: key, present: ids.size });
  }
  const poStatus = { Pending: 0, Ordered: 0, Received: 0, Cancelled: 0 };
  pos.forEach(p => { const s = String(p.Status || "").trim(); if (poStatus.hasOwnProperty(s)) poStatus[s]++; else if (s.indexOf("Received") === 0) poStatus.Received++; });
  const topDemandItems = products.map(p => ({ name: p["Item Name (Standardized)"], demand: num(p["Deman for canteen"]) })).filter(p => p.demand > 0).sort((a,b) => b.demand - a.demand).slice(0,5);
  return {
    totalStaff, presentToday: presentIds.size, absentToday: Math.max(totalStaff - presentIds.size, 0),
    totalProducts: products.length, lowStockCount: health.low,
    pendingDemands: demands.filter(d => d.Status === "Pending").length,
    pendingPOs: pos.filter(p => p.Status === "Pending").length,
    attendanceTrend: trend, categoryStock,
    lateToday: attendance.filter(a => String(a.Date || "").slice(0,10) === today && a["IN Status"] === "Late").length,
    stockHealth: health, poStatus, topDemandItems,
  };
}

async function handleAddDemand(p) {
  const products = await cached("products", readProducts);
  const item = products.find(x => normH(x["Item Name (Standardized)"]) === normH(p.itemName));
  if (!item) throw new Error("Item not found");
  const id = "DMD-" + Date.now();
  await appendRow(S.DEMAND, [id, new Date().toISOString(), p.raisedBy, p.itemName, p.category || item.Category, p.unit || item.Unit, p.qty, "Pending", "", p.notes || "", "", "", "", ""]);
  if (item._row) {
    const newDemand = num(item["Deman for canteen"]) + num(p.qty);
    await writeCell(S.PRODUCTS, item._row, 8, newDemand);
  }
  invalidate("demands", "products", "dash");
  return { id, row: 0 };
}
async function handleApproveDemand(p) {
  const row = Number(p.row); const approved = num(p.approvedQty);
  await writeCell(S.DEMAND, row, 8, "Approved");
  await writeCell(S.DEMAND, row, 11, approved);
  await writeCell(S.DEMAND, row, 12, new Date().toISOString());
  invalidate("demands", "products", "dash");
  return { ok: true };
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
    else if (action === "getLowStockAlerts") { const products = await cached("products", readProducts); data = products.filter(p => { const min = num(p.Min), max = num(p.Max), cur = num(p["Current Stock In Carton"]); return (min > 0 && cur <= min) || (max > 0 && cur >= max); }).map(p => ({ "Item Name": p["Item Name (Standardized)"], Category: p.Category, Unit: p.Unit, "Current Stock": p["Current Stock"], "Min Stock": p.Min, "Max Stock": p.Max, "Status": (num(p.Min) > 0 && num(p["Current Stock In Carton"]) <= num(p.Min)) ? "LOW" : "OVERSTOCK" })); }
    else if (action === "getWarehouseEntries") data = await cached("whlog", readWarehouseEntries);
    else if (action === "getCylinderEntries") data = await cached("cyl", readCylinderEntries);
    else if (action === "getCylinderPrice") data = Number(await getSetting("CylinderPrice", "2900")) || 2900;
    else if (action === "getPOBatches") data = await cached("po_batches", readPOBatches);
    else if (action === "getPurchaseOrders") data = await cached("po", readPO);
    else if (action === "getManualPurchases") data = await cached("mpo", readManualPurchases);
    else if (action === "getAttendance") {
      let att = await cached("attendance", readAttendance);
      if (payload.userId) att = att.filter(r => String(r["User ID"]) === String(payload.userId));
      data = att;
    }
    else if (action === "addDemand") data = await handleAddDemand(payload);
    else if (action === "approveDemand") data = await handleApproveDemand(payload);
    else data = { note: "Action not implemented yet: " + action };
    res.json({ ok: true, data });
  } catch (e) {
    console.error("Router error on action", action, ":", e.message);
    res.json({ ok: false, error: e.message });
  }
});
async function readPOBatches() {
  const poSh = S.PO;
  const pos = (await cached("po", readPO)).filter(po => po.Status !== "Cancelled");
  const batches = {};
  pos.forEach(po => {
    const vendor = po.Vendor || "Unassigned";
    const batchId = po["Batch ID"] || "Legacy";
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

app.get("/", (req, res) => res.json({ ok: true, service: "Peetambra Bridge Running" }));

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log("🚀 Bridge running on port " + PORT));
