const express = require("express");
const { google } = require("googleapis");
const NodeCache = require("node-cache");
const cors = require("cors");

const app = express();
app.use(cors());
app.use(express.json({ limit: "20mb" }));

// 2 minute ka RAM cache (isse speed 10x ho jayegi)
const cache = new NodeCache({ stdTTL: 120, checkperiod: 30 });

// Google Auth Setup
const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT),
  scopes: [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
  ],
});
const sheets = google.sheets({ version: "v4", auth });

const SHEET_ID = process.env.SHEET_ID;
const APPS_SCRIPT_URL = process.env.APPS_SCRIPT_URL;

const S = {
  PRODUCTS: "Product & Stock Master",
  ATTENDANCE: "Attendence Master",
  DEMAND: "Demand For Canteen",
  PO: "Purchase Orders",
  STAFF: "Staff & Permissions",
  WHLOG: "Warehouse Entry Log",
  CYLINDER: "Cylinder Entry",
  SETTINGS: "Settings",
};

function num(v) { const n = Number(v); return isNaN(n) ? 0 : n; }
function round3(v) { return Math.round(v * 1000) / 1000; }
function normH(s) { return String(s).toLowerCase().replace(/\s+/g, " ").trim(); }

async function readRange(sheetName, range = "A1:ZZ5000") {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `'${sheetName}'!${range}`,
    valueRenderOption: "UNFORMATTED_VALUE",
  });
  return res.data.values || [];
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

// =================== READERS (Cached) ===================
async function readStaff() { return rowsToObjects(await readRange(S.STAFF)); }
async function readDemands() { 
  const rows = await readRange(S.DEMAND, "A1:N500");
  return rows.slice(1).reverse().map((r, i) => ({
    _row: rows.length - i, "Demand ID": r[0], Timestamp: r[1], "Raised By": r[2], "Item Name": r[3],
    Category: r[4], Unit: r[5], "Qty Requested": r[6], Status: r[7] || "Pending",
    "Warehouse Stock After": r[8], Notes: r[9], "Approved Qty": r[10], "Approved At": r[11] || "",
  }));
}

async function readProducts() {
  const rows = await readRange(S.PRODUCTS);
  if (rows.length < 2) return [];
  const H = rows[0].map(normH);
  const find = (pred, def) => { for (let i = 0; i < H.length; i++) if (H[i] && pred(H[i])) return i; return def; };
  const C = { NAME: find(x => x.indexOf("item name") === 0, 1), CAT: find(x => x === "category", 2), OPENFULL: find(x => x.indexOf("full pack") !== -1, 5), OPENLOOSE: find(x => x.indexOf("loose") !== -1, 6), DEMAND: find(x => x.indexOf("deman") === 0 && x.indexOf("approved") === -1 && x.indexOf("unit") === -1, 7), APPROVED: find(x => x.indexOf("approved") !== -1, 9), PURCHASE: find(x => x.indexOf("purchase") !== -1, 10), MIN: find(x => x === "min", 13), MAX: find(x => x === "max", 14), VENDOR: find(x => x === "vendor", 15), UNIT: find(x => x === "unit", 4), STOCKUNITQTY: find(x => x.indexOf("stock unit") !== -1 || x.indexOf("base qty") !== -1, 20) };
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i]; if (!r || !String(r[C.NAME] || "").trim()) continue;
    const openFull = num(r[C.OPENFULL]), openLoose = num(r[C.OPENLOOSE]), purchase = num(r[C.PURCHASE]), approved = num(r[C.APPROVED]);
    const suq = num(r[C.STOCKUNITQTY]) || 1;
    const opening = round3(openFull * suq + openLoose);
    const currentPcs = round3(opening + purchase - approved);
    out.push({ _row: i + 1, "Item Name (Standardized)": String(r[C.NAME]).trim(), Category: r[C.CAT] || "", Unit: r[C.UNIT] || "", "W/H Opening Full Pack": openFull, "W/H Opening Loose": openLoose, "Deman for canteen": num(r[C.DEMAND]), "Approved Demand Qty by supervisor": approved, "Purchase Stock  QTy": purchase, "Current Stock": currentPcs, "Current Stock In Carton": round3(currentPcs / suq), Min: r[C.MIN], Max: r[C.MAX], Vendor: r[C.VENDOR] || "" });
  }
  return out;
}

async function cached(key, builder) {
  const hit = cache.get(key);
  if (hit) return hit;
  const data = await builder();
  cache.set(key, data);
  return data;
}
function invalidate(...keys) { keys.forEach(k => cache.del(k)); }

// =================== ROUTES ===================
app.get("/", (req, res) => res.json({ ok: true, service: "Peetambra Bridge Running" }));

app.post("/login", async (req, res) => {
  try {
    const { userId, password } = req.body;
    const staff = await cached("staff", readStaff);
    const match = staff.find(r => String(r.User || "").toLowerCase().trim() === String(userId).toLowerCase().trim() && String(r.Password || "").trim() === String(password).trim());
    if (!match) return res.json({ ok: false, error: "Invalid User ID or Password" });
    res.json({ ok: true, data: { userId: match.User, name: match.Name, designation: match.Designation, shift: match["Shift Timing"], permission: match.Permission } });
  } catch (e) { res.json({ ok: false, error: e.message }); }
});

app.post("/bootstrap", async (req, res) => {
  try {
    const { want = [], userId } = req.body;
    const out = {};
    const tasks = [];
    if (want.includes("products")) tasks.push(cached("products", readProducts).then(d => out.products = d));
    if (want.includes("demands")) tasks.push(cached("demands", readDemands).then(d => out.demands = d));
    if (want.includes("staff")) tasks.push(cached("staff", readStaff).then(d => out.staff = d));
    if (want.includes("myAttendance")) tasks.push(cached("attendance", async () => rowsToObjects(await readRange(S.ATTENDANCE))).then(d => out.myAttendance = d.filter(r => String(r["User ID"]) === String(userId))));
    await Promise.all(tasks);
    res.json({ ok: true, data: out });
  } catch (e) { res.json({ ok: false, error: e.message }); }
});

app.get("/products", async (req, res) => { res.json({ ok: true, data: await cached("products", readProducts) }); });
app.get("/demands", async (req, res) => { res.json({ ok: true, data: await cached("demands", readDemands) }); });
app.get("/staff", async (req, res) => { res.json({ ok: true, data: await cached("staff", readStaff) }); });

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log("🚀 Bridge running on port " + PORT));