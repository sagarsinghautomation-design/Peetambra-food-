const express = require("express");
const { google } = require("googleapis");
const NodeCache = require("node-cache");
const cors = require("cors");

const app = express();
app.use(cors());
// 👇 YEH LINE ZAROORI HAI - Form data aur URL-encoded data ke liye
app.use(express.urlencoded({ extended: true }));
app.use(express.json({ limit: "20mb" }));

const cache = new NodeCache({ stdTTL: 120, checkperiod: 30 });

const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT),
  scopes: ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"],
});
const sheets = google.sheets({ version: "v4", auth });
const SHEET_ID = process.env.SHEET_ID;

const S = {
  PRODUCTS: "Product & Stock Master", ATTENDANCE: "Attendence Master",
  DEMAND: "Demand For Canteen", PO: "Purchase Orders",
  STAFF: "Staff & Permissions", WHLOG: "Warehouse Entry Log",
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
  const hit = cache.get(key); if (hit) return hit;
  const data = await builder(); cache.set(key, data); return data;
}
function invalidate(...keys) { keys.forEach(k => cache.del(k)); }

// =================== HANDLERS ===================
async function handleLogin(p) {
  const staff = await cached("staff", readStaff);
  const match = staff.find(r => String(r.User || "").toLowerCase().trim() === String(p.userId).toLowerCase().trim() && String(r.Password || "").trim() === String(p.password).trim());
  if (!match) throw new Error("Invalid User ID or Password");
  return { userId: match.User, name: match.Name, designation: match.Designation, shift: match["Shift Timing"], permission: match.Permission };
}

async function handleBootstrap(p) {
  const want = p.want || []; const out = {}; const tasks = [];
  if (want.includes("products")) tasks.push(cached("products", readProducts).then(d => out.products = d));
  if (want.includes("demands")) tasks.push(cached("demands", readDemands).then(d => out.demands = d));
  if (want.includes("staff")) tasks.push(cached("staff", readStaff).then(d => out.staff = d));
  if (want.includes("myAttendance") || want.includes("allAttendance")) {
    tasks.push(cached("attendance", async () => rowsToObjects(await readRange(S.ATTENDANCE))).then(d => {
      if (want.includes("myAttendance")) out.myAttendance = d.filter(r => String(r["User ID"]) === String(p.userId));
      if (want.includes("allAttendance")) out.allAttendance = d;
    }));
  }
  await Promise.all(tasks);
  return out;
}

async function handleAddDemand(p) {
  const products = await cached("products", readProducts);
  const item = products.find(x => normH(x["Item Name (Standardized)"]) === normH(p.itemName));
  if (!item) throw new Error("Item not found");
  const id = "DMD-" + Date.now();
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID, range: `'${S.DEMAND}'!A1`,
    valueInputOption: "USER_ENTERED", insertDataOption: "INSERT_ROWS",
    requestBody: { values: [[id, new Date().toISOString(), p.raisedBy, p.itemName, p.category || item.Category, p.unit || item.Unit, p.qty, "Pending", "", p.notes || "", "", "", "", ""]] }
  });
  invalidate("demands", "products");
  return { id };
}

// =================== ROUTER (Updated to handle GET & POST both) ===================
// 👇 YEH ROUTER ZAROORI HAI - Yeh GET aur POST dono ko handle karega
app.all("/", async (req, res) => {
  const payload = { ...req.query, ...req.body };
  const action = payload.action;

  if (!action) return res.json({ ok: true, service: "Peetambra Bridge Running" });

  try {
    let data;
    if (action === "login") data = await handleLogin(payload);
    else if (action === "bootstrap") data = await handleBootstrap(payload);
    else if (action === "getProducts") data = await cached("products", readProducts);
    else if (action === "getStaff") data = await cached("staff", readStaff);
    else if (action === "getDemands") data = await cached("demands", readDemands);
    else if (action === "getAttendance") {
        let att = await cached("attendance", async () => rowsToObjects(await readRange(S.ATTENDANCE)));
        if (payload.userId) att = att.filter(r => String(r["User ID"]) === String(payload.userId));
        data = att;
    }
    else if (action === "addDemand") data = await handleAddDemand(payload);
    else data = { note: "Action not implemented yet: " + action };

    res.json({ ok: true, data });
  } catch (e) {
    res.json({ ok: false, error: e.message });
  }
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log("🚀 Bridge running on port " + PORT));
