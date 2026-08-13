// OPB QR check-in backend — tiny Express app on Azure App Service (Linux, Node 20).
//
// The browser PWA scans a QR, extracts the order number, and POSTs it here. This
// server holds the ONLY sensitive credential (a delegated Graph refresh token for
// the personal-OneDrive workbook owner) server-side, enforces all business rules,
// and patches Status/DateTime in the Excel Online workbook via Microsoft Graph.
//
// Security posture mirrors the NorkappTrip proxy: CORS fails closed to the allowed
// origin(s), per-IP rate limits stop scripted abuse, and framework details are hidden.
// Volunteer/admin access is gated by social sign-in (Google / Microsoft) verified in
// auth.js against an email allowlist; the scan and sheet endpoints require a session.

import express from "express";
import rateLimit from "express-rate-limit";
import {
  getAccessToken, workbookBase, openSession, closeSession, readTable, patchRow,
  listWorksheets, firstTableName, readWorksheetUsedRange, listWorkbooks,
  addTableRows, tableHeaders, ensureLogTable,
  readUsedRangeRaw, writeRange, deleteTable, createTable, colLetter,
} from "./graph.js";
import { evaluateScan, nowInZone, normalizeCutoff, normalizeEventDate, passDateYMD } from "./rules.js";
import { verifyProviderToken, resolveRoleMerged, issueSession, requireAuth } from "./auth.js";
import { listStoredUsers, upsertUser, removeUser } from "./userstore.js";
import { getConfig, setConfig } from "./configstore.js";

const TABLE_NAME = process.env.TABLE_NAME || "Table1";
const TZ = process.env.TZ_NAME || "Europe/Oslo";
// Meal cut-off (CET, HHmm): strictly after this only Dinner passes are valid, else Lunch.
// Ported from OPB_Excel_QRCodeScannerFlow (was 1400; OPB now uses 1600). Configurable.
const DEFAULT_SESSION_CUTOFF = normalizeCutoff(process.env.SESSION_CUTOFF || 1600);
// Default false = apply the flow's date + meal validation (a pass is valid only for its own
// date and the current meal window). Set SHEET_SCOPED=true to fall back to order-only matching.
const SHEET_SCOPED = /^(1|true|yes)$/i.test(process.env.SHEET_SCOPED || "");

async function validationSettings() {
  const cfg = await getConfig();
  const cutoff = normalizeCutoff(cfg.cutoff || process.env.SESSION_CUTOFF || DEFAULT_SESSION_CUTOFF);
  const eventDate = normalizeEventDate(cfg.eventDate || process.env.EVENT_DATE || null);
  return { cutoff, eventDate };
}

// Friendly workbook (spreadsheet) name derived from the env-configured locator.
function workbookName() {
  const loc = process.env.GRAPH_WORKBOOK || "";
  const raw = loc.startsWith("path:") ? loc.slice(5) : loc;
  const base = raw.split("/").filter(Boolean).pop() || "";
  return base.replace(/\.xlsx$/i, "") || "Workbook";
}

// The workbook to operate on: the admin-selected file from config, else the env default.
async function wbContext() {
  const cfg = await getConfig();
  if (cfg.workbook && cfg.workbook.id) return { loc: "id:" + cfg.workbook.id, name: cfg.workbook.name || "Workbook" };
  return { loc: undefined, name: workbookName() };
}

const ALLOWED = (process.env.ALLOWED_ORIGINS || "*")
  .split(",").map((s) => s.trim()).filter(Boolean);

function applyCors(req, res) {
  const origin = req.headers.origin || "";
  let allow = null;
  if (ALLOWED.includes("*")) allow = "*";
  else if (origin && ALLOWED.includes(origin)) allow = origin;
  if (allow) {
    res.set("Access-Control-Allow-Origin", allow);
    res.set("Vary", "Origin");
    res.set("Access-Control-Allow-Methods", "POST, GET, OPTIONS");
    res.set("Access-Control-Allow-Headers", "Content-Type, Authorization");
    res.set("Access-Control-Max-Age", "600");
  }
}

const app = express();
app.set("trust proxy", 1);
app.disable("x-powered-by");
app.use(express.json({ limit: "64kb" }));
app.use((req, res, next) => {
  applyCors(req, res);
  // Defence-in-depth headers. This is a JSON API (never framed, no referrer needed);
  // responses under /api carry guest PII, so forbid any shared/intermediary caching.
  res.setHeader("X-Content-Type-Options", "nosniff");
  res.setHeader("X-Frame-Options", "DENY");
  res.setHeader("Referrer-Policy", "no-referrer");
  res.setHeader("Cross-Origin-Opener-Policy", "same-origin");
  res.setHeader("Strict-Transport-Security", "max-age=31536000; includeSubDomains");
  if (req.path.startsWith("/api/")) res.setHeader("Cache-Control", "no-store");
  if (req.method === "OPTIONS") return res.status(204).end();
  next();
});

function clientKey(req) {
  let ip = req.ip || "";
  ip = ip.replace(/^\[(.+)\]:\d+$/, "$1");
  if (/^\d{1,3}(\.\d{1,3}){3}:\d+$/.test(ip)) ip = ip.replace(/:\d+$/, "");
  return ip || "unknown";
}
const rlOpts = { standardHeaders: true, legacyHeaders: false, keyGenerator: clientKey,
  message: { error: "Too many requests — slow down and retry shortly." } };
app.use(rateLimit({ windowMs: 60_000, max: 120, ...rlOpts }));
const scanLimiter = rateLimit({ windowMs: 60_000, max: 60, ...rlOpts });
// Tight limiter for sign-in; only FAILED attempts count so a valid user is never blocked.
const authLimiter = rateLimit({ windowMs: 60_000, max: 15, skipSuccessfulRequests: true, ...rlOpts });

app.get("/health", (_req, res) => res.json({ ok: true }));

// POST /api/auth { provider, credential } -> { token, name, role, email }
// Verifies the Google/Microsoft ID token, checks the email allowlist, issues a session.
app.post("/api/auth", authLimiter, async (req, res) => {
  const { provider, credential } = req.body || {};
  if (!provider || !credential) return res.status(400).json({ error: "provider and credential required" });
  try {
    const { email, name } = await verifyProviderToken(provider, credential);
    const role = await resolveRoleMerged(email);
    if (!role) return res.status(403).json({ error: "This account is not approved for check-in access." });
    const token = await issueSession({ email, name, role });
    res.json({ token, name, role, email });
  } catch (e) {
    console.warn("auth failed:", e?.message || e);
    res.status(401).json({ error: "Sign-in could not be verified." });
  }
});

// GET /api/me -> current session (used by the SPA to restore a session on reload)
app.get("/api/me", requireAuth(), (req, res) => res.json(req.user));

// ---- shared check-in config (admin-set, everyone reads) ----
// scanSheet = active event session; guestSheets = worksheets volunteers may view.
app.get("/api/config", requireAuth(), async (_req, res) => {
  try { res.json(await getConfig()); }
  catch (e) { res.status(500).json({ error: "Could not load configuration." }); }
});
app.post("/api/config", requireAuth("admin"), async (req, res) => {
  const patch = {};
  if ("workbook" in (req.body || {})) {
    const w = req.body.workbook;
    patch.workbook = w && typeof w.id === "string" ? { id: w.id, name: String(w.name || "") } : null;
  }
  if ("scanSheet" in (req.body || {})) patch.scanSheet = String(req.body.scanSheet || "");
  if ("guestSheets" in (req.body || {})) patch.guestSheets = Array.isArray(req.body.guestSheets) ? req.body.guestSheets.map((s) => String(s)) : [];
  if ("cutoff" in (req.body || {})) patch.cutoff = req.body.cutoff;
  if ("eventDate" in (req.body || {})) patch.eventDate = req.body.eventDate;
  try { tabsCache = { key: "", at: 0, data: null }; res.json(await setConfig(patch)); }
  catch (e) { res.status(500).json({ error: "Could not save configuration." }); }
});

// GET /api/workbooks -> { workbooks: [{id,name}] }  — Excel files in the owner's OneDrive.
app.get("/api/workbooks", requireAuth("admin"), async (_req, res) => {
  try {
    const token = await getAccessToken();
    res.json({ workbooks: await listWorkbooks(token) });
  } catch (e) {
    console.error("workbooks load failed:", e?.message || e);
    res.status(500).json({ error: "Could not list spreadsheets.", detail: e?.message || String(e) });
  }
});

// ---- Admin-only volunteer/role management ----
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const bootstrapAdmins = () => (process.env.ADMIN_EMAILS || "").split(/[,;\s]+/).map((s) => s.trim().toLowerCase()).filter(Boolean);

app.get("/api/users", requireAuth("admin"), async (req, res) => {
  try {
    const stored = await listStoredUsers();
    const storedEmails = new Set(stored.map((u) => u.email));
    const config = bootstrapAdmins().filter((e) => !storedEmails.has(e)).map((e) => ({ email: e, role: "admin", source: "config" }));
    const users = stored.map((u) => ({ email: u.email, role: u.role, source: "app", addedBy: u.addedBy || null, addedAt: u.addedAt || null }));
    res.json({ me: req.user.email, users: [...config, ...users].sort((a, b) => a.email.localeCompare(b.email)) });
  } catch (e) { res.status(500).json({ error: "Could not load users." }); }
});

app.post("/api/users", requireAuth("admin"), async (req, res) => {
  const email = String(req.body?.email || "").trim().toLowerCase();
  const role = req.body?.role === "admin" ? "admin" : "user";
  if (!EMAIL_RE.test(email)) return res.status(400).json({ error: "Enter a valid email address." });
  if (bootstrapAdmins().includes(email)) return res.status(409).json({ error: "That account is the built-in admin (managed in configuration)." });
  try { await upsertUser(email, role, req.user.email); res.json({ ok: true }); }
  catch (e) { res.status(500).json({ error: "Could not save the user." }); }
});

app.post("/api/users/remove", requireAuth("admin"), async (req, res) => {
  const email = String(req.body?.email || "").trim().toLowerCase();
  if (!email) return res.status(400).json({ error: "email required" });
  if (bootstrapAdmins().includes(email)) return res.status(409).json({ error: "The built-in admin can't be removed here." });
  if (email === req.user.email) return res.status(409).json({ error: "You can't remove your own access." });
  try { const removed = await removeUser(email); res.json({ ok: true, removed }); }
  catch (e) { res.status(500).json({ error: "Could not remove the user." }); }
});

// Resolve the Excel table that backs a worksheet (event) tab. Falls back to the
// configured TABLE_NAME for the legacy single-sheet workbook when no tab is given.
async function tableForSheet(token, base, session, sheet) {
  if (!sheet) return TABLE_NAME;
  const t = await firstTableName(token, base, session, sheet);
  if (!t) throw new Error(`Worksheet "${sheet}" has no table to check in against.`);
  return t;
}

// GET /api/tabs -> { tabs: [names] }  — the workbook's worksheets, for the UI pickers.
let tabsCache = { key: "", at: 0, data: null };
app.get("/api/tabs", requireAuth(), async (req, res) => {
  let token, base, session;
  try {
    const { loc, name } = await wbContext();
    const key = loc || "__env__";
    if (tabsCache.data && tabsCache.key === key && Date.now() - tabsCache.at < 300_000) return res.json(tabsCache.data);
    token = await getAccessToken();
    base = workbookBase(loc);
    session = await openSession(token, base);
    const tabs = await listWorksheets(token, base, session);
    const data = { tabs, workbook: name, updatedAt: new Date().toISOString() };
    tabsCache = { key, at: Date.now(), data };
    res.json(data);
  } catch (e) {
    console.error("tabs load failed:", e?.message || e);
    res.status(500).json({ error: "Could not list worksheets.", detail: e?.message || String(e) });
  } finally {
    if (token && base && session) await closeSession(token, base, session);
  }
});

// GET /api/ping-sheet?sheet=<name> -> { ok, workbook, sheet, table? }
// Lightweight connection test: can we reach the workbook and (if named) the worksheet's
// check-in table? Always 200 so the UI can render a red/green light without throwing.
app.get("/api/ping-sheet", requireAuth(), async (req, res) => {
  const sheet = String(req.query.sheet || "").trim();
  try {
    const { loc, name } = await wbContext();
    const token = await getAccessToken();
    const base = workbookBase(loc);
    if (!sheet) { await listWorksheets(token, base, null); return res.json({ ok: true, workbook: name }); }
    const table = await firstTableName(token, base, null, sheet);
    if (!table) return res.json({ ok: false, workbook: name, sheet, error: "No check-in table on that worksheet." });
    res.json({ ok: true, workbook: name, sheet, table });
  } catch (e) {
    res.json({ ok: false, sheet, error: e?.message || "Cannot reach the spreadsheet." });
  }
});

// POST /api/register { orderNumber, sheet }  ->  { response, customername, decision }
// `sheet` names the event-day worksheet to check the guest into (e.g. "Saturday Dinner").
app.post("/api/register", requireAuth(), scanLimiter, async (req, res) => {
  const raw = (req.body && req.body.orderNumber) != null ? String(req.body.orderNumber) : "";
  // Accept the raw QR too; take the part before the first ';' like the Power App did.
  const orderNumber = raw.split(";")[0].trim();
  const sheet = String(req.body?.sheet || "").trim();
  if (!orderNumber) return res.status(400).json({ error: "orderNumber required" });

  let token, base, session;
  try {
    const { loc } = await wbContext();
    token = await getAccessToken();
    base = workbookBase(loc);
    session = await openSession(token, base);

    const table = await tableForSheet(token, base, session, sheet);
    const { headers, rows } = await readTable(token, base, session, table);
    const { cutoff, eventDate } = await validationSettings();
    const result = evaluateScan({ headers, rows, orderNumber, tz: TZ, cutoff, eventDate, sheetScoped: SHEET_SCOPED, registeredBy: req.user.email, comments: "ScannedByApp" });

    let passesRegistered = result.orderRegistered || 0;
    if (result.decision === "SUCCESS") {
      for (const item of result.patch) await patchRow(token, base, session, table, item);
      passesRegistered += result.patch.length; // rows just flipped to REGISTERED
    }
    res.json({
      response: result.response, customername: result.customerName, decision: result.decision,
      sheet: sheet || table, reason: result.reason || null, session: result.session,
      passesRegistered, passesTotal: result.orderTotal || 0,
    });
  } catch (e) {
    console.error("register failed:", e?.message || e);
    res.status(500).json({ error: "Registration failed — please retry.", detail: e?.message || String(e) });
  } finally {
    if (token && base && session) await closeSession(token, base, session);
  }
});

// GET /api/lookup?q=<name or order> -> distinct orders on the active scan sheet, so a
// volunteer can register a guest (e.g. a season-pass holder) by name instead of a QR.
app.get("/api/lookup", requireAuth(), scanLimiter, async (req, res) => {
  const q = String(req.query.q || "").trim().toLowerCase();
  if (q.length < 2) return res.json({ matches: [] });
  let token, base, session;
  try {
    const { loc } = await wbContext();
    const cfg = await getConfig();
    const sheet = String(cfg.scanSheet || "").trim();
    token = await getAccessToken(); base = workbookBase(loc); session = await openSession(token, base);
    const table = await tableForSheet(token, base, session, sheet);
    const { headers, rows } = await readTable(token, base, session, table);
    const c = colFinder(headers);
    // Scope to the day being registered: real "today" if it is an event date, else the
    // configured eventDate override, else today (mirrors evaluateScan's date precedence).
    const { eventDate } = await validationSettings();
    const eventDates = new Set(rows.map((r) => (c.date !== -1 ? passDateYMD(r.values[c.date]) : null)).filter(Boolean));
    const todayYmd = nowInZone(TZ).ymd;
    const targetYmd = eventDates.has(todayYmd) ? todayYmd : (normalizeEventDate(eventDate) || todayYmd);
    const tierOf = (item) => { const s = String(item ?? ""); const m = s.match(/premium|standard/i); return m ? m[0][0].toUpperCase() + m[0].slice(1).toLowerCase() : ""; };
    const titleCase = (s) => String(s).replace(/\S+/g, (w) => w[0].toUpperCase() + w.slice(1).toLowerCase()); // collapse mixed-case values (Lunch/lunch)
    const byOrder = new Map();
    for (const r of rows) {
      const v = r.values;
      const order = String(v[c.order] ?? "").trim();
      if (!order) continue;
      const name = `${c.first !== -1 ? String(v[c.first] ?? "") : ""} ${c.last !== -1 ? String(v[c.last] ?? "") : ""}`.trim();
      if (!`${order} ${name}`.toLowerCase().includes(q)) continue;
      if (c.date !== -1 && passDateYMD(v[c.date]) !== targetYmd) continue; // only passes valid for the day
      let m = byOrder.get(order);
      if (!m) {
        if (byOrder.size >= 25) continue;
        m = { order, name, pass: c.item !== -1 ? passCategory(v[c.item]) : "", tier: c.item !== -1 ? tierOf(v[c.item]) : "", count: 0, meals: new Set(), foods: new Set() };
        byOrder.set(order, m);
      }
      m.count += c.quantity !== -1 ? (parseInt(v[c.quantity], 10) || 1) : 1; // sum the Quantity column, not rows
      if (c.meal !== -1) { const p = String(v[c.meal] ?? "").trim(); if (p) m.meals.add(titleCase(p)); }
      if (c.food !== -1) { const f = String(v[c.food] ?? "").trim(); if (f) m.foods.add(titleCase(f)); }
    }
    const matches = [...byOrder.values()].map((m) => ({
      order: m.order, name: m.name, pass: m.pass, tier: m.tier, count: m.count,
      date: targetYmd, meals: [...m.meals], foods: [...m.foods],
    }));
    res.json({ date: targetYmd, matches });
  } catch (e) {
    console.error("lookup failed:", e?.message || e);
    res.status(500).json({ error: "Lookup failed — please retry." });
  } finally {
    if (token && base && session) await closeSession(token, base, session);
  }
});

// GET /api/sheet?tabs=A,B -> { sheets: [{ name, table, headers, rows, count }], updatedAt }
// One entry per requested worksheet. With no ?tabs, reads the configured TABLE_NAME
// (legacy single-sheet mode). Short-cached per tab-set so opening the list is cheap.
// NOTE: returns guest names/PII — gated behind the volunteer/admin login.

// Run an async mapper over items with bounded concurrency (keeps Graph fast but unthrottled).
async function mapLimit(items, limit, fn) {
  const out = new Array(items.length);
  let i = 0;
  const worker = async () => { while (i < items.length) { const idx = i++; out[idx] = await fn(items[idx], idx); } };
  await Promise.all(Array.from({ length: Math.min(limit, items.length || 1) }, worker));
  return out;
}

let sheetCache = { key: "", at: 0, data: null };
app.get("/api/sheet", requireAuth(), async (req, res) => {
  const fresh = req.query.refresh === "1";
  const requested = String(req.query.tabs || "").split(",").map((s) => s.trim()).filter(Boolean);
  let token, base, session;
  try {
    const { loc } = await wbContext();
    const cacheKey = (loc || "__env__") + "|" + (requested.join("|") || "__legacy__");
    if (!fresh && sheetCache.data && sheetCache.key === cacheKey && Date.now() - sheetCache.at < 15_000)
      return res.json(sheetCache.data);
    token = await getAccessToken();
    base = workbookBase(loc);
    session = await openSession(token, base);

    const names = requested.length ? requested : [null]; // null => legacy TABLE_NAME
    const sheets = await mapLimit(names, 6, async (name) => {
      const table = name ? await firstTableName(token, base, session, name) : TABLE_NAME;
      const read = table
        ? await readTable(token, base, session, table)
        : await readWorksheetUsedRange(token, base, session, name);
      return {
        name: name || table, table: table || null,
        headers: read.headers, rows: read.rows.map((r) => r.values), count: read.rows.length,
      };
    });
    const data = { sheets, updatedAt: new Date().toISOString() };
    sheetCache = { key: cacheKey, at: Date.now(), data };
    res.json(data);
  } catch (e) {
    console.error("sheet load failed:", e?.message || e);
    res.status(500).json({ error: "Could not load the sheet.", detail: e?.message || String(e) });
  } finally {
    if (token && base && session) await closeSession(token, base, session);
  }
});

// GET /api/summary -> event-wide analytics: per-session counts + a recent check-in log.
// Aggregates every worksheet that looks like a check-in sheet (has Status + order cols).
function colFinder(headers) {
  const low = headers.map((h) => String(h ?? "").trim().toLowerCase());
  const find = (cands) => { for (const c of cands) { const i = low.indexOf(c); if (i !== -1) return i; } return -1; };
  return {
    order: find(["registrationid", "order number", "ordernumber", "order"]),
    status: find(["status"]), dt: find(["datetime", "date time"]),
    first: find(["first name", "firstname"]), last: find(["last name", "lastname"]),
    item: find(["item", "pass type", "passtype"]),
    date: find(["date"]), meal: find(["passtype", "pass type"]), food: find(["foodoption", "food option"]),
    quantity: find(["quantity", "qty"]),
  };
}
// Bucket a free-text item/pass into a tidy category for aggregation.
function passCategory(item) {
  const s = String(item ?? "").trim().toLowerCase();
  if (!s) return "Unspecified";
  if (/season/.test(s)) return "Season pass";
  if (/free/.test(s)) return "Free entry";
  if (/child|kid/.test(s)) return "Child";
  if (/entry|lunch|dinner|meal/.test(s)) return "Entry pass";
  return "Other";
}
let summaryCache = { key: "", at: 0, data: null };
app.get("/api/summary", requireAuth(), async (req, res) => {
  const fresh = req.query.refresh === "1";
  let token, base, session;
  try {
    const { loc, name } = await wbContext();
    const cfg = await getConfig();
    const scan = String(cfg.scanSheet || "").trim();
    const key = (loc || "__env__") + "|" + scan;
    if (!fresh && summaryCache.data && summaryCache.key === key && Date.now() - summaryCache.at < 30_000) return res.json(summaryCache.data);
    token = await getAccessToken();
    base = workbookBase(loc);
    session = await openSession(token, base);
    // The dashboard reflects the ACTIVE event session (config scanSheet). Only when none
    // is set do we fall back to the admin's guest sheets, then every worksheet.
    const allow = new Set((cfg.guestSheets || []).map((s) => String(s)));
    const allTabs = await listWorksheets(token, base, session);
    const tabs = scan ? [scan] : (allow.size ? allTabs.filter((t) => allow.has(t)) : allTabs);
    const sessions = []; const recent = []; const passCounts = {}; let total = 0, registered = 0;
    const byDate = {}; // date (YYYY-MM-DD) -> { valid, registered } — season-pass footfall per day
    // Read the check-in sheet(s) plus the food/parking aggregates concurrently.
    const [reads, food, parking] = await Promise.all([
      mapLimit(tabs, 6, async (t) => {
        try {
          const table = await firstTableName(token, base, session, t);
          if (!table) return null;
          const { headers, rows } = await readTable(token, base, session, table);
          const c = colFinder(headers);
          if (c.status === -1 || c.order === -1) return null; // only check-in style sheets
          return { t, rows, c };
        } catch { return null; }
      }),
      readFoodDuesSummary(token, base, session).catch(() => ({ items: [], revenue: 0, guests: 0 })),
      readParkingSummary(token, base, session).catch(() => ({ total: 0, byMake: [] })),
    ]);
    for (const rd of reads) {
      if (!rd) continue;
      const { t, rows, c } = rd;
      let tot = 0, reg = 0;
      for (const r of rows) {
        const v = r.values;
        if (String(v[c.order] ?? "").trim() === "") continue; // skip empty pre-numbered slots
        tot++;
        const ymd = c.date !== -1 ? passDateYMD(v[c.date]) : null;
        const isReg = String(v[c.status] ?? "").trim().toUpperCase() === "REGISTERED";
        if (ymd) { const b = (byDate[ymd] = byDate[ymd] || { valid: 0, registered: 0 }); b.valid++; if (isReg) b.registered++; }
        if (isReg) {
          reg++;
          const cat = passCategory(c.item !== -1 ? v[c.item] : "");
          passCounts[cat] = (passCounts[cat] || 0) + 1;
          recent.push({
            session: t, order: String(v[c.order] ?? ""), pass: cat,
            name: `${c.first !== -1 ? String(v[c.first] ?? "") : ""} ${c.last !== -1 ? String(v[c.last] ?? "") : ""}`.trim(),
            time: c.dt !== -1 ? String(v[c.dt] ?? "") : "",
          });
        }
      }
      if (tot > 0) { sessions.push({ name: t, total: tot, registered: reg }); total += tot; registered += reg; }
    }
    recent.sort((a, b) => String(b.time).localeCompare(String(a.time)));
    const byPass = Object.entries(passCounts).map(([type, count]) => ({ type, count })).sort((a, b) => b.count - a.count);
    const byDateArr = Object.entries(byDate).map(([date, x]) => ({ date, valid: x.valid, registered: x.registered })).sort((a, b) => a.date.localeCompare(b.date));
    const data = { workbook: name, sessions, byPass, byDate: byDateArr, totals: { total, registered }, recent, food, parking, updatedAt: new Date().toISOString() };
    summaryCache = { key, at: Date.now(), data };
    res.json(data);
  } catch (e) {
    console.error("summary failed:", e?.message || e);
    res.status(500).json({ error: "Could not build the dashboard.", detail: e?.message || String(e) });
  } finally {
    if (token && base && session) await closeSession(token, base, session);
  }
});

// ---- Parking + Food Stall entry (append-only writes to the workbook) ----
// These endpoints only ADD rows (and, once, create a log worksheet/table if missing).
// They never edit or delete existing data, so a mistake can't corrupt the register.
const FOOD_PRICE_SHEET = "FoodStallPriceList";
const FOOD_DUES_SHEET = "Food Stall-Dues";
const FOOD_LOG_SHEET = "FoodStallLog";
const FOOD_LOG_TABLE = "FoodStallLog";
const FOOD_LOG_HEADERS = ["DateTime", "Day", "Name", "Item", "Qty", "UnitPrice", "Amount", "RecordedBy"];
const FOOD_SETTLE_SHEET = "FoodSettlements";
const FOOD_SETTLE_TABLE = "FoodSettlements";
const FOOD_SETTLE_HEADERS = ["DateTime", "Name", "Amount", "Method", "PaidTotal", "Outstanding", "SettledBy"];
const PARKING_SHEET = "Parking";
const PARKING_TABLE = "ParkingLog";
const PARKING_HEADERS = ["Sl No", "Timestamp", "Name", "Mobile Number", "Car Registration number", "Car Make", "Car Model", "Car Colour"];
const MAXLEN = 120;
// Neutralize spreadsheet formula injection: a cell starting with = or @ (or a non-numeric
// + / -) can be executed as a formula by Excel. Prefix such text with an apostrophe so it
// is stored as literal text. Phone numbers like "+47…" (sign then digit) are left intact.
export function deFormula(s) {
  return (/^[=@\t\r]/.test(s) || /^[+\-](?![0-9])/.test(s)) ? "'" + s : s;
}
const clean = (v) => deFormula(String(v == null ? "" : v).replace(/[\r\n\t]+/g, " ").trim().slice(0, MAXLEN));

// Weekday name (e.g. "Saturday") for the event-day column, in the configured timezone.
function weekdayName(tz, when = new Date()) {
  try { return new Intl.DateTimeFormat("en-US", { timeZone: tz, weekday: "long" }).format(when); } catch { return ""; }
}

// Compose a workbook timestamp: the calendar-picked date + the current wall-clock time.
export function parkingStamp(dateYmd, clock) {
  const two = (n) => String(n).padStart(2, "0");
  return `${dateYmd} ${two(clock.h)}:${two(clock.mi)}:${two(clock.s)}`;
}

// Build a row in an existing table's own column order by matching each header to a field.
// fields = [[synonyms[], value], …]; unmatched headers get "".
export function rowForHeaders(headers, fields) {
  return headers.map((h) => {
    const low = String(h ?? "").trim().toLowerCase();
    for (const [syns, val] of fields) if (syns.some((s) => low === s || low.includes(s))) return val;
    return "";
  });
}

export function findCol(headers, cands) {
  const low = headers.map((h) => String(h ?? "").trim().toLowerCase());
  for (const c of cands) { const i = low.findIndex((h) => h.includes(c)); if (i !== -1) return i; }
  return -1;
}
// Parse a price cell that may read "120", "120,-", "120 kr", "kr 120,50".
export function parsePrice(v) {
  if (typeof v === "number") return isFinite(v) ? v : 0;
  const s = String(v ?? "").replace(/[^0-9.,]/g, "").replace(/\.(?=\d{3}\b)/g, "").replace(",", ".");
  const n = parseFloat(s); return isFinite(n) ? n : 0;
}

// Parse a price list into [{ item, price }], handling BOTH sheet orientations:
//   Vertical  — rows are items with an item column + a price column.
//   Horizontal — each column header is an item and its price sits in a row below
//                (the layout the OPB workbook actually uses).
export function parseFoodMenu(headers, dataRows) {
  const stripCur = (v) => String(v ?? "").replace(/\b(kr|nok|kroner|pris)\b|,-/gi, "");
  const looksText = (v) => /[a-z]/i.test(stripCur(v));           // an item name (letters beyond a currency word)
  const RESERVED = /^(item|dish|food|name|product|menu|particular|particulars|description|price|amount|rate|cost)$/i;
  const items = [];

  // Vertical: find a text item column and a numeric price column.
  let itemCol = findCol(headers, ["item", "dish", "food", "product", "name", "particular", "description", "menu"]);
  let priceCol = findCol(headers, ["price", "amount", "rate", "cost", "nok", "kr", "kroner", "pris"]);
  if (itemCol === -1) itemCol = headers.findIndex((_, c) => dataRows.some((r) => looksText(r[c])));
  if (priceCol === -1) priceCol = headers.findIndex((_, c) => c !== itemCol && dataRows.some((r) => parsePrice(r[c]) > 0));
  if (itemCol !== -1 && priceCol !== -1) {
    for (const r of dataRows) {
      const item = clean(r[itemCol]); const price = parsePrice(r[priceCol]);
      if (item && price > 0 && !RESERVED.test(item)) items.push({ item, price });
    }
  }

  // Horizontal fallback: header = item name, price is the first numeric cell in its column.
  if (!items.length) {
    headers.forEach((h, c) => {
      const item = clean(h);
      if (!item || !looksText(h) || RESERVED.test(item)) return;
      let price = 0;
      for (const r of dataRows) { const p = parsePrice(r[c]); if (p > 0) { price = p; break; } }
      if (price > 0) items.push({ item, price });
    });
  }

  return items;
}

// Read the price list into [{ item, price }], auto-detecting the layout.
async function readFoodMenu(token, base, session) {
  const { headers, rows } = await readWorksheetUsedRange(token, base, session, FOOD_PRICE_SHEET);
  return parseFoodMenu(headers, rows.map((r) => r.values));
}

const lc = (v) => String(v ?? "").trim().toLowerCase();

// Apply a food order to the "Food Stall-Dues" matrix (guest rows x item columns + Total).
// Pure: returns the (possibly widened) header row, the 0-based data-row index to write,
// its full row values, whether the header changed, and the row Total. priceOf(name)->number.
export function applyFoodDues(headers, dataRows, name, orders, priceOf) {
  const H = headers.map((h) => String(h ?? ""));
  const isSl = (h) => /^(sl\.?\s*no\.?|serial|s\.?\s*no\.?|#)$/i.test(String(h ?? "").trim());
  const slCol = H.findIndex(isSl);
  const nameCol = H.findIndex((h) => lc(h) === "name");
  let totalCol = H.findIndex((h) => lc(h) === "total");
  let paidCol = H.findIndex((h) => lc(h) === "paid");
  let outCol = H.findIndex((h) => ["outstanding", "due", "dues", "balance"].includes(lc(h)));
  const isMeta = (i) => i === slCol || i === nameCol || i === totalCol || i === paidCol || i === outCol;
  const itemColOf = (item) => H.findIndex((h, i) => !isMeta(i) && lc(h) !== "" && lc(h) === lc(item));

  const rows = dataRows.map((r) => (r || []).slice());
  let maxSl = 0;
  for (const r of rows) { const n = parseInt(r[slCol], 10); if (Number.isFinite(n)) maxSl = Math.max(maxSl, n); }

  // Find the guest's row: by name, else the first pre-numbered empty-name row, else append.
  let ri = nameCol === -1 ? -1 : rows.findIndex((r) => lc(r[nameCol]) === lc(name));
  if (ri === -1 && nameCol !== -1) ri = rows.findIndex((r) => lc(r[nameCol]) === "");
  if (ri === -1) { ri = rows.length; rows.push([]); }

  const row = rows[ri];
  while (row.length < H.length) row.push("");
  if (nameCol !== -1) row[nameCol] = name;
  if (slCol !== -1 && !(parseInt(row[slCol], 10) > 0)) row[slCol] = maxSl + 1;

  let headerChanged = false;
  for (const o of orders) {
    const item = String(o.item ?? "").trim(); const qty = Math.max(0, parseInt(o.qty, 10) || 0);
    if (!item || qty <= 0) continue;
    let ci = itemColOf(item);
    if (ci === -1) { H.push(item); ci = H.length - 1; row.push(""); headerChanged = true; }
    row[ci] = (parseInt(row[ci], 10) || 0) + qty;
  }
  while (row.length < H.length) row.push("");

  // Recompute Total (cumulative consumed) across every item column, excluding meta cols.
  let total = 0;
  H.forEach((h, i) => { if (!isMeta(i) && lc(h) !== "") { const q = parseInt(row[i], 10) || 0; if (q > 0) total += q * (priceOf(h) || 0); } });
  if (totalCol === -1) { H.push("Total"); totalCol = H.length - 1; row.push(""); headerChanged = true; }
  row[totalCol] = Math.round(total * 100) / 100;
  // Maintain Paid + Outstanding (= Total - Paid). Paid is only changed by settlements.
  if (paidCol === -1) { H.push("Paid"); paidCol = H.length - 1; row.push(""); headerChanged = true; }
  if (outCol === -1) { H.push("Outstanding"); outCol = H.length - 1; row.push(""); headerChanged = true; }
  const paid = parseFloat(row[paidCol]) || 0;
  row[outCol] = Math.round((row[totalCol] - paid) * 100) / 100;

  return { headers: H, headerChanged, rowIndex: ri, rowValues: row, total: row[totalCol], paid, outstanding: row[outCol] };
}

// Apply a payment to a guest's Food Stall-Dues row: Paid += amount, Outstanding = Total - Paid.
// Pure. Returns null if the guest has no dues row yet.
export function applyFoodPayment(headers, dataRows, name, amount) {
  const H = headers.map((h) => String(h ?? ""));
  const nameCol = H.findIndex((h) => lc(h) === "name");
  const totalCol = H.findIndex((h) => lc(h) === "total");
  let paidCol = H.findIndex((h) => lc(h) === "paid");
  let outCol = H.findIndex((h) => ["outstanding", "due", "dues", "balance"].includes(lc(h)));
  const rows = dataRows.map((r) => (r || []).slice());
  const ri = nameCol === -1 ? -1 : rows.findIndex((r) => lc(r[nameCol]) === lc(name));
  if (ri === -1) return null;
  let headerChanged = false;
  if (paidCol === -1) { H.push("Paid"); paidCol = H.length - 1; headerChanged = true; }
  if (outCol === -1) { H.push("Outstanding"); outCol = H.length - 1; headerChanged = true; }
  const row = rows[ri];
  while (row.length < H.length) row.push("");
  const total = totalCol !== -1 ? (parseFloat(row[totalCol]) || 0) : 0;
  const paid = Math.round(((parseFloat(row[paidCol]) || 0) + (parseFloat(amount) || 0)) * 100) / 100;
  row[paidCol] = paid;
  row[outCol] = Math.round((total - paid) * 100) / 100;
  return { headers: H, headerChanged, rowIndex: ri, rowValues: row, total, paid, outstanding: row[outCol] };
}

// Normalize a raw Parking row to [Sl No, Timestamp, Name, Mobile, Reg, Make, Model, Colour].
// App rows were written as [Timestamp, Sl No, Name, ...]; manual rows as [Sl No, "", Name, ...].
export function normalizeParkingRow(r) {
  const v = (r || []).slice(); while (v.length < 8) v.push("");
  const a = v[0];
  const isStamp = (x) => (typeof x === "number" && x > 40000) || /^\d{4}-\d{2}-\d{2}[ T]/.test(String(x)) || /^\d{1,2}\/\d{1,2}\/\d{4}/.test(String(x));
  if (isStamp(a)) return [v[1], v[0], v[2], v[3], v[4], v[5], v[6], v[7]];
  return [v[0], v[1], v[2], v[3], v[4], v[5], v[6], v[7]];
}

// Dashboard aggregate: quantity sold per item + revenue from the Food Stall-Dues matrix.
async function readFoodDuesSummary(token, base, session) {
  const raw = await readUsedRangeRaw(token, base, session, FOOD_DUES_SHEET);
  const values = raw.values || [];
  let hRow = values.findIndex((r) => r.some((c) => lc(c) === "name") && r.some((c) => lc(c) === "total"));
  if (hRow === -1) hRow = values.findIndex((r) => r.some((c) => lc(c) === "name"));
  if (hRow === -1) return { items: [], revenue: 0, guests: 0 };
  const headers = values[hRow].map((h) => String(h ?? ""));
  const isSl = (h) => /^(sl\.?\s*no\.?|serial|s\.?\s*no\.?|#)$/i.test(h.trim());
  const slCol = headers.findIndex(isSl), nameCol = headers.findIndex((h) => lc(h) === "name"), totalCol = headers.findIndex((h) => lc(h) === "total");
  const itemCols = headers.map((h, i) => ({ h, i })).filter(({ h, i }) => i !== slCol && i !== nameCol && i !== totalCol && lc(h) !== "");
  const items = itemCols.map(({ h }) => ({ item: h, qty: 0 }));
  let revenue = 0, guests = 0;
  for (const r of values.slice(hRow + 1)) {
    if (nameCol !== -1 && String(r[nameCol] ?? "").trim() === "") continue;
    guests++;
    itemCols.forEach(({ i }, k) => { const q = parseInt(r[i], 10); if (Number.isFinite(q)) items[k].qty += q; });
    if (totalCol !== -1) revenue += parsePrice(r[totalCol]);
  }
  return { items: items.filter((x) => x.qty > 0).sort((a, b) => b.qty - a.qty), revenue: Math.round(revenue * 100) / 100, guests };
}

// Dashboard aggregate: car count and breakdown by make from the Parking sheet.
async function readParkingSummary(token, base, session) {
  const raw = await readUsedRangeRaw(token, base, session, PARKING_SHEET);
  const values = raw.values || [];
  let hRow = values.findIndex((r) => r.some((c) => lc(c) === "name"));
  if (hRow === -1) return { total: 0, byMake: [] };
  const headers = values[hRow].map((h) => String(h ?? ""));
  const nameCol = headers.findIndex((h) => lc(h) === "name"), makeCol = headers.findIndex((h) => lc(h).includes("make"));
  const counts = {}; let total = 0;
  for (const r of values.slice(hRow + 1)) {
    if (nameCol !== -1 && String(r[nameCol] ?? "").trim() === "") continue;
    total++;
    const key = (makeCol !== -1 ? String(r[makeCol] ?? "").trim() : "") || "Unspecified";
    counts[key] = (counts[key] || 0) + 1;
  }
  const byMake = Object.entries(counts).map(([label, count]) => ({ label, count })).sort((a, b) => b.count - a.count);
  return { total, byMake };
}

// GET /api/foodmenu -> { workbook, items:[{item,price}] }
app.get("/api/foodmenu", requireAuth(), async (_req, res) => {
  let token, base, session;
  try {
    const { loc, name } = await wbContext();
    token = await getAccessToken(); base = workbookBase(loc); session = await openSession(token, base);
    const items = await readFoodMenu(token, base, session);
    res.json({ workbook: name, items });
  } catch (e) {
    console.error("foodmenu failed:", e?.message || e);
    res.status(500).json({ error: "Could not load the food menu.", detail: e?.message || String(e) });
  } finally { if (token && base && session) await closeSession(token, base, session); }
});

// POST /api/food-entry { name, items:[{item,qty}], day? }
// Primary store: the "Food Stall-Dues" matrix (guest rows x item columns + Total) the
// organisers use. Also appended to the FoodStallLog audit sheet for a full history.
app.post("/api/food-entry", requireAuth(), async (req, res) => {
  const name = clean(req.body?.name);
  // Accept a picked date (YYYY-MM-DD) and record its weekday in the log's Day column.
  const dateStr = clean(req.body?.date);
  const day = (/^\d{4}-\d{2}-\d{2}$/.test(dateStr) ? weekdayName(TZ, new Date(dateStr + "T12:00:00")) : clean(req.body?.day)) || weekdayName(TZ);
  const raw = Array.isArray(req.body?.items) ? req.body.items : [];
  const items = raw.map((i) => ({ item: clean(i?.item), qty: Math.max(0, Math.min(999, parseInt(i?.qty, 10) || 0)) })).filter((i) => i.item && i.qty > 0);
  if (!name) return res.status(400).json({ error: "Guest name is required." });
  if (!items.length) return res.status(400).json({ error: "Add at least one item with a quantity." });
  let token, base, session;
  try {
    const { loc } = await wbContext();
    token = await getAccessToken(); base = workbookBase(loc); session = await openSession(token, base);
    const [menu, rawRange] = await Promise.all([
      readFoodMenu(token, base, session),
      readUsedRangeRaw(token, base, session, FOOD_DUES_SHEET),
    ]);
    const priceByName = (nm) => { const hit = menu.find((m) => m.item.toLowerCase() === String(nm).toLowerCase()); return hit ? hit.price : 0; };
    const addedAmount = items.reduce((s, i) => s + priceByName(i.item) * i.qty, 0);

    // --- Primary: update the Food Stall-Dues matrix ---
    const values = rawRange.values;
    let hRow = values.findIndex((r) => r.some((c) => lc(c) === "name") && r.some((c) => lc(c) === "total"));
    if (hRow === -1) hRow = values.findIndex((r) => r.some((c) => lc(c) === "name"));
    if (hRow === -1) hRow = 0;
    const duesHeaders = (values[hRow] || []).map((h) => String(h ?? ""));
    const dataRows = values.slice(hRow + 1);
    const upd = applyFoodDues(duesHeaders, dataRows, name, items, priceByName);
    const c0 = rawRange.columnIndex, headerRow1 = rawRange.rowIndex + hRow + 1;
    const startL = colLetter(c0 + 1), endL = colLetter(c0 + upd.headers.length);
    if (upd.headerChanged) await writeRange(token, base, session, FOOD_DUES_SHEET, `${startL}${headerRow1}:${endL}${headerRow1}`, [upd.headers]);
    const rowNo = headerRow1 + 1 + upd.rowIndex;
    await writeRange(token, base, session, FOOD_DUES_SHEET, `${startL}${rowNo}:${endL}${rowNo}`, [upd.rowValues]);

    // --- Audit: append line items to FoodStallLog ---
    try {
      const table = await ensureLogTable(token, base, session, FOOD_LOG_SHEET, FOOD_LOG_HEADERS, FOOD_LOG_TABLE);
      const headers = await tableHeaders(token, base, session, table);
      const iso = nowInZone(TZ).iso; const by = req.user.email;
      const logRows = items.map((i) => rowForHeaders(headers, [
        [["datetime", "date time", "time"], iso], [["day"], day], [["name"], name],
        [["item", "dish", "food"], i.item], [["qty", "quantity"], i.qty],
        [["unitprice", "unit price", "price", "rate"], priceByName(i.item)],
        [["amount", "total"], Math.round(priceByName(i.item) * i.qty * 100) / 100],
        [["recordedby", "recorded by", "volunteer", "by"], by],
      ]));
      await addTableRows(token, base, session, table, logRows);
    } catch (e) { console.warn("food audit log failed:", e?.message || e); }

    res.json({ ok: true, added: items.length, amount: Math.round(addedAmount * 100) / 100, personTotal: Number(upd.total) || 0, paid: Number(upd.paid) || 0, outstanding: Number(upd.outstanding) || 0 });
  } catch (e) {
    console.error("food-entry failed:", e?.message || e);
    res.status(500).json({ error: "Could not save the purchase.", detail: e?.message || String(e) });
  } finally { if (token && base && session) await closeSession(token, base, session); }
});

// GET /api/food-dues?name= -> a guest's current dues snapshot (Total consumed, Paid, Outstanding).
app.get("/api/food-dues", requireAuth(), async (req, res) => {
  const name = clean(req.query.name);
  if (!name) return res.json({ found: false });
  let token, base, session;
  try {
    const { loc } = await wbContext();
    token = await getAccessToken(); base = workbookBase(loc); session = await openSession(token, base);
    const raw = await readUsedRangeRaw(token, base, session, FOOD_DUES_SHEET);
    const values = raw.values || [];
    let hRow = values.findIndex((r) => r.some((c) => lc(c) === "name") && r.some((c) => lc(c) === "total"));
    if (hRow === -1) hRow = values.findIndex((r) => r.some((c) => lc(c) === "name"));
    if (hRow === -1) return res.json({ found: false });
    const H = (values[hRow] || []).map((h) => String(h ?? ""));
    const nameCol = H.findIndex((h) => lc(h) === "name");
    const totalCol = H.findIndex((h) => lc(h) === "total");
    const paidCol = H.findIndex((h) => lc(h) === "paid");
    const outCol = H.findIndex((h) => ["outstanding", "due", "dues", "balance"].includes(lc(h)));
    const slCol = H.findIndex((h) => /^(sl\.?\s*no\.?|serial|#)$/i.test(String(h).trim()));
    const row = values.slice(hRow + 1).find((r) => lc(r[nameCol]) === lc(name));
    if (!row) return res.json({ found: false, name });
    const total = totalCol !== -1 ? parsePrice(row[totalCol]) : 0;
    const paid = paidCol !== -1 ? parsePrice(row[paidCol]) : 0;
    const outstanding = outCol !== -1 ? parsePrice(row[outCol]) : Math.round((total - paid) * 100) / 100;
    const items = [];
    H.forEach((h, i) => { if (i !== slCol && i !== nameCol && i !== totalCol && i !== paidCol && i !== outCol && lc(h) !== "") { const q = parseInt(row[i], 10); if (Number.isFinite(q) && q > 0) items.push({ item: h, qty: q }); } });
    res.json({ found: true, name, total, paid, outstanding, items });
  } catch (e) {
    console.error("food-dues failed:", e?.message || e);
    res.status(500).json({ error: "Could not load dues." });
  } finally { if (token && base && session) await closeSession(token, base, session); }
});

// POST /api/food-settle { name, amount, method? } -> record a payment; log to FoodSettlements.
app.post("/api/food-settle", requireAuth(), async (req, res) => {
  const name = clean(req.body?.name);
  const amount = Math.round((parseFloat(req.body?.amount) || 0) * 100) / 100;
  const method = clean(req.body?.method);
  if (!name) return res.status(400).json({ error: "Guest name is required." });
  if (!(amount > 0)) return res.status(400).json({ error: "Enter a payment amount greater than 0." });
  let token, base, session;
  try {
    const { loc } = await wbContext();
    token = await getAccessToken(); base = workbookBase(loc); session = await openSession(token, base);
    const raw = await readUsedRangeRaw(token, base, session, FOOD_DUES_SHEET);
    const values = raw.values || [];
    let hRow = values.findIndex((r) => r.some((c) => lc(c) === "name") && r.some((c) => lc(c) === "total"));
    if (hRow === -1) hRow = values.findIndex((r) => r.some((c) => lc(c) === "name"));
    if (hRow === -1) return res.status(404).json({ error: "No dues sheet found." });
    const duesHeaders = (values[hRow] || []).map((h) => String(h ?? ""));
    const upd = applyFoodPayment(duesHeaders, values.slice(hRow + 1), name, amount);
    if (!upd) return res.status(404).json({ error: `No food dues found for ${name}.` });
    const c0 = raw.columnIndex, headerRow1 = raw.rowIndex + hRow + 1;
    const startL = colLetter(c0 + 1), endL = colLetter(c0 + upd.headers.length);
    if (upd.headerChanged) await writeRange(token, base, session, FOOD_DUES_SHEET, `${startL}${headerRow1}:${endL}${headerRow1}`, [upd.headers]);
    const rowNo = headerRow1 + 1 + upd.rowIndex;
    await writeRange(token, base, session, FOOD_DUES_SHEET, `${startL}${rowNo}:${endL}${rowNo}`, [upd.rowValues]);
    // Append an audit line to FoodSettlements.
    try {
      const table = await ensureLogTable(token, base, session, FOOD_SETTLE_SHEET, FOOD_SETTLE_HEADERS, FOOD_SETTLE_TABLE);
      const headers = await tableHeaders(token, base, session, table);
      const logRow = rowForHeaders(headers, [
        [["datetime", "date time", "time"], nowInZone(TZ).iso], [["name"], name],
        [["amount", "paid amount"], amount], [["method", "mode"], method],
        [["paidtotal", "paid total", "paid"], upd.paid], [["outstanding", "due", "dues", "balance"], upd.outstanding],
        [["settledby", "recorded by", "recordedby", "by"], req.user.email],
      ]);
      await addTableRows(token, base, session, table, [logRow]);
    } catch (e) { console.warn("settle log failed:", e?.message || e); }
    res.json({ ok: true, name, amount, total: upd.total, paid: upd.paid, outstanding: upd.outstanding });
  } catch (e) {
    console.error("food-settle failed:", e?.message || e);
    res.status(500).json({ error: "Could not record the payment.", detail: e?.message || String(e) });
  } finally { if (token && base && session) await closeSession(token, base, session); }
});

// POST /api/parking-entry { name, mobile, reg, make, model, colour, date? } -> appends to Parking
app.post("/api/parking-entry", requireAuth(), async (req, res) => {
  const f = {
    name: clean(req.body?.name), mobile: clean(req.body?.mobile), reg: clean(req.body?.reg).toUpperCase(),
    make: clean(req.body?.make), model: clean(req.body?.model), colour: clean(req.body?.colour),
    date: clean(req.body?.date) || nowInZone(TZ).ymd,
  };
  if (!f.name) return res.status(400).json({ error: "Guest name is required." });
  if (!f.reg && !f.mobile) return res.status(400).json({ error: "Add a car registration or a mobile number." });
  if (!/^[\p{L}][\p{L} .'-]*$/u.test(f.name)) return res.status(400).json({ error: "Guest name should be letters only." });
  if (f.mobile && !/^\+?[0-9][0-9 ]{5,15}$/.test(f.mobile)) return res.status(400).json({ error: "Mobile number is not valid." });
  if (f.reg && (!/^[A-Za-z0-9 -]{2,12}$/.test(f.reg) || f.reg.replace(/[^A-Za-z0-9]/g, "").length < 2)) return res.status(400).json({ error: "Car registration is not valid." });
  if (f.make && !/^[\p{L}][\p{L} .-]*$/u.test(f.make)) return res.status(400).json({ error: "Make should be letters only." });
  if (f.model && !(/^[\p{L}\p{N} .-]+$/u.test(f.model) && /\p{L}/u.test(f.model))) return res.status(400).json({ error: "Model is not valid." });
  if (f.colour && !/^[\p{L}][\p{L} .-]*$/u.test(f.colour)) return res.status(400).json({ error: "Colour should be letters only." });
  let token, base, session;
  try {
    const { loc } = await wbContext();
    token = await getAccessToken(); base = workbookBase(loc); session = await openSession(token, base);
    const table = await ensureLogTable(token, base, session, PARKING_SHEET, PARKING_HEADERS, PARKING_TABLE);
    const headers = await tableHeaders(token, base, session, table);
    // Next Sl No = one past the highest already on the sheet (works whether rows live
    // inside or outside the table).
    let slNo = 1;
    try {
      const rr = await readUsedRangeRaw(token, base, session, PARKING_SHEET);
      const hr = rr.values.findIndex((r) => r.some((c) => lc(c) === "name"));
      const slIdx = (rr.values[hr === -1 ? 0 : hr] || []).findIndex((c) => /^sl/i.test(String(c ?? "").trim()));
      let maxSl = 0;
      for (const r of rr.values.slice((hr === -1 ? 0 : hr) + 1)) { const n = parseInt(r[slIdx], 10); if (Number.isFinite(n)) maxSl = Math.max(maxSl, n); }
      slNo = maxSl + 1;
    } catch { /* optional */ }
    const stamp = parkingStamp(f.date, nowInZone(TZ));
    const row = rowForHeaders(headers, [
      [["timestamp", "date", "time"], stamp], [["sl", "serial", "s.no", "sno", "#"], slNo], [["name"], f.name],
      [["mobile", "phone", "contact"], f.mobile], [["registration", "reg", "plate", "number plate", "car number"], f.reg],
      [["make"], f.make], [["model"], f.model], [["colour", "color"], f.colour],
    ]);
    await addTableRows(token, base, session, table, [row]);
    res.json({ ok: true, slNo });
  } catch (e) {
    console.error("parking-entry failed:", e?.message || e);
    res.status(500).json({ error: "Could not save the car.", detail: e?.message || String(e) });
  } finally { if (token && base && session) await closeSession(token, base, session); }
});

const PORT = process.env.PORT || 8080;
if (process.env.PORT !== "0" && process.env.npm_lifecycle_event !== "test") {
  app.listen(PORT, () => console.log(`OPB check-in backend on :${PORT} (tz=${TZ}, table=${TABLE_NAME})`));
}

export { app };
