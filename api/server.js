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
} from "./graph.js";
import { evaluateScan } from "./rules.js";
import { verifyProviderToken, resolveRoleMerged, issueSession, requireAuth } from "./auth.js";
import { listStoredUsers, upsertUser, removeUser } from "./userstore.js";
import { getConfig, setConfig } from "./configstore.js";

const TABLE_NAME = process.env.TABLE_NAME || "Table1";
const TZ = process.env.TZ_NAME || "Europe/Oslo";

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
  res.setHeader("X-Content-Type-Options", "nosniff");
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
    const result = evaluateScan({ headers, rows, orderNumber, tz: TZ, sheetScoped: !!sheet });

    if (result.decision === "SUCCESS") {
      for (const item of result.patch) await patchRow(token, base, session, table, item);
    }
    res.json({ response: result.response, customername: result.customerName, decision: result.decision, sheet: sheet || table });
  } catch (e) {
    console.error("register failed:", e?.message || e);
    res.status(500).json({ error: "Registration failed — please retry.", detail: e?.message || String(e) });
  } finally {
    if (token && base && session) await closeSession(token, base, session);
  }
});

// GET /api/sheet?tabs=A,B -> { sheets: [{ name, table, headers, rows, count }], updatedAt }
// One entry per requested worksheet. With no ?tabs, reads the configured TABLE_NAME
// (legacy single-sheet mode). Short-cached per tab-set so opening the list is cheap.
// NOTE: returns guest names/PII — gated behind the volunteer/admin login.
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
    const sheets = [];
    for (const name of names) {
      const table = name ? await firstTableName(token, base, session, name) : TABLE_NAME;
      const read = table
        ? await readTable(token, base, session, table)
        : await readWorksheetUsedRange(token, base, session, name);
      sheets.push({
        name: name || table, table: table || null,
        headers: read.headers, rows: read.rows.map((r) => r.values), count: read.rows.length,
      });
    }
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

const PORT = process.env.PORT || 8080;
if (process.env.PORT !== "0") {
  app.listen(PORT, () => console.log(`OPB check-in backend on :${PORT} (tz=${TZ}, table=${TABLE_NAME})`));
}

export { app };
