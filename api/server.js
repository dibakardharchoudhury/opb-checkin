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
} from "./graph.js";
import { evaluateScan } from "./rules.js";
import { verifyProviderToken, resolveRole, issueSession, requireAuth } from "./auth.js";

const TABLE_NAME = process.env.TABLE_NAME || "Table1";
const TZ = process.env.TZ_NAME || "Europe/Oslo";

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
    const role = resolveRole(email);
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

// POST /api/register { orderNumber }  ->  { response, customername, decision }
app.post("/api/register", requireAuth(), scanLimiter, async (req, res) => {
  const raw = (req.body && req.body.orderNumber) != null ? String(req.body.orderNumber) : "";
  // Accept the raw QR too; take the part before the first ';' like the Power App did.
  const orderNumber = raw.split(";")[0].trim();
  if (!orderNumber) return res.status(400).json({ error: "orderNumber required" });

  let token, base, session;
  try {
    token = await getAccessToken();
    base = workbookBase();
    session = await openSession(token, base);

    const { headers, rows } = await readTable(token, base, session, TABLE_NAME);
    const result = evaluateScan({ headers, rows, orderNumber, tz: TZ });

    if (result.decision === "SUCCESS") {
      for (const item of result.patch) await patchRow(token, base, session, TABLE_NAME, item);
    }
    res.json({ response: result.response, customername: result.customerName, decision: result.decision });
  } catch (e) {
    console.error("register failed:", e?.message || e);
    res.status(500).json({ error: "Registration failed — please retry.", detail: e?.message || String(e) });
  } finally {
    if (token && base && session) await closeSession(token, base, session);
  }
});

// GET /api/sheet -> { headers, rows, table, count, updatedAt }  (the live workbook)
// Short-cached so opening the Guest List doesn't hammer Graph. NOTE: this returns
// guest names/PII — gate it behind the volunteer/admin login when that phase lands.
let sheetCache = { at: 0, data: null };
app.get("/api/sheet", requireAuth(), async (req, res) => {
  const fresh = req.query.refresh === "1";
  if (!fresh && sheetCache.data && Date.now() - sheetCache.at < 15_000) return res.json(sheetCache.data);

  let token, base, session;
  try {
    token = await getAccessToken();
    base = workbookBase();
    session = await openSession(token, base);
    const { headers, rows } = await readTable(token, base, session, TABLE_NAME);
    const data = { headers, rows: rows.map((r) => r.values), table: TABLE_NAME, count: rows.length, updatedAt: new Date().toISOString() };
    sheetCache = { at: Date.now(), data };
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
