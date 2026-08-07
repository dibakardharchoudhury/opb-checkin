// Microsoft Graph client for the OPB personal-OneDrive workbook.
//
// Consumer OneDrive has NO app-only / service-principal access, so we act DELEGATED
// as the workbook owner (osloprobaseebangali@outlook.com). The one-time consent
// (see authorize.js) yields a refresh token; here we exchange it for short-lived
// access tokens and call the Excel workbook API. No user password is ever stored.
//
// Required env:
//   OPB_CLIENT_ID       app (client) id of the registration (personal accounts enabled)
//   OPB_CLIENT_SECRET   client secret (confidential Web app) — kept server-side only
//   OPB_REFRESH_TOKEN   from the one-time authorize.js consent
//   GRAPH_WORKBOOK      workbook locator: either "id:<driveItemId>" or "path:/Folder/File.xlsx"
//   TABLE_NAME          Excel table to read/update (default "Table1")

const TOKEN_URL = "https://login.microsoftonline.com/consumers/oauth2/v2.0/token";
const GRAPH = "https://graph.microsoft.com/v1.0";
const SCOPE = "Files.ReadWrite offline_access openid profile";

let cached = { token: null, exp: 0 };

export async function getAccessToken() {
  if (cached.token && cached.exp - Date.now() > 60_000) return cached.token;
  const params = {
    client_id: process.env.OPB_CLIENT_ID,
    grant_type: "refresh_token",
    refresh_token: process.env.OPB_REFRESH_TOKEN,
    scope: SCOPE,
  };
  // Confidential clients send a secret; a public client (no secret) must omit it.
  if (process.env.OPB_CLIENT_SECRET) params.client_secret = process.env.OPB_CLIENT_SECRET;
  const body = new URLSearchParams(params);
  const r = await fetch(TOKEN_URL, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });
  const j = await r.json();
  if (!r.ok) throw new Error(`token refresh failed: ${j.error || r.status} ${j.error_description || ""}`);
  cached = { token: j.access_token, exp: Date.now() + (j.expires_in || 3600) * 1000 };
  // Providers may rotate the refresh token; surface it so the operator can update the secret.
  if (j.refresh_token && j.refresh_token !== process.env.OPB_REFRESH_TOKEN) {
    console.warn("[graph] refresh token rotated — update OPB_REFRESH_TOKEN app setting");
    process.env.OPB_REFRESH_TOKEN = j.refresh_token;
  }
  return cached.token;
}

function workbookBase(loc) {
  loc = loc || process.env.GRAPH_WORKBOOK || "";
  if (loc.startsWith("id:")) return `${GRAPH}/me/drive/items/${encodeURIComponent(loc.slice(3))}/workbook`;
  if (loc.startsWith("path:")) {
    const p = loc.slice(5).replace(/^\/+/, "");
    return `${GRAPH}/me/drive/root:/${p.split("/").map(encodeURIComponent).join("/")}:/workbook`;
  }
  throw new Error('Workbook locator must be "id:<driveItemId>" or "path:/Folder/File.xlsx"');
}

// List Excel workbooks (.xlsx) in the owner's OneDrive so an admin can pick the file.
// Enumerates the drive deterministically (root children, paged) and also merges a
// search pass to catch files kept in subfolders — search alone omits results.
export async function listWorkbooks(token) {
  const map = new Map();
  const add = (it) => {
    if (it && it.id && it.file && /\.xlsx$/i.test(it.name || ""))
      map.set(it.id, { id: it.id, name: String(it.name).replace(/\.xlsx$/i, "") });
  };
  let url = `${GRAPH}/me/drive/root/children?$select=id,name,file&$top=200`;
  try { while (url) { const j = await gfetch(url, { token }); (j.value || []).forEach(add); url = j["@odata.nextLink"] || null; } } catch { /* ignore */ }
  try {
    let s = `${GRAPH}/me/drive/root/search(q='xlsx')?$select=id,name,file&$top=200`;
    while (s) { const j = await gfetch(s, { token }); (j.value || []).forEach(add); s = j["@odata.nextLink"] || null; }
  } catch { /* ignore */ }
  return [...map.values()].sort((a, b) => a.name.localeCompare(b.name));
}

async function gfetch(url, { token, session, method = "GET", body } = {}) {
  const headers = { Authorization: `Bearer ${token}` };
  if (session) headers["workbook-session-id"] = session;
  if (body !== undefined) headers["Content-Type"] = "application/json";
  const r = await fetch(url, { method, headers, body: body === undefined ? undefined : JSON.stringify(body) });
  const text = await r.text();
  const data = text ? JSON.parse(text) : {};
  if (!r.ok) {
    const err = data?.error?.message || r.statusText;
    throw new Error(`Graph ${method} ${r.status}: ${err}`);
  }
  return data;
}

// A persisted session keeps reads and the subsequent patch consistent.
export async function openSession(token, base) {
  const j = await gfetch(`${base}/createSession`, { token, method: "POST", body: { persistChanges: true } });
  return j.id;
}
export async function closeSession(token, base, session) {
  try { await gfetch(`${base}/closeSession`, { token, session, method: "POST", body: {} }); } catch { /* best effort */ }
}

// List the workbook's worksheets (tabs). Hidden sheets are skipped; order follows
// the tab order shown in Excel so the UI can present them naturally.
export async function listWorksheets(token, base, session) {
  const j = await gfetch(`${base}/worksheets?$select=name,position,visibility`, { token, session });
  return (j.value || [])
    .filter((w) => (w.visibility || "Visible") === "Visible")
    .sort((a, b) => (a.position ?? 0) - (b.position ?? 0))
    .map((w) => w.name);
}

// The first Excel table on a worksheet (event tabs carry exactly one). Null if none.
export async function firstTableName(token, base, session, worksheet) {
  const j = await gfetch(`${base}/worksheets/${encodeURIComponent(worksheet)}/tables?$select=name`, { token, session });
  const t = (j.value || [])[0];
  return t ? t.name : null;
}

// Fallback reader for worksheets without a table: use the used range and detect the
// real header row (many of these sheets have a title banner and/or a merged label
// column above the actual headers). Only used for viewing.
export async function readWorksheetUsedRange(token, base, session, worksheet) {
  const j = await gfetch(`${base}/worksheets/${encodeURIComponent(worksheet)}/usedRange(valuesOnly=true)?$select=values`, { token, session });
  return shapeUsedRange(j.values || []);
}

// Pure shaping of a raw value grid into { headers, rows }. Exported for testing.
// Picks the densest of the first rows as the header (skips title banners) and keeps
// only columns that have both a header and at least one data value (drops merged
// day-labels and empty columns).
export function shapeUsedRange(values) {
  if (!values || !values.length) return { headers: [], rows: [] };
  const filled = (v) => String(v ?? "").trim() !== "";
  const count = (row) => row.reduce((n, v) => n + (filled(v) ? 1 : 0), 0);
  const scan = Math.min(values.length, 15);
  let h = 0, best = -1;
  for (let i = 0; i < scan; i++) { const c = count(values[i]); if (c > best) { best = c; h = i; } }
  const headerRow = values[h] || [];
  const dataRows = values.slice(h + 1);
  const keep = [];
  for (let c = 0; c < headerRow.length; c++) {
    if (filled(headerRow[c]) && dataRows.some((r) => filled(r[c]))) keep.push(c);
  }
  const cols = keep.length ? keep : headerRow.map((_, c) => c);
  const headers = cols.map((c) => headerRow[c]);
  const rows = dataRows.map((r, i) => ({ index: i, values: cols.map((c) => r[c]) }));
  return { headers, rows };
}

export async function readTable(token, base, session, tableName) {
  const header = await gfetch(`${base}/tables/${encodeURIComponent(tableName)}/headerRowRange`, { token, session });
  const headers = (header.values && header.values[0]) || [];
  // Page through rows (Graph caps page size).
  let url = `${base}/tables/${encodeURIComponent(tableName)}/rows`;
  const rows = [];
  while (url) {
    const page = await gfetch(url, { token, session });
    for (const row of page.value || []) rows.push({ index: row.index, values: row.values[0] });
    url = page["@odata.nextLink"] || null;
  }
  return { headers, rows };
}

// Patch one table row by its 0-based table index, replacing only Status/DateTime.
export async function patchRow(token, base, session, tableName, item) {
  const values = item.values.slice();
  values[item.statusCol] = item.newStatus;
  values[item.dateTimeCol] = item.newDateTime;
  await gfetch(`${base}/tables/${encodeURIComponent(tableName)}/rows/itemAt(index=${item.index})`, {
    token, session, method: "PATCH", body: { values: [values] },
  });
}

export { workbookBase };
