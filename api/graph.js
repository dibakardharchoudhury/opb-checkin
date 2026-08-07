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
  const body = new URLSearchParams({
    client_id: process.env.OPB_CLIENT_ID,
    client_secret: process.env.OPB_CLIENT_SECRET || "",
    grant_type: "refresh_token",
    refresh_token: process.env.OPB_REFRESH_TOKEN,
    scope: SCOPE,
  });
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

function workbookBase() {
  const loc = process.env.GRAPH_WORKBOOK || "";
  if (loc.startsWith("id:")) return `${GRAPH}/me/drive/items/${encodeURIComponent(loc.slice(3))}/workbook`;
  if (loc.startsWith("path:")) {
    const p = loc.slice(5).replace(/^\/+/, "");
    return `${GRAPH}/me/drive/root:/${p.split("/").map(encodeURIComponent).join("/")}:/workbook`;
  }
  throw new Error('GRAPH_WORKBOOK must be "id:<driveItemId>" or "path:/Folder/File.xlsx"');
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
