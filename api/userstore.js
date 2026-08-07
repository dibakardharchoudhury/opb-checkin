// Persistent volunteer/role allowlist.
//
// In Azure it lives in Azure Table Storage, reached with the App Service's
// system-assigned managed identity (no keys or connection strings on the box, data
// encrypted at rest, access audited). When USERS_TABLE_ACCOUNT is unset (local dev
// and unit tests) it falls back to a JSON file so nothing external is required.
//
// The bootstrap admin(s) still come from ADMIN_EMAILS and are never written here,
// so the app can't be locked out by a store mistake.

import { promises as fs } from "node:fs";
import path from "node:path";
import os from "node:os";

const ACCOUNT = process.env.USERS_TABLE_ACCOUNT || "";
const TABLE = process.env.USERS_TABLE_NAME || "opbusers";
const PK = "user";
const useTable = !!ACCOUNT;

// ---- Table Storage backend (managed identity) ----
let _client = null;
async function tableClient() {
  if (_client) return _client;
  const { TableClient } = await import("@azure/data-tables");
  const { DefaultAzureCredential } = await import("@azure/identity");
  _client = new TableClient(`https://${ACCOUNT}.table.core.windows.net`, TABLE, new DefaultAzureCredential());
  try { await _client.createTable(); } catch { /* already exists */ }
  return _client;
}

// ---- file backend (local dev / tests) ----
// On Azure App Service the persistent, shared storage is /home (regardless of $HOME,
// which at runtime is not /home). Writing under it is what survives restarts/deploys.
const DATA_DIR = process.env.DATA_DIR
  || (process.env.WEBSITE_INSTANCE_ID ? "/home/data"
    : process.env.HOME ? path.join(process.env.HOME, "data")
    : path.join(os.tmpdir(), "opb-data"));
const FILE = path.join(DATA_DIR, "opb-users.json");
let cache = null;
async function fileLoad() {
  if (cache) return cache;
  try { cache = JSON.parse(await fs.readFile(FILE, "utf8")); }
  catch { cache = { users: [] }; }
  if (!Array.isArray(cache.users)) cache.users = [];
  return cache;
}
async function filePersist() {
  await fs.mkdir(DATA_DIR, { recursive: true });
  await fs.writeFile(FILE, JSON.stringify(cache, null, 2));
}

export async function listStoredUsers() {
  if (!useTable) { const d = await fileLoad(); return d.users.slice(); }
  const c = await tableClient();
  const out = [];
  for await (const e of c.listEntities({ queryOptions: { filter: `PartitionKey eq '${PK}'` } })) {
    out.push({ email: e.rowKey, role: e.role, addedBy: e.addedBy || null, addedAt: e.addedAt || null });
  }
  return out;
}

export async function upsertUser(email, role, addedBy) {
  const e = String(email).toLowerCase();
  if (!useTable) {
    const d = await fileLoad();
    const existing = d.users.find((u) => u.email === e);
    if (existing) existing.role = role;
    else d.users.push({ email: e, role, addedBy: addedBy || null, addedAt: new Date().toISOString() });
    await filePersist();
    return d.users;
  }
  const c = await tableClient();
  await c.upsertEntity({ partitionKey: PK, rowKey: e, role, addedBy: addedBy || null, addedAt: new Date().toISOString() }, "Merge");
  return listStoredUsers();
}

export async function removeUser(email) {
  const e = String(email).toLowerCase();
  if (!useTable) {
    const d = await fileLoad();
    const before = d.users.length;
    d.users = d.users.filter((u) => u.email !== e);
    await filePersist();
    return d.users.length < before;
  }
  const c = await tableClient();
  try { await c.deleteEntity(PK, e); return true; } catch { return false; }
}

// Test hook only — drop caches so a fresh DATA_DIR is re-read.
export function _resetCache() { cache = null; _client = null; }
