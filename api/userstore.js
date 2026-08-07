// Tiny persistent user store for in-app volunteer/role management.
//
// Stored as a JSON file under the App Service persistent share (/home/data on
// Linux App Service survives restarts & scale). No database needed for a small
// volunteer list. The bootstrap admin(s) still come from ADMIN_EMAILS and are
// never written here, so the app can't be locked out by a UI mistake.

import { promises as fs } from "node:fs";
import path from "node:path";
import os from "node:os";

const DATA_DIR = process.env.DATA_DIR
  || (process.env.HOME ? path.join(process.env.HOME, "data") : path.join(os.tmpdir(), "opb-data"));
const FILE = path.join(DATA_DIR, "opb-users.json");

let cache = null;

async function load() {
  if (cache) return cache;
  try { cache = JSON.parse(await fs.readFile(FILE, "utf8")); }
  catch { cache = { users: [] }; }
  if (!Array.isArray(cache.users)) cache.users = [];
  return cache;
}
async function persist() {
  await fs.mkdir(DATA_DIR, { recursive: true });
  await fs.writeFile(FILE, JSON.stringify(cache, null, 2));
}

export async function listStoredUsers() {
  const d = await load();
  return d.users.slice();
}

export async function upsertUser(email, role, addedBy) {
  const d = await load();
  const e = String(email).toLowerCase();
  const existing = d.users.find((u) => u.email === e);
  if (existing) existing.role = role;
  else d.users.push({ email: e, role, addedBy: addedBy || null, addedAt: new Date().toISOString() });
  await persist();
  return d.users;
}

export async function removeUser(email) {
  const d = await load();
  const e = String(email).toLowerCase();
  const before = d.users.length;
  d.users = d.users.filter((u) => u.email !== e);
  await persist();
  return d.users.length < before;
}

// Test hook only — drop the in-memory cache so a fresh DATA_DIR is re-read.
export function _resetCache() { cache = null; }
