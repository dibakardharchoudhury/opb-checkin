// Shared app configuration set by admins and read by everyone.
//
// Persisted as JSON on the App Service persistent share (same place as the user
// store), so an admin's choices apply to every volunteer device — not just their
// own browser. Two settings today:
//   scanSheet    the active event-session worksheet volunteers check guests into
//   guestSheets  the worksheets volunteers may view in the Guest List

import { promises as fs } from "node:fs";
import path from "node:path";
import os from "node:os";

// On Azure App Service the persistent, shared storage is /home (regardless of $HOME,
// which at runtime is not /home). Writing under it is what survives restarts/deploys.
const DATA_DIR = process.env.DATA_DIR
  || (process.env.WEBSITE_INSTANCE_ID ? "/home/data"
    : process.env.HOME ? path.join(process.env.HOME, "data")
    : path.join(os.tmpdir(), "opb-data"));
const FILE = path.join(DATA_DIR, "opb-config.json");

const DEFAULT = { scanSheet: "", guestSheets: [] };
let cache = null;

function normalize(c) {
  return {
    scanSheet: typeof c?.scanSheet === "string" ? c.scanSheet : "",
    guestSheets: Array.isArray(c?.guestSheets) ? c.guestSheets.filter((s) => typeof s === "string") : [],
  };
}

async function load() {
  if (cache) return cache;
  try { cache = normalize(JSON.parse(await fs.readFile(FILE, "utf8"))); }
  catch { cache = { ...DEFAULT }; }
  return cache;
}

export async function getConfig() {
  return { ...(await load()) };
}

export async function setConfig(patch) {
  const c = await load();
  if (patch && typeof patch === "object") {
    if ("scanSheet" in patch) c.scanSheet = typeof patch.scanSheet === "string" ? patch.scanSheet : "";
    if ("guestSheets" in patch) c.guestSheets = Array.isArray(patch.guestSheets) ? patch.guestSheets.filter((s) => typeof s === "string") : [];
  }
  await fs.mkdir(DATA_DIR, { recursive: true });
  await fs.writeFile(FILE, JSON.stringify(c, null, 2));
  return { ...c };
}

// Test hook only — drop the in-memory cache so a fresh DATA_DIR is re-read.
export function _resetCache() { cache = null; }
