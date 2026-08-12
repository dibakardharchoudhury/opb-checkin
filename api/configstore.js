// Shared app configuration set by admins and read by everyone.
//
// Persisted as JSON on the App Service persistent share (same place as the user
// store), so an admin's choices apply to every volunteer device — not just their
// own browser. The shared settings cover the active event, the current cut-off,
// and the volunteer-visible tabs.

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

const DEFAULT = { workbook: null, scanSheet: "", guestSheets: [], cutoff: "", eventDate: "" };
let cache = null;

function normalizeCutoff(v) {
  if (v == null || v === "") return "";
  const s = String(v).trim().replace(/:/g, "");
  if (!/^\d{3,4}$/.test(s)) return "";
  const n = Number.parseInt(s, 10);
  return Number.isFinite(n) ? String(n).padStart(4, "0") : "";
}

function normalizeEventDate(v) {
  if (v == null || v === "") return "";
  const s = String(v).trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return "";
  return s;
}

function normalize(c) {
  const wb = c?.workbook && typeof c.workbook.id === "string"
    ? { id: c.workbook.id, name: typeof c.workbook.name === "string" ? c.workbook.name : "" }
    : null;
  return {
    workbook: wb,
    scanSheet: typeof c?.scanSheet === "string" ? c.scanSheet : "",
    guestSheets: Array.isArray(c?.guestSheets) ? c.guestSheets.filter((s) => typeof s === "string") : [],
    cutoff: normalizeCutoff(c?.cutoff),
    eventDate: normalizeEventDate(c?.eventDate),
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
    if ("workbook" in patch) c.workbook = patch.workbook && typeof patch.workbook.id === "string"
      ? { id: patch.workbook.id, name: typeof patch.workbook.name === "string" ? patch.workbook.name : "" } : null;
    if ("scanSheet" in patch) c.scanSheet = typeof patch.scanSheet === "string" ? patch.scanSheet : "";
    if ("guestSheets" in patch) c.guestSheets = Array.isArray(patch.guestSheets) ? patch.guestSheets.filter((s) => typeof s === "string") : [];
    if ("cutoff" in patch) c.cutoff = normalizeCutoff(patch.cutoff);
    if ("eventDate" in patch) c.eventDate = normalizeEventDate(patch.eventDate);
  }
  await fs.mkdir(DATA_DIR, { recursive: true });
  await fs.writeFile(FILE, JSON.stringify(c, null, 2));
  return { ...c };
}

// Test hook only — drop the in-memory cache so a fresh DATA_DIR is re-read.
export function _resetCache() { cache = null; }
