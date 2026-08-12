import { test } from "node:test";
import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import { promises as fs } from "node:fs";

const dir = await fs.mkdtemp(path.join(os.tmpdir(), "opb-cfg-"));
process.env.DATA_DIR = dir;
const { getConfig, setConfig, _resetCache } = await import("./configstore.js");

test("defaults to empty config", async () => {
  _resetCache();
  const c = await getConfig();
  assert.equal(c.scanSheet, "");
  assert.deepEqual(c.guestSheets, []);
});

test("sets and merges scanSheet and guestSheets", async () => {
  await setConfig({ scanSheet: "Saturday Dinner" });
  let c = await getConfig();
  assert.equal(c.scanSheet, "Saturday Dinner");
  assert.deepEqual(c.guestSheets, []);
  await setConfig({ guestSheets: ["Friday Lunch", "Saturday Dinner"] });
  c = await getConfig();
  assert.equal(c.scanSheet, "Saturday Dinner"); // unchanged by partial patch
  assert.deepEqual(c.guestSheets, ["Friday Lunch", "Saturday Dinner"]);
});

test("accepts event-date and cutoff settings", async () => {
  await setConfig({ cutoff: "1700", eventDate: "2025-09-26" });
  const c = await getConfig();
  assert.equal(c.cutoff, "1700");
  assert.equal(c.eventDate, "2025-09-26");
});

test("ignores non-string entries and bad types", async () => {
  await setConfig({ guestSheets: ["A", 5, null, "B"], scanSheet: 123, cutoff: 1700, eventDate: 1234 });
  const c = await getConfig();
  assert.equal(c.scanSheet, "");
  assert.equal(c.cutoff, "1700");
  assert.equal(c.eventDate, "");
  assert.deepEqual(c.guestSheets, ["A", "B"]);
});
