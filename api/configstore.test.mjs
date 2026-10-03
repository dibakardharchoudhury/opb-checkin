import { after, test } from "node:test";
import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import { promises as fs } from "node:fs";

const dir = await fs.mkdtemp(path.join(os.tmpdir(), "opb-cfg-"));
process.env.DATA_DIR = dir;
const { getConfig, setConfig, setFoodPlanItem, setTransportItem, _resetCache } = await import("./configstore.js");

after(() => fs.rm(dir, { recursive: true, force: true }));

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

test("stores only validated public transport card fields", async () => {
  const saved = await setTransportItem("sat-1730", {
    time: " 17:30 ",
    topic: "  Star pickup  ",
    note: " Airbnb to venue ",
    names: [" Dibakar ", 42, "Mayukh", ""],
    backupNames: [" Arunavo ", null, "Suprakash da"],
  });

  assert.deepEqual(saved, {
    id: "sat-1730", topic: "Star pickup", names: ["Dibakar", "Mayukh"],
    time: "17:30", note: "Airbnb to venue", backupNames: ["Arunavo", "Suprakash da"],
  });
  _resetCache();
  assert.deepEqual((await getConfig()).transport["sat-1730"], {
    topic: "Star pickup", names: ["Dibakar", "Mayukh"], time: "17:30",
    note: "Airbnb to venue", backupNames: ["Arunavo", "Suprakash da"],
  });
  await assert.rejects(() => setTransportItem("../bad", { topic: "No", names: [] }));
  await assert.rejects(() => setTransportItem("sat-1730", { topic: "", names: [] }));
});

test("stores only validated public food plan fields", async () => {
  const saved = await setFoodPlanItem("mon-dinner", {
    meal: " Dinner ",
    time: " 17:00 ",
    provider: " Internal / outsource ",
    menu: [" Lamb + chana kofta ", 42, "Option 2: Outsource", ""],
    names: [" Prasenjit ", "Sudip", null, "Anirban Singha"],
    note: " Decision pending ",
  });
  assert.deepEqual(saved, {
    id: "mon-dinner", meal: "Dinner", time: "17:00", provider: "Internal / outsource",
    menu: ["Lamb + chana kofta", "Option 2: Outsource"],
    names: ["Prasenjit", "Sudip", "Anirban Singha"], note: "Decision pending",
  });
  _resetCache();
  assert.deepEqual((await getConfig()).foodPlan["mon-dinner"], {
    meal: "Dinner", time: "17:00", provider: "Internal / outsource",
    menu: ["Lamb + chana kofta", "Option 2: Outsource"],
    names: ["Prasenjit", "Sudip", "Anirban Singha"], note: "Decision pending",
  });
  await assert.rejects(() => setFoodPlanItem("../bad", { meal: "Dinner", provider: "Internal", menu: [] }));
  await assert.rejects(() => setFoodPlanItem("mon-dinner", { meal: "", provider: "", menu: [] }));
});
