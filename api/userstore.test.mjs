import { test } from "node:test";
import assert from "node:assert/strict";
import { promises as fs } from "node:fs";
import path from "node:path";
import os from "node:os";

// Isolate the store to a temp dir before importing the modules that read env.
const DIR = await fs.mkdtemp(path.join(os.tmpdir(), "opb-users-"));
process.env.DATA_DIR = DIR;

const { listStoredUsers, upsertUser, removeUser, _resetCache } = await import("./userstore.js");
const { resolveRoleMerged } = await import("./auth.js");

test("upsert adds, updates role, and lists", async () => {
  await upsertUser("Vol1@Gmail.com", "user", "admin@opb.no");
  let all = await listStoredUsers();
  assert.equal(all.length, 1);
  assert.equal(all[0].email, "vol1@gmail.com");
  assert.equal(all[0].role, "user");
  await upsertUser("vol1@gmail.com", "admin", "admin@opb.no"); // promote
  all = await listStoredUsers();
  assert.equal(all.length, 1);
  assert.equal(all[0].role, "admin");
});

test("resolveRoleMerged: config admin > store > USER_EMAILS > reject", async () => {
  process.env.ADMIN_EMAILS = "boss@opb.no";
  process.env.USER_EMAILS = "legacy@live.com";
  assert.equal(await resolveRoleMerged("boss@opb.no"), "admin");   // config bootstrap
  assert.equal(await resolveRoleMerged("vol1@gmail.com"), "admin"); // from store (promoted above)
  assert.equal(await resolveRoleMerged("legacy@live.com"), "user"); // env fallback
  assert.equal(await resolveRoleMerged("nobody@x.com"), null);      // rejected
});

test("removeUser deletes", async () => {
  await upsertUser("temp@x.com", "user", "admin@opb.no");
  assert.equal(await removeUser("temp@x.com"), true);
  const all = await listStoredUsers();
  assert.ok(!all.find((u) => u.email === "temp@x.com"));
});
