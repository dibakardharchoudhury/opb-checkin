import { after, test } from "node:test";
import assert from "node:assert/strict";

process.env.PORT = "0";
process.env.SESSION_SECRET = "test-secret-please-rotate";
delete process.env.ALLOWED_ORIGINS;

const { app } = await import("./server.js");
const { issueSession } = await import("./auth.js");
const server = app.listen(0);
await new Promise((resolve) => server.once("listening", resolve));
const base = `http://127.0.0.1:${server.address().port}`;

after(() => new Promise((resolve, reject) => server.close((error) => error ? reject(error) : resolve())));

test("CORS fails closed when ALLOWED_ORIGINS is not configured", async () => {
  const response = await fetch(`${base}/health`, { headers: { Origin: "https://evil.example" } });
  assert.equal(response.status, 200);
  assert.equal(response.headers.get("access-control-allow-origin"), null);
});

test("API security headers are applied to unauthorized responses", async () => {
  const response = await fetch(`${base}/api/me`);
  assert.equal(response.status, 401);
  assert.match(response.headers.get("cache-control") || "", /no-store/);
  assert.equal(response.headers.get("x-content-type-options"), "nosniff");
  assert.equal(response.headers.get("x-frame-options"), "DENY");
  assert.equal(response.headers.get("referrer-policy"), "no-referrer");
});

test("downstream failures do not expose internal error details", async () => {
  const token = await issueSession({ email: "admin@opb.no", name: "Admin", role: "admin" });
  const realFetch = globalThis.fetch;
  globalThis.fetch = (url, options) => String(url).startsWith(base)
    ? realFetch(url, options)
    : Promise.reject(new Error("sensitive-internal-test-detail"));
  const response = await globalThis.fetch(`${base}/api/workbooks`, { headers: { Authorization: `Bearer ${token}` } });
  globalThis.fetch = realFetch;
  assert.equal(response.status, 500);
  const body = await response.json();
  assert.deepEqual(body, { error: "Could not list spreadsheets." });
  assert.equal(Object.hasOwn(body, "detail"), false);
});