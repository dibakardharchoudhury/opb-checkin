import { test } from "node:test";
import assert from "node:assert/strict";
import { evaluateScan, excelSerial, sessionFor, passDateYMD, nowInZone } from "./rules.js";

// Variant A — discrete columns (Oslo Durgotsav 2025 App_Source shape).
const headersA = ["RegistrationID", "UniqueKey", "First Name", "Last Name", "Item", "Date", "PassType", "FoodOption", "Quantity", "Status", "DateTime"];
const rowA = (reg, dateSerial, pass, food, status = "") =>
  ({ index: 0, values: [reg, `${reg}${dateSerial}${pass}${food}`, "Jeet", "Sarkar", "01. Season pass", dateSerial, pass, food, "3", status, ""] });

// A fixed "now": 2025-09-26 12:00 Oslo -> Lunch session, serial 45926.
const noonOslo = new Date("2025-09-26T10:00:00Z"); // 12:00 CEST
const serial = excelSerial(nowInZone("Europe/Oslo", noonOslo).dateOnly);

test("excel serial + session helpers", () => {
  assert.equal(serial, 45926);
  assert.equal(sessionFor(1200), "Lunch");
  assert.equal(sessionFor(1401), "Dinner");
  assert.equal(sessionFor(1400), "Lunch"); // flow used > 1400
  assert.equal(passDateYMD(45926), "2025-09-26");
  assert.equal(passDateYMD("2025-09-26 00:00:00"), "2025-09-26");
});

test("SUCCESS: valid lunch pass today gets registered", () => {
  const rows = [{ ...rowA(3000, serial, "Lunch", "Yes"), index: 5 }];
  const r = evaluateScan({ headers: headersA, rows, orderNumber: "3000", now: noonOslo });
  assert.equal(r.decision, "SUCCESS");
  assert.match(r.response, /SUCCESS/);
  assert.equal(r.customerName, "Jeet Sarkar");
  assert.equal(r.patch.length, 1);
  assert.equal(r.patch[0].newStatus, "REGISTERED");
});

test("ALREADY: pass already REGISTERED", () => {
  const rows = [{ ...rowA(3000, serial, "Lunch", "Yes", "REGISTERED"), index: 5 }];
  const r = evaluateScan({ headers: headersA, rows, orderNumber: "3000", now: noonOslo });
  assert.equal(r.decision, "ALREADY");
  assert.match(r.response, /already registered/i);
});

test("INVALID: dinner pass scanned at lunchtime", () => {
  const rows = [rowA(3000, serial, "Dinner", "Yes")];
  const r = evaluateScan({ headers: headersA, rows, orderNumber: "3000", now: noonOslo });
  assert.equal(r.decision, "INVALID");
});

test("INVALID: unknown order", () => {
  const rows = [rowA(3000, serial, "Lunch", "Yes")];
  const r = evaluateScan({ headers: headersA, rows, orderNumber: "9999", now: noonOslo });
  assert.equal(r.decision, "INVALID");
});

test("SUCCESS: multiple rows (veg + non-veg) both updated", () => {
  const rows = [
    { ...rowA(3000, serial, "Lunch", "Yes"), index: 1 },
    { ...rowA(3000, serial, "Lunch", "No"), index: 2 },
    { ...rowA(3000, serial, "Dinner", "Yes"), index: 3 },
  ];
  const r = evaluateScan({ headers: headersA, rows, orderNumber: "3000", now: noonOslo });
  assert.equal(r.decision, "SUCCESS");
  assert.equal(r.patch.length, 2); // both Lunch rows, not the Dinner one
});

// Variant B — collapsed AppKey (Oslo Durgotsav 2026 shape). No Date/PassType columns.
const headersB = ["Order Number", "UniqueKey", "First Name", "Last Name", "Quantity", "AppKey", "Status", "DateTime"];
const rowB = (order, dateSerial, pass, food, status = "") =>
  ({ index: 0, values: [order, `${order}${dateSerial}${pass}${food}`, "Riya", "Roy", "1", `${order}${dateSerial}${pass}${food}`, status, ""] });

test("Variant B: AppKey prefix match registers", () => {
  const rows = [
    { ...rowB(3000, serial, "Lunch", "Yes"), index: 7 },
    { ...rowB(3000, serial, "Dinner", "Yes"), index: 8 },
  ];
  const r = evaluateScan({ headers: headersB, rows, orderNumber: "3000", now: noonOslo });
  assert.equal(r.decision, "SUCCESS");
  assert.equal(r.patch.length, 1);
  assert.equal(r.patch[0].index, 7);
});
