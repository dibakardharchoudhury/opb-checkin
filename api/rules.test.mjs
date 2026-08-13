import { test } from "node:test";
import assert from "node:assert/strict";
import { evaluateScan, excelSerial, sessionFor, passDateYMD, nowInZone } from "./rules.js";
import { rowForHeaders, parsePrice, parkingStamp, parseFoodMenu, applyFoodDues, normalizeParkingRow, deFormula } from "./server.js";

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
  // Default cut-off is now 1600 (was 1400): only strictly after 1600 is Dinner.
  assert.equal(sessionFor(1401), "Lunch");
  assert.equal(sessionFor(1600), "Lunch"); // exactly at the cut-off stays Lunch
  assert.equal(sessionFor(1601), "Dinner");
  assert.equal(sessionFor(2100), "Dinner");
  // Custom cut-off still honours the greater-than semantics.
  assert.equal(sessionFor(1500, 1400), "Dinner");
  assert.equal(sessionFor(1500), "Lunch");
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

// ---- Flow fidelity: a pass is valid ONLY for its own date and the current meal window ----
// (Ported from OPB_Excel_QRCodeScannerFlow Filter_array; cut-off now 1600 CET.)
const friSerial = serial - 1;                          // 2025-09-25 (Friday)
const satLunch = noonOslo;                             // 2025-09-26 12:00 CEST -> Lunch
const satEve = new Date("2025-09-26T15:00:00Z");       // 2025-09-26 17:00 CEST -> Dinner (>1600)
const sat3pm = new Date("2025-09-26T13:00:00Z");       // 2025-09-26 15:00 CEST -> HHmm 1500
// One order with a pass for every day/meal — exactly the case the UniqueKey was built for.
const fullOrderA = [
  { ...rowA(4000, friSerial, "Lunch", "Yes"), index: 1 },
  { ...rowA(4000, friSerial, "Dinner", "Yes"), index: 2 },
  { ...rowA(4000, serial, "Lunch", "Yes"), index: 3 },
  { ...rowA(4000, serial, "Dinner", "Yes"), index: 4 },
];

test("flow: only TODAY's Dinner row registers after the 1600 cut-off", () => {
  const r = evaluateScan({ headers: headersA, rows: fullOrderA, orderNumber: "4000", now: satEve });
  assert.equal(r.decision, "SUCCESS");
  assert.equal(r.patch.length, 1);
  assert.equal(r.patch[0].index, 4); // Saturday Dinner only — not Friday, not Lunch
});

test("flow: only TODAY's Lunch row registers before the cut-off", () => {
  const r = evaluateScan({ headers: headersA, rows: fullOrderA, orderNumber: "4000", now: satLunch });
  assert.equal(r.decision, "SUCCESS");
  assert.equal(r.patch.length, 1);
  assert.equal(r.patch[0].index, 3); // Saturday Lunch only
});

test("flow: 1500 is Lunch under the 1600 cut-off, Dinner under a 1400 cut-off", () => {
  const rL = evaluateScan({ headers: headersA, rows: fullOrderA, orderNumber: "4000", now: sat3pm });
  assert.equal(rL.patch[0].index, 3); // default 1600 -> Lunch
  const rD = evaluateScan({ headers: headersA, rows: fullOrderA, orderNumber: "4000", now: sat3pm, cutoff: 1400 });
  assert.equal(rD.patch[0].index, 4); // cut-off 1400 -> Dinner
});

test("flow: a pass only for Friday is INVALID when scanned on Saturday", () => {
  const friOnly = [
    { ...rowA(5000, friSerial, "Lunch", "Yes"), index: 1 },
    { ...rowA(5000, friSerial, "Dinner", "Yes"), index: 2 },
  ];
  const r = evaluateScan({ headers: headersA, rows: friOnly, orderNumber: "5000", now: satLunch });
  assert.equal(r.decision, "INVALID");
  assert.match(r.response, /NOT valid/i);
});

test("flow: veg + non-veg for the same date/meal both register, other meals untouched", () => {
  const rows = [
    { ...rowA(6000, serial, "Dinner", "Yes"), index: 10 },
    { ...rowA(6000, serial, "Dinner", "No"), index: 11 },
    { ...rowA(6000, serial, "Lunch", "Yes"), index: 12 },
  ];
  const r = evaluateScan({ headers: headersA, rows, orderNumber: "6000", now: satEve });
  assert.equal(r.decision, "SUCCESS");
  assert.deepEqual(r.patch.map((p) => p.index).sort((a, b) => a - b), [10, 11]);
});

test("flow: already-registered today's pass returns ALREADY, not a re-register", () => {
  const rows = [
    { ...rowA(7000, serial, "Dinner", "Yes", "REGISTERED"), index: 20 },
    { ...rowA(7000, serial, "Lunch", "Yes"), index: 21 },
  ];
  const r = evaluateScan({ headers: headersA, rows, orderNumber: "7000", now: satEve });
  assert.equal(r.decision, "ALREADY");
  assert.match(r.response, /already registered/i);
});

test("configurable event date + cutoff override the current day", () => {
  const rows = [{ ...rowA(8000, serial, "Lunch", "Yes"), index: 50 }];
  const r = evaluateScan({
    headers: headersA,
    rows,
    orderNumber: "8000",
    now: new Date("2026-08-12T12:00:00Z"),
    eventDate: "2025-09-26",
    cutoff: 1600,
  });
  assert.equal(r.decision, "SUCCESS");
  assert.equal(r.patch.length, 1);
});

// Date precedence: real "today" wins when it's an event date; else the override; else today.
const octSerial = excelSerial(new Date(2026, 9, 17));   // 2026-10-17
const octNoon = new Date("2026-10-17T10:00:00Z");        // 12:00 Oslo -> Lunch (cutoff 1600)

test("live event: today (an event date) wins over the eventDate override", () => {
  const rows = [{ ...rowA(9100, octSerial, "Lunch", "Yes"), index: 60 }];
  const r = evaluateScan({ headers: headersA, rows, orderNumber: "9100", now: octNoon, eventDate: "2025-09-26" });
  assert.equal(r.decision, "SUCCESS");     // today matches an event date -> override ignored
  assert.equal(r.patch[0].index, 60);
});

test("testing: today outside the event falls back to the eventDate override", () => {
  const rows = [{ ...rowA(9200, octSerial, "Lunch", "Yes"), index: 61 }];
  const r = evaluateScan({ headers: headersA, rows, orderNumber: "9200", now: new Date("2026-08-13T10:00:00Z"), eventDate: "2026-10-17" });
  assert.equal(r.decision, "SUCCESS");
  assert.equal(r.patch[0].index, 61);
});

test("no override + today outside the event -> INVALID", () => {
  const rows = [{ ...rowA(9300, octSerial, "Lunch", "Yes"), index: 62 }];
  const r = evaluateScan({ headers: headersA, rows, orderNumber: "9300", now: new Date("2026-08-13T10:00:00Z") });
  assert.equal(r.decision, "INVALID");
});

test("parking and food stall sample rows format and price correctly", () => {
  const foodHeaders = ["DateTime", "Day", "Name", "Item", "Qty", "UnitPrice", "Amount", "RecordedBy"];
  const foodRow = rowForHeaders(foodHeaders, [
    [["datetime", "date time", "time"], "2025-09-26T17:15:00.000"],
    [["day"], "Saturday"],
    [["name"], "Test Guest One"],
    [["item", "dish", "food"], "Veg Thali"],
    [["qty", "quantity"], 2],
    [["unitprice", "unit price", "price", "rate"], 120],
    [["amount", "total"], 240],
    [["recordedby", "recorded by", "volunteer", "by"], "volunteer@opb.no"],
  ]);
  assert.equal(foodRow[2], "Test Guest One");
  assert.equal(foodRow[3], "Veg Thali");
  assert.equal(foodRow[4], 2);
  assert.equal(parsePrice("kr 120,50"), 120.5);

  const parkingHeaders = ["Timestamp", "Sl No", "Name", "Mobile Number", "Car Registration number", "Car Make", "Car Model", "Car Colour"];
  const stamp = parkingStamp("2025-09-26", { h: 17, mi: 5, s: 9 });
  assert.equal(stamp, "2025-09-26 17:05:09");
  const parkingRow = rowForHeaders(parkingHeaders, [
    [["timestamp", "date", "time"], stamp],
    [["sl", "serial", "s.no", "sno", "#"], 1],
    [["name"], "Test Guest One"],
    [["mobile", "phone", "contact"], "+47 400 00 000"],
    [["registration", "reg", "plate", "number plate", "car number"], "EL12345"],
    [["make"], "Volvo"],
    [["model"], "XC60"],
    [["colour", "color"], "Silver"],
  ]);
  assert.equal(parkingRow[0], "2025-09-26 17:05:09");
  assert.equal(parkingRow[1], 1);
  assert.equal(parkingRow[4], "EL12345");
  assert.equal(parkingRow[7], "Silver");

  // The timestamp also lands in a legacy "Date"-named first column.
  const legacy = rowForHeaders(["Date", "Sl No", "Name"], [
    [["timestamp", "date", "time"], stamp], [["sl", "serial", "s.no", "sno", "#"], 2], [["name"], "A"],
  ]);
  assert.equal(legacy[0], "2025-09-26 17:05:09");
});

test("food menu parses the horizontal OPB price list (item headers, prices below)", () => {
  const headers = ["Veg Chop (1 stk)", "Singara (1 stk)", "Ghugni (1 plate)", "Cold Drink (1 stk)"];
  const dataRows = [[25, 30, 25, 25]];
  const items = parseFoodMenu(headers, dataRows);
  assert.equal(items.length, 4);
  assert.deepEqual(items[0], { item: "Veg Chop (1 stk)", price: 25 });
  assert.deepEqual(items[1], { item: "Singara (1 stk)", price: 30 });
});

test("food menu still parses a vertical Item/Price price list", () => {
  const headers = ["Item", "Price"];
  const dataRows = [["Samosa", 40], ["Jalebi", "kr 30"], ["Header-ish", 0]];
  const items = parseFoodMenu(headers, dataRows);
  assert.deepEqual(items, [{ item: "Samosa", price: 40 }, { item: "Jalebi", price: 30 }]);
});

const DUES_HEADERS = ["Sl No.", "Name", "Veg Chop (1 stk)", "Mochar Chop (1 stk)", "Malpoa (1 stk)", "Kheermohan (1 stk)", "Ghugni (1 plate)", "Dahi Vada (2 stk)", "Cold Drink (1 stk)", "Total"];
const PRICES = { "veg chop (1 stk)": 25, "mochar chop (1 stk)": 30, "malpoa (1 stk)": 25, "kheermohan (1 stk)": 30, "ghugni (1 plate)": 25, "dahi vada (2 stk)": 30, "cold drink (1 stk)": 25, "cup cake (1 stk)": 25 };
const priceOf = (n) => PRICES[String(n).toLowerCase()] || 0;

test("food dues: new guest fills first empty pre-numbered row and totals", () => {
  const rows = [
    [1, "Test Guest One", 1, "", "", 2, 2, 1, 3, 240],
    [2, "Test Guest Two", 1, 3, 2, "", 1, "", "", 190],
    [3, "", "", "", "", "", "", "", "", 0],
  ];
  const u = applyFoodDues(DUES_HEADERS, rows, "Test Guest", [
    { item: "Veg Chop (1 stk)", qty: 1 }, { item: "Mochar Chop (1 stk)", qty: 1 },
    { item: "Cup Cake (1 stk)", qty: 1 }, { item: "Cold Drink (1 stk)", qty: 1 },
  ], priceOf);
  assert.equal(u.rowIndex, 2);                       // the empty "3" row
  assert.equal(u.rowValues[0], 3);                   // keeps its Sl No.
  assert.equal(u.rowValues[1], "Test Guest");
  assert.equal(u.rowValues[2], 1);                   // Veg Chop
  assert.equal(u.rowValues[3], 1);                   // Mochar Chop
  assert.equal(u.rowValues[8], 1);                   // Cold Drink
  assert.equal(u.headerChanged, true);               // Cup Cake column appended
  assert.equal(u.headers[u.headers.length - 1], "Cup Cake (1 stk)");
  assert.equal(u.rowValues[u.rowValues.length - 1], 1);
  assert.equal(u.total, 105);                        // 25+30+25+25
});

test("food dues: existing guest accumulates quantities and recomputes Total", () => {
  const rows = [[1, "Test Guest One", 1, "", "", 2, 2, 1, 3, 240]];
  const u = applyFoodDues(DUES_HEADERS, rows, "test guest one", [{ item: "Veg Chop (1 stk)", qty: 1 }], priceOf);
  assert.equal(u.rowIndex, 0);
  assert.equal(u.rowValues[2], 2);                   // 1 + 1
  assert.equal(u.total, 265);                        // 2*25+2*30+2*25+1*30+3*25
});

test("parking: normalize aligns app rows and leaves manual rows intact", () => {
  assert.deepEqual(
    normalizeParkingRow([46282.9, 1, "Test Guest", 123456, "AB12345", "Toyota", "Corolla", "Black"]),
    [1, 46282.9, "Test Guest", 123456, "AB12345", "Toyota", "Corolla", "Black"]);
  assert.deepEqual(
    normalizeParkingRow([1, "", "Test Guest Two", 11111111, "AA11111", "Tesla", "Model X", "Blue"]),
    [1, "", "Test Guest Two", 11111111, "AA11111", "Tesla", "Model X", "Blue"]);
});

test("security: deFormula neutralizes spreadsheet formula injection, keeps phones", () => {
  assert.equal(deFormula("=HYPERLINK(\"http://evil\")"), "'=HYPERLINK(\"http://evil\")");
  assert.equal(deFormula("@SUM(A1)"), "'@SUM(A1)");
  assert.equal(deFormula("-cmd|calc"), "'-cmd|calc");
  assert.equal(deFormula("+47 400 00 000"), "+47 400 00 000"); // phone: sign then digit stays
  assert.equal(deFormula("Test Guest"), "Test Guest");           // ordinary text untouched
});
