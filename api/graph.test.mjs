import { test } from "node:test";
import assert from "node:assert/strict";
import { shapeUsedRange } from "./graph.js";

test("shapeUsedRange skips a title banner and merged label column (Parking layout)", () => {
  const grid = [
    ["", "PARKING LOG", "", "", "", "", "", ""],                                    // title banner row
    ["Friday: 26th September 2025", "Sl No.", "Name", "Mobile Number", "Car Registration number", "Car Make", "Car Model", "Car Colour"], // real header row (col A = merged day label)
    ["", "1", "Test Guest One", "11111111", "AA11111", "Tesla", "Model X", "Blue"],
    ["", "2", "Test Guest Two", "22222222", "BB22222", "Tesla", "Model Y", "White"],
    ["", "17", "", "", "", "", "", ""],                                             // pre-numbered empty row
  ];
  const { headers, rows } = shapeUsedRange(grid);
  assert.deepEqual(headers, ["Sl No.", "Name", "Mobile Number", "Car Registration number", "Car Make", "Car Model", "Car Colour"]);
  assert.equal(rows.length, 3); // header sliced off; the merged-day column dropped
  assert.deepEqual(rows[0].values, ["1", "Test Guest One", "11111111", "AA11111", "Tesla", "Model X", "Blue"]);
});

test("shapeUsedRange keeps a normal first-row header sheet unchanged", () => {
  const grid = [
    ["RegistrationID", "First Name", "Last Name", "Item"],
    ["3000", "Jeet", "Sarkar", "Season pass"],
    ["3001", "Sandeep", "Biswal", "Season pass"],
  ];
  const { headers, rows } = shapeUsedRange(grid);
  assert.deepEqual(headers, ["RegistrationID", "First Name", "Last Name", "Item"]);
  assert.equal(rows.length, 2);
  assert.deepEqual(rows[1].values, ["3001", "Sandeep", "Biswal", "Season pass"]);
});

test("shapeUsedRange returns empty for an empty grid", () => {
  assert.deepEqual(shapeUsedRange([]), { headers: [], rows: [] });
});
