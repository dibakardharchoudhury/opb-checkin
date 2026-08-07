// Business rules for OPB check-in, ported faithfully from the Power Automate flow
// OPB_Excel_QRCodeScannerFlow and made schema-flexible so it works with BOTH
// workbook variants:
//   A) discrete columns: RegistrationID/Order Number, Date, PassType, FoodOption, UniqueKey
//   B) collapsed key:     Order Number/RegistrationID, AppKey (= order+dateSerial+PassType+FoodOption)
//
// The scan supplies only the order/registration number (from the QR "<order>;").
// Date is "today" and the meal session is derived from the current Oslo time,
// exactly as the original flow did.

const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30); // 1899-12-30; Excel serial day 0

// Column name candidates (case-insensitive) for each logical field.
const COLS = {
  order: ["RegistrationID", "Order Number", "OrderNumber", "Order"],
  key: ["UniqueKey", "AppKey"],
  first: ["First Name", "FirstName"],
  last: ["Last Name", "LastName"],
  date: ["Date"],
  passType: ["PassType", "Pass Type"],
  foodOption: ["FoodOption", "Food Option"],
  status: ["Status"],
  dateTime: ["DateTime", "Date Time"],
};

function findCol(headers, candidates) {
  const lower = headers.map((h) => String(h ?? "").trim().toLowerCase());
  for (const c of candidates) {
    const i = lower.indexOf(c.toLowerCase());
    if (i !== -1) return i;
  }
  return -1;
}

// Resolve the logical->physical column index map for a given header row.
export function mapColumns(headers) {
  const m = {};
  for (const [logical, cands] of Object.entries(COLS)) m[logical] = findCol(headers, cands);
  if (m.order === -1) throw new Error("Workbook has no Order Number / RegistrationID column");
  if (m.status === -1) throw new Error("Workbook has no Status column");
  if (m.dateTime === -1) throw new Error("Workbook has no DateTime column");
  return m;
}

// Excel serial date (integer) for a given local calendar date in a timezone.
export function excelSerial(date) {
  const utcMidnight = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate());
  return Math.round((utcMidnight - EXCEL_EPOCH_UTC) / 86400000);
}

// The wall-clock calendar parts (Y/M/D/HHmm) in a timezone, without pulling in a tz lib.
export function nowInZone(tz, when = new Date()) {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: tz, year: "numeric", month: "2-digit", day: "2-digit",
    hour: "2-digit", minute: "2-digit", second: "2-digit", hour12: false,
  }).formatToParts(when).reduce((a, p) => ((a[p.type] = p.value), a), {});
  const y = +parts.year, mo = +parts.month, d = +parts.day;
  let h = +parts.hour; if (h === 24) h = 0;
  const mi = +parts.minute, s = +parts.second;
  return {
    y, mo, d, h, mi, s,
    dateOnly: new Date(y, mo - 1, d),               // local Date at midnight for serial math
    ymd: `${parts.year}-${parts.month}-${parts.day}`,
    hhmm: h * 100 + mi,
    // Sheet's DateTime string style, e.g. 2025-09-26T13:03:52.090
    iso: `${parts.year}-${parts.month}-${parts.day}T${parts.hour}:${parts.minute}:${parts.second}.` +
         String(when.getMilliseconds()).padStart(3, "0"),
  };
}

// The flow rule: after 14:00 only Dinner passes are valid, otherwise Lunch.
export function sessionFor(hhmm) {
  return hhmm > 1400 ? "Dinner" : "Lunch";
}

// Normalize a cell that holds a pass date (may be an Excel serial number, a Date,
// or a datetime string) to a YYYY-MM-DD string. Returns null if not a date.
export function passDateYMD(cell) {
  if (cell == null || cell === "") return null;
  if (typeof cell === "number" && isFinite(cell)) {
    const ms = EXCEL_EPOCH_UTC + Math.round(cell) * 86400000;
    const dt = new Date(ms);
    return `${dt.getUTCFullYear()}-${String(dt.getUTCMonth() + 1).padStart(2, "0")}-${String(dt.getUTCDate()).padStart(2, "0")}`;
  }
  const s = String(cell);
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  const d = new Date(s);
  if (!isNaN(d)) return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  return null;
}

const norm = (v) => String(v ?? "").trim();

// Core evaluation. rows = [{ index, values: [...] }] straight from the table read.
// Returns a decision plus the rows to patch. Pure & synchronous for easy testing.
//
// sheetScoped=true: the target worksheet IS the event session (e.g. "Saturday Dinner"),
// so every row already belongs to it — match on the order number alone and skip the
// date/time-of-day derivation the single-sheet flow used.
export function evaluateScan({ headers, rows, orderNumber, tz = "Europe/Oslo", now = new Date(), sheetScoped = false }) {
  const col = mapColumns(headers);
  const clock = nowInZone(tz, now);
  const session = sessionFor(clock.hhmm);
  const todaySerial = excelSerial(clock.dateOnly);
  const order = norm(orderNumber);

  const get = (values, idx) => (idx === -1 ? "" : values[idx]);
  const candidates = rows.filter((r) => norm(get(r.values, col.order)) === order);

  const nameOf = (r) => `${norm(get(r.values, col.first))} ${norm(get(r.values, col.last))}`.trim();
  const customerName = candidates.length ? nameOf(candidates[0]) : "Unknown";

  const hasDiscrete = col.date !== -1 && col.passType !== -1;

  let valid;
  if (sheetScoped) {
    valid = candidates; // the tab defines the session; order match is sufficient
  } else if (hasDiscrete) {
    valid = candidates.filter(
      (r) => passDateYMD(get(r.values, col.date)) === clock.ymd &&
             norm(get(r.values, col.passType)).toLowerCase() === session.toLowerCase()
    );
  } else if (col.key !== -1) {
    // Collapsed variant: key = order + dateSerial + PassType (+ FoodOption). Match on
    // the prefix so both veg/non-veg variants of the same session are caught.
    const prefix = `${order}${todaySerial}${session}`.toLowerCase();
    valid = candidates.filter((r) => norm(get(r.values, col.key)).toLowerCase().startsWith(prefix));
  } else {
    valid = candidates; // last resort: order only
  }

  if (candidates.length === 0)
    return { decision: "INVALID", response: "ERROR!!! This Pass is NOT valid at this moment!", customerName, patch: [], col, session };

  if (valid.length === 0)
    return { decision: "INVALID", response: "ERROR!!! This Pass is NOT valid at this moment!", customerName, patch: [], col, session };

  // Flow checks the first matched row's status.
  const alreadyReg = norm(get(valid[0].values, col.status)).toUpperCase() === "REGISTERED";
  if (alreadyReg)
    return { decision: "ALREADY", response: `ERROR!!!  ${customerName} is already registered!!!`, customerName, patch: [], col, session };

  const patch = valid.map((r) => ({
    index: r.index,
    values: r.values,
    statusCol: col.status,
    dateTimeCol: col.dateTime,
    newStatus: "REGISTERED",
    newDateTime: clock.iso,
  }));

  return {
    decision: "SUCCESS",
    response: `SUCCESS!! ${customerName} has been registered successfully!`,
    customerName,
    patch,
    col,
    session,
  };
}
