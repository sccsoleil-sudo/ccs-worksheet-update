/** Excel helpers using SheetJS / xlsx-js-style (global XLSX) */

const AMOUNT_KEYS = ["SUBI $", "SUBI$", "Amount (CoCode Crcy)", "Amount"];

/** Map S4 / enriched column names → worksheet column names */
const FIELD_ALIASES = {
  "SUBI $": AMOUNT_KEYS,
  "SUBI$": AMOUNT_KEYS,
  Amount: AMOUNT_KEYS,
  "Amount (CoCode Crcy)": AMOUNT_KEYS,
  Text: ["Text", "Item Text"],
  "Item Text": ["Item Text", "Text"],
  Name: ["Name", "Customer Name"],
  "Customer Name": ["Customer Name", "Name"],
  RC: ["RC", "Reason Code"],
  "Reason Code": ["Reason Code", "RC"],
};

const FILL_TODAY = {
  patternType: "solid",
  fgColor: { rgb: "FFF2CC" }, // light yellow
};
const FILL_YESTERDAY = {
  patternType: "solid",
  fgColor: { rgb: "BDD7EE" }, // light blue
};
const FILL_NEW = FILL_TODAY;

const AMOUNT_NUM_FMT = "#,##0.00"; // number only — no CAD text

function isAmountHeader(header) {
  const h = String(header || "");
  return AMOUNT_KEYS.includes(h) || /^subi\s*\$$/i.test(h) || /^amount\b/i.test(h);
}

/** Strip CAD / $ / commas and return a JS number (or null). */
export function toPlainNumber(value) {
  if (value == null || value === "") return null;
  if (typeof value === "number" && Number.isFinite(value)) return value;
  let s = String(value).trim();
  if (!s) return null;
  s = s
    .replace(/CAD/gi, "")
    .replace(/USD/gi, "")
    .replace(/€/g, "")
    .replace(/£/g, "")
    .replace(/\$/g, "")
    .replace(/\s/g, "")
    .replace(/,/g, "");
  s = s.replace(/[^0-9.\-]/g, "");
  if (!s || s === "-" || s === ".") return null;
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

export function sheetToRows(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return [];
  // raw:true keeps numeric cells as numbers (not "1,234.00 CAD" text)
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: null, raw: true });
  return rows.map((row) => {
    const out = { ...row };
    for (const key of Object.keys(out)) {
      if (isAmountHeader(key)) {
        const n = toPlainNumber(out[key]);
        if (n != null) out[key] = n;
      }
    }
    return out;
  });
}

export function firstSheetRows(workbook) {
  const name = workbook.SheetNames[0];
  return { name, rows: sheetToRows(workbook, name) };
}

export async function readWorkbook(file) {
  const buf = await file.arrayBuffer();
  return XLSX.read(buf, { type: "array", cellDates: true, bookVBA: true });
}

export function downloadWorkbook(sheets, filename) {
  const wb = XLSX.utils.book_new();
  const entries = Object.entries(sheets);
  if (!entries.length) {
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet([{ Status: "No rows" }]),
      "Summary"
    );
  } else {
    for (const [name, rows] of entries) {
      const safe = String(name).slice(0, 31);
      const data = rows.length ? rows.map(normalizeAmountFields) : [{ Status: "Empty" }];
      const ws = XLSX.utils.json_to_sheet(data);
      forceAmountCellsNumeric(ws);
      XLSX.utils.book_append_sheet(wb, ws, safe);
    }
  }
  XLSX.writeFile(wb, filename);
}

/**
 * Download workbook. If source was .xlsm (or preferXlsm), write .xlsm with VBA blob.
 */
export function downloadWorkbookFile(workbook, filename, { preferXlsm = false } = {}) {
  const isXlsm = preferXlsm || /\.xlsm$/i.test(filename);
  const outName = isXlsm
    ? filename.replace(/\.(xlsx|xls|xlsm)$/i, ".xlsm")
    : filename.replace(/\.(xlsx|xls|xlsm)$/i, ".xlsx");
  XLSX.writeFile(workbook, outName, {
    bookType: isXlsm ? "xlsm" : "xlsx",
    bookVBA: true,
    cellStyles: true,
  });
  return outName;
}

export function getSheetHeaders(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return [];
  const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
  if (!aoa.length) return [];
  return aoa[0]
    .map((h) => (h == null ? "" : String(h).trim()))
    .filter((h) => h !== "");
}

export function getAmount(row) {
  for (const key of AMOUNT_KEYS) {
    if (row[key] != null && row[key] !== "") {
      const n = toPlainNumber(row[key]);
      if (n != null) return n;
    }
  }
  return null;
}

function valueForHeader(row, header) {
  if (
    Object.prototype.hasOwnProperty.call(row, header) &&
    row[header] != null &&
    row[header] !== ""
  ) {
    return isAmountHeader(header) ? toPlainNumber(row[header]) ?? row[header] : row[header];
  }
  const aliases = FIELD_ALIASES[header] || [header];
  for (const key of aliases) {
    if (row[key] != null && row[key] !== "") {
      return isAmountHeader(header) || isAmountHeader(key)
        ? toPlainNumber(row[key]) ?? row[key]
        : row[key];
    }
  }
  if (isAmountHeader(header)) {
    return getAmount(row);
  }
  return null;
}

function normalizeAmountFields(row) {
  const out = { ...row };
  for (const key of Object.keys(out)) {
    if (isAmountHeader(key)) {
      const n = toPlainNumber(out[key]);
      if (n != null) out[key] = n;
    }
  }
  return out;
}

/** Map an extraction row onto worksheet column headers. */
export function mapRowToWorksheetHeaders(row, headers) {
  const out = {};
  for (const h of headers) {
    const v = valueForHeader(row, h);
    if (v != null && v !== "") out[h] = v;
  }
  return normalizeAmountFields(out);
}

function sheetLastRow(ws) {
  if (!ws["!ref"]) return 0;
  return XLSX.utils.decode_range(ws["!ref"]).e.r;
}

function forceAmountCellsNumeric(ws, headers = null, startRow = 0, rowCount = null) {
  if (!ws["!ref"]) return;
  const range = XLSX.utils.decode_range(ws["!ref"]);
  let headerList = headers;
  if (!headerList) {
    headerList = [];
    for (let C = range.s.c; C <= range.e.c; C++) {
      const addr = XLSX.utils.encode_cell({ r: range.s.r, c: C });
      const cell = ws[addr];
      headerList.push(cell && cell.v != null ? String(cell.v).trim() : "");
    }
  }
  const amountCols = [];
  headerList.forEach((h, i) => {
    if (isAmountHeader(h)) amountCols.push(i);
  });
  if (!amountCols.length) return;

  const r0 = headers ? startRow : Math.max(startRow, range.s.r + 1);
  const r1 = rowCount != null ? startRow + rowCount - 1 : range.e.r;
  for (let R = r0; R <= r1; R++) {
    for (const C of amountCols) {
      const addr = XLSX.utils.encode_cell({ r: R, c: C });
      const cell = ws[addr];
      if (!cell) continue;
      const n = toPlainNumber(cell.v);
      if (n == null) continue;
      cell.t = "n";
      cell.v = n;
      delete cell.w;
      cell.z = AMOUNT_NUM_FMT;
    }
  }
}

function highlightRowRange(ws, startRow, rowCount, sourceRows, colCount) {
  for (let i = 0; i < rowCount; i++) {
    const R = startRow + i;
    const src = sourceRows[i];
    const fill =
      src?.Source === "Yesterday"
        ? FILL_YESTERDAY
        : src?.Source === "Today"
          ? FILL_TODAY
          : FILL_NEW;
    for (let C = 0; C < colCount; C++) {
      const addr = XLSX.utils.encode_cell({ r: R, c: C });
      if (!ws[addr]) ws[addr] = { t: "z" };
      ws[addr].s = { ...(ws[addr].s || {}), fill };
    }
  }
}

/**
 * Append rows to an existing sheet (matched to that sheet's header row).
 * Amount columns are written as numbers (no CAD text).
 * Highlights new rows (yellow; discrepancy Yesterday rows in blue).
 */
export function appendRowsToSheet(workbook, sheetName, rows) {
  if (!rows?.length) return { count: 0, startRow: -1 };
  if (!workbook.Sheets[sheetName]) {
    throw new Error(`Sheet not found in worksheet: ${sheetName}`);
  }
  const ws = workbook.Sheets[sheetName];
  const headers = getSheetHeaders(workbook, sheetName);
  if (!headers.length) {
    throw new Error(`Sheet "${sheetName}" has no header row to map columns.`);
  }

  const startRow = sheetLastRow(ws) + 1;
  const mapped = rows.map((r) => mapRowToWorksheetHeaders(r, headers));
  XLSX.utils.sheet_add_json(ws, mapped, {
    header: headers,
    skipHeader: true,
    origin: -1,
  });

  const colCount = Math.max(headers.length, 1);
  forceAmountCellsNumeric(ws, headers, startRow, mapped.length);
  highlightRowRange(ws, startRow, mapped.length, rows, colCount);

  return { count: mapped.length, startRow };
}

export function amountAssignmentKey(row) {
  return `${getAmount(row)}|${row.Assignment ?? ""}`;
}

export function normalizeDisputeId(v) {
  if (v == null || v === "") return null;
  let s = String(v).trim();
  if (s.endsWith(".0")) s = s.slice(0, -2);
  if (s.toLowerCase() === "nan") return null;
  return s;
}

export function hasCcsReference(row) {
  const ref = row.Reference;
  if (ref == null) return false;
  return String(ref).toLowerCase().includes("ccs");
}

/** Default append target: preset sheet, or first tab when null. */
export function defaultAppendSheet(workbook, preset) {
  const preferred = preset?.appendSheet;
  if (preferred && workbook.SheetNames.includes(preferred)) return preferred;
  return workbook.SheetNames[0];
}

export function defaultDiscAppendSheet(workbook, preset) {
  const preferred = preset?.discAppendSheet;
  if (preferred && workbook.SheetNames.includes(preferred)) return preferred;
  return workbook.SheetNames[0];
}
