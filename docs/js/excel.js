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

export function sheetToRows(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return [];
  return XLSX.utils.sheet_to_json(sheet, { defval: null, raw: false });
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
      const data = rows.length ? rows : [{ Status: "Empty" }];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data), safe);
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
    if (row[key] != null && row[key] !== "") return row[key];
  }
  return null;
}

function valueForHeader(row, header) {
  if (
    Object.prototype.hasOwnProperty.call(row, header) &&
    row[header] != null &&
    row[header] !== ""
  ) {
    return row[header];
  }
  const aliases = FIELD_ALIASES[header] || [header];
  for (const key of aliases) {
    if (row[key] != null && row[key] !== "") return row[key];
  }
  if (/^subi\s*\$$/i.test(header) || /^amount/i.test(header)) {
    return getAmount(row);
  }
  return null;
}

/** Map an extraction row onto worksheet column headers. */
export function mapRowToWorksheetHeaders(row, headers) {
  const out = {};
  for (const h of headers) {
    const v = valueForHeader(row, h);
    if (v != null && v !== "") out[h] = v;
  }
  return out;
}

function sheetLastRow(ws) {
  if (!ws["!ref"]) return 0; // 0-based: no data → next write at row 0
  return XLSX.utils.decode_range(ws["!ref"]).e.r;
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
 * Highlights new rows (yellow; discrepancy Yesterday rows in blue).
 * Mutates workbook in place. Returns { count, startRow }.
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

  // Color the appended block (xlsx-js-style)
  const colCount = Math.max(headers.length, sheetLastRow(ws) >= 0 ? headers.length : 1);
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
