/** Excel helpers using SheetJS (global XLSX) */

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
  // bookVBA keeps macros so we can write back .xlsm
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

export function downloadWorkbookFile(workbook, filename) {
  const isXlsm = /\.xlsm$/i.test(filename);
  XLSX.writeFile(workbook, filename, {
    bookType: isXlsm ? "xlsm" : "xlsx",
    bookVBA: true,
  });
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
  for (const key of ["SUBI $", "Amount (CoCode Crcy)", "Amount"]) {
    if (row[key] != null && row[key] !== "") return row[key];
  }
  return null;
}

/** Map an extraction row onto worksheet column headers. */
export function mapRowToWorksheetHeaders(row, headers) {
  const out = {};
  for (const h of headers) {
    if (Object.prototype.hasOwnProperty.call(row, h) && row[h] != null && row[h] !== "") {
      out[h] = row[h];
      continue;
    }
    // Common amount aliases used across S4 export vs worksheet
    if (h === "SUBI $" || h === "Amount (CoCode Crcy)" || h === "Amount") {
      const amt = getAmount(row);
      if (amt != null && amt !== "") out[h] = amt;
    }
  }
  return out;
}

/**
 * Append rows to an existing sheet (matched to that sheet's header row).
 * Mutates workbook in place. Returns count appended.
 */
export function appendRowsToSheet(workbook, sheetName, rows) {
  if (!rows?.length) return 0;
  if (!workbook.Sheets[sheetName]) {
    throw new Error(`Sheet not found in worksheet: ${sheetName}`);
  }
  const headers = getSheetHeaders(workbook, sheetName);
  if (!headers.length) {
    throw new Error(`Sheet "${sheetName}" has no header row to map columns.`);
  }
  const mapped = rows.map((r) => mapRowToWorksheetHeaders(r, headers));
  XLSX.utils.sheet_add_json(workbook.Sheets[sheetName], mapped, {
    header: headers,
    skipHeader: true,
    origin: -1,
  });
  return mapped.length;
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
