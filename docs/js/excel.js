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
  return XLSX.read(buf, { type: "array", cellDates: true });
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

export function getAmount(row) {
  for (const key of ["SUBI $", "Amount (CoCode Crcy)", "Amount"]) {
    if (row[key] != null && row[key] !== "") return row[key];
  }
  return null;
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
