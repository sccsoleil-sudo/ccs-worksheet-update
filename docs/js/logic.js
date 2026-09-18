import {
  ALL_DIVS,
  AMOUNT_COL,
  BRAND_AXE_MAPPING,
  CPD_CUSTOMER_GROUPS,
  DIV_MAPPING,
  WORKSHEET_PRESETS,
} from "./config.js";
import {
  amountAssignmentKey,
  getAmount,
  hasCcsReference,
  normalizeDisputeId,
  sheetToRows,
} from "./excel.js";

function brandAxe(row) {
  const text = row["Item Text"] != null ? String(row["Item Text"]).slice(0, 4).toUpperCase() : "";
  const div = row.Div;
  for (const e of BRAND_AXE_MAPPING) {
    if (e.acronym === text && e.div === div) return { brand: e.brand, axe: e.axe };
  }
  return { brand: "[UNIDENTIFIED]", axe: "*General*" };
}

export function enrichRows(rows) {
  return rows.map((r) => {
    const div = DIV_MAPPING[r["Business Area"]] || "[UNIDENTIFIED]";
    const { brand, axe } = brandAxe({ ...r, Div: div });
    const { Div: _d, Brand: _b, Axe: _a, ...rest } = r;
    return { Div: div, Brand: brand, Axe: axe, ...rest };
  });
}

function customerGroup(customer) {
  const c = String(customer).trim();
  for (const [group, ids] of Object.entries(CPD_CUSTOMER_GROUPS)) {
    if (ids.includes(c)) return group;
  }
  return null;
}

/** New lines for one division (Amount + Assignment; skip CCS Reference). */
export function findNewByDiv(todayRows, yesterdayRows, div) {
  const lastKeys = new Set(yesterdayRows.map(amountAssignmentKey));
  let result = todayRows.filter((r) => {
    if (r.Div !== div) return false;
    if (hasCcsReference(r)) return false;
    return !lastKeys.has(amountAssignmentKey(r));
  });

  if (div === "CPD") {
    result = result.filter((r) => r.Customer != null && r.Customer !== "");
  }

  const sheets = {};
  const unassigned = [];
  if (!result.length) return { rows: result, sheets, unassigned };

  sheets.All_Results = result;
  sheets[div] = result;

  if (div === "CPD") {
    for (const group of Object.keys(CPD_CUSTOMER_GROUPS)) sheets[group] = [];
    for (const row of result) {
      const g = customerGroup(row.Customer);
      if (g) sheets[g].push(row);
      else unassigned.push(row);
    }
    for (const g of Object.keys(CPD_CUSTOMER_GROUPS)) {
      if (!sheets[g].length) delete sheets[g];
    }
  }

  return { rows: result, sheets, unassigned };
}

/** New lines for every division → { CPD: {...}, LDB: {...}, ... } */
export function findNewAllDivs(todayRows, yesterdayRows) {
  const byDiv = {};
  for (const div of ALL_DIVS) {
    byDiv[div] = findNewByDiv(todayRows, yesterdayRows, div);
  }
  return byDiv;
}

/** Discrepancies for one division (same Dispute ID, Assignment / Item Text / Div change). */
export function findDiscrepanciesByDiv(todayRows, yesterdayRows, div) {
  const newF = todayRows.filter(
    (r) => r.Div === div && normalizeDisputeId(r["Dispute ID"])
  );
  const lastF = yesterdayRows.filter(
    (r) => r.Div === div && normalizeDisputeId(r["Dispute ID"])
  );

  const lastById = new Map();
  for (const r of lastF) {
    const id = normalizeDisputeId(r["Dispute ID"]);
    if (!lastById.has(id)) lastById.set(id, []);
    lastById.get(id).push(r);
  }

  const out = [];
  const seen = new Set();
  for (const n of newF) {
    const id = normalizeDisputeId(n["Dispute ID"]);
    const lasts = lastById.get(id) || [];
    for (const l of lasts) {
      const a1 = String(n.Assignment ?? "").trim();
      const a2 = String(l.Assignment ?? "").trim();
      const t1 = String(n["Item Text"] ?? "").trim();
      const t2 = String(l["Item Text"] ?? "").trim();
      const d1 = String(n.Div ?? "");
      const d2 = String(l.Div ?? "");
      if (a1 !== a2 || t1 !== t2 || d1 !== d2) {
        const pairKey = `${id}|${a1}|${t1}|${a2}|${t2}|${d1}|${d2}`;
        if (seen.has(pairKey)) continue;
        seen.add(pairKey);
        out.push({ ...n, Source: "Today" });
        out.push({ ...l, Source: "Yesterday" });
      }
    }
  }

  const sheets = {};
  if (!out.length) return sheets;
  sheets.All_Discrepancies = out;
  sheets[div] = out;

  if (div === "CPD") {
    for (const [group, ids] of Object.entries(CPD_CUSTOMER_GROUPS)) {
      const part = out.filter((r) => ids.includes(String(r.Customer ?? "").trim()));
      if (part.length) sheets[group] = part;
    }
  }
  return sheets;
}

export function findDiscrepanciesAllDivs(todayRows, yesterdayRows) {
  const byDiv = {};
  for (const div of ALL_DIVS) {
    byDiv[div] = findDiscrepanciesByDiv(todayRows, yesterdayRows, div);
  }
  return byDiv;
}

function ensureSubi(row) {
  if (row["SUBI $"] != null && row["SUBI $"] !== "") return row;
  const amt = getAmount(row);
  return { ...row, "SUBI $": amt };
}

export function findNotOnWorksheet(extractionRows, worksheetWb, presetKey, divFilter = null) {
  const preset = WORKSHEET_PRESETS[presetKey];
  const sheetsRead = [];
  const keys = new Set();

  for (const name of preset.sheets) {
    if (!worksheetWb.SheetNames.includes(name)) continue;
    const rows = sheetToRows(worksheetWb, name).map(ensureSubi);
    sheetsRead.push(name);
    for (const r of rows) {
      if (r.Assignment == null) continue;
      keys.add(`${r["SUBI $"]}|${r.Assignment}`);
    }
  }
  if (!sheetsRead.length) {
    throw new Error(`None of the expected sheets found: ${preset.sheets.join(", ")}`);
  }

  let filtered = extractionRows.map(ensureSubi).filter((r) => r.Assignment != null);
  filtered = filtered.filter((r) => !hasCcsReference(r));
  if (divFilter) filtered = filtered.filter((r) => r.Div === divFilter);

  const missing = filtered.filter((r) => !keys.has(`${r["SUBI $"]}|${r.Assignment}`));

  const sheets = { Not_on_worksheet: missing };
  for (const div of ALL_DIVS) {
    const part = missing.filter((r) => r.Div === div);
    if (part.length) sheets[div] = part;
  }

  return {
    missing,
    sheets,
    meta: {
      sheetsRead,
      extractionAfterFilter: filtered.length,
      worksheetKeys: keys.size,
      notOnWorksheet: missing.length,
    },
  };
}

void AMOUNT_COL;
