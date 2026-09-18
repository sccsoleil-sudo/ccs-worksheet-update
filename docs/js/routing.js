/**
 * Route extraction / discrepancy rows to the correct worksheet by file name.
 *
 * Rules (from business naming):
 * - "Percentage Deduction" → Reason Code R15 (Version 1)
 * - "CCS Deduction(s)"     → Reason Code R07 (Advance V2)
 * - CPD files also match client group in the name (LCL, SDM, AMZ, …)
 * - LDB / LPD / PPD: division appears in the worksheet file name
 */

import { ALL_DIVS, CPD_CUSTOMER_GROUPS, WORKSHEET_PRESETS } from "./config.js";

export function customerGroupOf(customer) {
  const c = String(customer ?? "").trim();
  for (const [group, ids] of Object.entries(CPD_CUSTOMER_GROUPS)) {
    if (ids.includes(c)) return group;
  }
  return null;
}

export function reasonCodeOf(row) {
  const v = row["Reason Code"] ?? row.RC ?? "";
  return String(v).trim().toUpperCase();
}

/**
 * Parse worksheet file name → routing metadata.
 */
export function classifyWorksheetFilename(filename) {
  const n = String(filename || "").toLowerCase();

  let reason = null; // R15 | R07
  let version = null; // Version 1 | Advance V2
  if (n.includes("percentage deduction") || n.includes("percentage_deduction")) {
    reason = "R15";
    version = "Version 1";
  } else if (
    n.includes("ccs deduction") ||
    n.includes("ccs deductions") ||
    n.includes("ccs_deduction")
  ) {
    reason = "R07";
    version = "Advance V2";
  }

  let div = null;
  // Prefer explicit division tokens in the name
  for (const d of ALL_DIVS) {
    // word-ish match: " CPD " or "-CPD" or start/end
    const re = new RegExp(`(^|[^a-z])${d.toLowerCase()}([^a-z]|$)`);
    if (re.test(n)) {
      div = d;
      break;
    }
  }

  let clientGroup = null;
  // Longer names first so MCKESSON wins over partials
  const groups = Object.keys(CPD_CUSTOMER_GROUPS).sort((a, b) => b.length - a.length);
  for (const g of groups) {
    if (n.includes(g.toLowerCase())) {
      clientGroup = g;
      break;
    }
  }

  return {
    filename,
    reason,
    version,
    div,
    clientGroup,
    label: describeRoute({ reason, version, div, clientGroup }),
  };
}

export function describeRoute({ reason, version, div, clientGroup }) {
  const parts = [];
  if (reason) parts.push(reason);
  if (version) parts.push(version);
  if (div) parts.push(div);
  if (clientGroup) parts.push(`client ${clientGroup}`);
  return parts.length ? parts.join(" · ") : "unclassified (manual)";
}

/**
 * Does this data row belong on this worksheet (by name rules)?
 */
export function rowMatchesWorksheet(row, meta) {
  if (meta.reason) {
    if (reasonCodeOf(row) !== meta.reason) return false;
  }
  if (meta.div) {
    if (row.Div !== meta.div) return false;
  }
  // CPD client split only when filename indicates a client group
  if (meta.clientGroup) {
    const g = customerGroupOf(row.Customer);
    if (g !== meta.clientGroup) return false;
  }
  return true;
}

export function filterRowsForWorksheet(rows, meta) {
  return (rows || []).filter((r) => rowMatchesWorksheet(r, meta));
}

export function presetForWorksheetMeta(meta) {
  if (meta.version && WORKSHEET_PRESETS[meta.version]) return meta.version;
  if (meta.reason === "R15") return "Version 1";
  if (meta.reason === "R07") return "Advance V2";
  return "Advance V2";
}
