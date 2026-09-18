import { WORKSHEET_PRESETS } from "./config.js";
import { downloadWorkbook, firstSheetRows, readWorkbook } from "./excel.js";
import {
  enrichRows,
  findDiscrepanciesCpd,
  findDiscrepanciesNonCpd,
  findNewCpd,
  findNewNonCpd,
  findNotOnWorksheet,
} from "./logic.js";

const state = {
  todayRows: null,
  prevRows: null,
  todayName: "",
  prevName: "",
};

function todayStr() {
  return new Date().toISOString().slice(0, 10);
}

function $(id) {
  return document.getElementById(id);
}

function setStatus(el, kind, text) {
  el.innerHTML = `<div class="status ${kind}">${text}</div>`;
}

function metricsHtml(items) {
  return `<div class="metrics">${items
    .map(([label, val]) => `<div class="metric"><b>${val}</b><span>${label}</span></div>`)
    .join("")}</div>`;
}

function previewTable(rows, limit = 40) {
  if (!rows?.length) return "";
  const cols = Object.keys(rows[0]).slice(0, 12);
  const slice = rows.slice(0, limit);
  const head = cols.map((c) => `<th>${escapeHtml(c)}</th>`).join("");
  const body = slice
    .map(
      (r) =>
        `<tr>${cols.map((c) => `<td>${escapeHtml(r[c])}</td>`).join("")}</tr>`
    )
    .join("");
  return `<div class="preview"><table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>`;
}

function escapeHtml(v) {
  if (v == null) return "";
  return String(v)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function dlBtn(label, sheets, filename) {
  const id = `dl_${Math.random().toString(36).slice(2, 9)}`;
  queueMicrotask(() => {
    const btn = document.getElementById(id);
    if (btn) {
      btn.addEventListener("click", () => downloadWorkbook(sheets, filename));
    }
  });
  return `<button class="secondary" id="${id}">${label}</button>`;
}

async function loadPairFile(input, metaEl, which) {
  const file = input.files?.[0];
  if (!file) return;
  metaEl.textContent = `Loading ${file.name}…`;
  try {
    const wb = await readWorkbook(file);
    const { rows } = firstSheetRows(wb);
    const enriched = enrichRows(rows);
    if (which === "today") {
      state.todayRows = enriched;
      state.todayName = file.name;
    } else {
      state.prevRows = enriched;
      state.prevName = file.name;
    }
    metaEl.textContent = `${file.name} · ${enriched.length} rows (auto-enriched)`;
  } catch (e) {
    metaEl.textContent = `Error: ${e.message}`;
  }
}

function needBoth(outEl) {
  if (!state.todayRows || !state.prevRows) {
    setStatus(outEl, "err", "Upload Today and Previous day files first.");
    return false;
  }
  return true;
}

// Tabs
document.querySelectorAll(".tab").forEach((tab) => {
  tab.addEventListener("click", () => {
    document.querySelectorAll(".tab").forEach((t) => t.classList.remove("active"));
    document.querySelectorAll(".section").forEach((s) => s.classList.remove("active"));
    tab.classList.add("active");
    $(tab.dataset.tab).classList.add("active");
  });
});

$("fileToday").addEventListener("change", () =>
  loadPairFile($("fileToday"), $("metaToday"), "today")
);
$("filePrev").addEventListener("change", () =>
  loadPairFile($("filePrev"), $("metaPrev"), "prev")
);

$("btnEnrich").addEventListener("click", () => {
  const out = $("enrichOut");
  if (!state.todayRows && !state.prevRows) {
    setStatus(out, "err", "Upload at least one file.");
    return;
  }
  let html = "";
  if (state.todayRows) {
    const unid = state.todayRows.filter((r) => r.Brand === "[UNIDENTIFIED]").length;
    html += `<p><strong>Today:</strong> ${escapeHtml(state.todayName)}</p>`;
    html += metricsHtml([
      ["Rows", state.todayRows.length],
      ["Unidentified brand", unid],
    ]);
    html += `<div class="actions">${dlBtn(
      "Download today enriched",
      { Enriched: state.todayRows },
      `CCS_today_enriched_${todayStr()}.xlsx`
    )}${dlBtn(
      `Save for tomorrow`,
      { Enriched: state.todayRows },
      `CCS_for_tomorrow_${todayStr()}.xlsx`
    )}</div>`;
    html += previewTable(state.todayRows);
  }
  if (state.prevRows) {
    html += `<p style="margin-top:1rem"><strong>Previous:</strong> ${escapeHtml(state.prevName)}</p>`;
    html += `<div class="actions">${dlBtn(
      "Download previous enriched",
      { Enriched: state.prevRows },
      `CCS_previous_enriched_${todayStr()}.xlsx`
    )}</div>`;
  }
  out.innerHTML = html;
});

$("btnNewNon").addEventListener("click", () => {
  const out = $("newOut");
  if (!needBoth(out)) return;
  const rows = findNewNonCpd(state.todayRows, state.prevRows);
  const sheets = {};
  if (rows.length) {
    sheets.All_New = rows;
    for (const div of ["LDB", "LPD", "PPD"]) {
      const part = rows.filter((r) => r.Div === div);
      if (part.length) sheets[div] = part;
    }
  }
  out.innerHTML =
    metricsHtml([["New non-CPD", rows.length]]) +
    `<div class="actions">${dlBtn(
      "Download CCS_Differences_nonCPD.xlsx",
      sheets,
      "CCS_Differences_nonCPD.xlsx"
    )}</div>` +
    previewTable(rows);
});

$("btnNewCpd").addEventListener("click", () => {
  const out = $("newOut");
  if (!needBoth(out)) return;
  const { sheets, unassigned } = findNewCpd(state.todayRows, state.prevRows);
  const n = sheets.All_Results?.length || 0;
  let html =
    metricsHtml([
      ["New CPD", n],
      ["Unassigned customer", unassigned.length],
    ]) +
    `<div class="actions">${dlBtn(
      "Download CCS_Difference_CPD.xlsx",
      sheets,
      "CCS_Difference_CPD.xlsx"
    )}`;
  if (unassigned.length) {
    html += dlBtn(
      "Download unassigned",
      { Unassigned: unassigned },
      "CPD_Unassigned_Rows.xlsx"
    );
  }
  html += `</div>` + previewTable(sheets.All_Results || []);
  out.innerHTML = html;
});

$("btnDiscCpd").addEventListener("click", () => {
  const out = $("discOut");
  if (!needBoth(out)) return;
  const sheets = findDiscrepanciesCpd(state.todayRows, state.prevRows);
  const n = sheets.All_Differing_Rows?.length || 0;
  out.innerHTML =
    metricsHtml([["CPD discrepancy rows", n]]) +
    `<div class="actions">${dlBtn(
      "Download CCS_Discrepancies_CPD.xlsx",
      sheets,
      "CCS_Discrepancies_CPD.xlsx"
    )}</div>` +
    previewTable(sheets.All_Differing_Rows || []);
});

$("btnDiscNon").addEventListener("click", () => {
  const out = $("discOut");
  if (!needBoth(out)) return;
  const sheets = findDiscrepanciesNonCpd(state.todayRows, state.prevRows);
  const n = sheets.All_Discrepancies?.length || 0;
  out.innerHTML =
    metricsHtml([["Non-CPD discrepancy rows", n]]) +
    `<div class="actions">${dlBtn(
      "Download CCS_Discrepancies_By_Div.xlsx",
      sheets,
      "CCS_Discrepancies_By_Div.xlsx"
    )}</div>` +
    previewTable(sheets.All_Discrepancies || []);
});

$("btnRunAll").addEventListener("click", () => {
  const out = $("runAllOut");
  if (!needBoth(out)) return;
  const non = findNewNonCpd(state.todayRows, state.prevRows);
  const { sheets: cpdSheets, unassigned } = findNewCpd(state.todayRows, state.prevRows);
  const discCpd = findDiscrepanciesCpd(state.todayRows, state.prevRows);
  const discNon = findDiscrepanciesNonCpd(state.todayRows, state.prevRows);

  const nonSheets = {};
  if (non.length) {
    nonSheets.All_New = non;
    for (const div of ["LDB", "LPD", "PPD"]) {
      const part = non.filter((r) => r.Div === div);
      if (part.length) nonSheets[div] = part;
    }
  }

  out.innerHTML =
    metricsHtml([
      ["New non-CPD", non.length],
      ["New CPD", cpdSheets.All_Results?.length || 0],
      ["CPD discrepancies", discCpd.All_Differing_Rows?.length || 0],
      ["Non-CPD discrepancies", discNon.All_Discrepancies?.length || 0],
    ]) +
    `<p class="hint">Download <strong>Save for tomorrow</strong> and keep it for the next day’s Previous file.</p>` +
    `<div class="actions">${[
      dlBtn(
        `Save for tomorrow`,
        { Enriched: state.todayRows },
        `CCS_for_tomorrow_${todayStr()}.xlsx`
      ),
      dlBtn("New lines · non-CPD", nonSheets, "CCS_Differences_nonCPD.xlsx"),
      dlBtn("New lines · CPD", cpdSheets, "CCS_Difference_CPD.xlsx"),
      dlBtn("Discrepancies · CPD", discCpd, "CCS_Discrepancies_CPD.xlsx"),
      dlBtn("Discrepancies · non-CPD", discNon, "CCS_Discrepancies_By_Div.xlsx"),
      unassigned.length
        ? dlBtn("CPD unassigned", { Unassigned: unassigned }, "CPD_Unassigned_Rows.xlsx")
        : "",
    ].join("")}</div>`;
});

// Worksheet check
let extractRows = null;
let wsWb = null;

$("fileExtract").addEventListener("change", async () => {
  const f = $("fileExtract").files?.[0];
  if (!f) return;
  $("metaExtract").textContent = `Loading ${f.name}…`;
  try {
    const wb = await readWorkbook(f);
    extractRows = enrichRows(firstSheetRows(wb).rows);
    $("metaExtract").textContent = `${f.name} · ${extractRows.length} rows`;
  } catch (e) {
    $("metaExtract").textContent = e.message;
  }
});

$("fileWs").addEventListener("change", async () => {
  const f = $("fileWs").files?.[0];
  if (!f) return;
  $("metaWs").textContent = `Loading ${f.name}…`;
  try {
    wsWb = await readWorkbook(f);
    $("metaWs").textContent = `${f.name} · sheets: ${wsWb.SheetNames.join(", ")}`;
  } catch (e) {
    $("metaWs").textContent = e.message;
  }
});

$("btnWs").addEventListener("click", () => {
  const out = $("wsOut");
  const preset = $("wsPreset").value;
  if (!extractRows || !wsWb) {
    setStatus(out, "err", "Upload extraction and worksheet files.");
    return;
  }
  try {
    const { missing, meta } = findNotOnWorksheet(extractRows, wsWb, preset);
    out.innerHTML =
      `<p class="hint">Sheets used: ${meta.sheetsRead.join(", ")} · Preset ${WORKSHEET_PRESETS[preset].label}</p>` +
      metricsHtml([
        ["Extraction (after CCS filter)", meta.extractionAfterFilter],
        ["Worksheet keys", meta.worksheetKeys],
        ["Not on worksheet", meta.notOnWorksheet],
      ]) +
      `<div class="actions">${dlBtn(
        `Download not_on_worksheet_${preset}.xlsx`,
        { Not_on_worksheet: missing },
        `not_on_worksheet_${preset.replace(/\s+/g, "_")}_${todayStr()}.xlsx`
      )}</div>` +
      previewTable(missing);
  } catch (e) {
    setStatus(out, "err", e.message);
  }
});
