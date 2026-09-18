import { ALL_DIVS, WORKSHEET_PRESETS } from "./config.js";
import {
  appendRowsToSheet,
  downloadWorkbook,
  downloadWorkbookFile,
  firstSheetRows,
  readWorkbook,
} from "./excel.js";
import {
  enrichRows,
  findDiscrepanciesAllDivs,
  findDiscrepanciesByDiv,
  findNewAllDivs,
  findNewByDiv,
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

$("btnNew").addEventListener("click", () => {
  const out = $("newOut");
  if (!needBoth(out)) return;
  const divSel = $("newDiv").value;

  if (divSel === "ALL") {
    const byDiv = findNewAllDivs(state.todayRows, state.prevRows);
    const metrics = ALL_DIVS.map((d) => [`New ${d}`, byDiv[d].rows.length]);
    const buttons = ALL_DIVS.map((d) =>
      dlBtn(
        `New lines · ${d}`,
        byDiv[d].sheets,
        `CCS_NewLines_${d}.xlsx`
      )
    );
    const cpdUn = byDiv.CPD.unassigned;
    if (cpdUn.length) {
      buttons.push(
        dlBtn("CPD unassigned", { Unassigned: cpdUn }, "CPD_Unassigned_Rows.xlsx")
      );
    }
    const previewDiv = ALL_DIVS.find((d) => byDiv[d].rows.length) || "CPD";
    out.innerHTML =
      metricsHtml(metrics) +
      `<div class="actions">${buttons.join("")}</div>` +
      previewTable(byDiv[previewDiv].rows);
    return;
  }

  const { rows, sheets, unassigned } = findNewByDiv(
    state.todayRows,
    state.prevRows,
    divSel
  );
  let html =
    metricsHtml([
      [`New ${divSel}`, rows.length],
      ...(divSel === "CPD" ? [["Unassigned customer", unassigned.length]] : []),
    ]) +
    `<div class="actions">${dlBtn(
      `Download CCS_NewLines_${divSel}.xlsx`,
      sheets,
      `CCS_NewLines_${divSel}.xlsx`
    )}`;
  if (unassigned.length) {
    html += dlBtn(
      "Download unassigned",
      { Unassigned: unassigned },
      "CPD_Unassigned_Rows.xlsx"
    );
  }
  html += `</div>` + previewTable(rows);
  out.innerHTML = html;
});

$("btnDisc").addEventListener("click", () => {
  const out = $("discOut");
  if (!needBoth(out)) return;
  const divSel = $("discDiv").value;
  let allRows = [];

  if (divSel === "ALL") {
    const byDiv = findDiscrepanciesAllDivs(state.todayRows, state.prevRows);
    const metrics = ALL_DIVS.map((d) => [
      `${d} discrepancies`,
      byDiv[d].All_Discrepancies?.length || 0,
    ]);
    const buttons = ALL_DIVS.map((d) =>
      dlBtn(
        `Discrepancies · ${d}`,
        byDiv[d],
        `CCS_Discrepancies_${d}.xlsx`
      )
    );
    for (const d of ALL_DIVS) {
      allRows = allRows.concat(byDiv[d].All_Discrepancies || []);
    }
    const previewDiv =
      ALL_DIVS.find((d) => byDiv[d].All_Discrepancies?.length) || "CPD";
    out.innerHTML =
      metricsHtml(metrics) +
      `<div class="actions">${buttons.join("")}</div>` +
      previewTable(byDiv[previewDiv].All_Discrepancies || []);
  } else {
    const sheets = findDiscrepanciesByDiv(
      state.todayRows,
      state.prevRows,
      divSel
    );
    allRows = sheets.All_Discrepancies || [];
    const n = allRows.length;
    out.innerHTML =
      metricsHtml([[`${divSel} discrepancy rows`, n]]) +
      `<div class="actions">${dlBtn(
        `Download CCS_Discrepancies_${divSel}.xlsx`,
        sheets,
        `CCS_Discrepancies_${divSel}.xlsx`
      )}</div>` +
      previewTable(allRows);
  }

  pendingDiscRows = allRows;
  if (allRows.length) {
    showDiscAppendPanel();
  } else {
    $("discAppendPanel").style.display = "none";
  }
});

function preferredDiscSheet(sheetNames) {
  const preset = WORKSHEET_PRESETS[$("discWsPreset").value];
  const preferred = preset?.discAppendSheet;
  if (preferred && sheetNames.includes(preferred)) return preferred;
  const fallbacks = [
    "KAMs - To correct",
    "To correct",
    "Cleared",
    "Export",
    "KAMs",
  ];
  return fallbacks.find((n) => sheetNames.includes(n)) || sheetNames[0];
}

function fillDiscAppendSheetSelect() {
  const wb = discWsWb || wsWb;
  const sel = $("discAppendSheet");
  sel.innerHTML = "";
  if (!wb) return;
  for (const name of wb.SheetNames) {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    sel.appendChild(opt);
  }
  sel.value = preferredDiscSheet(wb.SheetNames);
}

function showDiscAppendPanel() {
  $("discAppendPanel").style.display = "block";
  const wb = discWsWb || wsWb;
  if (wb) {
    fillDiscAppendSheetSelect();
    setStatus(
      $("discAppendOut"),
      "info",
      `${pendingDiscRows.length} discrepancy row(s) ready. Choose Today-only or both, pick the sheet, then download.`
    );
  } else {
    $("discAppendSheet").innerHTML = "";
    setStatus(
      $("discAppendOut"),
      "info",
      `${pendingDiscRows.length} discrepancy row(s) found. Upload a worksheet below (or under vs Worksheet), then append.`
    );
  }
}

let pendingDiscRows = [];
let discWsWb = null;
let discWsFileName = "";

$("fileDiscWs").addEventListener("change", async () => {
  const f = $("fileDiscWs").files?.[0];
  if (!f) return;
  $("metaDiscWs").textContent = `Loading ${f.name}…`;
  try {
    discWsWb = await readWorkbook(f);
    discWsFileName = f.name;
    $("metaDiscWs").textContent = `${f.name} · sheets: ${discWsWb.SheetNames.join(", ")}${
      discWsWb.vbaraw ? " · macros detected" : ""
    }`;
    if (pendingDiscRows.length) fillDiscAppendSheetSelect();
  } catch (e) {
    $("metaDiscWs").textContent = e.message;
  }
});

$("discWsPreset").addEventListener("change", () => {
  if (discWsWb || wsWb) fillDiscAppendSheetSelect();
});

$("btnAppendDisc").addEventListener("click", async () => {
  const out = $("discAppendOut");
  if (!pendingDiscRows.length) {
    setStatus(out, "err", "Check discrepancies first.");
    return;
  }
  const mode = $("discRowMode").value;
  let rows = pendingDiscRows;
  if (mode === "today") {
    rows = pendingDiscRows.filter((r) => r.Source === "Today");
  }
  if (!rows.length) {
    setStatus(out, "err", "No rows to append for the selected mode.");
    return;
  }

  const file =
    $("fileDiscWs").files?.[0] || $("fileWs").files?.[0] || null;
  if (!file) {
    setStatus(
      out,
      "err",
      "Upload a worksheet (.xlsx / .xlsm) here or under vs Worksheet."
    );
    return;
  }

  const sheetName = $("discAppendSheet").value;
  if (!sheetName) {
    setStatus(out, "err", "Select a target sheet.");
    return;
  }

  try {
    const freshWb = await readWorkbook(file);
    // Drop Source helper col if worksheet has no such header — mapRow only keeps matching headers
    const n = appendRowsToSheet(freshWb, sheetName, rows);
    const nameHint = discWsFileName || wsFileName || file.name;
    const base = nameHint.replace(/\.(xlsx|xls|xlsm)$/i, "");
    const outName = freshWb.vbaraw
      ? `${base}_discrepancies_added_${todayStr()}.xlsm`
      : `${base}_discrepancies_added_${todayStr()}.xlsx`;
    downloadWorkbookFile(freshWb, outName);
    setStatus(
      out,
      "ok",
      `Added ${n} discrepancy row(s) to “${sheetName}”. Downloaded ${outName}. Confirm macros before replacing your working file.`
    );
  } catch (e) {
    setStatus(out, "err", e.message);
  }
});

$("btnRunAll").addEventListener("click", () => {
  const out = $("runAllOut");
  if (!needBoth(out)) return;
  const newByDiv = findNewAllDivs(state.todayRows, state.prevRows);
  const discByDiv = findDiscrepanciesAllDivs(state.todayRows, state.prevRows);

  const metrics = [];
  for (const d of ALL_DIVS) {
    metrics.push([`New ${d}`, newByDiv[d].rows.length]);
    metrics.push([
      `${d} disc.`,
      discByDiv[d].All_Discrepancies?.length || 0,
    ]);
  }

  const buttons = [
    dlBtn(
      `Save for tomorrow`,
      { Enriched: state.todayRows },
      `CCS_for_tomorrow_${todayStr()}.xlsx`
    ),
  ];
  for (const d of ALL_DIVS) {
    buttons.push(
      dlBtn(`New · ${d}`, newByDiv[d].sheets, `CCS_NewLines_${d}.xlsx`)
    );
    buttons.push(
      dlBtn(
        `Disc · ${d}`,
        discByDiv[d],
        `CCS_Discrepancies_${d}.xlsx`
      )
    );
  }
  if (newByDiv.CPD.unassigned.length) {
    buttons.push(
      dlBtn(
        "CPD unassigned",
        { Unassigned: newByDiv.CPD.unassigned },
        "CPD_Unassigned_Rows.xlsx"
      )
    );
  }

  out.innerHTML =
    metricsHtml(metrics) +
    `<p class="hint">Results are split by <strong>division</strong> (CPD / LDB / LPD / PPD). Download <strong>Save for tomorrow</strong> for the next day.</p>` +
    `<div class="actions">${buttons.join("")}</div>`;
});

// Worksheet check
let extractRows = null;
let wsWb = null;
let wsFileName = "";
let pendingAppendRows = [];

function fillAppendSheetSelect(preferred) {
  const sel = $("wsAppendSheet");
  sel.innerHTML = "";
  for (const name of wsWb.SheetNames) {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    sel.appendChild(opt);
  }
  if (preferred && wsWb.SheetNames.includes(preferred)) {
    sel.value = preferred;
  }
}

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
    wsFileName = f.name;
    const hasVba = Boolean(wsWb.vbaraw);
    $("metaWs").textContent = `${f.name} · sheets: ${wsWb.SheetNames.join(", ")}${
      hasVba ? " · macros detected" : ""
    }`;
    $("wsAppendPanel").style.display = "none";
    pendingAppendRows = [];
  } catch (e) {
    $("metaWs").textContent = e.message;
  }
});

$("btnWs").addEventListener("click", () => {
  const out = $("wsOut");
  const preset = $("wsPreset").value;
  const divFilter = $("wsDiv").value || null;
  if (!extractRows || !wsWb) {
    setStatus(out, "err", "Upload extraction and worksheet files.");
    return;
  }
  try {
    const { missing, sheets, meta } = findNotOnWorksheet(
      extractRows,
      wsWb,
      preset,
      divFilter
    );
    pendingAppendRows = missing;
    const divMetrics = ALL_DIVS.map((d) => [d, (sheets[d] || []).length]);
    out.innerHTML =
      `<p class="hint">Sheets used: ${meta.sheetsRead.join(", ")} · ${WORKSHEET_PRESETS[preset].label}${divFilter ? ` · filter ${divFilter}` : ""}</p>` +
      metricsHtml([
        ["Extraction (after filter)", meta.extractionAfterFilter],
        ["Worksheet keys", meta.worksheetKeys],
        ["Not on worksheet", meta.notOnWorksheet],
        ...divMetrics,
      ]) +
      `<div class="actions">${dlBtn(
        `Download missing rows only`,
        sheets,
        `not_on_worksheet_${preset.replace(/\s+/g, "_")}_${todayStr()}.xlsx`
      )}</div>` +
      previewTable(missing);

    if (missing.length) {
      fillAppendSheetSelect(WORKSHEET_PRESETS[preset].appendSheet);
      $("wsAppendPanel").style.display = "block";
      setStatus(
        $("wsAppendOut"),
        "info",
        `${missing.length} row(s) ready to append. Choose the target sheet, then download the updated worksheet.`
      );
    } else {
      $("wsAppendPanel").style.display = "none";
    }
  } catch (e) {
    setStatus(out, "err", e.message);
    $("wsAppendPanel").style.display = "none";
  }
});

$("btnAppendWs").addEventListener("click", async () => {
  const out = $("wsAppendOut");
  if (!wsWb || !pendingAppendRows.length) {
    setStatus(out, "err", "Find missing rows first.");
    return;
  }
  const sheetName = $("wsAppendSheet").value;
  try {
    // Re-read original file so repeated clicks don’t double-append
    const file = $("fileWs").files?.[0];
    if (!file) throw new Error("Worksheet file missing — upload again.");
    const freshWb = await readWorkbook(file);
    const n = appendRowsToSheet(freshWb, sheetName, pendingAppendRows);
    const base = (wsFileName || "worksheet").replace(/\.(xlsx|xls|xlsm)$/i, "");
    const outName = freshWb.vbaraw
      ? `${base}_updated_${todayStr()}.xlsm`
      : `${base}_updated_${todayStr()}.xlsx`;
    downloadWorkbookFile(freshWb, outName);
    setStatus(
      out,
      "ok",
      `Added ${n} row(s) to “${sheetName}”. Downloaded ${outName}. Open it and confirm macros/formulas look correct before replacing your working file.`
    );
  } catch (e) {
    setStatus(out, "err", e.message);
  }
});
