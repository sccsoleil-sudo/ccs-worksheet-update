import { ALL_DIVS, WORKSHEET_PRESETS } from "./config.js?v=20260921cad";
import {
  appendRowsToSheet,
  defaultAppendSheet,
  defaultDiscAppendSheet,
  downloadWorkbook,
  downloadWorkbookFile,
  firstSheetRows,
  readWorkbook,
} from "./excel.js?v=20260921cad";
import {
  enrichRows,
  findDiscrepanciesAllDivs,
  findDiscrepanciesByDiv,
  findNewAllDivs,
  findNewByDiv,
  findNotOnWorksheet,
} from "./logic.js?v=20260921cad";
import {
  classifyWorksheetFilename,
  filterRowsForWorksheet,
  presetForWorksheetMeta,
} from "./routing.js?v=20260921cad";

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
  const wb = discWsWb || wsWb;
  if (!wb) return sheetNames[0];
  const preset = WORKSHEET_PRESETS[$("discWsPreset").value];
  return defaultDiscAppendSheet(wb, preset);
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
    const meta = classifyWorksheetFilename(file.name);
    let routed = filterRowsForWorksheet(rows, meta);
    // If filename has no route hints, keep all selected discrepancy rows
    if (!meta.reason && !meta.div && !meta.clientGroup) routed = rows;
    if (!routed.length) {
      setStatus(
        out,
        "err",
        `No discrepancy rows match this worksheet’s name rules (${meta.label}).`
      );
      return;
    }
    const { count: n } = appendRowsToSheet(freshWb, sheetName, routed);
    const nameHint = discWsFileName || wsFileName || file.name;
    const base = nameHint.replace(/\.(xlsx|xls|xlsm)$/i, "");
    const wasXlsm = /\.xlsm$/i.test(file.name) || Boolean(freshWb.vbaraw);
    const outName = downloadWorkbookFile(
      freshWb,
      wasXlsm
        ? `${base}_discrepancies_added_${todayStr()}.xlsm`
        : `${base}_discrepancies_added_${todayStr()}.xlsx`,
      { preferXlsm: wasXlsm }
    );
    setStatus(
      out,
      "ok",
      `Added ${n} discrepancy row(s) to “${sheetName}” for <em>${escapeHtml(
        file.name
      )}</em> (${meta.label}). Yellow = Today, blue = Yesterday. Saved as <strong>${outName}</strong>${
        wasXlsm
          ? ". Confirm macros before replacing your working XLSM."
          : "."
      }`
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

// Worksheet check (multi-file + name-based routing)
let extractRows = null;
/** @type {{ file: File, wb: any, meta: any, missing: any[] }[]} */
let wsTargets = [];
let wsWb = null; // first / selected — used by discrepancy panel fallback
let wsFileName = "";
let pendingAppendRows = [];

function resolvePreset(meta) {
  const override = $("wsPreset").value;
  if (override && override !== "auto") return override;
  return presetForWorksheetMeta(meta);
}

function fillTargetFileSelect() {
  const sel = $("wsTargetFile");
  sel.innerHTML = "";
  wsTargets.forEach((t, i) => {
    const opt = document.createElement("option");
    opt.value = String(i);
    opt.textContent = `${t.file.name} (${t.missing.length} to add · ${t.meta.label})`;
    sel.appendChild(opt);
  });
}

function fillAppendSheetSelectForTarget(target) {
  const sel = $("wsAppendSheet");
  sel.innerHTML = "";
  if (!target) return;
  for (const name of target.wb.SheetNames) {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    sel.appendChild(opt);
  }
  const preset = WORKSHEET_PRESETS[resolvePreset(target.meta)];
  sel.value = defaultAppendSheet(target.wb, preset);
  wsWb = target.wb;
  wsFileName = target.file.name;
  pendingAppendRows = target.missing;
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
  const files = [...($("fileWs").files || [])];
  if (!files.length) return;
  $("metaWs").textContent = `Loading ${files.length} file(s)…`;
  $("wsRouteSummary").textContent = "";
  wsTargets = [];
  try {
    const lines = [];
    for (const f of files) {
      const wb = await readWorkbook(f);
      const meta = classifyWorksheetFilename(f.name);
      wsTargets.push({ file: f, wb, meta, missing: [] });
      lines.push(`• ${f.name} → ${meta.label}`);
    }
    wsWb = wsTargets[0].wb;
    wsFileName = wsTargets[0].file.name;
    $("metaWs").textContent = `${files.length} worksheet(s) loaded`;
    $("wsRouteSummary").innerHTML = `<strong>Name routing:</strong><br>${lines.join("<br>")}`;
    $("wsAppendPanel").style.display = "none";
  } catch (e) {
    $("metaWs").textContent = e.message;
  }
});

$("wsTargetFile").addEventListener("change", () => {
  const t = wsTargets[Number($("wsTargetFile").value)];
  fillAppendSheetSelectForTarget(t);
});

$("btnWs").addEventListener("click", () => {
  const out = $("wsOut");
  const extraDiv = $("wsDiv").value || null;
  if (!extractRows || !wsTargets.length) {
    setStatus(out, "err", "Upload extraction and at least one worksheet.");
    return;
  }
  try {
    let html = "";
    let anyMissing = false;
    const allMissingCombined = [];

    for (const t of wsTargets) {
      const presetKey = resolvePreset(t.meta);
      let scoped = filterRowsForWorksheet(extractRows, t.meta);
      if (extraDiv) scoped = scoped.filter((r) => r.Div === extraDiv);

      const { missing, sheets, meta } = findNotOnWorksheet(
        scoped,
        t.wb,
        presetKey,
        null
      );
      t.missing = missing;
      if (missing.length) {
        anyMissing = true;
        allMissingCombined.push(...missing.map((r) => ({ ...r, _Worksheet: t.file.name })));
      }

      html += `<h3 style="margin:1rem 0 0.35rem;font-size:1rem">${escapeHtml(t.file.name)}</h3>`;
      html += `<p class="hint">${escapeHtml(t.meta.label)} · type ${presetKey} · compare sheets: ${meta.sheetsRead.join(", ")}</p>`;
      html += metricsHtml([
        ["Rows in scope", meta.extractionAfterFilter],
        ["Worksheet keys", meta.worksheetKeys],
        ["Not on worksheet", meta.notOnWorksheet],
        ...ALL_DIVS.map((d) => [d, (sheets[d] || []).length]),
      ]);
      html += `<div class="actions">${dlBtn(
        `Download missing · ${t.file.name}`,
        sheets,
        `not_on_worksheet_${t.file.name.replace(/\.(xlsx|xls|xlsm)$/i, "")}_${todayStr()}.xlsx`
      )}</div>`;
      html += previewTable(missing);
    }

    if (allMissingCombined.length) {
      html =
        `<div class="actions" style="margin-bottom:0.75rem">${dlBtn(
          "Download all missing (all files)",
          { All_Missing: allMissingCombined },
          `not_on_worksheet_ALL_${todayStr()}.xlsx`
        )}</div>` + html;
    }

    out.innerHTML = html;
    pendingAppendRows = wsTargets.flatMap((t) => t.missing);

    if (anyMissing) {
      fillTargetFileSelect();
      fillAppendSheetSelectForTarget(wsTargets[0]);
      $("wsAppendPanel").style.display = "block";
      setStatus(
        $("wsAppendOut"),
        "info",
        `Ready to append. Each worksheet only gets rows matching its name (R15/R07, division, CPD client).`
      );
    } else {
      $("wsAppendPanel").style.display = "none";
    }
  } catch (e) {
    setStatus(out, "err", e.message);
    $("wsAppendPanel").style.display = "none";
  }
});

async function appendOneTarget(target, sheetNameOverride = null) {
  const freshWb = await readWorkbook(target.file);
  const preset = WORKSHEET_PRESETS[resolvePreset(target.meta)];
  const sheetName =
    sheetNameOverride || defaultAppendSheet(freshWb, preset);
  const { count } = appendRowsToSheet(freshWb, sheetName, target.missing);
  const base = target.file.name.replace(/\.(xlsx|xls|xlsm)$/i, "");
  const wasXlsm = /\.xlsm$/i.test(target.file.name) || Boolean(freshWb.vbaraw);
  const outName = downloadWorkbookFile(
    freshWb,
    wasXlsm
      ? `${base}_updated_${todayStr()}.xlsm`
      : `${base}_updated_${todayStr()}.xlsx`,
    { preferXlsm: wasXlsm }
  );
  return { count, sheetName, outName, wasXlsm, fileName: target.file.name };
}

$("btnAppendWs").addEventListener("click", async () => {
  const out = $("wsAppendOut");
  const target = wsTargets[Number($("wsTargetFile").value)];
  if (!target?.missing?.length) {
    setStatus(out, "err", "Find missing rows first / select a file with rows to add.");
    return;
  }
  try {
    const sheetName = $("wsAppendSheet").value;
    const r = await appendOneTarget(target, sheetName);
    setStatus(
      out,
      "ok",
      `Added ${r.count} row(s) to “${r.sheetName}” in <em>${escapeHtml(
        r.fileName
      )}</em>. Saved <strong>${r.outName}</strong>${
        r.wasXlsm ? ". Confirm macros before replacing your working XLSM." : "."
      }`
    );
  } catch (e) {
    setStatus(out, "err", e.message);
  }
});

$("btnAppendWsAll").addEventListener("click", async () => {
  const out = $("wsAppendOut");
  const withRows = wsTargets.filter((t) => t.missing?.length);
  if (!withRows.length) {
    setStatus(out, "err", "No matching missing rows to append.");
    return;
  }
  try {
    const notes = [];
    for (const t of withRows) {
      const r = await appendOneTarget(t);
      notes.push(
        `${r.fileName}: +${r.count} → ${r.outName}`
      );
      // small gap so browsers don't block multiple downloads
      await new Promise((res) => setTimeout(res, 400));
    }
    setStatus(
      out,
      "ok",
      `Updated ${withRows.length} worksheet(s):<br>${notes.map(escapeHtml).join("<br>")}`
    );
  } catch (e) {
    setStatus(out, "err", e.message);
  }
});
