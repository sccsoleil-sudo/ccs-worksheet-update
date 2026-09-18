"""
CCS Worksheet Update Tool
Daily workflow: enrich S4 export → find new lines → check discrepancies.
"""

from __future__ import annotations

from datetime import date

import streamlit as st

from ccs_tool.discrepancies import find_discrepancies_cpd, find_discrepancies_non_cpd
from ccs_tool.enrich import enrich_div_brand_axe, enrichment_summary
from ccs_tool.io_utils import read_excel, sheets_to_excel_bytes
from ccs_tool.new_lines import find_new_cpd, find_new_non_cpd
from ccs_tool.worksheet_check import WORKSHEET_PRESETS, find_not_on_worksheet, list_excel_files

st.set_page_config(page_title="CCS Worksheet Update", page_icon="📋", layout="wide")

st.title("CCS Worksheet Update")
st.caption("Enrich S4 files → find new lines → check Assignment / Item Text changes")

TODAY_STR = date.today().isoformat()  # e.g. 2026-09-18
SAVE_FOR_TOMORROW_NAME = f"CCS_for_tomorrow_{TODAY_STR}.xlsx"

# ---------------------------------------------------------------------------
# Sidebar: upload files once, reuse everywhere
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("1. Load files")
    st.info(
        "Any Excel file name is fine — it does **not** need to be "
        "`CCS NEW` or `CCS LAST`."
    )
    today_file = st.file_uploader(
        "Today's S4 download (any name)",
        type=["xlsx", "xls"],
        key="today",
        help="Fresh export from S4 today. Filename can be anything.",
    )
    yesterday_file = st.file_uploader(
        "Previous day file (any name)",
        type=["xlsx", "xls"],
        key="yesterday",
        help=(
            "Use yesterday's S4 download, OR the file you saved from this tool "
            f"as '{SAVE_FOR_TOMORROW_NAME}' (or any prior CCS_for_tomorrow_*.xlsx)."
        ),
    )
    if today_file:
        st.caption(f"Today: `{today_file.name}`")
    if yesterday_file:
        st.caption(f"Previous: `{yesterday_file.name}`")

    st.markdown("---")
    st.markdown("**For tomorrow**")
    st.markdown(
        "After you run the tool, download **Save for tomorrow**. "
        "Keep that file. Tomorrow, upload it in the **Previous day** slot, "
        "and upload the new S4 export as **Today**."
    )
    st.markdown("---")
    st.markdown(
        "**Daily steps**\n"
        "1. Enrich both files (Div / Brand / Axe)\n"
        "2. Find **new lines** vs previous day\n"
        "3. Check **discrepancies** on existing Dispute IDs\n"
        "4. Check extraction **vs worksheet** (folder path OK)"
    )


def _need_both():
    if not today_file or not yesterday_file:
        st.warning(
            "Upload **Today's S4 download** and the **Previous day** file "
            "in the sidebar first (any file names are OK)."
        )
        return False
    return True


tab_enrich, tab_new, tab_disc, tab_ws, tab_all = st.tabs(
    [
        "Step 1 · Enrich",
        "Step 2 · New lines",
        "Step 3 · Discrepancies",
        "Step 4 · vs Worksheet",
        "Run all",
    ]
)

# ---------------------------------------------------------------------------
# Step 1 — Enrich
# ---------------------------------------------------------------------------
with tab_enrich:
    st.subheader("Populate Div · Brand · Axe")
    st.write(
        "Maps Business Area → Div and Item Text acronym → Brand / Axe. "
        "Works with any Excel file name."
    )
    if st.button("Enrich uploaded files", type="primary", key="btn_enrich"):
        if not today_file and not yesterday_file:
            st.warning("Upload at least one file in the sidebar.")
        else:
            for label, f in [("Today", today_file), ("Yesterday", yesterday_file)]:
                if not f:
                    continue
                try:
                    raw = read_excel(f)
                    enriched = enrich_div_brand_axe(raw)
                    by_div, n_unid = enrichment_summary(enriched)
                    st.markdown(f"### {label}: `{f.name}`")
                    c1, c2 = st.columns(2)
                    c1.metric("Rows", len(enriched))
                    c2.metric("Unidentified brand", n_unid)
                    st.dataframe(by_div, hide_index=True, use_container_width=True)
                    data = sheets_to_excel_bytes({"Enriched": enriched}, "No data")
                    out_name = f.name.rsplit(".", 1)[0] + "_enriched.xlsx"
                    st.download_button(
                        f"Download enriched · {label}",
                        data=data,
                        file_name=out_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_enrich_{label}",
                    )
                    if label == "Today":
                        st.download_button(
                            f"Save for tomorrow · {SAVE_FOR_TOMORROW_NAME}",
                            data=data,
                            file_name=SAVE_FOR_TOMORROW_NAME,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_enrich_for_tomorrow",
                            help="Keep this file. Tomorrow upload it as Previous day.",
                        )
                    st.session_state[f"enriched_{label}"] = enriched
                except Exception as e:
                    st.error(f"{label}: {e}")

# ---------------------------------------------------------------------------
# Step 2 — New lines
# ---------------------------------------------------------------------------
with tab_new:
    st.subheader("New lines to add to the worksheet")
    st.write(
        "Compares **Amount + Assignment**. Rows whose Reference contains `CCS` are skipped."
    )
    col_a, col_b = st.columns(2)

    with col_a:
        st.markdown("#### LDB · LPD · PPD")
        if st.button("Find non-CPD new lines", type="primary", key="btn_new_non"):
            if _need_both():
                try:
                    df_new = st.session_state.get("enriched_Today")
                    df_last = st.session_state.get("enriched_Yesterday")
                    if df_new is None:
                        df_new = enrich_div_brand_axe(read_excel(today_file))
                    if df_last is None:
                        df_last = enrich_div_brand_axe(read_excel(yesterday_file))
                    sheets = find_new_non_cpd(df_new, df_last)
                    total = len(sheets.get("All_New", []))
                    st.success(f"Found **{total}** new non-CPD row(s).")
                    if total:
                        st.dataframe(
                            sheets["All_New"].head(50),
                            use_container_width=True,
                            hide_index=True,
                        )
                    st.download_button(
                        "Download CCS_Differences_nonCPD.xlsx",
                        data=sheets_to_excel_bytes(
                            sheets, "No differences found"
                        ),
                        file_name="CCS_Differences_nonCPD.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_new_non",
                    )
                except Exception as e:
                    st.error(str(e))

    with col_b:
        st.markdown("#### CPD (by customer group)")
        if st.button("Find CPD new lines", type="primary", key="btn_new_cpd"):
            if _need_both():
                try:
                    df_new = st.session_state.get("enriched_Today")
                    df_last = st.session_state.get("enriched_Yesterday")
                    if df_new is None:
                        df_new = enrich_div_brand_axe(read_excel(today_file))
                    if df_last is None:
                        df_last = enrich_div_brand_axe(read_excel(yesterday_file))
                    sheets, unassigned = find_new_cpd(df_new, df_last)
                    total = len(sheets.get("All_Results", []))
                    st.success(f"Found **{total}** new CPD row(s).")
                    if total:
                        st.dataframe(
                            sheets["All_Results"].head(50),
                            use_container_width=True,
                            hide_index=True,
                        )
                    st.download_button(
                        "Download CCS_Difference_CPD.xlsx",
                        data=sheets_to_excel_bytes(
                            sheets,
                            "No new CPD items found with unique Amt/Asgn",
                        ),
                        file_name="CCS_Difference_CPD.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_new_cpd",
                    )
                    if not unassigned.empty:
                        st.warning(f"{len(unassigned)} row(s) with unassigned Customer.")
                        st.download_button(
                            "Download CPD_Unassigned_Rows.xlsx",
                            data=sheets_to_excel_bytes(
                                {"Unassigned": unassigned}, "None"
                            ),
                            file_name="CPD_Unassigned_Rows.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_unassigned",
                        )
                except Exception as e:
                    st.error(str(e))

# ---------------------------------------------------------------------------
# Step 3 — Discrepancies
# ---------------------------------------------------------------------------
with tab_disc:
    st.subheader("Discrepancies on existing Dispute IDs")
    st.write(
        "Flags rows that already exist (same Dispute ID) but **Assignment**, "
        "**Item Text**, or **Div** changed vs yesterday."
    )
    col_c, col_d = st.columns(2)

    with col_c:
        st.markdown("#### CPD")
        if st.button("Check CPD discrepancies", type="primary", key="btn_disc_cpd"):
            if _need_both():
                try:
                    df_new = st.session_state.get("enriched_Today")
                    df_last = st.session_state.get("enriched_Yesterday")
                    if df_new is None:
                        df_new = enrich_div_brand_axe(read_excel(today_file))
                    if df_last is None:
                        df_last = enrich_div_brand_axe(read_excel(yesterday_file))
                    sheets = find_discrepancies_cpd(df_new, df_last)
                    total = len(sheets.get("All_Differing_Rows", []))
                    st.success(
                        f"Found **{total}** discrepancy row(s)."
                        if total
                        else "No CPD discrepancies."
                    )
                    if total:
                        st.dataframe(
                            sheets["All_Differing_Rows"].head(50),
                            use_container_width=True,
                            hide_index=True,
                        )
                    st.download_button(
                        "Download CCS_Discrepancies_CPD.xlsx",
                        data=sheets_to_excel_bytes(
                            sheets,
                            "No Div, Assignment, or Text changes found",
                        ),
                        file_name="CCS_Discrepancies_CPD.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_disc_cpd",
                    )
                except Exception as e:
                    st.error(str(e))

    with col_d:
        st.markdown("#### LDB · LPD · PPD")
        if st.button("Check non-CPD discrepancies", type="primary", key="btn_disc_non"):
            if _need_both():
                try:
                    df_new = st.session_state.get("enriched_Today")
                    df_last = st.session_state.get("enriched_Yesterday")
                    if df_new is None:
                        df_new = enrich_div_brand_axe(read_excel(today_file))
                    if df_last is None:
                        df_last = enrich_div_brand_axe(read_excel(yesterday_file))
                    sheets = find_discrepancies_non_cpd(df_new, df_last)
                    total = len(sheets.get("All_Discrepancies", []))
                    st.success(
                        f"Found **{total}** discrepancy row(s)."
                        if total
                        else "No non-CPD discrepancies."
                    )
                    if total:
                        st.dataframe(
                            sheets["All_Discrepancies"].head(50),
                            use_container_width=True,
                            hide_index=True,
                        )
                    st.download_button(
                        "Download CCS_Discrepancies_By_Div.xlsx",
                        data=sheets_to_excel_bytes(sheets, "No discrepancies found"),
                        file_name="CCS_Discrepancies_By_Div.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_disc_non",
                    )
                except Exception as e:
                    st.error(str(e))

# ---------------------------------------------------------------------------
# Step 4 — Extraction vs existing worksheet (folder path OK)
# ---------------------------------------------------------------------------
with tab_ws:
    st.subheader("Not yet on the worksheet")
    st.write(
        "Compare **today's extraction** to your **existing worksheet** "
        "(multi-sheet). Finds rows whose `SUBI $` + `Assignment` are not on the worksheet. "
        "Amount column can also be named `Amount (CoCode Crcy)`."
    )

    source_mode = st.radio(
        "How to load files",
        ["Folder path on this laptop", "Upload two files"],
        horizontal=True,
        key="ws_source_mode",
    )

    preset_name = st.selectbox(
        "Worksheet type",
        list(WORKSHEET_PRESETS.keys()),
        format_func=lambda k: f"{k} — {WORKSHEET_PRESETS[k]['label']}",
        key="ws_preset",
    )
    preset = WORKSHEET_PRESETS[preset_name]
    st.caption(
        f"Sheets checked: {', '.join(preset['worksheet_sheets'])}. "
        f"Skip Reference containing CCS: {'yes' if preset['filter_ccs_in_reference'] else 'no'}."
    )

    extraction_df = None
    worksheet_path = None
    extraction_label = ""
    worksheet_label = ""

    if source_mode == "Folder path on this laptop":
        folder = st.text_input(
            "Folder path (Finder → right-click folder → hold Option → Copy as Pathname)",
            placeholder="/Users/you/Documents/CCS",
            key="ws_folder",
        )
        if folder:
            try:
                excel_files = list_excel_files(folder)
                names = [p.name for p in excel_files]
                if not names:
                    st.warning("No Excel files (.xlsx / .xls) found in that folder.")
                else:
                    st.success(f"Found {len(names)} Excel file(s) in folder.")
                    c1, c2 = st.columns(2)
                    with c1:
                        ext_name = st.selectbox(
                            f"Today's extraction ({preset['hint_extraction']})",
                            names,
                            key="ws_ext_pick",
                        )
                    with c2:
                        ws_name = st.selectbox(
                            f"Existing worksheet ({preset['hint_worksheet']})",
                            names,
                            index=min(1, len(names) - 1),
                            key="ws_ws_pick",
                        )
                    from pathlib import Path as _Path

                    folder_path = _Path(folder).expanduser()
                    extraction_df = read_excel(folder_path / ext_name)
                    worksheet_path = folder_path / ws_name
                    extraction_label = ext_name
                    worksheet_label = ws_name
            except Exception as e:
                st.error(str(e))
    else:
        c1, c2 = st.columns(2)
        with c1:
            up_ext = st.file_uploader(
                f"Today's extraction ({preset['hint_extraction']})",
                type=["xlsx", "xls"],
                key="ws_up_ext",
            )
        with c2:
            up_ws = st.file_uploader(
                f"Existing worksheet ({preset['hint_worksheet']})",
                type=["xlsx", "xls"],
                key="ws_up_ws",
            )
        if up_ext and up_ws:
            import tempfile
            from pathlib import Path as _Path

            extraction_df = read_excel(up_ext)
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            tmp.write(up_ws.getvalue())
            tmp.close()
            worksheet_path = _Path(tmp.name)
            extraction_label = up_ext.name
            worksheet_label = up_ws.name

    custom_sheets = st.text_input(
        "Sheet names (comma-separated) — leave blank to use preset",
        value="",
        key="ws_custom_sheets",
    )

    if st.button("Find rows not on worksheet", type="primary", key="btn_ws"):
        if extraction_df is None or worksheet_path is None:
            st.warning("Select/upload both the extraction and the worksheet first.")
        else:
            try:
                sheets = (
                    [s.strip() for s in custom_sheets.split(",") if s.strip()]
                    if custom_sheets.strip()
                    else preset["worksheet_sheets"]
                )
                missing, meta = find_not_on_worksheet(
                    extraction_df,
                    worksheet_path,
                    sheets,
                    filter_ccs_in_reference=preset["filter_ccs_in_reference"],
                )
                st.write(
                    f"**{extraction_label}** vs **{worksheet_label}** · "
                    f"sheets used: {', '.join(meta['sheets_read'])}"
                )
                m1, m2, m3 = st.columns(3)
                m1.metric("Extraction rows (after filter)", meta["extraction_rows_after_filter"])
                m2.metric("Worksheet unique keys", meta["worksheet_unique_keys"])
                m3.metric("Not on worksheet", meta["not_on_worksheet"])

                if missing.empty:
                    st.success("All extraction rows are already on the worksheet.")
                else:
                    st.dataframe(missing.head(100), use_container_width=True, hide_index=True)
                    st.download_button(
                        f"Download not_on_worksheet_{preset_name}.xlsx",
                        data=sheets_to_excel_bytes(
                            {"Not_on_worksheet": missing}, "None"
                        ),
                        file_name=f"not_on_worksheet_{preset_name}_{TODAY_STR}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_ws_missing",
                    )
            except Exception as e:
                st.error(str(e))

# ---------------------------------------------------------------------------
# Run all
# ---------------------------------------------------------------------------
with tab_all:
    st.subheader("One-click daily run")
    st.write(
        "Enriches both files (any names), then produces all result workbooks. "
        "Also creates a **Save for tomorrow** file to reuse as Previous day next time."
    )
    if st.button("Run full daily check", type="primary", key="btn_all"):
        if not _need_both():
            st.stop()
        try:
            df_new = enrich_div_brand_axe(read_excel(today_file))
            df_last = enrich_div_brand_axe(read_excel(yesterday_file))
            st.session_state["enriched_Today"] = df_new
            st.session_state["enriched_Yesterday"] = df_last

            non_new = find_new_non_cpd(df_new, df_last)
            cpd_new, unassigned = find_new_cpd(df_new, df_last)
            disc_cpd = find_discrepancies_cpd(df_new, df_last)
            disc_non = find_discrepancies_non_cpd(df_new, df_last)

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("New non-CPD", len(non_new.get("All_New", [])))
            m2.metric("New CPD", len(cpd_new.get("All_Results", [])))
            m3.metric("CPD discrepancies", len(disc_cpd.get("All_Differing_Rows", [])))
            m4.metric("Non-CPD discrepancies", len(disc_non.get("All_Discrepancies", [])))

            st.markdown("---")
            st.markdown("### Keep this for tomorrow")
            st.download_button(
                f"Save for tomorrow · {SAVE_FOR_TOMORROW_NAME}",
                data=sheets_to_excel_bytes({"Enriched": df_new}, "No data"),
                file_name=SAVE_FOR_TOMORROW_NAME,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_all_for_tomorrow",
                type="primary",
                help="Tomorrow: upload this as Previous day + new S4 export as Today.",
            )
            st.caption(
                "Browser cannot keep files after you close the page. "
                "Download this file to your computer (e.g. Downloads folder)."
            )

            st.markdown("### Result downloads")
            st.download_button(
                "Today enriched",
                data=sheets_to_excel_bytes({"Enriched": df_new}, "No data"),
                file_name=f"CCS_today_enriched_{TODAY_STR}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_all_new_enr",
            )
            st.download_button(
                "Previous day enriched",
                data=sheets_to_excel_bytes({"Enriched": df_last}, "No data"),
                file_name=f"CCS_previous_enriched_{TODAY_STR}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_all_last_enr",
            )
            st.download_button(
                "New lines · non-CPD",
                data=sheets_to_excel_bytes(non_new, "No differences found"),
                file_name="CCS_Differences_nonCPD.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_all_non_new",
            )
            st.download_button(
                "New lines · CPD",
                data=sheets_to_excel_bytes(
                    cpd_new, "No new CPD items found with unique Amt/Asgn"
                ),
                file_name="CCS_Difference_CPD.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_all_cpd_new",
            )
            st.download_button(
                "Discrepancies · CPD",
                data=sheets_to_excel_bytes(
                    disc_cpd, "No Div, Assignment, or Text changes found"
                ),
                file_name="CCS_Discrepancies_CPD.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_all_disc_cpd",
            )
            st.download_button(
                "Discrepancies · LDB/LPD/PPD",
                data=sheets_to_excel_bytes(disc_non, "No discrepancies found"),
                file_name="CCS_Discrepancies_By_Div.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_all_disc_non",
            )
            if not unassigned.empty:
                st.download_button(
                    "CPD unassigned customers",
                    data=sheets_to_excel_bytes({"Unassigned": unassigned}, "None"),
                    file_name="CPD_Unassigned_Rows.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_all_unassigned",
                )
            st.success("Done. Download **Save for tomorrow** plus any result files you need.")
        except Exception as e:
            st.error(str(e))
