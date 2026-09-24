
import io
import re
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st
from rapidfuzz import fuzz, process
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter


st.set_page_config(
    page_title="Excel Intelligence & Matcher",
    page_icon="🔎",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
.block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
.app-title {font-size: 2.1rem; font-weight: 750; margin-bottom: .1rem;}
.app-subtitle {color:#667085; margin-bottom:1rem;}
.metric-card {
    border:1px solid #e5e7eb; border-radius:12px; padding:14px;
    background:#fff; min-height:105px;
}
.small-note {color:#667085; font-size:.88rem;}
[data-testid="stMetricValue"] {font-size:1.55rem;}
</style>
""", unsafe_allow_html=True)


# ----------------------------- State -----------------------------

DEFAULT_STATE = {
    "workbooks": {},
    "selected_result": None,
    "analysis_run": False,
    "analysis_mode": "Common Entries",
}
for k, v in DEFAULT_STATE.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ----------------------------- Helpers -----------------------------

def clean_text(v):
    if pd.isna(v):
        return ""
    return str(v).strip()


def normalize_value(v):
    s = clean_text(v).lower()
    s = s.replace("_", " ").replace("-", " ")
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[^\w\s]", "", s)
    return re.sub(r"\s+", " ", s).strip()


def safe_sheet_name(name):
    return str(name)[:31] if name else "Sheet1"


def load_uploaded_workbooks(files):
    workbooks = {}
    for f in files:
        try:
            raw = f.getvalue()
            xls = pd.ExcelFile(io.BytesIO(raw))
            sheets = {}
            for sheet in xls.sheet_names:
                df = pd.read_excel(io.BytesIO(raw), sheet_name=sheet)
                df.columns = [
                    str(c).strip() if not str(c).startswith("Unnamed:")
                    else f"Column_{i+1}"
                    for i, c in enumerate(df.columns)
                ]
                df = df.dropna(axis=1, how="all")
                sheets[sheet] = df
            workbooks[f.name] = {
                "raw": raw,
                "sheets": sheets,
            }
        except Exception as e:
            st.error(f"Could not read **{f.name}**: {e}")
    return workbooks


def source_rows(selections):
    rows = []
    for item in selections:
        df = item["df"]
        col = item["column"]
        for idx, value in df[col].items():
            rows.append({
                "File": item["file"],
                "Sheet": item["sheet"],
                "Row": idx + 2,
                "Column": col,
                "Value": value,
                "Normalized": normalize_value(value),
            })
    return pd.DataFrame(rows)


def selected_column_stats(selections):
    out = []
    for item in selections:
        s = item["df"][item["column"]]
        nonblank = s.notna() & (s.astype(str).str.strip() != "")
        out.append({
            "File": item["file"],
            "Sheet": item["sheet"],
            "Column": item["column"],
            "Rows": len(s),
            "Non-Blank": int(nonblank.sum()),
            "Blank": int((~nonblank).sum()),
            "Distinct": int(s[nonblank].nunique()),
            "Duplicate Rows": int(nonblank.sum() - s[nonblank].nunique()),
        })
    return pd.DataFrame(out)


def workbook_stats(workbooks):
    rows = []
    for file, info in workbooks.items():
        for sheet, df in info["sheets"].items():
            rows.append({
                "File": file,
                "Sheet": sheet,
                "Rows": len(df),
                "Columns": len(df.columns),
                "Cells": int(df.shape[0] * df.shape[1]),
                "Blank Cells": int(df.isna().sum().sum()),
                "Distinct Rows": int(len(df.drop_duplicates())),
            })
    return pd.DataFrame(rows)


def build_value_index(selections):
    idx = {}
    for item in selections:
        s = item["df"][item["column"]]
        for row_no, value in s.items():
            norm = normalize_value(value)
            if not norm:
                continue
            idx.setdefault(norm, {
                "display": clean_text(value),
                "sources": set(),
                "occurrences": 0,
                "locations": [],
            })
            src = f"{item['file']} | {item['sheet']} | {item['column']}"
            idx[norm]["sources"].add(src)
            idx[norm]["occurrences"] += 1
            idx[norm]["locations"].append(
                f"{item['file']} | {item['sheet']} | Row {row_no + 2}"
            )
    return idx


def common_analysis(selections):
    idx = build_value_index(selections)
    n_sources = len(selections)
    rows = []
    for norm, d in idx.items():
        if len(d["sources"]) == n_sources:
            rows.append({
                "Value": d["display"],
                "Sources": len(d["sources"]),
                "Occurrences": d["occurrences"],
                "Locations": "; ".join(d["locations"]),
            })
    return pd.DataFrame(rows).sort_values(
        ["Sources", "Value"], ascending=[False, True]
    ) if rows else pd.DataFrame(
        columns=["Value", "Sources", "Occurrences", "Locations"]
    )


def distinct_analysis(selections):
    idx = build_value_index(selections)
    rows = []
    for norm, d in idx.items():
        rows.append({
            "Value": d["display"],
            "Sources": len(d["sources"]),
            "Occurrences": d["occurrences"],
            "Locations": "; ".join(d["locations"]),
        })
    return pd.DataFrame(rows).sort_values(
        ["Sources", "Value"], ascending=[False, True]
    ) if rows else pd.DataFrame(
        columns=["Value", "Sources", "Occurrences", "Locations"]
    )


def similarity_score(a, b, method):
    if method == "Token Sort":
        return fuzz.token_sort_ratio(a, b)
    if method == "Token Set":
        return fuzz.token_set_ratio(a, b)
    return fuzz.WRatio(a, b)


def fuzzy_analysis(selections, threshold, method):
    idx = build_value_index(selections)
    values = list(idx.keys())
    if len(values) < 2:
        return pd.DataFrame(
            columns=[
                "Group", "Representative", "Matched Value",
                "Similarity %", "Sources", "Occurrences", "Locations"
            ]
        )

    groups = []
    assigned = set()

    for value in values:
        if value in assigned:
            continue

        candidates = process.extract(
            value, values,
            scorer=lambda a, b, **kwargs: similarity_score(a, b, method),
            score_cutoff=threshold,
            limit=None,
        )

        group = [m[0] for m in candidates]
        if len(group) > 1:
            groups.append(group)
            assigned.update(group)

    rows = []
    for gid, group in enumerate(groups, 1):
        representative = max(
            group,
            key=lambda x: (len(idx[x]["display"]), len(idx[x]["sources"]))
        )
        for value in group:
            score = 100 if value == representative else similarity_score(
                representative, value, method
            )
            rows.append({
                "Group": gid,
                "Representative": idx[representative]["display"],
                "Matched Value": idx[value]["display"],
                "Similarity %": round(score, 1),
                "Sources": len(idx[value]["sources"]),
                "Occurrences": idx[value]["occurrences"],
                "Locations": "; ".join(idx[value]["locations"]),
            })

    return pd.DataFrame(rows)


def missing_by_source(selections):
    idx = build_value_index(selections)
    rows = []
    for norm, d in idx.items():
        if len(d["sources"]) < len(selections):
            rows.append({
                "Value": d["display"],
                "Present In": len(d["sources"]),
                "Missing From": len(selections) - len(d["sources"]),
                "Locations": "; ".join(d["locations"]),
            })
    return pd.DataFrame(rows).sort_values(
        ["Present In", "Value"], ascending=[False, True]
    ) if rows else pd.DataFrame(
        columns=["Value", "Present In", "Missing From", "Locations"]
    )


def highlight_dataframe(df, mode, threshold):
    def style(row):
        styles = pd.Series("", index=row.index)
        if mode == "Similar Entries" and "Similarity %" in row.index:
            score = float(row["Similarity %"])
            if score >= 95:
                styles[:] = "background-color:#c6efce"
            elif score >= threshold:
                styles[:] = "background-color:#fff2cc"
        else:
            styles[:] = "background-color:#e8f1fb"
        return styles
    return df.style.apply(style, axis=1)


def export_excel(sheets):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=safe_sheet_name(name), index=False)

    output.seek(0)
    wb = load_workbook(output)

    header_fill = PatternFill(
        start_color="1F4E78", end_color="1F4E78", fill_type="solid"
    )
    header_font = Font(color="FFFFFF", bold=True)

    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

        for col in ws.columns:
            letter = get_column_letter(col[0].column)
            max_len = min(
                max(len(str(c.value)) if c.value is not None else 0 for c in col) + 2,
                55,
            )
            ws.column_dimensions[letter].width = max_len

    final = io.BytesIO()
    wb.save(final)
    final.seek(0)
    return final.getvalue()


# ----------------------------- Header -----------------------------

st.markdown('<div class="app-title">🔎 Excel Intelligence & Matcher</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="app-subtitle">'
    "Load multiple Excel workbooks, organize their data, compare selected columns, "
    "detect exact/common/distinct values, and discover spelling variations."
    "</div>",
    unsafe_allow_html=True,
)


# ----------------------------- Sidebar -----------------------------

with st.sidebar:
    st.header("⚙️ Analysis Controls")

    mode = st.radio(
        "Analysis",
        ["Common Entries", "Distinct Entries", "Similar Entries", "Missing by Source"],
        index=["Common Entries", "Distinct Entries", "Similar Entries", "Missing by Source"].index(
            st.session_state.analysis_mode
        ),
    )
    st.session_state.analysis_mode = mode

    if mode == "Similar Entries":
        threshold = st.slider(
            "Similarity threshold",
            50, 100, 85, 1,
            help=(
                "Controls how close two normalized values must be. "
                "100 requires an exact match; lower values allow more variation."
            ),
        )

        method = st.selectbox(
            "Similarity method",
            ["Token Sort", "Token Set", "WRatio"],
            help="Token Set handles additional words well; Token Sort handles word-order changes.",
        )

        st.caption(
            f"Current threshold: **{threshold}%**"
        )
        if threshold >= 95:
            st.caption("Very strict: only highly similar values.")
        elif threshold >= 90:
            st.caption("Strict: good for minor spelling/format differences.")
        elif threshold >= 80:
            st.caption("Balanced: catches many spelling variations; review results.")
        else:
            st.caption("Broad: may produce more false matches; manual verification recommended.")

    st.divider()
    st.caption("Tip: use the Analysis tab after selecting the columns you want to compare.")


# ----------------------------- Tabs -----------------------------

tab_load, tab_overview, tab_mapping, tab_analysis, tab_explore, tab_export = st.tabs(
    [
        "📥 1. Load Data",
        "📊 2. Data Overview",
        "🔗 3. Column Mapping",
        "🔎 4. Analysis",
        "🎨 5. Explore",
        "📤 6. Export",
    ]
)


# ============================================================
# TAB 1 — LOAD
# ============================================================

with tab_load:
    st.subheader("Upload and organize your Excel data")

    files = st.file_uploader(
        "Upload one or more Excel workbooks",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        help="Multiple workbooks and multiple sheets per workbook are supported.",
    )

    if files:
        if st.button("📥 Load / Refresh Workbooks", type="primary"):
            with st.spinner("Reading workbooks and sheets..."):
                st.session_state.workbooks = load_uploaded_workbooks(files)
                st.session_state.analysis_run = False
                st.session_state.selected_result = None
            st.success(
                f"Loaded {len(st.session_state.workbooks)} workbook(s)."
            )

    if st.session_state.workbooks:
        stats = workbook_stats(st.session_state.workbooks)

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Workbooks", stats["File"].nunique())
        c2.metric("Sheets", len(stats))
        c3.metric("Total Rows", f"{stats['Rows'].sum():,}")
        c4.metric("Total Columns", f"{stats['Columns'].sum():,}")

        st.dataframe(stats, use_container_width=True, height=420)

        st.markdown("### Sheet preview")
        for file, info in st.session_state.workbooks.items():
            with st.expander(f"📄 {file}"):
                for sheet, df in info["sheets"].items():
                    st.markdown(f"**{sheet}** — {len(df):,} rows × {len(df.columns):,} columns")
                    st.dataframe(df.head(8), use_container_width=True)


# ============================================================
# TAB 2 — OVERVIEW
# ============================================================

with tab_overview:
    st.subheader("Data quality and structure")

    if not st.session_state.workbooks:
        st.info("Load Excel files in Tab 1 first.")
    else:
        stats = workbook_stats(st.session_state.workbooks)

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Files", stats["File"].nunique())
        c2.metric("Sheets", len(stats))
        c3.metric("Rows", f"{stats['Rows'].sum():,}")
        c4.metric("Columns", f"{stats['Columns'].sum():,}")
        c5.metric("Blank Cells", f"{stats['Blank Cells'].sum():,}")

        st.markdown("### Sheet-level statistics")
        st.dataframe(stats, use_container_width=True)

        st.markdown("### Column-level statistics")
        col_stats = []
        for file, info in st.session_state.workbooks.items():
            for sheet, df in info["sheets"].items():
                for col in df.columns:
                    s = df[col]
                    nonblank = s.notna() & (s.astype(str).str.strip() != "")
                    col_stats.append({
                        "File": file,
                        "Sheet": sheet,
                        "Column": col,
                        "Rows": len(s),
                        "Distinct": int(s[nonblank].nunique()),
                        "Blank": int((~nonblank).sum()),
                        "Duplicate Values": int(nonblank.sum() - s[nonblank].nunique()),
                        "Data Type": str(s.dtype),
                    })
        st.dataframe(pd.DataFrame(col_stats), use_container_width=True, height=500)


# ============================================================
# TAB 3 — MAPPING
# ============================================================

with tab_mapping:
    st.subheader("Select the columns you want to compare")

    if not st.session_state.workbooks:
        st.info("Load Excel files in Tab 1 first.")
    else:
        st.caption(
            "Choose one sheet and one column from each workbook. "
            "The application compares the selected columns across all sources."
        )

        selections = []

        for i, (file, info) in enumerate(st.session_state.workbooks.items()):
            with st.expander(f"📄 {file}", expanded=True):
                sheets = list(info["sheets"].keys())
                sheet = st.selectbox(
                    "Sheet",
                    sheets,
                    key=f"map_sheet_{i}",
                )
                df = info["sheets"][sheet]
                columns = list(df.columns)

                column = st.selectbox(
                    "Column",
                    columns,
                    key=f"map_col_{i}",
                )

                a, b, c = st.columns(3)
                a.metric("Rows", f"{len(df):,}")
                b.metric("Distinct", f"{df[column].nunique(dropna=True):,}")
                c.metric("Blank", f"{df[column].isna().sum():,}")

                st.dataframe(
                    df[[column]].head(10),
                    use_container_width=True,
                    height=220,
                )

                selections.append({
                    "file": file,
                    "sheet": sheet,
                    "column": column,
                    "df": df,
                })

        st.session_state.selections = selections

        st.success(
            f"{len(selections)} source column(s) configured for comparison."
        )

        selected_stats = selected_column_stats(selections)
        st.dataframe(selected_stats, use_container_width=True)


# ============================================================
# TAB 4 — ANALYSIS
# ============================================================

with tab_analysis:
    st.subheader("Run comparison and matching analysis")

    selections = st.session_state.get("selections", [])

    if len(selections) < 2:
        st.info("Configure at least two source columns in Tab 3.")
    else:
        st.info(
            f"Mode: **{mode}**. "
            "Results are generated only from the columns selected in Tab 3."
        )

        if st.button(
            "🚀 Run Analysis",
            type="primary",
            use_container_width=True,
        ):
            with st.spinner("Analyzing values across selected sources..."):
                if mode == "Common Entries":
                    result = common_analysis(selections)
                elif mode == "Distinct Entries":
                    result = distinct_analysis(selections)
                elif mode == "Missing by Source":
                    result = missing_by_source(selections)
                else:
                    result = fuzzy_analysis(
                        selections,
                        threshold,
                        method,
                    )

                st.session_state.selected_result = result
                st.session_state.analysis_run = True

        result = st.session_state.selected_result

        if result is not None:
            if mode == "Similar Entries" and not result.empty:
                groups = result["Group"].nunique()
                matches = len(result)
                avg_score = result["Similarity %"].mean()

                c1, c2, c3 = st.columns(3)
                c1.metric("Similarity Groups", groups)
                c2.metric("Matched Values", matches)
                c3.metric("Average Similarity", f"{avg_score:.1f}%")
            else:
                c1, c2, c3 = st.columns(3)
                c1.metric("Result Rows", f"{len(result):,}")

                if "Sources" in result.columns:
                    c2.metric(
                        "Max Sources",
                        f"{result['Sources'].max() if len(result) else 0}"
                    )
                if "Occurrences" in result.columns:
                    c3.metric(
                        "Occurrences",
                        f"{result['Occurrences'].sum() if len(result) else 0:,}"
                    )

            search = st.text_input(
                "🔍 Search within results",
                placeholder="Search value, file, source, group...",
            )

            display = result.copy()
            if search:
                mask = display.astype(str).apply(
                    lambda c: c.str.contains(
                        search, case=False, na=False, regex=False
                    )
                ).any(axis=1)
                display = display[mask]

            if display.empty:
                st.warning("No results match the current filter.")
            else:
                st.dataframe(
                    highlight_dataframe(display, mode, threshold if mode == "Similar Entries" else 100),
                    use_container_width=True,
                    height=520,
                )


# ============================================================
# TAB 5 — EXPLORE
# ============================================================

with tab_explore:
    st.subheader("Explore selected values and source locations")

    selections = st.session_state.get("selections", [])
    result = st.session_state.selected_result

    if not selections:
        st.info("Configure columns in Tab 3 first.")
    else:
        selected = source_rows(selections)

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Selected Rows", f"{len(selected):,}")
        c2.metric("Distinct Values", f"{selected['Normalized'].nunique():,}")
        c3.metric("Source Files", selected["File"].nunique())
        c4.metric("Source Sheets", selected["Sheet"].nunique())

        st.markdown("### Value explorer")

        query = st.text_input(
            "Find a value",
            placeholder="e.g. Goa, Forensic Science, Maharashtra...",
        )

        explorer = selected.copy()

        if query:
            explorer = explorer[
                explorer["Value"].astype(str).str.contains(
                    query, case=False, na=False, regex=False
                )
            ]

        st.dataframe(
            explorer,
            use_container_width=True,
            height=500,
        )

        if result is not None and mode == "Similar Entries" and not result.empty:
            st.markdown("### Similarity groups")

            group_ids = sorted(result["Group"].unique())
            group_id = st.selectbox("Match group", group_ids)

            group = result[result["Group"] == group_id]
            st.dataframe(
                highlight_dataframe(group, mode, threshold),
                use_container_width=True,
            )

            st.caption(
                "Yellow highlights indicate accepted similarity; green indicates very high similarity."
            )


# ============================================================
# TAB 6 — EXPORT
# ============================================================

with tab_export:
    st.subheader("Export analysis and source data")

    selections = st.session_state.get("selections", [])
    result = st.session_state.selected_result

    if not selections:
        st.info("Configure columns in Tab 3 first.")
    else:
        selected = source_rows(selections)
        selected_stats = selected_column_stats(selections)

        export_sheets = {
            "Analysis": result if result is not None else pd.DataFrame(),
            "Selected_Data": selected,
            "Column_Stats": selected_stats,
            "Workbook_Stats": workbook_stats(st.session_state.workbooks),
        }

        if result is not None and mode == "Similar Entries":
            export_sheets["Similarity_Groups"] = result

        excel_bytes = export_excel(export_sheets)

        c1, c2 = st.columns(2)

        with c1:
            st.download_button(
                "📊 Download Complete Excel Report",
                data=excel_bytes,
                file_name="Excel_Intelligence_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with c2:
            csv = (
                result.to_csv(index=False).encode("utf-8")
                if result is not None
                else b""
            )
            st.download_button(
                "⬇️ Download Analysis CSV",
                data=csv,
                file_name="analysis_results.csv",
                mime="text/csv",
                use_container_width=True,
                disabled=result is None,
            )

        st.markdown("### Report contents")
        st.write(
            "The Excel report contains the analysis results, all selected source "
            "rows, column-level statistics, workbook/sheet statistics, and "
            "similarity groups when fuzzy matching is used."
        )

        st.dataframe(
            pd.DataFrame({
                "Report Sheet": list(export_sheets.keys()),
                "Rows": [len(v) for v in export_sheets.values()],
                "Columns": [len(v.columns) for v in export_sheets.values()],
            }),
            use_container_width=True,
        )


st.divider()
st.caption(
    "Excel Intelligence & Matcher • Multi-workbook • Multi-sheet • "
    "Exact + normalized + fuzzy comparison"
)
