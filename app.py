import streamlit as st
import pandas as pd
import numpy as np
import re
import io
from rapidfuzz import fuzz, process
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Excel Multi-Sheet Data Matcher",
    page_icon="🔎",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ============================================================
# CUSTOM CSS
# ============================================================

st.markdown("""
<style>

.main-title {
    font-size: 34px;
    font-weight: 700;
    margin-bottom: 5px;
}

.subtitle {
    color: #6b7280;
    font-size: 16px;
    margin-bottom: 20px;
}

.section-title {
    font-size: 22px;
    font-weight: 650;
    margin-top: 15px;
    margin-bottom: 10px;
}

.file-card {
    padding: 15px;
    border-radius: 10px;
    border: 1px solid #ddd;
    margin-bottom: 10px;
}

.metric-card {
    padding: 15px;
    border-radius: 10px;
    border: 1px solid #ddd;
    background-color: #fafafa;
}

.stButton > button {
    border-radius: 8px;
    font-weight: 600;
}

</style>
""", unsafe_allow_html=True)


# ============================================================
# SESSION STATE
# ============================================================

if "uploaded_files" not in st.session_state:
    st.session_state.uploaded_files = {}

if "analysis_result" not in st.session_state:
    st.session_state.analysis_result = None


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def normalize_value(value):
    """
    Normalize text for comparison.
    """

    if pd.isna(value):
        return ""

    value = str(value)

    # Convert to lowercase
    value = value.lower().strip()

    # Remove extra whitespace
    value = re.sub(r"\s+", " ", value)

    # Replace common separators
    value = value.replace("_", " ")
    value = value.replace("-", " ")

    # Remove punctuation
    value = re.sub(r"[^\w\s]", "", value)

    # Remove repeated whitespace again
    value = re.sub(r"\s+", " ", value).strip()

    return value


def display_value(value):
    if pd.isna(value):
        return ""

    return str(value).strip()


def get_unique_values(df, column, include_blank=False):

    values = df[column].drop_duplicates()

    if not include_blank:
        values = values[
            values.notna() &
            (values.astype(str).str.strip() != "")
        ]

    return values.tolist()


def build_selection_dataframe(selections):

    records = []

    for item in selections:

        df = item["df"]
        file_name = item["file"]
        sheet_name = item["sheet"]
        column = item["column"]

        for idx, value in df[column].items():

            records.append({
                "File": file_name,
                "Sheet": sheet_name,
                "Row": idx + 2,
                "Column": column,
                "Value": value,
                "Normalized": normalize_value(value)
            })

    return pd.DataFrame(records)


def exact_common_analysis(selections, include_blank=False):

    all_values = {}

    for item in selections:

        values = get_unique_values(
            item["df"],
            item["column"],
            include_blank
        )

        for value in values:

            normalized = normalize_value(value)

            if not normalized and not include_blank:
                continue

            if normalized not in all_values:
                all_values[normalized] = {
                    "display": value,
                    "sources": set(),
                    "count": 0
                }

            all_values[normalized]["sources"].add(
                f"{item['file']} | {item['sheet']} | {item['column']}"
            )

            all_values[normalized]["count"] += 1

    total_sources = len(selections)

    result = []

    for normalized, data in all_values.items():

        source_count = len(data["sources"])

        if source_count == total_sources:

            result.append({
                "Value": data["display"],
                "Normalized Value": normalized,
                "Files/Sources": source_count,
                "Total Occurrences": data["count"],
                "Sources": "; ".join(sorted(data["sources"]))
            })

    return pd.DataFrame(result)


def distinct_analysis(selections, include_blank=False):

    all_values = {}

    for item in selections:

        values = get_unique_values(
            item["df"],
            item["column"],
            include_blank
        )

        for value in values:

            normalized = normalize_value(value)

            if not normalized and not include_blank:
                continue

            if normalized not in all_values:
                all_values[normalized] = {
                    "display": value,
                    "sources": set(),
                    "occurrences": 0
                }

            all_values[normalized]["sources"].add(
                f"{item['file']} | {item['sheet']} | {item['column']}"
            )

            all_values[normalized]["occurrences"] += 1

    result = []

    for normalized, data in all_values.items():

        result.append({
            "Value": data["display"],
            "Normalized Value": normalized,
            "Number of Sources": len(data["sources"]),
            "Total Occurrences": data["occurrences"],
            "Sources": "; ".join(sorted(data["sources"]))
        })

    return pd.DataFrame(result).sort_values(
        ["Number of Sources", "Value"],
        ascending=[False, True]
    )


def fuzzy_analysis(
    selections,
    threshold=85,
    include_blank=False
):

    # --------------------------------------------------------
    # Collect all unique normalized values
    # --------------------------------------------------------

    values = {}

    for item in selections:

        df = item["df"]
        column = item["column"]

        for value in get_unique_values(
            df,
            column,
            include_blank
        ):

            normalized = normalize_value(value)

            if not normalized:
                continue

            if normalized not in values:

                values[normalized] = {
                    "display": value,
                    "sources": set(),
                    "occurrences": 0
                }

            values[normalized]["sources"].add(
                f"{item['file']} | {item['sheet']} | {item['column']}"
            )

            values[normalized]["occurrences"] += 1

    unique_values = list(values.keys())

    # --------------------------------------------------------
    # Build fuzzy groups
    # --------------------------------------------------------

    groups = []
    assigned = set()

    for value in unique_values:

        if value in assigned:
            continue

        matches = process.extract(
            value,
            unique_values,
            scorer=fuzz.token_sort_ratio,
            score_cutoff=threshold,
            limit=None
        )

        group = []

        for match_value, score, _ in matches:

            if match_value not in group:
                group.append(match_value)

        if len(group) > 1:

            for v in group:
                assigned.add(v)

            groups.append(group)

    # --------------------------------------------------------
    # Generate result
    # --------------------------------------------------------

    rows = []

    group_id = 1

    for group in groups:

        canonical = max(
            group,
            key=lambda x: len(values[x]["display"])
        )

        for value in group:

            if value == canonical:
                similarity = 100
            else:
                similarity = fuzz.token_sort_ratio(
                    canonical,
                    value
                )

            rows.append({
                "Group": group_id,
                "Representative": values[canonical]["display"],
                "Matched Value": values[value]["display"],
                "Similarity %": round(similarity, 1),
                "Number of Sources": len(
                    values[value]["sources"]
                ),
                "Occurrences": values[value]["occurrences"],
                "Sources": "; ".join(
                    sorted(values[value]["sources"])
                )
            })

        group_id += 1

    return pd.DataFrame(rows)


def create_detailed_matches(
    selections,
    fuzzy_groups,
    threshold
):

    if fuzzy_groups.empty:
        return pd.DataFrame()

    rows = []

    for _, group_row in fuzzy_groups.iterrows():

        representative = group_row["Representative"]

        matched_value = group_row["Matched Value"]

        for item in selections:

            df = item["df"]
            column = item["column"]

            for idx, value in df[column].items():

                if pd.isna(value):
                    continue

                similarity = fuzz.token_sort_ratio(
                    normalize_value(representative),
                    normalize_value(value)
                )

                if similarity >= threshold:

                    rows.append({
                        "Match Group": group_row["Group"],
                        "Representative": representative,
                        "Matched Value": value,
                        "Similarity %": round(similarity, 1),
                        "File": item["file"],
                        "Sheet": item["sheet"],
                        "Column": column,
                        "Excel Row": idx + 2
                    })

    return pd.DataFrame(rows)


def dataframe_to_excel(
    df,
    highlight_column=None,
    highlight_values=None
):

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="Results"
        )

    output.seek(0)

    wb = load_workbook(output)

    ws = wb["Results"]

    # Header formatting
    header_fill = PatternFill(
        start_color="1F4E78",
        end_color="1F4E78",
        fill_type="solid"
    )

    header_font = Font(
        color="FFFFFF",
        bold=True
    )

    for cell in ws[1]:

        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(
            horizontal="center"
        )

    # Highlight matched values
    if highlight_column and highlight_values:

        headers = [
            cell.value for cell in ws[1]
        ]

        if highlight_column in headers:

            col_index = headers.index(
                highlight_column
            ) + 1

            highlight_fill = PatternFill(
                start_color="FFF2CC",
                end_color="FFF2CC",
                fill_type="solid"
            )

            for row in range(2, ws.max_row + 1):

                cell = ws.cell(
                    row=row,
                    column=col_index
                )

                if normalize_value(cell.value) in highlight_values:

                    cell.fill = highlight_fill

    # Auto width
    for column_cells in ws.columns:

        max_length = 0

        column_letter = get_column_letter(
            column_cells[0].column
        )

        for cell in column_cells:

            try:
                max_length = max(
                    max_length,
                    len(str(cell.value))
                )
            except:
                pass

        ws.column_dimensions[
            column_letter
        ].width = min(max_length + 2, 50)

    final_output = io.BytesIO()

    wb.save(final_output)

    final_output.seek(0)

    return final_output


# ============================================================
# HEADER
# ============================================================

st.markdown(
    '<div class="main-title">🔎 Multi-Excel Data Matcher</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitle">'
    'Compare, find common values, identify distinct entries, '
    'and detect spelling variations across multiple Excel files.'
    '</div>',
    unsafe_allow_html=True
)


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.header("⚙️ Analysis Settings")

    analysis_mode = st.radio(
        "Analysis Mode",
        [
            "Common Entries",
            "Distinct Entries",
            "Similar Entries"
        ]
    )

    st.divider()

    if analysis_mode == "Similar Entries":

        similarity_threshold = st.slider(
            "Similarity Threshold",
            min_value=50,
            max_value=100,
            value=85,
            step=1
        )

        st.caption(
            "Higher values require closer spelling matches."
        )

    include_blank = st.checkbox(
        "Include blank values",
        value=False
    )

    st.divider()

    st.subheader("🔤 Normalization")

    normalize_case = st.checkbox(
        "Ignore upper/lower case",
        value=True,
        disabled=True
    )

    normalize_spaces = st.checkbox(
        "Ignore extra spaces",
        value=True,
        disabled=True
    )

    normalize_punctuation = st.checkbox(
        "Ignore punctuation",
        value=True,
        disabled=True
    )


# ============================================================
# FILE UPLOAD
# ============================================================

st.markdown(
    '<div class="section-title">1️⃣ Upload Excel Files</div>',
    unsafe_allow_html=True
)

uploaded_files = st.file_uploader(
    "Upload one or more Excel files",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)


if not uploaded_files:

    st.info(
        "Upload two or more Excel files to begin comparison."
    )

    st.stop()


# ============================================================
# READ FILES
# ============================================================

workbooks = {}

for uploaded_file in uploaded_files:

    try:

        excel_file = pd.ExcelFile(
            uploaded_file
        )

        sheets = {}

        for sheet in excel_file.sheet_names:

            df = pd.read_excel(
                uploaded_file,
                sheet_name=sheet
            )

            # Remove completely empty columns
            df = df.dropna(
                axis=1,
                how="all"
            )

            sheets[sheet] = df

        workbooks[
            uploaded_file.name
        ] = sheets

    except Exception as e:

        st.error(
            f"Error reading {uploaded_file.name}: {e}"
        )


# ============================================================
# WORKBOOK OVERVIEW
# ============================================================

st.markdown(
    '<div class="section-title">📁 Uploaded Files</div>',
    unsafe_allow_html=True
)

cols = st.columns(
    min(len(workbooks), 4)
)

for i, (file_name, sheets) in enumerate(
    workbooks.items()
):

    with cols[i % len(cols)]:

        total_rows = sum(
            len(df)
            for df in sheets.values()
        )

        st.metric(
            file_name,
            f"{len(sheets)} sheets"
        )

        st.caption(
            f"{total_rows:,} total rows"
        )


# ============================================================
# COLUMN SELECTION
# ============================================================

st.markdown(
    '<div class="section-title">'
    '2️⃣ Select Columns to Compare'
    '</div>',
    unsafe_allow_html=True
)

st.caption(
    "You can select a different sheet and column from each uploaded workbook."
)

selections = []


for file_name, sheets in workbooks.items():

    with st.expander(
        f"📄 {file_name}",
        expanded=True
    ):

        sheet_names = list(sheets.keys())

        selected_sheet = st.selectbox(
            "Select Sheet",
            sheet_names,
            key=f"sheet_{file_name}"
        )

        df = sheets[selected_sheet]

        columns = list(df.columns)

        selected_column = st.selectbox(
            "Select Column",
            columns,
            key=f"column_{file_name}"
        )

        preview_col1, preview_col2 = st.columns(
            [3, 1]
        )

        with preview_col1:

            st.dataframe(
                df[
                    [selected_column]
                ].head(10),
                use_container_width=True,
                height=250
            )

        with preview_col2:

            st.metric(
                "Rows",
                f"{len(df):,}"
            )

            st.metric(
                "Unique",
                f"{df[selected_column].nunique(dropna=True):,}"
            )

        selections.append({
            "file": file_name,
            "sheet": selected_sheet,
            "column": selected_column,
            "df": df
        })


# ============================================================
# VALIDATION
# ============================================================

if len(selections) < 2:

    st.warning(
        "Upload at least two Excel files for comparison."
    )

    st.stop()


# ============================================================
# ANALYZE BUTTON
# ============================================================

st.markdown(
    '<div class="section-title">3️⃣ Run Analysis</div>',
    unsafe_allow_html=True
)

run_analysis = st.button(
    "🚀 Analyze Selected Columns",
    type="primary",
    use_container_width=True
)


if run_analysis:

    with st.spinner(
        "Analyzing selected columns..."
    ):

        if analysis_mode == "Common Entries":

            result = exact_common_analysis(
                selections,
                include_blank
            )

        elif analysis_mode == "Distinct Entries":

            result = distinct_analysis(
                selections,
                include_blank
            )

        else:

            result = fuzzy_analysis(
                selections,
                similarity_threshold,
                include_blank
            )

        st.session_state.analysis_result = result


# ============================================================
# RESULTS
# ============================================================

result = st.session_state.analysis_result


if result is None:

    st.info(
        "Select your columns and click **Analyze Selected Columns**."
    )

    st.stop()


st.markdown(
    '<div class="section-title">4️⃣ Results</div>',
    unsafe_allow_html=True
)


# ============================================================
# METRICS
# ============================================================

if analysis_mode == "Common Entries":

    total_results = len(result)

    total_files = len(selections)

    c1, c2, c3 = st.columns(3)

    c1.metric(
        "Common Entries",
        f"{total_results:,}"
    )

    c2.metric(
        "Files Compared",
        f"{total_files:,}"
    )

    c3.metric(
        "Matching Across",
        f"{total_files} files"
    )


elif analysis_mode == "Distinct Entries":

    c1, c2, c3 = st.columns(3)

    c1.metric(
        "Distinct Values",
        f"{len(result):,}"
    )

    c2.metric(
        "Total Sources",
        f"{len(selections):,}"
    )

    if len(result):

        c3.metric(
            "Appearing in All Sources",
            f"{(result['Number of Sources'] == len(selections)).sum():,}"
        )


else:

    c1, c2, c3 = st.columns(3)

    c1.metric(
        "Similar Groups",
        f"{result['Group'].nunique() if not result.empty else 0:,}"
    )

    c2.metric(
        "Matched Values",
        f"{len(result):,}"
    )

    c3.metric(
        "Threshold",
        f"{similarity_threshold}%"
    )


# ============================================================
# RESULT FILTER
# ============================================================

if not result.empty:

    search_text = st.text_input(
        "🔍 Search results",
        placeholder="Type to filter results..."
    )

    filtered_result = result.copy()

    if search_text:

        mask = filtered_result.astype(
            str
        ).apply(
            lambda col: col.str.contains(
                search_text,
                case=False,
                na=False,
                regex=False
            )
        ).any(axis=1)

        filtered_result = filtered_result[
            mask
        ]

else:

    filtered_result = result


# ============================================================
# HIGHLIGHT FUNCTION
# ============================================================

def highlight_matches(row):

    styles = pd.Series(
        "",
        index=row.index
    )

    if analysis_mode == "Similar Entries":

        if "Similarity %" in row.index:

            similarity = row["Similarity %"]

            if similarity >= 95:
                styles[:] = "background-color: #c6efce"

            elif similarity >= similarity_threshold:
                styles[:] = "background-color: #fff2cc"

    else:

        styles[:] = "background-color: #e8f4ff"

    return styles


if not filtered_result.empty:

    st.dataframe(
        filtered_result.style.apply(
            highlight_matches,
            axis=1
        ),
        use_container_width=True,
        height=500
    )

else:

    st.warning(
        "No matching results found."
    )


# ============================================================
# FUZZY DETAIL VIEW
# ============================================================

if (
    analysis_mode == "Similar Entries"
    and not result.empty
):

    st.markdown(
        '<div class="section-title">'
        '🔬 Detailed Match Locations'
        '</div>',
        unsafe_allow_html=True
    )

    selected_group = st.selectbox(
        "Select Match Group",
        sorted(
            result["Group"].unique()
        )
    )

    group_data = result[
        result["Group"] == selected_group
    ]

    st.dataframe(
        group_data,
        use_container_width=True
    )

    if st.button(
        "🔎 Find These Matches in Original Files"
    ):

        detail = create_detailed_matches(
            selections,
            group_data,
            similarity_threshold
        )

        if not detail.empty:

            st.dataframe(
                detail.style.apply(
                    lambda row: pd.Series(
                        "background-color: #fff2cc",
                        index=row.index
                    ),
                    axis=1
                ),
                use_container_width=True
            )

        else:

            st.info(
                "No detailed matches found."
            )


# ============================================================
# SOURCE-WISE ANALYSIS
# ============================================================

st.markdown(
    '<div class="section-title">'
    '📊 Source-wise Summary'
    '</div>',
    unsafe_allow_html=True
)

source_summary = []

for item in selections:

    df = item["df"]
    column = item["column"]

    values = df[column]

    source_summary.append({
        "File": item["file"],
        "Sheet": item["sheet"],
        "Column": column,
        "Rows": len(df),
        "Non-Blank": values.notna().sum(),
        "Unique Values": values.nunique(
            dropna=True
        )
    })

source_summary_df = pd.DataFrame(
    source_summary
)

st.dataframe(
    source_summary_df,
    use_container_width=True
)


# ============================================================
# EXPORT
# ============================================================

st.markdown(
    '<div class="section-title">'
    '5️⃣ Export Results'
    '</div>',
    unsafe_allow_html=True
)

export_col1, export_col2, export_col3 = st.columns(3)


# ------------------------------------------------------------
# CSV
# ------------------------------------------------------------

with export_col1:

    csv_data = filtered_result.to_csv(
        index=False
    ).encode("utf-8")

    st.download_button(
        "⬇️ Download CSV",
        data=csv_data,
        file_name="excel_match_results.csv",
        mime="text/csv",
        use_container_width=True
    )


# ------------------------------------------------------------
# EXCEL
# ------------------------------------------------------------

with export_col2:

    excel_data = dataframe_to_excel(
        filtered_result
    )

    st.download_button(
        "📊 Download Excel",
        data=excel_data,
        file_name="excel_match_results.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )


# ------------------------------------------------------------
# SOURCE DATA EXPORT
# ------------------------------------------------------------

with export_col3:

    source_data = build_selection_dataframe(
        selections
    )

    source_excel = dataframe_to_excel(
        source_data
    )

    st.download_button(
        "📦 Export All Selected Data",
        data=source_excel,
        file_name="all_selected_source_data.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )


# ============================================================
# COMPLETE RAW DATA EXPORT
# ============================================================

st.markdown(
    '<div class="section-title">'
    '📚 Complete Source Data'
    '</div>',
    unsafe_allow_html=True
)

st.caption(
    "All selected source columns are combined below. "
    "This is useful for auditing the matching process."
)

source_data = build_selection_dataframe(
    selections
)

st.dataframe(
    source_data,
    use_container_width=True,
    height=400
)


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "Multi-Excel Data Matcher • Exact matching + "
    "normalization + fuzzy matching • Streamlit"
)