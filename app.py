"""Column Description Review — Streamlit app."""

import io
import csv

import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import catalog
import ai_gen
import humanize as hz

st.set_page_config(
    page_title="Column Descriptions",
    page_icon="🗂️",
    layout="wide",
)

# ── Custom styling ────────────────────────────────────────────────────────

st.markdown("""
<style>
  .type-badge {
    display: inline-block;
    padding: 2px 10px;
    border-radius: 12px;
    font-size: 12px;
    font-weight: 600;
    margin-left: 10px;
  }
  .badge-dlt    { background: #14532d; color: #4ade80; }
  .badge-mv     { background: #451a03; color: #fbbf24; }
  .badge-delta  { background: #1e1b4b; color: #818cf8; }
  .col-label    { font-size: 12px; color: #8b90a5; margin-bottom: 2px; }
  .current-desc { font-size: 13px; color: #c4c8d8; padding: 6px 0; min-height: 40px; }
  .divider      { border-top: 1px solid #2e3345; margin: 8px 0; }
</style>
""", unsafe_allow_html=True)


# ── Excel export / import helpers ─────────────────────────────────────────

_HEADERS = [
    "table", "column_name", "data_type",
    "current_description", "proposed_description",
    "reviewer_approved",   # reviewer fills: Yes / No
    "reviewer_notes",      # reviewer fills: free text
]

_REVIEWER_COLS = {"reviewer_approved", "reviewer_notes"}
_FILL_HEADER   = PatternFill("solid", fgColor="1E3A5F")
_FILL_REVIEWER = PatternFill("solid", fgColor="FFF9C4")
_FILL_READONLY = PatternFill("solid", fgColor="F5F5F5")
_FONT_HEADER   = Font(bold=True, color="FFFFFF", size=11)
_FONT_READONLY = Font(color="555555")
_BORDER        = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="thin"),
)


def _build_excel(full_name: str, columns: list[dict], suggestions: dict) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Column Review"

    # Instruction banner (row 1)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(_HEADERS))
    instr = ws.cell(row=1, column=1,
        value=(
            "Instructions: Fill in 'reviewer_approved' (Yes or No) for each row. "
            "You may also edit 'proposed_description' or add 'reviewer_notes'. "
            "Do not change any other columns. Save and return this file."
        )
    )
    instr.fill = PatternFill("solid", fgColor="E8F4FD")
    instr.font = Font(italic=True, color="1A4A6E", size=10)
    instr.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 36

    # Header row (row 2)
    for col_idx, header in enumerate(_HEADERS, 1):
        cell = ws.cell(row=2, column=col_idx, value=header)
        cell.fill = _FILL_HEADER
        cell.font = _FONT_HEADER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _BORDER
    ws.row_dimensions[2].height = 28

    # Data rows
    for row_idx, col in enumerate(columns, 3):
        row_data = {
            "table": full_name,
            "column_name": col["name"],
            "data_type": col["type"],
            "current_description": col["current_comment"],
            "proposed_description": suggestions.get(col["name"], ""),
            "reviewer_approved": "",
            "reviewer_notes": "",
        }
        for col_idx, header in enumerate(_HEADERS, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=row_data[header])
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = _BORDER
            if header in _REVIEWER_COLS:
                cell.fill = _FILL_REVIEWER
            else:
                cell.fill = _FILL_READONLY
                cell.font = _FONT_READONLY

    # Column widths
    for i, w in enumerate([30, 20, 15, 45, 45, 18, 35], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A3"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _parse_excel(raw: bytes) -> dict[str, dict]:
    wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=True)
    ws = wb.active

    # Find the header row (contains "column_name")
    header_row_idx = None
    headers = []
    for row in ws.iter_rows():
        vals = [str(c.value or "").strip().lower() for c in row]
        if "column_name" in vals:
            header_row_idx = row[0].row
            headers = vals
            break

    if header_row_idx is None:
        raise ValueError("Could not find a header row with 'column_name'.")

    idx = {h: i for i, h in enumerate(headers)}
    result = {}
    for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
        name = str(row[idx["column_name"]] or "").strip()
        if not name:
            continue
        approved_raw = str(row[idx.get("reviewer_approved", -1)] or "").strip().lower()
        result[name] = {
            "proposed_description": str(row[idx.get("proposed_description", -1)] or "").strip(),
            "approved": approved_raw in ("yes", "y", "true", "1"),
            "notes": str(row[idx.get("reviewer_notes", -1)] or "").strip(),
        }
    return result


def _parse_csv(raw: bytes) -> dict[str, dict]:
    reader = csv.DictReader(io.StringIO(raw.decode("utf-8", errors="replace")))
    result = {}
    for row in reader:
        name = (row.get("column_name") or "").strip()
        if not name:
            continue
        approved_raw = (row.get("reviewer_approved") or "").strip().lower()
        result[name] = {
            "proposed_description": (row.get("proposed_description") or "").strip(),
            "approved": approved_raw in ("yes", "y", "true", "1"),
            "notes": (row.get("reviewer_notes") or "").strip(),
        }
    return result


# ── Session state defaults ────────────────────────────────────────────────

def _init_state():
    defaults = {
        "catalogs": [], "schemas": [], "tables": [], "columns": [],
        "selected_catalog": None, "selected_schema": None, "selected_table": None,
        "table_meta": {}, "suggestions": {}, "apply_results": {},
        "generated": False,
        "review_import": {},  # {col_name: {approved, notes, proposed_description}}
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init_state()


# ── Apply helper ──────────────────────────────────────────────────────────

def _do_apply(full_name: str, suggestions: dict) -> None:
    items = [(k, v) for k, v in suggestions.items() if v.strip()]
    if not items:
        st.warning("Nothing to apply.")
        return
    results = {}
    progress = st.progress(0, text="Applying…")
    for i, (col_name, comment) in enumerate(items):
        try:
            catalog.apply_column_comment(full_name, col_name, comment)
            results[col_name] = "ok"
        except Exception as e:
            results[col_name] = f"error: {e}"
        progress.progress((i + 1) / len(items), text=f"Saved {i + 1}/{len(items)}")
    st.session_state.apply_results = results
    progress.empty()
    failed = [k for k, v in results.items() if v != "ok"]
    if failed:
        st.warning(f"Some columns failed to save: {', '.join(failed)}")
    else:
        st.success(f"All {len(results)} column descriptions saved.")
    st.rerun()


# ── Sidebar ───────────────────────────────────────────────────────────────

def _reset(**extra):
    st.session_state.update(dict(
        columns=[], suggestions={}, apply_results={},
        generated=False, review_import={}, **extra
    ))


with st.sidebar:
    st.title("🗂️ Column Descriptions")
    st.caption("Browse Unity Catalog and manage column-level descriptions.")
    st.divider()

    if not st.session_state.catalogs:
        with st.spinner("Loading catalogs…"):
            try:
                st.session_state.catalogs = catalog.list_catalogs()
            except Exception as e:
                st.error(f"Could not connect to Databricks: {e}")
                st.stop()

    cat = st.selectbox("Catalog",
        options=["— select —"] + st.session_state.catalogs, key="ui_catalog")

    if cat and cat != "— select —" and cat != st.session_state.selected_catalog:
        _reset(selected_catalog=cat, selected_schema=None, selected_table=None,
               schemas=[], tables=[], table_meta={})
        with st.spinner("Loading schemas…"):
            try:
                st.session_state.schemas = catalog.list_schemas(cat)
            except Exception as e:
                st.error(str(e))

    schema = st.selectbox("Schema",
        options=["— select —"] + st.session_state.schemas,
        disabled=not st.session_state.schemas, key="ui_schema")

    if schema and schema != "— select —" and schema != st.session_state.selected_schema:
        _reset(selected_schema=schema, selected_table=None, tables=[], table_meta={})
        with st.spinner("Loading tables…"):
            try:
                st.session_state.tables = catalog.list_tables(cat, schema)
            except Exception as e:
                st.error(str(e))

    table_name = st.selectbox("Table",
        options=["— select —"] + [t["name"] for t in st.session_state.tables],
        disabled=not st.session_state.tables, key="ui_table")

    if table_name and table_name != "— select —":
        tbl_meta = next((t for t in st.session_state.tables if t["name"] == table_name), None)
        if tbl_meta and tbl_meta["full_name"] != (st.session_state.table_meta or {}).get("full_name"):
            _reset(selected_table=table_name, table_meta=tbl_meta)
            with st.spinner("Loading columns…"):
                try:
                    st.session_state.columns = catalog.get_columns(tbl_meta["full_name"])
                except Exception as e:
                    st.error(str(e))

    st.divider()

    if st.button("⚡ Generate & Humanize",
            disabled=not st.session_state.columns,
            use_container_width=True, type="primary"):
        tbl = st.session_state.table_meta
        cols = st.session_state.columns
        with st.spinner("Generating descriptions…"):
            try:
                raw = ai_gen.generate_column_descriptions(tbl["full_name"], cols)
            except Exception as e:
                st.error(str(e))
                raw = {}
        with st.spinner("Humanizing…"):
            humanized = hz.humanize_all(raw)
        st.session_state.suggestions = {
            c["name"]: humanized.get(c["name"], c["current_comment"]) for c in cols
        }
        st.session_state.apply_results = {}
        st.session_state.review_import = {}
        st.session_state.generated = True
        st.rerun()

    if st.session_state.generated:
        st.caption("✅ Descriptions generated and humanized.")


# ── Main area ─────────────────────────────────────────────────────────────

tbl = st.session_state.table_meta
cols = st.session_state.columns

if not tbl:
    st.markdown("### Select a table from the sidebar to get started.")
    st.stop()

type_label = tbl.get("type_label", "Delta Table")
badge_class = (
    "badge-dlt" if "Streaming" in type_label
    else "badge-mv" if "Materialized" in type_label
    else "badge-delta"
)
st.markdown(
    f"### `{tbl['full_name']}`"
    f"<span class='type-badge {badge_class}'>{type_label}</span>",
    unsafe_allow_html=True,
)

if not cols:
    st.info("No columns found for this table, or you may not have SELECT access.")
    st.stop()

if not st.session_state.generated:
    st.info(
        "Click **⚡ Generate & Humanize** in the sidebar to generate AI descriptions, "
        "or scroll down to review and edit existing descriptions."
    )

st.divider()

# ── Share for Review / Import Reviewed File ───────────────────────────────

export_col, import_col = st.columns(2)

with export_col:
    st.markdown("**📤 Share for Review**")
    st.caption(
        "Download an Excel file to share with your team. "
        "They fill in the yellow columns (approval + notes) and return it."
    )
    if cols:
        xlsx = _build_excel(tbl["full_name"], cols, st.session_state.suggestions)
        st.download_button(
            label="Download Review File (.xlsx)",
            data=xlsx,
            file_name=f"{tbl['full_name'].replace('.', '_')}_review.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

with import_col:
    st.markdown("**📥 Import Reviewed File**")
    st.caption(
        "Upload the completed file (.xlsx or .csv). "
        "Approved rows load into the grid. Use **Apply Approved Only** to publish just those."
    )
    uploaded = st.file_uploader(
        "Upload reviewed file", type=["xlsx", "csv"],
        label_visibility="collapsed", key="review_upload",
    )
    if uploaded:
        try:
            raw = uploaded.read()
            review_data = (
                _parse_csv(raw) if uploaded.name.endswith(".csv")
                else _parse_excel(raw)
            )
            approved_count = sum(1 for v in review_data.values() if v["approved"])
            rejected_count = len(review_data) - approved_count

            for col_name, data in review_data.items():
                if data["approved"] and data["proposed_description"]:
                    st.session_state.suggestions[col_name] = data["proposed_description"]
            st.session_state.review_import = review_data
            st.session_state.apply_results = {}

            st.success(f"✅ {approved_count} approved  •  ⛔ {rejected_count} not approved")
            st.rerun()
        except Exception as e:
            st.error(f"Could not read file: {e}")

st.divider()

# ── Column grid ───────────────────────────────────────────────────────────

h1, h2, h3, h4, h5 = st.columns([2, 1.2, 3, 3, 1])
h1.markdown("<div class='col-label'>Column</div>", unsafe_allow_html=True)
h2.markdown("<div class='col-label'>Type</div>", unsafe_allow_html=True)
h3.markdown("<div class='col-label'>Current Description</div>", unsafe_allow_html=True)
h4.markdown("<div class='col-label'>AI Suggestion (editable)</div>", unsafe_allow_html=True)
h5.markdown("<div class='col-label'>Humanize</div>", unsafe_allow_html=True)
st.markdown("<div class='divider'></div>", unsafe_allow_html=True)

for col in cols:
    name = col["name"]
    review = st.session_state.review_import.get(name)
    c1, c2, c3, c4, c5 = st.columns([2, 1.2, 3, 3, 1])

    c1.markdown(f"**{name}**")
    if review:
        c1.caption("✅ Approved" if review["approved"] else "⛔ Not approved")
        if review["notes"]:
            c1.caption(f"💬 {review['notes']}")

    c2.caption(col["type"])
    c3.markdown(
        f"<div class='current-desc'>{col['current_comment'] or '<em>—</em>'}</div>",
        unsafe_allow_html=True,
    )

    new_text = c4.text_area(
        label=f"suggestion_{name}",
        value=st.session_state.suggestions.get(name, col["current_comment"]),
        height=90, label_visibility="collapsed", key=f"text_{name}",
    )
    st.session_state.suggestions[name] = new_text

    with c5:
        st.markdown("<div style='padding-top:4px'></div>", unsafe_allow_html=True)
        if st.button("✨", key=f"hz_{name}", help=f"Re-humanize {name}"):
            st.session_state.suggestions[name] = hz.humanize(
                st.session_state.suggestions.get(name, "")
            )
            st.rerun()

    result = st.session_state.apply_results.get(name)
    if result == "ok":
        c4.success("Saved", icon="✅")
    elif result and result.startswith("error"):
        c4.error(result[7:])

    st.markdown("<div class='divider'></div>", unsafe_allow_html=True)


# ── Action buttons ────────────────────────────────────────────────────────

st.markdown("")
has_approved = any(v["approved"] for v in st.session_state.review_import.values())
btn1, btn2, btn3, _ = st.columns([1.5, 1.8, 1, 3.7])

if btn1.button("💾 Apply All", type="primary", use_container_width=True):
    _do_apply(tbl["full_name"], st.session_state.suggestions)

if btn2.button("💾 Apply Approved Only",
        disabled=not has_approved, use_container_width=True,
        help="Only applies columns marked 'Yes' by the reviewer"):
    approved_only = {
        name: st.session_state.suggestions[name]
        for name, data in st.session_state.review_import.items()
        if data["approved"] and name in st.session_state.suggestions
    }
    _do_apply(tbl["full_name"], approved_only)

if btn3.button("Clear", use_container_width=True):
    st.session_state.suggestions = {}
    st.session_state.apply_results = {}
    st.session_state.review_import = {}
    st.session_state.generated = False
    st.rerun()
