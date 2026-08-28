"""Column Description Review — Streamlit app."""

import io
import os
import csv
from datetime import datetime

import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import config
import ai_gen
import humanize as hz
from databricks_client import DatabricksClient

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


# ── Code export helpers ───────────────────────────────────────────────────

def _build_pyspark_script(full_name: str, suggestions: dict) -> str:
    quoted = ".".join(f"`{p}`" for p in full_name.split("."))
    lines = [
        f"# Column descriptions for {full_name}",
        f"# Generated {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "",
        "from pyspark.sql import SparkSession",
        "spark = SparkSession.builder.getOrCreate()",
        "",
    ]
    for col_name, desc in suggestions.items():
        if desc.strip():
            col_safe = col_name.replace("`", "``")
            escaped  = desc.replace('"', '\\"')
            lines.append(
                f'spark.sql("""COMMENT ON COLUMN {quoted}.`{col_safe}` IS \\"{escaped}\\"""")'
            )
    return "\n".join(lines)


# ── Session state defaults ────────────────────────────────────────────────

def _init_state():
    defaults = {
        # Connection — host/token pre-seeded from .env if set
        "connected": False,
        "db_client": None,
        "current_user": "",
        "db_host": config.DEFAULT_HOST,
        "db_token": config.DEFAULT_TOKEN,
        # Browse
        "catalogs": [], "schemas": [], "tables": [], "columns": [],
        "selected_catalog": None, "selected_schema": None, "selected_table": None,
        "table_meta": {}, "suggestions": {},
        "generated": False,
        "gen_error": None,    # persists generation errors across reruns
        "review_import": {},  # {col_name: {approved, notes, proposed_description}}
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init_state()


# ── Connect helper ────────────────────────────────────────────────────────

def _connect_databricks(host: str, token: str) -> None:
    try:
        client = DatabricksClient(host, token)
        user = client.test_connection()
    except Exception as e:
        st.error(f"Connection failed: {e}")
        return

    st.session_state.db_client = client
    st.session_state.connected = True
    st.session_state.current_user = user
    # Reset browse state so catalogs reload with new credentials
    for key in ("catalogs", "schemas", "tables", "columns", "table_meta",
                "suggestions", "review_import"):
        st.session_state[key] = [] if isinstance(st.session_state[key], list) else {}
    st.session_state.generated = False
    st.session_state.selected_catalog = None
    st.session_state.selected_schema = None
    st.session_state.selected_table = None
    st.toast(f"Connected as {user}", icon="✅")


# ── Sidebar ───────────────────────────────────────────────────────────────

def _reset(**extra):
    # Clear text widget keys for the current table so stale values don't bleed
    # into the next table's widgets when column names happen to match.
    for col in st.session_state.get("columns", []):
        st.session_state.pop(f"text_{col['name']}", None)
    st.session_state.update(dict(
        columns=[], suggestions={},
        generated=False, gen_error=None, review_import={}, **extra
    ))


with st.sidebar:
    st.title("🗂️ Column Descriptions")
    st.caption("Browse Unity Catalog and manage column-level descriptions.")

    # ── Connect form ──────────────────────────────────────────────────────
    connected = st.session_state.connected
    badge = "🟢 Live" if connected else "🔴 Off"
    with st.expander(f"☁️ Databricks Connection  {badge}", expanded=not connected):
        st.text_input(
            "Workspace URL",
            key="db_host",
            placeholder="https://your-workspace.azuredatabricks.net",
        )
        st.text_input(
            "Personal Access Token",
            key="db_token",
            type="password",
            placeholder="dapi…",
        )
        if st.button(
            "Connect",
            type="primary" if not connected else "secondary",
            use_container_width=True,
        ):
            host = st.session_state.db_host.strip()
            token = st.session_state.db_token.strip()
            if not host or not token:
                st.error("Both Workspace URL and token are required.")
            else:
                _connect_databricks(host, token)
                st.rerun()

        if connected:
            st.caption(f"Connected as **{st.session_state.current_user}**")

    if not connected:
        st.stop()

    st.divider()

    # ── Catalog browser (only shown after connect) ────────────────────────
    client: DatabricksClient = st.session_state.db_client

    if not st.session_state.catalogs:
        with st.spinner("Loading catalogs…"):
            st.session_state.catalogs = client.list_catalogs()
            if not st.session_state.catalogs:
                st.warning("No catalogs found. Check your USE CATALOG privilege.")

    cat = st.selectbox("Catalog",
        options=["— select —"] + st.session_state.catalogs, key="ui_catalog")

    if cat and cat != "— select —" and cat != st.session_state.selected_catalog:
        _reset(selected_catalog=cat, selected_schema=None, selected_table=None,
               schemas=[], tables=[], table_meta={})
        with st.spinner("Loading schemas…"):
            st.session_state.schemas = client.list_schemas(cat)
            if not st.session_state.schemas:
                st.warning(f"No schemas found in `{cat}`. Check your USE SCHEMA privilege.")

    schema = st.selectbox("Schema",
        options=["— select —"] + st.session_state.schemas,
        disabled=not st.session_state.schemas, key="ui_schema")

    if schema and schema != "— select —" and schema != st.session_state.selected_schema:
        _reset(selected_schema=schema, selected_table=None, tables=[], table_meta={})
        with st.spinner("Loading tables…"):
            st.session_state.tables = client.list_tables(cat, schema)
            if not st.session_state.tables:
                st.warning(f"No tables found in `{cat}.{schema}`.")

    table_name = st.selectbox("Table",
        options=["— select —"] + [t["name"] for t in st.session_state.tables],
        disabled=not st.session_state.tables, key="ui_table")

    if table_name and table_name != "— select —":
        tbl_meta = next((t for t in st.session_state.tables if t["name"] == table_name), None)
        if tbl_meta and tbl_meta["full_name"] != (st.session_state.table_meta or {}).get("full_name"):
            _reset(selected_table=table_name, table_meta=tbl_meta)
            with st.spinner("Loading columns…"):
                try:
                    st.session_state.columns = client.get_columns(tbl_meta["full_name"])
                except PermissionError as e:
                    parts = str(e).split("Raw error:", 1)
                    st.error(f"**Access denied** — {parts[0].strip()}")
                    if len(parts) > 1:
                        with st.expander("Raw error (share with your admin)"):
                            st.code(parts[1].strip())
                except Exception as e:
                    st.error(str(e))

    st.divider()

    if st.button("⚡ Generate & Humanize",
            disabled=not st.session_state.columns,
            use_container_width=True, type="primary"):
        tbl = st.session_state.table_meta
        cols = st.session_state.columns
        ws = st.session_state.db_client._ws
        with st.spinner("Generating descriptions…"):
            try:
                raw = ai_gen.generate_column_descriptions(ws, tbl["full_name"], cols)
            except Exception as e:
                st.session_state.gen_error = str(e)
                st.rerun()  # raises RerunException — nothing below executes

        # Only reached if generation succeeded
        st.session_state.gen_error = None
        with st.spinner("Humanizing…"):
            humanized = hz.humanize_all(raw)
        for c in cols:
            val = humanized.get(c["name"], c["current_comment"])
            st.session_state.suggestions[c["name"]] = val
            st.session_state[f"text_{c['name']}"] = val  # keep widget in sync
        st.session_state.review_import = {}
        st.session_state.generated = True
        st.rerun()

    if st.session_state.get("gen_error"):
        st.error(st.session_state.gen_error)

    if st.session_state.generated:
        st.caption("✅ Descriptions generated and humanized.")

    st.divider()
    if st.button("⏹ Stop Server", use_container_width=True,
                 help="Shuts down the Streamlit process. Restart with: python -m streamlit run app.py"):
        os._exit(0)


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
                    val = data["proposed_description"]
                    st.session_state.suggestions[col_name] = val
                    st.session_state[f"text_{col_name}"] = val  # keep widget in sync
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
            val = hz.humanize(st.session_state.suggestions.get(name, ""))
            st.session_state.suggestions[name] = val
            st.session_state[f"text_{name}"] = val  # keep widget in sync
            st.rerun()

    st.markdown("<div class='divider'></div>", unsafe_allow_html=True)


# ── Action buttons ────────────────────────────────────────────────────────

st.markdown("")
has_approved = any(v["approved"] for v in st.session_state.review_import.values())
approved_only = {
    name: st.session_state.suggestions[name]
    for name, data in st.session_state.review_import.items()
    if data["approved"] and name in st.session_state.suggestions
}

if st.button("🗑 Clear", use_container_width=False):
    for col in st.session_state.columns:
        st.session_state.pop(f"text_{col['name']}", None)
    st.session_state.suggestions = {}
    st.session_state.review_import = {}
    st.session_state.generated = False
    st.session_state.gen_error = None
    st.rerun()

# ── Export as code ────────────────────────────────────────────────────────

st.markdown("")
st.markdown("**📋 Export as code** — paste into a Databricks notebook to apply descriptions manually")

safe_name = tbl["full_name"].replace(".", "_")
py_all = _build_pyspark_script(tbl["full_name"], st.session_state.suggestions)

exp1, exp2, _ = st.columns([1.8, 1.8, 4.4])
exp1.download_button("⬇ PySpark (All)", data=py_all,
    file_name=f"{safe_name}_all.py", mime="text/plain", use_container_width=True)

if has_approved:
    py_app = _build_pyspark_script(tbl["full_name"], approved_only)
    exp2.download_button("⬇ PySpark (Approved)", data=py_app,
        file_name=f"{safe_name}_approved.py", mime="text/plain", use_container_width=True)
