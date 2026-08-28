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
    page_title="CEDP Semantic Studio",
    page_icon="📖",
    layout="wide",
)

# ── Custom styling ────────────────────────────────────────────────────────

st.markdown("""
<style>
  /* ── Table type badges ─────────────────────────────────── */
  .type-badge {
    display: inline-block; padding: 2px 10px; border-radius: 12px;
    font-size: 11px; font-weight: 700; letter-spacing: 0.5px;
    text-transform: uppercase; margin-left: 10px; vertical-align: middle;
  }
  .badge-dlt   { background: #14532d; color: #4ade80; }
  .badge-mv    { background: #451a03; color: #fbbf24; }
  .badge-delta { background: #1e1b4b; color: #818cf8; }

  /* ── Column name pill ──────────────────────────────────── */
  .col-name {
    display: inline-block; background: #141824; border: 1px solid #2e3755;
    border-radius: 5px; padding: 3px 9px;
    font-family: 'Courier New', monospace; font-size: 13px; color: #c7d2fe;
    word-break: break-all; line-height: 1.5;
  }

  /* ── Data type pills ───────────────────────────────────── */
  .dt-pill {
    display: inline-block; padding: 2px 8px; border-radius: 4px;
    font-size: 11px; font-weight: 600; font-family: monospace; white-space: nowrap;
  }
  .dt-string  { background: #0f2744; color: #7dd3fc; }
  .dt-int     { background: #0a2818; color: #86efac; }
  .dt-double  { background: #132010; color: #a3e635; }
  .dt-bool    { background: #2d1503; color: #fb923c; }
  .dt-date    { background: #1a0f38; color: #c084fc; }
  .dt-other   { background: #161b2e; color: #94a3b8; }

  /* ── Column grid ───────────────────────────────────────── */
  .col-label {
    font-size: 10px; font-weight: 700; letter-spacing: 0.9px;
    text-transform: uppercase; color: #4f5a7a; margin-bottom: 4px;
  }
  .current-desc {
    font-size: 13px; color: #8895b3; padding: 6px 2px;
    min-height: 44px; line-height: 1.55;
  }
  .divider { border-top: 1px solid #1a2035; margin: 6px 0; }

  /* ── Review status chips ───────────────────────────────── */
  .chip {
    display: inline-block; padding: 2px 9px; border-radius: 10px;
    font-size: 11px; font-weight: 600; margin: 3px 2px 0 0; line-height: 1.6;
  }
  .chip-approved { background: #0d3320; color: #4ade80; border: 1px solid #166534; }
  .chip-rejected { background: #2d0a0a; color: #f87171; border: 1px solid #7f1d1d; }
  .chip-note     { background: #111827; color: #94a3b8; border: 1px solid #1e2a3a; }

  /* ── Hint / info card ──────────────────────────────────── */
  .hint-card {
    background: rgba(99,102,241,0.07); border: 1px solid rgba(99,102,241,0.22);
    border-radius: 8px; padding: 11px 16px; font-size: 13px; color: #a5b4fc;
    margin: 4px 0 16px 0; line-height: 1.5;
  }

  /* ── Section headers ───────────────────────────────────── */
  .section-title {
    font-size: 13px; font-weight: 700; color: #cbd5e1;
    margin-bottom: 4px; letter-spacing: 0.2px;
  }
  .section-caption {
    font-size: 12px; color: #4f5a7a; margin-bottom: 10px; line-height: 1.5;
  }

  /* ── Overview label ────────────────────────────────────── */
  .overview-label {
    font-size: 10px; font-weight: 700; letter-spacing: 0.9px;
    text-transform: uppercase; color: #6366f1; margin-bottom: 6px;
  }

  /* ── Empty state ───────────────────────────────────────── */
  .empty-state {
    text-align: center; padding: 64px 32px; color: #4f5a7a;
  }
  .empty-state .es-icon { font-size: 52px; margin-bottom: 18px; line-height: 1; }
  .empty-state h3 { color: #7c8db5; font-size: 20px; margin: 0 0 10px 0; font-weight: 600; }
  .empty-state p  { font-size: 14px; line-height: 1.75; margin: 0; }
  .empty-state .step {
    display: inline-flex; align-items: center; gap: 8px;
    background: #0f1422; border: 1px solid #1e2a3a;
    border-radius: 8px; padding: 8px 16px; margin: 6px 4px;
    font-size: 13px; color: #64748b;
  }
  .empty-state .step-num {
    background: #1e2a45; color: #818cf8; border-radius: 50%;
    width: 20px; height: 20px; display: inline-flex;
    align-items: center; justify-content: center;
    font-size: 11px; font-weight: 700; flex-shrink: 0;
  }
</style>
""", unsafe_allow_html=True)


def _type_pill(dtype: str) -> str:
    dl = dtype.lower()
    if any(x in dl for x in ("string", "varchar", "char", "text")):
        cls = "dt-string"
    elif any(x in dl for x in ("bigint", "int", "long", "short", "byte", "tinyint", "smallint")):
        cls = "dt-int"
    elif any(x in dl for x in ("double", "float", "decimal", "numeric", "real")):
        cls = "dt-double"
    elif "bool" in dl:
        cls = "dt-bool"
    elif any(x in dl for x in ("date", "timestamp")):
        cls = "dt-date"
    else:
        cls = "dt-other"
    return f"<span class='dt-pill {cls}'>{dtype}</span>"


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


def _build_excel(full_name: str, columns: list[dict], suggestions: dict, table_description: str = "") -> bytes:
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

    # Table overview row (row 2)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(_HEADERS))
    td_cell = ws.cell(row=2, column=1,
        value=f"Table Overview: {table_description or '(not generated)'}")
    td_cell.fill = PatternFill("solid", fgColor="E8F0FE")
    td_cell.font = Font(italic=True, color="1A237E", size=10)
    td_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 30

    # Header row (row 3)
    for col_idx, header in enumerate(_HEADERS, 1):
        cell = ws.cell(row=3, column=col_idx, value=header)
        cell.fill = _FILL_HEADER
        cell.font = _FONT_HEADER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _BORDER
    ws.row_dimensions[3].height = 28

    # Data rows (row 4+)
    for row_idx, col in enumerate(columns, 4):
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

    ws.freeze_panes = "A4"

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

def _build_pyspark_script(full_name: str, suggestions: dict, table_description: str = "") -> str:
    quoted = ".".join(f"`{p}`" for p in full_name.split("."))
    lines = [
        f"# Column descriptions for {full_name}",
        f"# Generated {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "",
        "from pyspark.sql import SparkSession",
        "spark = SparkSession.builder.getOrCreate()",
        "",
    ]
    if table_description.strip():
        escaped = table_description.replace('"', '\\"')
        lines.append(f'spark.sql("""COMMENT ON TABLE {quoted} IS \\"{escaped}\\"""")')
        lines.append("")
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
        "table_meta": {}, "suggestions": {}, "table_description": "",
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
    st.session_state.pop("text_table_desc", None)
    st.session_state.update(dict(
        columns=[], suggestions={}, table_description="",
        generated=False, gen_error=None, review_import={}, **extra
    ))


with st.sidebar:
    st.title("📖 CEDP Semantic Studio")
    st.caption("Generate and export column descriptions for Unity Catalog tables.")

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
        with st.spinner("Generating table overview…"):
            try:
                raw_td = ai_gen.generate_table_description(
                    ws, tbl["full_name"], cols, tbl.get("comment", "")
                )
                td = hz.humanize(raw_td)
            except Exception:
                td = st.session_state.table_description  # keep existing on failure
        st.session_state.table_description = td
        st.session_state["text_table_desc"] = td
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
    st.markdown("""
<div class='empty-state'>
  <div class='es-icon'>📖</div>
  <h3>Select a table to get started</h3>
  <p>Browse your Unity Catalog from the sidebar, then generate AI descriptions with one click.</p>
  <br>
  <div>
    <span class='step'><span class='step-num'>1</span>Connect to Databricks</span>
    <span class='step'><span class='step-num'>2</span>Pick a catalog → schema → table</span>
    <span class='step'><span class='step-num'>3</span>Click ⚡ Generate &amp; Humanize</span>
    <span class='step'><span class='step-num'>4</span>Review, export &amp; apply in Databricks</span>
  </div>
</div>
""", unsafe_allow_html=True)
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

# ── Table overview ────────────────────────────────────────────────────────
st.markdown("<div class='overview-label'>Table Overview</div>", unsafe_allow_html=True)
td_left, td_right = st.columns([6, 1])
table_desc = td_left.text_area(
    "table_overview",
    value=st.session_state.table_description or tbl.get("comment", ""),
    height=80,
    label_visibility="collapsed",
    key="text_table_desc",
    placeholder="Click ⚡ Generate & Humanize to produce an AI overview, or type one manually.",
)
st.session_state.table_description = table_desc
with td_right:
    st.markdown("<div style='padding-top:4px'></div>", unsafe_allow_html=True)
    if st.button("✨", key="hz_table_desc", help="Re-humanize table description"):
        humanized_td = hz.humanize(table_desc)
        st.session_state.table_description = humanized_td
        st.session_state["text_table_desc"] = humanized_td
        st.rerun()
if tbl.get("comment"):
    st.caption(f"Current in Databricks: {tbl['comment']}")

if not cols:
    st.info("No columns found for this table, or you may not have SELECT access.")
    st.stop()

if not st.session_state.generated:
    st.markdown(
        "<div class='hint-card'>⚡ Click <strong>Generate &amp; Humanize</strong> in the sidebar "
        "to produce AI descriptions — or type directly in the suggestion fields below.</div>",
        unsafe_allow_html=True,
    )

st.divider()

# ── Share for Review / Import Reviewed File ───────────────────────────────

export_col, import_col = st.columns(2)

with export_col:
    st.markdown("<div class='section-title'>📤 Share for Review</div><div class='section-caption'>Download an Excel file for your team. They fill in the approval + notes columns and return it.</div>", unsafe_allow_html=True)
    if cols:
        xlsx = _build_excel(tbl["full_name"], cols, st.session_state.suggestions, st.session_state.table_description)
        st.download_button(
            label="Download Review File (.xlsx)",
            data=xlsx,
            file_name=f"{tbl['full_name'].replace('.', '_')}_review.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

with import_col:
    st.markdown("<div class='section-title'>📥 Import Reviewed File</div><div class='section-caption'>Upload the completed file (.xlsx or .csv). Approved rows load back into the grid automatically.</div>", unsafe_allow_html=True)
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
h5.markdown("<div class='col-label'>✨</div>", unsafe_allow_html=True)
st.markdown("<div class='divider'></div>", unsafe_allow_html=True)

for col in cols:
    name = col["name"]
    review = st.session_state.review_import.get(name)
    c1, c2, c3, c4, c5 = st.columns([2, 1.2, 3, 3, 1])

    review_html = ""
    if review:
        if review["approved"]:
            review_html += "<span class='chip chip-approved'>✓ Approved</span>"
        else:
            review_html += "<span class='chip chip-rejected'>✗ Rejected</span>"
        if review["notes"]:
            review_html += f"<br><span class='chip chip-note'>💬 {review['notes']}</span>"

    c1.markdown(
        f"<div class='col-name'>{name}</div>{review_html}",
        unsafe_allow_html=True,
    )

    c2.markdown(_type_pill(col["type"]), unsafe_allow_html=True)
    c3.markdown(
        f"<div class='current-desc'>{col['current_comment'] or '<span style=\"color:#2e3755\">—</span>'}</div>",
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
    st.session_state.pop("text_table_desc", None)
    st.session_state.suggestions = {}
    st.session_state.table_description = ""
    st.session_state.review_import = {}
    st.session_state.generated = False
    st.session_state.gen_error = None
    st.rerun()

# ── Export as code ────────────────────────────────────────────────────────

st.markdown("<div class='section-title'>📋 Export as PySpark</div><div class='section-caption'>Download or copy the script — paste into a Databricks notebook to apply descriptions.</div>", unsafe_allow_html=True)

safe_name = tbl["full_name"].replace(".", "_")
table_desc_export = st.session_state.table_description
py_all = _build_pyspark_script(tbl["full_name"], st.session_state.suggestions, table_desc_export)

if has_approved:
    py_app = _build_pyspark_script(tbl["full_name"], approved_only, table_desc_export)

exp1, exp2, _ = st.columns([1.8, 1.8, 4.4])
exp1.download_button("⬇ Download (All)", data=py_all,
    file_name=f"{safe_name}_all.py", mime="text/plain", use_container_width=True)
if has_approved:
    exp2.download_button("⬇ Download (Approved only)", data=py_app,
        file_name=f"{safe_name}_approved.py", mime="text/plain", use_container_width=True)

with st.expander("Preview PySpark (All)", expanded=False):
    st.code(py_all, language="python")

if has_approved:
    with st.expander("Preview PySpark (Approved only)", expanded=False):
        st.code(py_app, language="python")
