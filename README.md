# Column Description Review

A Streamlit app for reviewing, editing, and publishing AI-generated column descriptions to Databricks Unity Catalog tables. Supports Delta tables, DLT streaming tables, and materialized views.

## Features

- Browse any catalog → schema → table you have access to
- Generate AI descriptions for all columns in one click
- Descriptions are automatically humanized (AI-sounding language removed)
- Edit any description inline before saving
- Re-humanize individual columns with the ✨ button after manual edits
- Apply changes back to the table with a single "Apply All" click
- Works for Delta tables, DLT streaming tables, and materialized views

## Setup (first time)

### 1. Clone the repo

```
git clone <repo-url>
cd column-descriptions
```

### 2. Create a virtual environment and install dependencies

```
python -m venv .venv
.venv\Scripts\activate        # Windows
pip install -r requirements.txt
```

### 3. Configure your credentials

Copy the example config and fill in your values:

```
copy .env.example .env
```

Open `.env` and set these three values:

| Variable | Where to find it |
|---|---|
| `DATABRICKS_HOST` | Your workspace URL — e.g. `https://adb-1234567890.azuredatabricks.net` |
| `DATABRICKS_TOKEN` | Databricks UI → top-right avatar → Settings → Developer → Access tokens → Generate new token |
| `DATABRICKS_WAREHOUSE_ID` | Databricks UI → SQL → Warehouses → click your warehouse → Connection details → last segment of the HTTP path |

No separate API keys or model names needed — the token covers both catalog access and AI generation.

### 5. Run the app

```
streamlit run app.py
```

The app opens in your browser at `http://localhost:8501`.

## Team sharing

Every team member follows the same setup steps. Each person uses their own Databricks personal access token — this ensures the app respects their individual Unity Catalog permissions. They will only see catalogs, schemas, and tables they have access to.

## How the humanize step works

After AI generation, descriptions are automatically cleaned up by `humanize.py` — a rule-based post-processor (no extra LLM call) that:

- Swaps AI vocabulary for plain English (`bolstered` → `supported`, `delve` → `explore`, etc.)
- Removes filler phrases (`It is worth noting that…`, `…highlighting its significance`)
- Fixes unnatural constructions (`serves as` → `is`, `showcasing` → `showing`)

You can re-run humanize on any individual column after editing by clicking the ✨ button on that row.

## Permissions required in Databricks

- `USE CATALOG` on the catalog
- `USE SCHEMA` on the schema
- `SELECT` on the table (to read column metadata)
- `MODIFY` on the table (to write column comments)
