# CEDP Semantic Studio

A Streamlit app for generating, reviewing, and exporting AI-generated column descriptions for Databricks Unity Catalog tables. Built for the Clickstream Engagement Data Product (CEDP) team.

Supports Delta tables, DLT streaming tables, and materialized views.

---

## Workflow

```
Connect → Browse → Generate → Review → Export → Apply in Databricks
```

### 1. Connect
Enter your Databricks workspace URL and personal access token in the **☁️ Databricks Connection** sidebar panel. The app stores credentials in the session only — nothing is written to disk.

If you have a `.env` file with `DATABRICKS_HOST` and `DATABRICKS_TOKEN`, those values pre-fill the form automatically.

### 2. Browse
Select a **Catalog → Schema → Table** from the cascading dropdowns. The app uses the Unity Catalog REST API to list objects you have access to — tables or schemas you cannot see are silently skipped.

### 3. Generate & Humanize
Click **⚡ Generate & Humanize** in the sidebar. The app:
- Calls a Databricks Foundation Model endpoint (`databricks-meta-llama-3-3-70b-instruct`) to generate business-friendly column descriptions tailored to clickstream / digital behavioral analytics data
- Runs a rule-based post-processor (`humanize.py`) to remove AI-sounding language — no extra API call

Columns are batched in groups of 15 to stay within the model's output token limit. Large tables automatically split failing batches further.

### 4. Review and Edit
Descriptions appear in an editable grid. You can:
- Edit any description directly in the text area
- Click **✨** on a row to re-humanize that column after manual edits

#### Team review via Excel
1. Click **Download Review File (.xlsx)** to export a spreadsheet with proposed descriptions
2. Share with stakeholders — reviewers fill in the yellow `reviewer_approved` (Yes/No) and `reviewer_notes` columns
3. Upload the completed file using **Import Reviewed File** — approved descriptions load back into the grid

### 5. Export as PySpark
Once descriptions are ready, download a PySpark script and run it in a Databricks notebook:

| Button | What it exports |
|---|---|
| **⬇ PySpark (All)** | All non-empty descriptions |
| **⬇ PySpark (Approved)** | Only reviewer-approved descriptions (appears after import) |

### 6. Apply in Databricks
Open the downloaded `.py` file and paste it into a Databricks notebook cell. Run it to apply the column comments:

```python
# Example of generated file content:
# Column descriptions for my_catalog.clickstream.fact_page_views
# Generated 2026-08-28 09:14

from pyspark.sql import SparkSession
spark = SparkSession.builder.getOrCreate()

spark.sql("""COMMENT ON COLUMN `my_catalog`.`clickstream`.`fact_page_views`.`hitKey` IS \"Unique identifier for a hit or event.\"""")
spark.sql("""COMMENT ON COLUMN `my_catalog`.`clickstream`.`fact_page_views`.`sessionId` IS \"Surrogate key identifying the user session.\"""")
```

Each `spark.sql()` call applies one column comment using `COMMENT ON COLUMN` syntax, which works for Delta tables, DLT streaming tables, and materialized views.

---

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
source .venv/bin/activate     # Mac / Linux
pip install -r requirements.txt
```

### 3. Configure credentials (optional)

Copy the example config to pre-fill the Connect form:

```
copy .env.example .env        # Windows
cp .env.example .env          # Mac / Linux
```

| Variable | Where to find it |
|---|---|
| `DATABRICKS_HOST` | Your workspace URL — e.g. `https://adb-1234567890.azuredatabricks.net` |
| `DATABRICKS_TOKEN` | Databricks UI → avatar → Settings → Developer → Access tokens → Generate new token |
| `DATABRICKS_WAREHOUSE_ID` | Databricks UI → SQL Warehouses → your warehouse → Connection details → last segment of HTTP path |

`DATABRICKS_HOST` and `DATABRICKS_TOKEN` can also be entered directly in the app's Connect form — the `.env` file just pre-fills them.
`DATABRICKS_WAREHOUSE_ID` is required in `.env` (used internally for DDL operations).

### 4. Run the app

```
python -m streamlit run app.py
```

The app opens at `http://localhost:8501`.

### 5. Stop the app

Press **Ctrl+C** in the terminal, or click **⏹ Stop Server** at the bottom of the sidebar.

---

## Team sharing

Each team member follows the same setup steps and uses their own personal access token — this ensures the app respects individual Unity Catalog permissions. People only see catalogs, schemas, and tables they have access to.

---

## Why Databricks Foundation Model APIs instead of an Anthropic key

The app calls the LLM through the Databricks serving endpoint (`databricks-meta-llama-3-3-70b-instruct`) rather than the Anthropic API for three reasons:

1. **No separate key to manage** — the same personal access token used to browse the catalog also authenticates LLM calls. There is nothing extra to generate, rotate, or share with the team.
2. **Data stays inside the workspace** — prompts and column metadata never leave your Databricks environment. Using an external API (Anthropic, OpenAI) would send table and column names to a third-party endpoint, which may conflict with data governance or security policies.
3. **No additional cost centre** — Foundation Model API usage is billed through your existing Databricks contract, not a separate AI vendor account.

If you ever need to switch models, update `SERVING_ENDPOINT` in `config.py` to any other serving endpoint available in your workspace.

---

## How humanize works

After AI generation, `humanize.py` cleans up descriptions with a rule-based post-processor (no extra LLM call):

- Swaps AI vocabulary for plain English (`bolstered` → `supported`, `leverage` → `use`, etc.)
- Removes filler phrases (`It is worth noting that…`, `highlighting its significance`)
- Fixes unnatural constructions (`serves as` → `is`, `showcasing` → `showing`)

Click **✨** on any row to re-run humanize after a manual edit.

---

## Permissions required

| Privilege | Required for |
|---|---|
| `USE CATALOG` | Browsing schemas |
| `USE SCHEMA` | Browsing tables |
| `SELECT` on the table | Reading column metadata |

Write access (`MODIFY`) is no longer required — descriptions are applied manually via the exported PySpark script, using whichever account runs the notebook.
