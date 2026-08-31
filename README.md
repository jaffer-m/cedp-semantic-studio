# CEDP Semantic Studio

A Streamlit app for generating, reviewing, and exporting AI-generated descriptions for Databricks Unity Catalog tables. Built for the Clickstream Engagement Data Product (CEDP) team.

Supports Delta tables, DLT streaming tables, and materialized views.

---

## Workflow

```
Connect → Browse → Generate → Review → Export → Apply in Databricks
```

### 1. Connect
Enter your Databricks workspace URL and personal access token in the **☁️ Databricks Connection** sidebar panel. The app stores credentials in the session only — nothing is written to disk.

If you have a `.env` file with `DATABRICKS_HOST` and `DATABRICKS_TOKEN`, those values pre-fill the form automatically.

> **Using a different workspace (e.g. prod vs stage)?**  
> The **Model Endpoint** field in the Connect form lets you override the serving endpoint name per workspace without changing any config files. If your prod workspace uses a different model deployment, type its name there before connecting.

### 2. Browse
Select a **Catalog → Schema → Table** from the cascading dropdowns. The app uses the Unity Catalog REST API to list objects you have access to — tables or schemas you cannot see are silently skipped.

### 3. Generate & Humanize
Click **⚡ Generate & Humanize** in the sidebar. The app:
- Generates a **table-level overview** (2–3 sentence business description of what the table contains, its grain, and primary use)
- Generates **column descriptions** tailored to clickstream / digital behavioral analytics data
- Runs a rule-based post-processor (`humanize.py`) to remove AI-sounding language — no extra API call

Columns are batched in groups of 15 to stay within the model's output token limit. Large tables automatically split failing batches further.

### 4. Review and Edit
Descriptions appear in an editable grid. You can:
- Edit the **table overview** text area at the top
- Edit any column description directly in its text area
- Click **✨** on the table overview or any row to re-humanize that text after manual edits

#### Team review via Excel
1. Click **Download Review File (.xlsx)** — exports a spreadsheet with the table overview (row 2) and all proposed column descriptions
2. Share with stakeholders — reviewers fill in the yellow `reviewer_approved` (Yes/No) and `reviewer_notes` columns
3. Upload the completed file using **Import Reviewed File** — approved descriptions load back into the grid

### 5. Export as PySpark
Once descriptions are ready, use the **📋 Export as PySpark** section:

| Button | What it exports |
|---|---|
| **⬇ Download (All)** | All non-empty descriptions |
| **⬇ Download (Approved only)** | Only reviewer-approved descriptions (appears after import) |

Use **Preview PySpark** to expand and copy the script directly — no download needed for quick paste into a notebook.

### 6. Apply in Databricks
Paste the script into a Databricks notebook cell and run it:

```python
# Column descriptions for my_catalog.clickstream.fact_page_views
# Generated 2026-08-28 09:14

from pyspark.sql import SparkSession
spark = SparkSession.builder.getOrCreate()

spark.sql("""COMMENT ON TABLE `my_catalog`.`clickstream`.`fact_page_views` IS \"Captures page view events at hit level, including session context, page metadata, and visit attribution.\"""")

spark.sql("""COMMENT ON COLUMN `my_catalog`.`clickstream`.`fact_page_views`.`hitKey` IS \"Unique identifier for a hit or event.\"""")
spark.sql("""COMMENT ON COLUMN `my_catalog`.`clickstream`.`fact_page_views`.`sessionId` IS \"Surrogate key identifying the user session.\"""")
```

`COMMENT ON TABLE` and `COMMENT ON COLUMN` work for Delta tables, DLT streaming tables, and materialized views.

---

## Setup (first time)

### 1. Clone the repo

```
git clone <repo-url>
cd cedp-semantic-studio
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

| Variable | Required | Where to find it |
|---|---|---|
| `DATABRICKS_WAREHOUSE_ID` | Yes | Databricks UI → SQL Warehouses → your warehouse → Connection details → last segment of HTTP path |
| `DATABRICKS_HOST` | No | Your workspace URL — e.g. `https://adb-1234567890.azuredatabricks.net` (can also be typed in the UI) |
| `DATABRICKS_TOKEN` | No | Databricks UI → avatar → Settings → Developer → Access tokens → Generate new token (can also be typed in the UI) |
| `DATABRICKS_SERVING_ENDPOINT` | No | Override the model endpoint name if your workspace uses a different deployment. Defaults to `databricks-meta-llama-3-3-70b-instruct` |

### 4. Run the app

```
python -m streamlit run app.py
```

The app opens at `http://localhost:8501`.

### 5. Stop the app

Press **Ctrl+C** in the terminal, or click **⏹ Stop Server** at the bottom of the sidebar.

---

## Multi-workspace usage (stage vs prod)

The model endpoint name often differs between workspaces. To find the correct name in any workspace, run in a Databricks notebook:

```python
display(spark.sql("SHOW ENDPOINTS"))
```

Or check **Machine Learning → Serving** in the Databricks UI. Then either:
- Type the endpoint name into the **Model Endpoint** field in the Connect form, or
- Set `DATABRICKS_SERVING_ENDPOINT=<name>` in the `.env` file for that environment

---

## Team sharing

Each team member follows the same setup steps and uses their own personal access token — this ensures the app respects individual Unity Catalog permissions. People only see catalogs, schemas, and tables they have access to.

---

## Why Databricks Foundation Model APIs instead of an Anthropic key

The app calls the LLM through the Databricks serving endpoint rather than an external API for three reasons:

1. **No separate key to manage** — the same personal access token used to browse the catalog also authenticates LLM calls. There is nothing extra to generate, rotate, or share with the team.
2. **Data stays inside the workspace** — prompts and column metadata never leave your Databricks environment. Using an external API would send table and column names to a third-party endpoint, which may conflict with data governance or security policies.
3. **No additional cost centre** — Foundation Model API usage is billed through your existing Databricks contract, not a separate AI vendor account.

---

## How humanize works

After AI generation, `humanize.py` cleans up descriptions with a rule-based post-processor (no extra LLM call):

- Swaps AI vocabulary for plain English (`bolstered` → `supported`, `leverage` → `use`, etc.)
- Removes filler phrases (`It is worth noting that…`, `highlighting its significance`)
- Fixes unnatural constructions (`serves as` → `is`, `showcasing` → `showing`)

Click **✨** on the table overview or any column row to re-run humanize after a manual edit.

---

## Permissions required

| Privilege | Required for |
|---|---|
| `USE CATALOG` | Browsing schemas |
| `USE SCHEMA` | Browsing tables |
| `SELECT` on the table | Reading column metadata |

Write access (`MODIFY`) is not required — descriptions are applied manually via the exported PySpark script, using whichever account runs the notebook.
