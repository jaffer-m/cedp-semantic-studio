"""Configuration — only the warehouse ID is required.

Databricks host and token are resolved automatically from ~/.databrickscfg
by the SDK. No API keys or model names needed.
"""

import os
from dotenv import load_dotenv

load_dotenv()

_warehouse = os.getenv("DATABRICKS_WAREHOUSE_ID", "").strip()
if not _warehouse:
    raise EnvironmentError(
        "Missing DATABRICKS_WAREHOUSE_ID.\n"
        "Copy .env.example to .env and set your warehouse ID."
    )

WAREHOUSE_ID: str = _warehouse

# Default serving endpoint — uses Databricks Foundation Model APIs built into the workspace.
# No separate key or host required; the SDK handles auth automatically.
SERVING_ENDPOINT: str = "databricks-meta-llama-3-3-70b-instruct"
