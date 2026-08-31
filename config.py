"""Configuration — warehouse ID required; host + token optional (can be set in the UI).

DATABRICKS_HOST and DATABRICKS_TOKEN in .env pre-seed the Connect form in the sidebar.
If not set, users type them in the UI at runtime.
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

# Pre-seed values for the sidebar Connect form. Empty string = user must type them.
DEFAULT_HOST: str = os.getenv("DATABRICKS_HOST", "").strip()
DEFAULT_TOKEN: str = os.getenv("DATABRICKS_TOKEN", "").strip()

# Serving endpoint for AI generation. Override via DATABRICKS_SERVING_ENDPOINT in .env
# if your workspace uses a different model deployment name.
SERVING_ENDPOINT: str = os.getenv(
    "DATABRICKS_SERVING_ENDPOINT", "databricks-meta-llama-3-3-70b-instruct"
).strip()
