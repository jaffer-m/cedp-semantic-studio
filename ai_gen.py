"""AI generation of column descriptions via Databricks Foundation Model APIs.

Uses the Databricks SDK serving endpoint query — no separate API key, host,
or model name configuration needed. Auth comes from ~/.databrickscfg.

Columns are processed in batches of BATCH_SIZE to stay within the endpoint's
output token limit (empirically ~2500 tokens per call).
"""

import json
import logging
import re

from databricks.sdk import WorkspaceClient
from databricks.sdk.service.serving import ChatMessage, ChatMessageRole

import config

logger = logging.getLogger(__name__)

BATCH_SIZE = 25  # columns per API call — keeps output well under token limits

_SYSTEM_PROMPT = """You are a data catalog documentation expert. Generate specific, detailed,
business-friendly descriptions for database table columns.

Rules:
- Explain what each column represents in business terms.
- For categorical columns, list known valid values (e.g. RETAIL, PHARMACY, GAS STATION).
- For flag columns, explain what Y and N mean.
- For date columns calculated at query time, warn that the value should not be stored.
- Use plain English. Be specific — vague descriptions like "indicates the type" are not useful.
- Never include PII field names or example values in descriptions.
- Use business-friendly language suitable for a data catalog audience.
- Do not reference internal system names or implementation details.
- Do not include the column name or data type in the description.
- Return valid JSON only. No markdown fences, no extra text.

Output format:
{
  "column_descriptions": {
    "<column_name>": "<description>",
    ...
  }
}"""


def _call_llm(ws: WorkspaceClient, full_name: str, columns: list[dict]) -> dict[str, str]:
    """Single API call for a batch of columns. Returns {col_name: description}."""
    col_lines = "\n".join(
        f"  - {c['name']} ({c['type']})"
        + (f": currently '{c['current_comment']}'" if c.get("current_comment") else "")
        for c in columns
    )
    user_msg = (
        f"Generate column descriptions for the table `{full_name}`.\n\n"
        f"Columns:\n{col_lines}\n\n"
        f"Return JSON with column_descriptions for every column listed above."
    )

    response = ws.serving_endpoints.query(
        name=config.SERVING_ENDPOINT,
        messages=[
            ChatMessage(role=ChatMessageRole.SYSTEM, content=_SYSTEM_PROMPT),
            ChatMessage(role=ChatMessageRole.USER, content=user_msg),
        ],
        temperature=0.2,
        max_tokens=4096,
    )

    content = response.choices[0].message.content.strip()
    if content.startswith("```"):
        content = re.sub(r'^```[a-z]*\n?', '', content)
        content = re.sub(r'\n?```$', '', content)

    try:
        parsed = json.loads(content)
        return parsed.get("column_descriptions", {})
    except json.JSONDecodeError as e:
        logger.error("LLM returned non-JSON for batch ending at %s: %s",
                     columns[-1]["name"], content[:300])
        raise RuntimeError(
            f"LLM response was not valid JSON (batch ending at '{columns[-1]['name']}'): {e}"
        ) from e


def generate_column_descriptions(
    ws: WorkspaceClient,
    full_name: str,
    columns: list[dict],
) -> dict[str, str]:
    """Generate descriptions for all columns, batching to stay within token limits.

    Args:
        ws: authenticated WorkspaceClient from the Connect form session
        full_name: catalog.schema.table
        columns: list of {name, type, nullable, current_comment}

    Returns:
        {column_name: description} dict
    """
    results: dict[str, str] = {}
    batches = [columns[i:i + BATCH_SIZE] for i in range(0, len(columns), BATCH_SIZE)]
    for batch in batches:
        results.update(_call_llm(ws, full_name, batch))
    return results
