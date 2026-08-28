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

BATCH_SIZE = 15  # columns per API call — conservative default; auto-halves on truncation

_SYSTEM_PROMPT = """You are a data catalog documentation expert for a clickstream and
digital behavioral analytics team. The data captures how customers interact with digital
experiences across web and mobile applications, including page views, searches, clicks,
product interactions, add-to-cart actions, orders, sessions, and visits.

Tables follow a fact/dimension structure:
- Fact tables contain event-level or aggregated behavioral metrics.
- Dimension tables provide descriptive context such as page, component, scenario,
  channel, device platform, and application details.

Rules for generating column descriptions:
- Be concise and business-friendly. Target a data catalog audience, not engineers.
- Tailor each description to the table's grain (event-level, session-level, etc.).
- Explain what the column represents and how it is used.
- Classify the column type when evident: business identifier, surrogate key,
  timestamp, date, metric, flag (explain Y/N values), status, or technical audit field.
- For categorical columns, list known valid values when they can be inferred
  (e.g. WEB, MOBILE, RETAIL, PHARMACY).
- Do not make assumptions when the meaning cannot be confidently determined from
  the column name, table schema, or available metadata — write a neutral description
  or omit speculation.
- Never include PII field names or example values in descriptions.
- Do not reference internal system names or implementation details.
- Do not repeat the column name or data type in the description.
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


def _call_llm_resilient(ws: WorkspaceClient, full_name: str, columns: list[dict]) -> dict[str, str]:
    """Call the LLM for a batch; if JSON is truncated, split in half and retry."""
    try:
        return _call_llm(ws, full_name, columns)
    except RuntimeError:
        if len(columns) <= 1:
            raise  # can't split further — surface the error
        mid = len(columns) // 2
        left = _call_llm_resilient(ws, full_name, columns[:mid])
        right = _call_llm_resilient(ws, full_name, columns[mid:])
        return {**left, **right}


def generate_table_description(
    ws: WorkspaceClient,
    full_name: str,
    columns: list[dict],
    current_comment: str = "",
) -> str:
    """Generate a 2–3 sentence business overview for the table. Returns plain text."""
    col_names = ", ".join(c["name"] for c in columns[:30])
    user_msg = (
        f"Write a concise 2-3 sentence business description for the table `{full_name}`.\n\n"
        f"Column names (sample): {col_names}\n"
        + (f"Existing description: {current_comment}\n" if current_comment else "")
        + "\nDescribe what the table contains, its grain, and its primary use. "
        "Be specific to clickstream and digital behavioral analytics. "
        "Return only the description text — no JSON, no bullet points, no formatting."
    )
    response = ws.serving_endpoints.query(
        name=config.SERVING_ENDPOINT,
        messages=[
            ChatMessage(role=ChatMessageRole.SYSTEM, content=_SYSTEM_PROMPT),
            ChatMessage(role=ChatMessageRole.USER, content=user_msg),
        ],
        temperature=0.2,
        max_tokens=300,
    )
    return response.choices[0].message.content.strip()


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
        results.update(_call_llm_resilient(ws, full_name, batch))
    return results
