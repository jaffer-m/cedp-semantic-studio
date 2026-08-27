"""Unity Catalog operations — browse and update column descriptions.

Uses SQL Statement Execution (not the UC REST API) so it works with a plain
personal access token. All apply operations use COMMENT ON COLUMN syntax,
which works for Delta tables, DLT streaming tables, and materialized views.
Auth is resolved automatically from ~/.databrickscfg by the SDK.
"""

import re
from databricks.sdk import WorkspaceClient
from databricks.sdk.service.sql import StatementState, StatementParameterListItem

import config

_client: WorkspaceClient | None = None


def _get_client() -> WorkspaceClient:
    global _client
    if _client is None:
        _client = WorkspaceClient()  # reads ~/.databrickscfg automatically
    return _client


def _run_sql(sql: str, params: list | None = None) -> list[dict]:
    w = _get_client()
    kwargs: dict = dict(
        warehouse_id=config.WAREHOUSE_ID,
        statement=sql,
        wait_timeout="50s",
    )
    if params:
        kwargs["parameters"] = params

    resp = w.statement_execution.execute_statement(**kwargs)

    if not resp.status or resp.status.state != StatementState.SUCCEEDED:
        msg = (
            resp.status.error.message
            if resp.status and resp.status.error
            else "unknown error"
        )
        raise RuntimeError(f"SQL failed: {msg}")

    if not resp.result or not resp.result.data_array:
        return []

    cols = [c.name for c in resp.manifest.schema.columns]
    return [dict(zip(cols, row)) for row in resp.result.data_array]


def _validate(name: str) -> None:
    if not re.match(r'^[a-zA-Z0-9_.`\- ]+$', name):
        raise ValueError(f"Invalid identifier: {name!r}")


def _quote(name: str) -> str:
    parts = name.split(".")
    return ".".join(f"`{p.strip('`')}`" for p in parts)


def _escape(text: str) -> str:
    return text.replace("'", "\\'")


# ── Browse ────────────────────────────────────────────────────────────────

def list_catalogs() -> list[str]:
    rows = _run_sql("SHOW CATALOGS")
    return sorted(r["catalog"] for r in rows)


def list_schemas(catalog: str) -> list[str]:
    _validate(catalog)
    rows = _run_sql(
        f"SELECT schema_name FROM `{catalog}`.information_schema.schemata"
    )
    return sorted(r["schema_name"] for r in rows)


def list_tables(catalog: str, schema: str) -> list[dict]:
    _validate(f"{catalog}.{schema}")
    rows = _run_sql(
        f"SELECT table_name, table_type, comment "
        f"FROM `{catalog}`.information_schema.tables "
        f"WHERE table_schema = :schema",
        params=[StatementParameterListItem(name="schema", value=schema)],
    )
    result = []
    for r in rows:
        table_type = r.get("table_type") or ""
        if table_type == "STREAMING_TABLE":
            type_label = "DLT Streaming"
        elif table_type == "MATERIALIZED_VIEW":
            type_label = "Materialized View"
        else:
            type_label = "Delta Table"
        result.append({
            "name": r["table_name"],
            "full_name": f"{catalog}.{schema}.{r['table_name']}",
            "table_type": table_type,
            "type_label": type_label,
            "comment": r.get("comment") or "",
        })
    return sorted(result, key=lambda x: x["name"].lower())


def get_columns(full_name: str) -> list[dict]:
    _validate(full_name)
    parts = full_name.split(".")
    if len(parts) != 3:
        raise ValueError(f"Expected catalog.schema.table, got: {full_name}")
    catalog, schema, table = parts

    rows = _run_sql(
        f"SELECT column_name, data_type, is_nullable, comment "
        f"FROM `{catalog}`.information_schema.columns "
        f"WHERE table_schema = :schema AND table_name = :table "
        f"ORDER BY ordinal_position",
        params=[
            StatementParameterListItem(name="schema", value=schema),
            StatementParameterListItem(name="table", value=table),
        ],
    )
    return [
        {
            "name": r["column_name"],
            "type": r.get("data_type") or "",
            "nullable": (r.get("is_nullable") or "YES").upper() == "YES",
            "current_comment": r.get("comment") or "",
        }
        for r in rows
    ]


# ── Apply ─────────────────────────────────────────────────────────────────

def apply_column_comment(full_name: str, col_name: str, comment: str) -> None:
    """Write a column comment using COMMENT ON COLUMN.

    Works for Delta tables, DLT streaming tables, and materialized views.
    """
    _validate(full_name)
    col_safe = col_name.replace("`", "``")
    escaped = _escape(comment)
    _run_sql(
        f"COMMENT ON COLUMN {_quote(full_name)}.`{col_safe}` IS '{escaped}'"
    )
