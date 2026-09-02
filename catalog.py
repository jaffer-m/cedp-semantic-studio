"""Unity Catalog write operations — apply column comments via SQL.

Browse operations (list catalogs/schemas/tables/columns) are handled by
databricks_client.py via the UC REST API.

COMMENT ON COLUMN is DDL that must run through a SQL warehouse — that's why
this module still uses Statement Execution rather than the REST API.

list_tables_via_sql is a browse-op fallback: some schemas' tables are visible
to a SQL warehouse (information_schema) but not to the UC Tables REST API for
the same identity — a platform-side authorization inconsistency, not
something the caller can avoid. app.py falls back to it when the REST API
returns no tables for a schema.
"""

import re

from databricks.sdk import WorkspaceClient
from databricks.sdk.service.sql import StatementState

import config


def _run_sql(ws: WorkspaceClient, sql: str) -> None:
    resp = ws.statement_execution.execute_statement(
        warehouse_id=config.WAREHOUSE_ID,
        statement=sql,
        wait_timeout="50s",
    )
    if not resp.status or resp.status.state != StatementState.SUCCEEDED:
        msg = (
            resp.status.error.message
            if resp.status and resp.status.error
            else "unknown error"
        )
        raise RuntimeError(f"SQL failed: {msg}")


def _run_query(ws: WorkspaceClient, sql: str) -> list[list]:
    resp = ws.statement_execution.execute_statement(
        warehouse_id=config.WAREHOUSE_ID,
        statement=sql,
        wait_timeout="50s",
    )
    if not resp.status or resp.status.state != StatementState.SUCCEEDED:
        msg = (
            resp.status.error.message
            if resp.status and resp.status.error
            else "unknown error"
        )
        raise RuntimeError(f"SQL failed: {msg}")
    return resp.result.data_array if resp.result and resp.result.data_array else []


def _validate(name: str) -> None:
    if not re.match(r'^[a-zA-Z0-9_.`\- ]+$', name):
        raise ValueError(f"Invalid identifier: {name!r}")


def _quote(name: str) -> str:
    parts = name.split(".")
    return ".".join(f"`{p.strip('`')}`" for p in parts)


def _escape(text: str) -> str:
    return text.replace("'", "\\'")


_PERM_SIGNALS = [
    ("does not have USE CATALOG", "USE CATALOG", "CATALOG"),
    ("does not have USE SCHEMA",  "USE SCHEMA",  "SCHEMA"),
    ("does not have SELECT",      "SELECT",      "TABLE"),
    ("does not have MODIFY",      "MODIFY",      "TABLE"),
]


def _check_permission_error(msg: str, target: str) -> None:
    msg_upper = msg.upper()
    for signal, privilege, object_type in _PERM_SIGNALS:
        if signal.upper() in msg_upper:
            grant_sql = f"GRANT {privilege} ON {object_type} {target} TO `<user-or-group>`;"
            raise PermissionError(
                f"Missing privilege: {privilege} on {target}\n\n"
                f"Ask your Databricks admin to run:\n  {grant_sql}\n\n"
                f"Raw error: {msg}"
            )


def apply_column_comment(
    ws: WorkspaceClient,
    full_name: str,
    col_name: str,
    comment: str,
) -> None:
    """Write a column comment using COMMENT ON COLUMN.

    Works for Delta tables, DLT streaming tables, and materialized views.
    Requires a WorkspaceClient connected with the user's credentials.
    """
    _validate(full_name)
    col_safe = col_name.replace("`", "``")
    escaped = _escape(comment)
    try:
        _run_sql(
            ws,
            f"COMMENT ON COLUMN {_quote(full_name)}.`{col_safe}` IS '{escaped}'",
        )
    except RuntimeError as e:
        _check_permission_error(str(e), f"TABLE {full_name}")
        raise


def list_tables_via_sql(ws: WorkspaceClient, catalog_name: str, schema: str) -> list[dict]:
    """Fallback table listing via information_schema, for schemas where the
    UC Tables REST API returns nothing despite the SQL warehouse having access.
    """
    _validate(catalog_name)
    _validate(schema)
    rows = _run_query(
        ws,
        "SELECT table_name, table_type, comment "
        f"FROM {_quote(catalog_name)}.information_schema.tables "
        f"WHERE table_schema = '{_escape(schema)}'",
    )
    result = []
    for name, table_type, comment in rows:
        table_type = table_type or ""
        if table_type == "STREAMING_TABLE":
            type_label = "DLT Streaming"
        elif table_type == "MATERIALIZED_VIEW":
            type_label = "Materialized View"
        else:
            type_label = "Delta Table"
        result.append({
            "name": name or "",
            "full_name": f"{catalog_name}.{schema}.{name}",
            "table_type": table_type,
            "type_label": type_label,
            "comment": comment or "",
        })
    return sorted(result, key=lambda x: x["name"].lower())
