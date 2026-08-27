"""Unity Catalog write operations — apply column comments via SQL.

Browse operations (list catalogs/schemas/tables/columns) are handled by
databricks_client.py via the UC REST API.

COMMENT ON COLUMN is DDL that must run through a SQL warehouse — that's why
this module still uses Statement Execution rather than the REST API.
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
