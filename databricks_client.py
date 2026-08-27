"""Databricks Unity Catalog client — REST API browsing, no SQL warehouse needed.

Browse operations (catalogs, schemas, tables, columns) use the UC REST API via
the Databricks SDK. This avoids warehouse dependency for browsing and handles
PermissionDenied gracefully per call instead of crashing.

Write operations (apply_column_comment) remain in catalog.py which uses SQL
Statement Execution — COMMENT ON COLUMN requires a warehouse.
"""

from __future__ import annotations

from typing import Any

from databricks.sdk import WorkspaceClient
from databricks.sdk.errors import NotFound, PermissionDenied, ResourceDoesNotExist


class DatabricksClient:
    """Thin wrapper around the Databricks SDK for Unity Catalog browsing."""

    def __init__(self, host: str, token: str) -> None:
        self._ws = WorkspaceClient(host=host.rstrip("/"), token=token)

    def test_connection(self) -> str:
        """Return the current user's display name. Raises on auth failure."""
        me = self._ws.current_user.me()
        return me.display_name or me.user_name or "unknown"

    # ── Browse ────────────────────────────────────────────────────────────

    def list_catalogs(self) -> list[str]:
        try:
            return sorted(c.name for c in self._ws.catalogs.list() if c.name)
        except PermissionDenied:
            return []

    def list_schemas(self, catalog: str) -> list[str]:
        try:
            return sorted(
                s.name
                for s in self._ws.schemas.list(catalog_name=catalog)
                if s.name
            )
        except (PermissionDenied, NotFound, ResourceDoesNotExist):
            return []

    def list_tables(self, catalog: str, schema: str) -> list[dict]:
        try:
            result = []
            for t in self._ws.tables.list(catalog_name=catalog, schema_name=schema):
                table_type = t.table_type.value if t.table_type else ""
                if table_type == "STREAMING_TABLE":
                    type_label = "DLT Streaming"
                elif table_type == "MATERIALIZED_VIEW":
                    type_label = "Materialized View"
                else:
                    type_label = "Delta Table"
                result.append({
                    "name": t.name or "",
                    "full_name": t.full_name or f"{catalog}.{schema}.{t.name}",
                    "table_type": table_type,
                    "type_label": type_label,
                    "comment": t.comment or "",
                })
            return sorted(result, key=lambda x: x["name"].lower())
        except (PermissionDenied, NotFound, ResourceDoesNotExist):
            return []

    def get_columns(self, full_name: str) -> list[dict]:
        """Fetch column metadata for a table via the UC REST API."""
        try:
            t = self._ws.tables.get(full_name=full_name)
        except (NotFound, ResourceDoesNotExist) as e:
            raise ValueError(f"Table not found: {full_name}") from e
        except PermissionDenied as e:
            raise PermissionError(
                f"Missing SELECT privilege on TABLE {full_name}\n\n"
                f"Ask your Databricks admin to run:\n"
                f"  GRANT SELECT ON TABLE {full_name} TO `<user-or-group>`;\n\n"
                f"Raw error: {e}"
            ) from e

        return [
            {
                "name": col.name or "",
                "type": col.type_text or (col.type_name.value if col.type_name else "UNKNOWN"),
                "nullable": col.nullable if col.nullable is not None else True,
                "current_comment": col.comment or "",
            }
            for col in (t.columns or [])
        ]
