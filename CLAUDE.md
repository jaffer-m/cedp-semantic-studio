# CLAUDE.md

## Table & column description generation

When writing, reviewing, or tuning the AI description-generation logic in [ai_gen.py](ai_gen.py) (the system prompt, batching, or output rules), consult these Confluence pages for best-practice guidance before making changes:

- [What Data Should I Prioritize Cataloging](https://kroger.atlassian.net/wiki/spaces/DATACO/pages/1037271307/What+Data+Should+I+Prioritize+Cataloging) — guidance on which tables/columns are worth cataloging first.
- [Agentic and Automated Metadata Management Tools](https://kroger.atlassian.net/wiki/spaces/DATACO/pages/1486291297/Agentic+and+Automated+Metadata+Management+Tools) — org-wide standards for AI/automated metadata generation.
- [UCMeG: UC Metadata Generator (AI-Powered)](https://kroger.atlassian.net/wiki/spaces/DIENG/pages/1363410953/UCMeG+UC+Metadata+Generator+AI-Powered) — reference implementation/approach for AI-powered Unity Catalog metadata generation.

These pages require Kroger SSO and can't be fetched automatically — read them directly when updating description-generation behavior, and reconcile the `_SYSTEM_PROMPT` in [ai_gen.py](ai_gen.py) against them.
