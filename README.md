# USC Asset Disposal & Reconciliation Dashboard

Baker Tilly Third-Party Validation (TPV) dashboard for Utility Stores Corporation's
asset-disposal / auction reconciliation programme.

**Start here → [`PROGRESS.md`](./PROGRESS.md)** — the resumable build checkpoint: goal,
locked decisions, source-file map, the `#REF!` fix spec, the full data model, and the
remaining build plan.

## Layout
- `PROGRESS.md` — build status & spec (read this first).
- `source-data/` — extracted, machine-readable data pulled from the Drive source files
  (`consolidation_data.json`, `project_updates.json`) + human-readable summaries.
- `base/index_base_shell.html` — the clean, complete dashboard app shell used as the build base.

The final dashboard (`USC_Dashboard_v2.html`) is generated from the Drive source workbook and
served via GitHub Pages for a no-login shareable link; a copy is saved back to the Drive folder.
