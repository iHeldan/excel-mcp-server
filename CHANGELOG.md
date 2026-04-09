# Changelog

## 0.4.0 - 2026-04-09

This release turns SheetForge from an early MCP workbook helper into a production-ready Excel automation tool for agent workflows.

### Highlights

- Added compact write responses across mutation tools, with `include_changes=True` available for detailed diffs when needed.
- Added `quick_read`, `read_excel_as_table`, and `read_excel_table` support for `row_mode="objects"` and `infer_schema=True`.
- Added ASCII-safe transliteration for object-mode field names, so headers like `Näyttökerrat` become `nayttokerrat`.
- Added `list_tables`, `set_worksheet_visibility`, `set_column_widths`, `set_row_heights`, `autofit_columns`, worksheet protection helpers, and print setup helpers.
- Added `format_ranges` for batch formatting and improved conditional formatting DX.
- Unified chart authoring so `create_chart` now supports both contiguous `data_range` inputs and explicit `series` definitions.
- Fixed pivot grouping bugs around `columns` handling and case-insensitive field resolution.
- Fixed contiguous scatter chart creation so the first Y data point is no longer dropped.
- Expanded regression coverage from the original baseline to 172 passing tests.

### Packaging and Docs

- Added Ruff linting and build smoke checks to CI.
- Synced the manifest, README, and static landing page with the current 47-tool surface.
- Added this changelog for future release notes.
