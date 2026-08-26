# SpreadsheetTool — Architecture

Technical document for the SpreadsheetTool plugin: how it works, the **area-receipt**
technique every cell-writing method must follow, and the project rule that keeps it
consistent for future methods.

## Overview

SpreadsheetTool is a plugin (implements `AIOrchestrator.API.IAgentTool`) that exposes
spreadsheet operations to LLM agents. The public methods become LLM-callable tools via
the `GeToolDefinitions` reflection pipeline, which reads the XML documentation file.

The implementation is built on the **Aspose.Cells_FOSS** fork (MIT). The fork has no
calculation engine, no auto-fit API and no rendering: everything the agent needs that
the fork cannot do is provided deterministically by the tool itself (see
"Deterministic assists" below).

The agent works on a **sandbox**: the tool receives agent-facing paths
(`/folder/file.xlsx`) and resolves them with `SandboxPath.Resolve` to host paths. Every
result that names a file returns the **sandbox-relative path** (`SandboxPath.ToAgent`),
never a bare file name and never a host path — the agent needs it to attach the file via
its `done` method.

## The area-receipt technique

### Problem

The write methods (`set_cell_value`, `set_range`, `append_rows`, `set_cell_formula`,
`apply_style`, `format_header_row`, `set_page_setup`) used to answer `"true"`. That is a
**zero-information** result: the agent must trust its own memory of what it sent. When
that memory drifts, the agent silently overwrites cells, skips rows or duplicates whole
blocks without ever noticing — and the `done` message may then claim things that were
never written.

### Solution

Every method that creates or modifies cells returns an **area receipt**: a compact,
deterministic snapshot of the exact area it touched, produced by `DescribeArea`:

```json
{
  "sheet": "Coffee Shop Data",
  "range": "A1:C8",
  "rows": 8,
  "columns": 3,
  "values": [["Month", "Revenue ($)", "Costs ($)"], ["January", "8500", "6200"], ...],
  "formulas": { "B8": "=SUM(B2:B7)" },
  "types":    { "A1": "text" },
  "formats":  { "B2": "$#,##0.00" }
}
```

- `sheet` — the worksheet name (enables a working memory across multiple sheets);
- `range` — the exact **A1 position on the page** of the modified area;
- `rows` / `columns` — the area dimensions;
- `values` — the **display values** (number formats applied), so the agent sees what a
  human would see;
- `formulas` / `types` / `formats` — **SPARSE** sections that list ONLY the cells that
  carry that information. Cells not listed in `types` are numeric. Empty sections are
  omitted entirely, keeping the token cost low.

### Purpose

- **Feedback / memory of work**: after every write the agent has ground truth about what
  was actually stored (including formatted values and formulas), instead of `"true"`.
- **Self-verification**: the receipt lets the agent check for missing rows, duplicates
  or overwrites at the moment they happen — without a separate read round-trip.
- **Deterministic, zero extra LLM cost**: the receipt is built from the in-memory
  workbook by the tool itself; it costs only the few tokens of the JSON.

### Contract

- Every method that **writes cells or applies settings to cells** returns the area
  receipt of the modified area (single cell, block, styled range, header row, page
  setup), or an `"Error: ..."` string on invalid input — never a bare `"true"`.
- Read methods may use the same shape: `GetRange(detailed=true)` returns the same
  sparse JSON for an arbitrary area, so the agent can "zoom into" any region.
- The XML `<returns>` documentation of every write method states explicitly that the
  result is the receipt/feedback for the agent to verify its own work.

## 🗒 Project rule (MEMO) — future cell-writing methods

> **Every future method of this tool that works on cells MUST return the area receipt
> of the area it modified** — via `DescribeArea(...)` — exactly like the existing write
> methods. It must NEVER return a bare `"true"`, an empty string, or a plain count.
>
> - Call `DescribeArea(sheet, startRow, startCol, endRow, endCol)` on the modified
>   range and return its result (or a dictionary with `sheet` + `range` +
>   `applied:[...]` for settings-only methods like `apply_style` / `set_page_setup`).
> - Keep the sections **sparse**: include `formulas` / `types` / `formats` only when at
>   least one cell carries them; omit empty sections.
> - The receipt must identify the position on the page (`range` in A1) and the
>   `sheet` name, so the agent can maintain a working memory across sheets.
> - Document the contract in the method's XML `<returns>`.
>
> This preserves the feedback loop the tool provides to agents and prevents silent
> write errors from ever becoming invisible again.

## Deterministic assists (save-time pass)

At every persist (`Save`, `SaveAs`, `Dispose` auto-save) the tool applies a
deterministic, agent-invisible auto-format pass before validating and committing the
file:

- **bestFit columns**: columns the agent touched that still carry the default width are
  written with `bestFit="1"` and NO width value, so the OPENING APPLICATION auto-fits
  them (formula results included — the fork cannot compute them). User-set widths are
  never touched (the user is always right).
- **Table titles**: text cells forming a contiguous vertical run directly above a
  number/formula cell are styled bold + light pastel background, one palette color per
  table. Explicit user/agent styles are never overridden.
- **Format normalization**: agent-supplied number formats with surrounding quotes
  (`"$#,##0.00"`) are normalized — a quoted string would render as literal text in
  Excel format syntax.

The pass runs on a temp file, the XML parts are parsed for well-formedness, and only a
valid result replaces the real file (the on-disk workbook is never replaced by an
invalid one).

## Persistence & delivery

- The workbook is saved through `PersistValidated` (temp → validate → atomic replace),
  then a git snapshot is taken with `GitSupport.Snapshot` for `GitTool.restore`
  rollback.
- `Dispose()` auto-saves pending changes (the agent may forget to call `save`) — the
  delivery of the file to the user must never see a stale/empty on-disk version.
- Every create/save result returns the **sandbox-relative path** so the agent can
  attach the file via its `done` `attachments` field.

## Tool surface

The complete method catalog is generated by reflection from the public methods + the
XML docs. Verify the surface with the harness:

```
dotnet run --project SpreadsheetTool.Tests -- --methods
dotnet run --project SpreadsheetTool.Tests -- --autofmt      # deterministic auto-format + receipts
dotnet run --project SpreadsheetTool.Tests -- --charttest    # chart caches + bestFit regression
```
