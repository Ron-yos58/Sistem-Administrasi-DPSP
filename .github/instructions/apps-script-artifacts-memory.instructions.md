---
description: Project-specific persistence and serialization patterns for generated Apps Script outputs.
applyTo: "**/*.gs,Scripts.html"
---

# Apps Script Artifacts Memory

Keep spreadsheet values and generated output links stable across refreshes and regeneration.

## Spreadsheet Date And Time Boundaries

Normalize Spreadsheet `Date` and time-only values when building DTOs. Convert time-only cells to `HH:mm` before any generic string conversion so the serial date `30 Dec 1899` never reaches templates or the frontend.

## Generated Finance Registry

Store Honor and Perjadin outputs in `Generated Files` with:

- Artifact keys `FINANCE_HONOR` and `FINANCE_PERJADIN`.
- Types `SHEET`, `PDF`, and `XLSX`.
- Exactly one `ACTIVE` row per request, artifact key, and type; mark the previous row `SUPERSEDED`.
- Google Sheet editor URLs only for `SHEET`.
- Direct Drive download URLs derived from file IDs for `PDF` and `XLSX`.

Invalidate related active artifacts when source request or Perjadin cost data changes. Detail responses derive the six named finance URLs from active registry rows and verify Drive files still exist.

## Explicit Generation Workflow

Keep lifecycle transitions separate from artifact generation:

- Saving a Draft persists data only.
- Activating a request validates stored data and changes status only.
- Generate Docs, PDFs, Gmail drafts, and finance exports only from explicit user actions.
- Avoid automatic finance synchronization during application bootstrap or ordinary page navigation.

## Content-Based Invalidation

Invalidate generated artifacts only when generation-relevant request content changes. Status-only transitions and unchanged saves preserve document IDs, PDF IDs, Gmail draft IDs, and active registry entries.

Finance exports must pass server-side completeness checks. Export the existing editable generated sheet without regenerating it immediately before export, because regeneration would discard user-entered finance values.
