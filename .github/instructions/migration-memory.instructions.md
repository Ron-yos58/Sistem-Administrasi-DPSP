---
description: Project-specific procedures and safeguards for executing data migrations from legacy spreadsheets.
applyTo: "Migration.gs,docs/**/*.md"
---

# Migration Memory

Ensure consistent consolidation of legacy request records and metadata without data loss.

## Manual Pre-Migration Backup

Always create a full copy of the Google Spreadsheet manually before running any migration or cleanup command. Keep historical backups in hidden sheets inside the spreadsheet.

## Request Consolidation and Fingerprints

Group multiple legacy rows by their natural fingerprint (tipe kegiatan, narasumber, mitra, tanggal kegiatan, and tempat) to merge duplicates into a single request with multiple document entries in the `DOCUMENTS` sheet.

## Date and Time Robustness

Parse legacy Indonesian date ranges (e.g. using `s/d`, `sampai`, or full ranges with years in both dates) into start and end ISO dates before storing them. Verify date format validation matches master spreadsheet constraints.

## Financial Data Sync Safeguards

Ensure synchronization functions (like `syncTravelDataInternal_`) preserve cost data using stable participant keys. Avoid deleting archived requests or historical travel records during a full sync.

## Verification Checklist

Always execute `previewLegacyMigration` from the Admin panel and review the execution summary and blockers list before running the migration with the `MIGRATE` confirmation code.
