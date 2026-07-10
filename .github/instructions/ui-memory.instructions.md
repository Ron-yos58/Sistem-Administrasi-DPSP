---
description: Project-specific UI interaction and attribution preferences.
applyTo: "Index.html,Styles.html,Scripts.html"
---

# UI Memory

Keep navigation behavior and product attribution consistent across the DPSP web app.

## Navigation Drawer

Use a hamburger-triggered sidebar drawer on both desktop and mobile. Keep click-outside, backdrop, Escape-key, focus restoration, and ARIA state behavior synchronized.

## Product Credit

Keep the sidebar footer credit visible as "Designed and Developed by Ronald Sebastian" without competing with the user profile or primary navigation.

## Workflow Status Presentation

Treat request lifecycle and document processing as separate concepts:

- Request lifecycle uses `DRAFT`, `READY`, and `ARCHIVED`, displayed as `Draft`, `Siap Diproses`, and `Selesai`.
- Document progress uses one derived indicator: `Belum Dibuat`, `Doc & PDF Dibuat`, `Draft Email Dibuat`, `Gagal Generate`, or `Selesai`.
- Render each concept through its centralized helper. Do not add generic `AKTIF` badges or duplicate status blocks.
- Archived requests are read-only. Hide edit/process actions and require an explicit restore feature before allowing changes.

## Action And Content Ownership

Give every action and information group one clear owner in the interface:

- `Simpan Draft` only persists request data.
- `Proses Permohonan` validates a Draft and changes it to `READY`; it never generates files or Gmail drafts.
- Document, PDF, Gmail draft, Honor, and Perjadin generation remain separate explicit actions.
- Individual document and email actions live only in the document action section.
- Schedule sessions live inside the Detail `Ringkasan Permohonan`; do not add a second standalone session block.
- Successful form saves return to Ringkasan or Permohonan and never open Detail automatically.

## Detail Output Links

Present Honor and Perjadin as separate result groups without a kind selector. For each enabled group, show one explicit action for its dedicated editable spreadsheet file.

- `Buka Spreadsheet` opens the generated finance spreadsheet file, not a worksheet inside the main database spreadsheet.
- Treat that spreadsheet as the only editable finance workspace for Honor and Perjadin values.
- Display the backend completeness message so operators know whether the spreadsheet is already complete.
- Keep persisted links visible for archived requests and after refresh.
- Generated documents expose direct Google Doc and PDF download actions; do not restore a separate draft preview button in Detail.

## Request Listing

Show archived (`Selesai`) requests in the main `Permohonan` list by default so operators can review full operational history in one place. Use the status filter to narrow the list, not to reveal hidden archived items.

## Email Recipient Selection

Treat Master CC as the managed recipient catalog and automatic routing source. Let operators adjust recipients per request through To and CC checkboxes in the request form.

- Show automatic recipients as preselected recommendations and distinguish them from manual additions.
- Keep at least one To recipient before email preview or draft creation.
- Prevent the same email from appearing in both To and CC; To takes precedence.
- Show the final resolved To and CC lists in review and email preview before creating a Gmail draft.
- Store recipient choices as structured request data rather than embedding metadata in a human-readable notes field.
