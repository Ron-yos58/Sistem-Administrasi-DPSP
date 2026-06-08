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
- Document progress uses one derived indicator: `Belum Dibuat`, `Doc & PDF Dibuat`, `Draft Email Dibuat`, `Gagal Generate`, or `Selesai / Diarsipkan`.
- Render each concept through its centralized helper. Do not add generic `AKTIF` badges or duplicate status blocks.
- Archived requests are read-only. Hide edit/process actions and require an explicit restore feature before allowing changes.

## Action And Content Ownership

Give every action and information group one clear owner in the interface:

- Batch processing is labeled `Buat semua Doc, PDF & draft Gmail`.
- Individual document actions live only in `Proses satu dokumen`.
- Schedule sessions live inside the Detail `Ringkasan Permohonan`; do not add a second standalone session block.
- Successful form saves return to Ringkasan or Permohonan and never open Detail automatically.
