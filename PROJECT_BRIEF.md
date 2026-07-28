# Project Brief: Sistem Surat DPSP

## 1. Project Overview

Website Google Apps Script untuk pengelolaan permohonan surat, dokumen, draft email, honor, dan perjalanan dinas DPSP.

## 2. Concept / Product Description

Operator mengisi satu permohonan, memilih satu atau beberapa dokumen, memeriksa penerima, membuat Google Doc/PDF, lalu membuat draft Gmail. Semua operasi terlacak dan dapat dijalankan ulang tanpa duplikasi normal.

## 3. Tech Stack

- Google Apps Script V8
- HtmlService
- Google Sheets sebagai datastore
- Google Docs/Drive untuk dokumen dan PDF
- GmailApp untuk draft
- HTML, CSS, dan JavaScript tanpa framework

## 4. Architecture

```text
Browser HtmlService
        |
 google.script.run
        |
Apps Script services
  |-- DataService
  |-- DocumentService
  |-- EmailService
  |-- FinanceService
        |
Google Sheets / Docs / Drive / Gmail
```

## 5. Key Files Map

- `Config.gs`: ID, enum, header
- `Setup.gs`: bootstrap sheet
- `DataService.gs`: CRUD dan routing
- `DocumentService.gs`: Docs dan PDF
- `EmailService.gs`: preview dan draft
- `FinanceService.gs`: perjadin dan sheet keuangan
- `ExportService.gs`: PDF dan XLSX satu-sheet
- `Index.html`, `Styles.html`, `Scripts.html`: frontend

## 6. Team Roles

- Remy: scope dan integrasi
- Kira: alur produk
- Milo: sistem visual dan aksesibilitas
- Nova: frontend dan state
- Sage: data, backend, security
- Ivy: risiko dan E2E

## 7. Sprint Status

Sprint 1 selesai: audit, arsitektur, implementasi MVP penuh, dokumentasi, static check.

## 8. Current State

Kode siap diunggah. Runtime Google Workspace belum dapat diuji dari workspace lokal karena memerlukan otorisasi akun dan data Sheet aktual.

## 9. Security Rules

- Domain-only deployment.
- Execute as user accessing.
- Whitelist `Config_Access`.
- Role diverifikasi server-side.
- Tidak ada pengiriman email otomatis.
- Penghapusan hanya artefak bertag sistem.

## 10. How to Run Locally

Frontend GAS tidak mempunyai runtime lokal penuh. Pemeriksaan statis:

```powershell
npm run check
```

## 11. How to Deploy

Ikuti `PANDUAN_IMPLEMENTASI.md`.

## 12. Cross-Chat Handoff Protocol

Chat lanjutan membaca `PROJECT_BRIEF.md`, `ANALISIS_SISTEM.md`, dan `docs/sprint-1/done.md` sebelum mengubah kode.

## 13. Bug & Fix Tracking

Temuan dan keputusan disimpan pada `ANALISIS_SISTEM.md`. Bug runtime baru harus mencatat fungsi, ID permohonan, user, timestamp, dan output `Audit Log`.

## 14. Multi-Repo Setup

Project saat ini satu workspace dan belum menjadi Git repository. Bila Git dibuat, gunakan branch per sprint dan jangan memasukkan `.clasp.json` berisi Script ID ke repository publik.
