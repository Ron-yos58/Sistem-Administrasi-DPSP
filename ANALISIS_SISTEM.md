# Analisis Sistem Surat DPSP

## Keputusan Produk

- Entitas utama adalah `Permohonan`, bukan nomor baris atau sheet hasil.
- Satu permohonan dapat mempunyai beberapa dokumen.
- Email hanya dibuat sebagai draft Gmail. Sistem tidak mengirim otomatis.
- Data dihapus secara lunak dengan status `ARCHIVED`.
- Sheet keluaran honor/perjadin hanya dapat dihapus bila memiliki prefix dan metadata sistem.
- Konflik panduan `Edu Fair + Surat Balasan Campus Visit` diputuskan mengikuti tabel mapping: Edu Fair hanya menghasilkan `Surat Tugas`.

## Menu

1. Ringkasan
2. Permohonan
3. Buat Permohonan
4. Honor & Perjadin
5. Data Referensi

## Alur Kerja

1. Operator membuat draft permohonan.
2. Sistem memvalidasi kombinasi kegiatan, dokumen, penerima tugas, dan tanggal.
3. Operator menandai data `READY`.
4. Sistem menampilkan preview penerima To/CC per dokumen.
5. Sistem membuat Google Doc dan PDF secara idempotent.
6. Sistem membuat draft Gmail dengan PDF terlampir.
7. Operator meninjau dan mengirim email langsung dari Gmail.

## Perbaikan Atas Sistem Lama

| Risiko lama | Tindakan baru |
| --- | --- |
| Indeks kolom literal dan komentar tidak cocok | Header dan konstanta terpusat |
| `Data Pegawai` vs `Database Pegawai` | Alias didukung, nama kanonik `Data Pegawai` |
| Email pegawai/mitra tercampur | Routing dihitung per dokumen |
| Alamat Gmail dipisahkan newline | Array email divalidasi, dideduplikasi, lalu dipisahkan koma |
| Nominal perjadin berpindah saat urutan berubah | `Participant Key` stabil |
| Dokumen/draft ganda | `ScriptLock`, revision, file ID, draft ID, marker rekonsiliasi |
| `Log_Sheet` mempunyai banyak arti | `Audit Log` dan `Generated Files` terpisah |
| Delete menghapus semua sheet di luar allowlist | Hanya sheet `GEN-*` dengan developer metadata |
| Helper global bertabrakan | Nama helper tunggal dan fungsi internal bersufiks `_` |
| Bergantung pada UI Spreadsheet | API server untuk `google.script.run` |

## Struktur Data Baru

- `Master Permohonan`: satu baris per ID.
- `Dokumen Permohonan`: satu baris per dokumen.
- `Jadwal Kegiatan`: satu atau lebih sesi tanggal, waktu, dan tempat per ID.
- `Data Pegawai`: banyak orang per ID, dengan `Participant Key`.
- `Data Perjadin`: biaya per orang berdasarkan `Participant Key`.
- `Master CC`: sumber penerima internal.
- `Template`: subject dan body email.
- `Generated Files`: registry Doc, PDF, dan draft.
- `Audit Log`: jejak operasi.
- `Config_Access`: whitelist dan role.

## Kontrol Preventif

- Otorisasi server-side: `ADMIN` dan `OPERATOR`.
- Optimistic locking memakai `Revision`.
- Mutasi memakai `LockService.getScriptLock()`.
- Validasi enum, email, tanggal, jumlah penerima, dan nominal.
- Tidak memakai `getActiveSpreadsheet()` pada alur web app.
- Output Drive dicatat dengan ID, revision, user, dan timestamp.
- Draft email per dokumen, bukan per baris gabungan.
