# Panduan Implementasi Website DPSP

## 1. Backup

1. Buka Google Sheet sumber.
2. Buat salinan penuh sebelum menjalankan setup atau migrasi.
3. Pastikan akun implementasi memiliki akses ke spreadsheet, semua template Docs, dan folder output.

## 2. File yang Diunggah

Unggah file berikut ke satu project Google Apps Script:

1. `appsscript.json`
2. `Code.gs`
3. `Config.gs`
4. `Utils.gs`
5. `Setup.gs`
6. `DataService.gs`
7. `Migration.gs`
8. `DocumentService.gs`
9. `EmailService.gs`
10. `FinanceService.gs`
11. `ExportService.gs`
12. `Index.html`
13. `Styles.html`
14. `Scripts.html`

Jangan unggah script legacy. `.claspignore` sudah mengecualikannya.

## 3. Konfigurasi

Periksa `Config.gs`:

- `SPREADSHEET_ID`
- `OUTPUT_FOLDER_ID`
- sembilan ID template Google Docs
- nama sheet

Template Google Docs dapat memakai placeholder:

- `{namaKegiatan}`
- `{{namaKegiatan}}`
- `<<Nama Kegiatan>>`

Placeholder lain terdapat pada `buildDocumentPlaceholders_()` di `DocumentService.gs`.

## 4. Setup Awal

1. Pilih fungsi `setupSystem`.
2. Klik **Run**.
3. Beri izin Spreadsheet, Drive, Docs, Gmail Compose, dan email identity.
4. Fungsi membuat/melengkapi sheet sistem dan header.
5. Buka `Config_Access`, lalu isi:

| Email | Active | Role |
| --- | --- | --- |
| operator@unpar.ac.id | TRUE | ADMIN |

Role yang didukung: `ADMIN`, `OPERATOR`, `VIEWER`.

## 5. Migrasi Data Lama

Jalankan migrasi hanya bila `Master Permohonan` sudah berisi data lama.

1. Jalankan `previewLegacyMigration`.
2. Baca execution log.
3. Pastikan `ready: true`.
4. Jalankan `migrateLegacyDataConfirmed` dari menu Run (wrapper tanpa parameter), atau panggil `migrateLegacyData('MIGRATE')` dari eksekusi terprogram.

Jika data lama sudah terlanjur dimigrasi saat kolom `ID Permohonan` masih kosong, jalankan `repairMigratedMasterIds` sekali untuk mengisi ID pada `Master Permohonan` agar data muncul di Ringkasan/Permohonan.

Migrasi:

- membuat backup tersembunyi untuk Master, Pegawai, Perjadin, dan Dokumen,
- menggabungkan ID duplikat menjadi satu permohonan,
- membuat `LEGACY-xxxx` otomatis bila kolom ID Permohonan lama kosong,
- memindahkan setiap surat ke `Dokumen Permohonan`,
- membuat tanggal ISO dan `Participant Key`,
- mempertahankan link Doc lama yang dapat dikenali.

Migrasi menolak berjalan bila `Dokumen Permohonan` sudah berisi data.

## 6. Uji Sebelum Deploy

1. Buat masing-masing satu permohonan Edu Fair, Campus Visit, Narasumber Workshop, dan Narasumber Promosi.
2. Periksa preview To/CC setiap dokumen.
3. Buat Google Doc dan PDF.
4. Buat draft Gmail.
5. Pastikan draft muncul di akun operator yang sedang login.
6. Tambah/urutkan pegawai, sinkronkan perjadin, dan pastikan nominal tetap pada orang yang sama.
7. Jalankan `getSystemStatus`.

## 7. Deployment Web App

1. Apps Script: **Deploy > New deployment**.
2. Type: **Web app**.
3. Execute as: **User accessing the web app**.
4. Who has access: domain Universitas Katolik Parahyangan.
5. Deploy dan buka URL hasil deployment.

`Execute as me` tidak didukung untuk alur ini. Mode tersebut dapat menyembunyikan email pengguna dan membuat draft di akun deployer.

Setelah perubahan kode:

1. **Deploy > Manage deployments**.
2. Edit deployment.
3. Pilih **New version**.
4. Deploy.

## 8. Urutan Operasional

1. Isi form dan simpan `DRAFT`.
2. Lengkapi data lalu simpan `READY`.
3. Buka detail dan preview email.
4. Klik **Buat dokumen** atau **Buat dokumen & draft**.
5. Buka Google Doc/PDF untuk pemeriksaan.
6. Buka Gmail Drafts, tinjau, lalu kirim manual.
7. Isi biaya perjadin pada menu Honor & Perjadin.
8. Ekspor sheet honor/perjadin menjadi PDF atau XLSX dari detail permohonan.

## 9. Batas Google Apps Script

- Proses dibuat per permohonan untuk mengurangi risiko melewati runtime.
- Penerima email dibatasi maksimal 50 alamat per draft.
- Jangan menjalankan banyak proses dokumen bersamaan.
- Bila template atau folder tidak dapat diakses, proses berhenti dan mencatat error.

## 10. Troubleshooting

- **Akun tidak memiliki akses**: tambah email ke `Config_Access`.
- **Identitas pengguna tidak tersedia**: deploy ulang sebagai `User accessing the web app`.
- **Template tidak ditemukan**: cek ID dan izin template di `Config.gs`.
- **Master CC tidak menemukan jabatan**: samakan teks jabatan dengan panduan.
- **Dokumen ganda**: jangan gunakan `force=true` kecuali memang ingin versi baru.
- **Data diubah pengguna lain**: muat ulang detail; revision mencegah overwrite.
