# Panduan Implementasi Website DPSP

## 1. Backup

1. Buka Google Sheet sumber.
2. Buat salinan penuh sebelum menjalankan setup.
3. Pastikan akun implementasi memiliki akses ke spreadsheet, semua template Docs, dan folder output.

## 2. File yang Diunggah

Unggah file berikut ke satu project Google Apps Script:

1. `appsscript.json`
2. `Code.gs`
3. `Config.gs`
4. `Utils.gs`
5. `Setup.gs`
6. `DataService.gs`
7. `DocumentService.gs`
8. `EmailService.gs`
9. `FinanceService.gs`
10. `ExportService.gs`
11. `Index.html`
12. `Styles.html`
13. `Scripts.html`

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
4. Fungsi membuat/melengkapi sheet sistem dan header, termasuk `Jadwal Kegiatan` untuk banyak sesi dalam satu pengajuan.
5. Buka `Config_Access`, lalu isi:

| Email | Active | Role |
| --- | --- | --- |
| operator@unpar.ac.id | TRUE | ADMIN |

Role yang didukung: `ADMIN` dan `OPERATOR`.

`ADMIN` dipakai untuk setup, arsip tahunan, dan seluruh pekerjaan operator. `OPERATOR` dipakai untuk input permohonan, generate dokumen, draft email, dan ekspor laporan.

## 5. Uji Sebelum Deploy

1. Buat masing-masing satu permohonan Edu Fair, Campus Visit, Narasumber Workshop, dan Narasumber Promosi.
2. Periksa preview To/CC setiap dokumen.
3. Buat Google Doc dan PDF.
4. Buat draft Gmail.
5. Pastikan draft muncul di akun operator yang sedang login.
6. Tambah/urutkan pegawai, sinkronkan perjadin, dan pastikan nominal tetap pada orang yang sama.
7. Jalankan `getSystemStatus`.

### Akses data saat generate sheet

Saat operator membuat sheet honor/perjadin, web app tidak membuka database utama ke browser. Apps Script membaca `Master Permohonan`, `Dokumen Permohonan`, `Data Pegawai`, dan `Data Perjadin` di server menggunakan akun pengguna yang sedang login, lalu hanya membuka sheet hasil `GEN-*` kepada pengguna. Metadata `DPSP_GENERATED` dipakai untuk memastikan sistem tidak menimpa sheet manual.

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
2. Tambahkan seluruh sesi kegiatan pada pengajuan yang sama melalui tombol `Tambah sesi`. Gunakan satu sesi berjangka untuk kegiatan beruntun, atau beberapa sesi untuk tanggal/waktu yang terpisah.
3. Lengkapi data lalu simpan `READY`.
4. Buka detail dan preview email.
5. Klik **Buat dokumen** atau **Buat dokumen & draft**.
6. Buka Google Doc/PDF untuk pemeriksaan.
7. Buka Gmail Drafts, tinjau, lalu kirim manual.
8. Isi biaya perjadin pada menu Honor & Perjadin.
9. Ekspor sheet honor/perjadin menjadi PDF atau XLSX dari detail permohonan.

Semua sesi dalam satu pengajuan memakai satu `ID Permohonan`, daftar orang, dokumen, dan alur tanda tangan yang sama. Contoh tanggal 25 dan 30 Juni dengan waktu berbeda tetap menghasilkan satu Surat Tugas; tanggal dan waktu setiap sesi digabungkan ke placeholder dokumen.

### Definisi status workflow

- `Draft`: data belum final, masih dapat diedit, dan belum dapat diproses menjadi dokumen.
- `Siap Diproses`: data sudah tervalidasi. Status ini tidak berarti dokumen atau draft email sudah dibuat.
- `Selesai`: seluruh Doc/PDF dan draft Gmail sudah tersedia, lalu operator menutup permohonan secara manual. Permohonan menjadi hanya-baca.

Progres dokumen ditampilkan terpisah dari status permohonan: `Belum Dibuat`, `Doc & PDF Dibuat`, `Draft Email Dibuat`, `Gagal Generate`, atau `Selesai`. Aplikasi membuat draft Gmail, tetapi tidak dapat memastikan email benar-benar sudah dikirim oleh operator.

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
