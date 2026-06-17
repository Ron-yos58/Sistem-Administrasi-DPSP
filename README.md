# Sistem Administrasi & Dasbor Surat DPSP

Sistem ini digunakan untuk mengelola permohonan surat tugas, rekomendasi, perizinan, narasumber, serta honorarium dan perjalanan dinas (perjadin) di lingkungan Direktorat Perencanaan Strategis dan Pemasaran (DPSP). Aplikasi ini dibangun menggunakan Google Apps Script (GAS) dengan antarmuka web modern yang terintegrasi langsung dengan spreadsheet sebagai basis data utama.

---

## 📋 Status Permohonan & Indikator Siklus Hidup

Setiap permohonan surat melewati 3 status utama. Berikut adalah detail indikator, warna badge pada antarmuka, pesan sistem, dan perilaku pengeditan data untuk masing-masing status:

### 1. Draft (`DRAFT`)
* **Warna Badge:** 🟡 **Kuning / Oranye** (`.status-draft`)
* **Pesan Sistem:** *"Data masih dapat diubah. Lengkapi lalu tandai Siap Diproses sebelum membuat dokumen."*
* **Kondisi Data:** Permohonan baru dibuat atau belum divalidasi. 
* **Perilaku Pengeditan:** Data bebas diubah sewaktu-waktu. Fitur pembuatan dokumen Google Docs, PDF, dan draft email masih dikunci.

### 2. Siap Diproses (`READY`)
* **Warna Badge:** 🟢 **Hijau** (`.status-ready`)
* **Pesan Sistem:**
  * Sebelum dokumen siap: *"Data sudah tervalidasi dan siap dibuat menjadi Doc, PDF, serta draft Gmail."*
  * Setelah dokumen siap: *"Seluruh Doc, PDF, dan draft Gmail tersedia. Permohonan dapat ditandai selesai."*
* **Kondisi Data:** Data permohonan telah tervalidasi dan memenuhi kelengkapan minimum.
* **Perilaku Pengeditan:**
  * **Dapat diubah:** Data tetap dapat diedit oleh pengguna.
  * **Reset Penginstalan (*Fingerprint Reset*):** Jika Anda mengedit kolom penting (seperti jadwal kegiatan, detail mitra, tipe kegiatan, data pegawai/narasumber, dll.), sistem akan mendeteksi perubahan tersebut. Status dokumen yang telah dibuat akan di-reset kembali menjadi `PENDING` dan file sebelumnya ditandai sebagai *superseded* (usang) agar di-generate ulang dengan data terbaru.

### 3. Selesai (`ARCHIVED`)
* **Warna Badge:** ⚪ **Abu-abu** (`.status-archived`)
* **Pesan Sistem:** *"Permohonan telah selesai dan bersifat hanya-baca."*
* **Kondisi Data:** Seluruh proses pembuatan dokumen dan draft email telah sukses diselesaikan.
* **Perilaku Pengeditan:** **Terkunci Total (*Locked / Read-Only*)**. Data tidak bisa diubah baik dari sisi form dasbor maupun dimodifikasi langsung lewat backend API.

---

## 🗄️ Arsitektur Penyimpanan Status (Google Sheets)

Status dan log pelacakan disimpan pada sheet-sheet berikut di Google Spreadsheet:

1. **Sheet `Master Permohonan` (Sheet Name: `MASTER`)**
   * Menyimpan status utama permohonan pada kolom **`Status Permohonan`** (`DRAFT`, `READY`, `ARCHIVED`).
   * Menyimpan metadata lain seperti `Dibuat Oleh`, `Dibuat Pada`, `Diubah Oleh`, dan `Diubah Pada`.

2. **Sheet `Dokumen Permohonan` (Sheet Name: `DOCUMENTS`)**
   * Karena satu permohonan bisa memiliki beberapa dokumen surat (contoh: Surat Tugas sekaligus Surat Izin Pimpinan), status per-dokumen dilacak secara terpisah di sheet ini.
   * Kolom **`Status Dokumen`** melacak proses generate file Google Doc dan PDF (`PENDING`, `GENERATED`, `ERROR`).
   * Kolom **`Email Status`** melacak status pembuatan draft Gmail (`PENDING`, `DRAFTED`).

3. **Sheet `Audit Log` (Sheet Name: `AUDIT`)**
   * Mencatat log histori mutasi status (seperti `SAVE_REQUEST`, `ACTIVATE_REQUEST`, dan `ARCHIVE_REQUEST`) lengkap dengan timestamp, email pengguna yang memicu aksi, dan payload revisi data dalam bentuk JSON.

---

## 🛠️ Panduan Struktur Codebase

* [Config.gs](file:///c:/Users/Ronald_FTI/OneDrive/Documents/Dashbor%20Surat%20DPSP/Config.gs): Mengatur konfigurasi global aplikasi, daftar ID template Google Docs, jenis surat, daftar unit/fakultas, dan header-header kolom spreadsheet.
* [DataService.gs](file:///c:/Users/Ronald_FTI/OneDrive/Documents/Dashbor%20Surat%20DPSP/DataService.gs): Logika CRUD data permohonan, validasi, penguncian data, dan kalkulasi *fingerprint* perubahan data.
* [DocumentService.gs](file:///c:/Users/Ronald_FTI/OneDrive/Documents/Dashbor%20Surat%20DPSP/DocumentService.gs): Layanan pembuatan dokumen Google Docs, konversi ke PDF, dan penyimpanan file ke Google Drive.
* [EmailService.gs](file:///c:/Users/Ronald_FTI/OneDrive/Documents/Dashbor%20Surat%20DPSP/EmailService.gs): Logika penyusunan template email dan pembuatan otomatis draft Gmail untuk penerima (*To*, *CC*, *BCC*).
* [FinanceService.gs](file:///c:/Users/Ronald_FTI/OneDrive/Documents/Dashbor%20Surat%20DPSP/FinanceService.gs): Logika perhitungan honorarium narasumber dan nominal perjalanan dinas (perjadin).
* [Index.html](file:///c:/Users/Ronald_FTI/OneDrive/Documents/Dashbor%20Surat%20DPSP/Index.html): Struktur markup HTML dari antarmuka dasbor.
* [Scripts.html](file:///c:/Users/Ronald_FTI/OneDrive/Documents/Dashbor%20Surat%20DPSP/Scripts.html): Logika JavaScript interaktif di sisi klien (frontend), penanganan state dasbor, tombol aksi, dan pemanggilan fungsi Apps Script (*Google Server Callback*).
* [Styles.html](file:///c:/Users/Ronald_FTI/OneDrive/Documents/Dashbor%20Surat%20DPSP/Styles.html): Gaya tampilan CSS modern dasbor, termasuk pendefinisian warna badge untuk masing-masing status.

---

## ⚖️ Do's & Don'ts Alur Kerja (Workflow)

Untuk memastikan sistem berjalan dengan baik dan menghindari kerusakan data atau kegagalan otomatisasi, ikuti panduan berikut:

### 1. Permohonan Draft (`DRAFT`)
* **DO:**
  * Isi data awal kegiatan, mitra, jadwal, dan petugas sedetail mungkin.
  * Ubah data permohonan secara bebas selagi status masih `DRAFT`.
* **DON'T:**
  * Mencoba membuat/generate Google Docs, PDF, draft email, atau laporan keuangan honor/perjadin (fitur-fitur ini dikunci di backend dan tidak dapat dieksekusi).

### 2. Permohonan Siap Diproses (`READY`)
* **DO:**
  * Lakukan *Generate Google Docs* untuk membuat draf surat.
  * Masukkan nomor surat resmi yang valid setelah draf ditinjau.
  * Generate PDF surat dan draf email (Gmail Draft).
  * Lengkapi data keuangan (honorarium dan perjadin) langsung pada Google Sheet terpisah yang ter-generate otomatis.
* **DON'T:**
  * Melakukan perubahan data mayor pada form dasbor jika dokumen sudah final, kecuali Anda siap jika status pengerjaan di-reset kembali ke `PENDING` (*fingerprint reset* akan mendeteksi perubahan data dan menuntut regenerasi ulang dokumen agar file tetap sinkron).

### 3. Permohonan Selesai (`ARCHIVED`)
* **DO:**
  * Gunakan tautan yang tersedia di dasbor untuk mengunduh PDF, melihat dokumen, atau membuka draf email yang sudah siap di Gmail.
  * Jadikan data permohonan ini sebagai arsip acuan historis.
* **DON'T:**
  * Mencoba mengedit data permohonan, jadwal, mitra, atau petugas (seluruh data dikunci secara total).
  * Mencoba melakukan generate ulang dokumen, PDF, draft email, atau sheet keuangan (sistem memblokir seluruh fungsi pembuatan dokumen setelah masuk status `ARCHIVED` untuk mencegah modifikasi atau penimpaan file final).
