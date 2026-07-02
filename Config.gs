const APP_CONFIG = Object.freeze({
  APP_NAME: 'Sistem Surat DPSP',
  APP_VERSION: '1.0.0',
  TIME_ZONE: 'Asia/Jakarta',
  SPREADSHEET_ID: '1jDY5XW86HmDdzYTcPVxpQBULHcP5qW-cxqCuEJZeyVI',
  OUTPUT_FOLDER_ID: '1MsDCWr-xJRd2K91KZLyJago5uiV4L4Dx',
  CACHE_SECONDS: 120,
  MAX_LIST_ROWS: 1000,
  SHEETS: {
    MASTER: { name: 'Master Permohonan', aliases: [] },
    DOCUMENTS: { name: 'Dokumen Permohonan', aliases: [] },
    SCHEDULES: { name: 'Jadwal Kegiatan', aliases: [] },
    EMPLOYEES: { name: 'Data Pegawai', aliases: ['Database Pegawai'] },
    TRAVEL: { name: 'Data Perjadin', aliases: [] },
    CC: { name: 'Master CC', aliases: [] },
    EMAIL_TEMPLATE: { name: 'Template', aliases: [] },
    EXPORT: { name: 'Export_ID', aliases: [] },
    AUDIT: { name: 'Audit Log', aliases: [] },
    FILES: { name: 'Generated Files', aliases: [] },
    ACCESS: { name: 'Config_Access', aliases: [] },
    SIGNATURE: { name: 'Signature_Config', aliases: [] },
    TEMPLATE_CONFIG: { name: 'Template_Config', aliases: ['Templates', 'Master Template'] }
  },
  ACTIVITY_TYPES: ['Edu Fair', 'Campus Visit', 'Penugasan Narasumber'],
  DOCUMENT_TYPES: [
    'Surat Tugas',
    'Surat Balasan Campus Visit',
    'Surat Rekomendasi Campus Visit - SU',
    'Surat izin pimpinan - Campus Visit',
    'Surat Permohonan Narasumber kepada Dekan'
  ],
  FACULTIES: [
    'Fakultas Ekonomi',
    'Fakultas Filsafat',
    'Fakultas Hukum',
    'Fakultas Ilmu Sosial dan Ilmu Politik',
    'Fakultas Kedokteran',
    'Fakultas Keguruan dan Ilmu Pendidikan',
    'Fakultas Sains',
    'Fakultas Teknik',
    'Fakultas Teknologi Rekayasa',
    'Fakultas Vokasi',
    'Direktorat Akademik',
    'Direktorat Digitalisasi',
    'Direktorat Kemahasiswaan',
    'Direktorat Manajemen Aset, Keuangan, Dan Sarana Prasarana',
    'Direktorat Organisasi dan Sumber Daya Insani',
    'Direktorat Pemelajaran',
    'Direktorat Pengelolaan Bisnis, Inovasi dan Kewirausahaan',
    'Direktorat Perencanaan Strategis dan Pemasaran',
    'Direktorat Urusan Internasional, Kerja Sama, dan Alumni',
    'Kantor Legal',
    'Kantor Media Digital',
    'Kantor Sekretariat Rektorat',
    'Lembaga Penelitian dan Pengabdian kepada Masyarakat',
    'Lembaga Penjaminan Mutu',
    'Lembaga Pengembangan Humaniora',
    'Perpustakaan'
  ],
});


const OFFICIAL_LETTER_MAPPING = Object.freeze([
  {
    activityType: 'Edu Fair',
    subType: 'Surat Tugas',
    condition: '-',
    templateKey: 'EDU_FAIR_TASK'
  },
  {
    activityType: 'Campus Visit',
    subType: 'Surat Balasan Campus Visit',
    condition: '-',
    templateKey: 'CAMPUS_VISIT_REPLY'
  },
  {
    activityType: 'Campus Visit',
    subType: 'Surat izin pimpinan - Campus Visit',
    condition: '-',
    templateKey: 'CAMPUS_VISIT_PERMISSION'
  },
  {
    activityType: 'Campus Visit',
    subType: 'Surat Rekomendasi Campus Visit - SU',
    condition: '-',
    templateKey: 'CAMPUS_VISIT_RECOMMENDATION'
  },
  {
    activityType: 'Campus Visit',
    subType: 'Surat Tugas',
    condition: '-',
    templateKey: 'CAMPUS_VISIT_TASK'
  },
  {
    activityType: 'Penugasan Narasumber',
    subType: 'Surat Permohonan Narasumber kepada Dekan',
    condition: 'Tidak Dicarikan',
    templateKey: 'SPEAKER_REQUEST_KNOWN'
  },
  {
    activityType: 'Penugasan Narasumber',
    subType: 'Surat Permohonan Narasumber kepada Dekan',
    condition: 'Dicarikan',
    templateKey: 'SPEAKER_REQUEST_SEARCH'
  },
  {
    activityType: 'Penugasan Narasumber',
    subType: 'Surat Tugas',
    condition: 'Workshop',
    templateKey: 'SPEAKER_WORKSHOP_TASK'
  },
  {
    activityType: 'Penugasan Narasumber',
    subType: 'Surat Tugas',
    condition: 'Promosi',
    templateKey: 'SPEAKER_PROMOTION_TASK'
  }
]);

const AUTOCRAT_HEADERS = Object.freeze([
  'Merged Doc ID - Penugasan Edu Fair',
  'Merged Doc URL - Penugasan Edu Fair',
  'Link to merged Doc - Penugasan Edu Fair',
  'Document Merge Status - Penugasan Edu Fair',
  'Merged Doc ID - Penugasan Narasumber',
  'Merged Doc URL - Penugasan Narasumber',
  'Link to merged Doc - Penugasan Narasumber',
  'Document Merge Status - Penugasan Narasumber',
  'Merged Doc ID - Penugasan Narasumber (Promosi)',
  'Merged Doc URL - Penugasan Narasumber (Promosi)',
  'Link to merged Doc - Penugasan Narasumber (Promosi)',
  'Document Merge Status - Penugasan Narasumber (Promosi)',
  'Merged Doc ID - Penugasan Campus Visit',
  'Merged Doc URL - Penugasan Campus Visit',
  'Link to merged Doc - Penugasan Campus Visit',
  'Document Merge Status - Penugasan Campus Visit',
  'Merged Doc ID - Surat izin pimpinan - Campus Visit',
  'Merged Doc URL - Surat izin pimpinan - Campus Visit',
  'Link to merged Doc - Surat izin pimpinan - Campus Visit',
  'Document Merge Status - Surat izin pimpinan - Campus Visit',
  'Merged Doc ID - Surat Rekomendasi Campus Visit - SU',
  'Merged Doc URL - Surat Rekomendasi Campus Visit - SU',
  'Link to merged Doc - Surat Rekomendasi Campus Visit - SU',
  'Document Merge Status - Surat Rekomendasi Campus Visit - SU',
  'Merged Doc ID - Surat Balasan Campus Visit',
  'Merged Doc URL - Surat Balasan Campus Visit',
  'Link to merged Doc - Surat Balasan Campus Visit',
  'Document Merge Status - Surat Balasan Campus Visit',
  'Merged Doc ID - Permohonan Narasumber kepada Dekan (Belum ada Narasumber)',
  'Merged Doc URL - Permohonan Narasumber kepada Dekan (Belum ada Narasumber)',
  'Link to merged Doc - Permohonan Narasumber kepada Dekan (Belum ada Narasumber)',
  'Document Merge Status - Permohonan Narasumber kepada Dekan (Belum ada Narasumber)',
  'Merged Doc ID - Permohonan Narasumber kepada Dekan (Sudah Ada Narasumber)',
  'Merged Doc URL - Permohonan Narasumber kepada Dekan (Sudah Ada Narasumber)',
  'Link to merged Doc - Permohonan Narasumber kepada Dekan (Sudah Ada Narasumber)',
  'Document Merge Status - Permohonan Narasumber kepada Dekan (Sudah Ada Narasumber)'
]);

const MASTER_HEADERS = Object.freeze([
  'ID Permohonan',
  'Tipe Kegiatan',
  'Jenis Surat',
  'Sub-Tipe Kegiatan',
  'Status Narasumber',
  'Fakultas Asal Narasumber',
  'Nomor Surat',
  'Nomor Surat Masuk',
  'Tanggal Surat Masuk',
  'Pengirim Surat Masuk',
  'Perihal Surat Masuk',
  'Nama Kegiatan',
  'Nama Mitra',
  'Alamat Mitra',
  'Email Mitra',
  'Tanggal Surat Dibuat',
  'Hari Kegiatan',
  'Tanggal Kegiatan',
  'Waktu Kegiatan',
  'Tempat Kegiatan',
  'Nama Penandatangan',
  'NIK Penandatangan',
  'Jabatan Penandatangan',
  'Honor',
  'Perjalanan Dinas',
  'Nomor Urut Pegawai',
  'Nama Pegawai',
  'NIP/NPM Pegawai',
  'Jabatan Pegawai',
  'Prodi/Unit Pegawai',
  'Email Pegawai',
  'Email To',
  'Jabatan Email To',
  'Email CC',
  'Jabatan Email CC',
  'Keterangan Email',
  'Edit Surat',
  'Download PDF Surat',
  'Email Status',
  'Status Permohonan',
  'Dibuat Oleh',
  'Dibuat Pada',
  'Diubah Oleh',
  'Diubah Pada',
  'Client Token',
  'Tanggal Mulai ISO',
  'Tanggal Selesai ISO',
  'Revision',
  'Manual Recipients'
]);

const LEGACY_MASTER_HEADERS = Object.freeze(
  MASTER_HEADERS.slice(0, 38)
    .concat(AUTOCRAT_HEADERS)
    .concat(MASTER_HEADERS.slice(38))
);

const EMPLOYEE_HEADERS = Object.freeze([
  'ID Permohonan',
  'Nama Pegawai / Mahasiswa',
  'NIP/NPM',
  'Jabatan Struktural/Fungsional',
  'Prodi / Unit',
  'Email',
  'Dosen/Pangkat Penunjang',
  'Kategori Penerima Tugas',
  'Participant Key'
]);

const SCHEDULE_HEADERS = Object.freeze([
  'Schedule ID',
  'ID Permohonan',
  'Tanggal Mulai',
  'Tanggal Selesai',
  'Waktu Mulai',
  'Waktu Selesai',
  'Tempat',
  'Urutan'
]);

const DOCUMENT_HEADERS = Object.freeze([
  'Document ID',
  'ID Permohonan',
  'Jenis Surat',
  'Sub-Tipe Kegiatan',
  'Status Narasumber',
  'Nomor Surat',
  'Template Key',
  'Status Dokumen',
  'Google Doc ID',
  'Google Doc URL',
  'PDF ID',
  'PDF URL',
  'Email Draft ID',
  'Email Status',
  'Revision',
  'Dibuat Pada',
  'Diubah Pada',
  'Email To',
  'Email CC',
  'Email BCC'
]);

const TRAVEL_HEADERS = Object.freeze([
  'ID Permohonan',
  'Nama Pegawai/Mahasiswa',
  'NIP/NPM',
  'Nomor Surat Tugas',
  'Dosen/Pangkat Penunjang',
  'Kategori Penerima Tugas',
  'Tanggal Kegiatan',
  'Tempat Kegiatan',
  'Uang Kegiatan',
  'Uang Saku',
  'Uang Makan',
  'Uang Penginapan',
  'Transportasi Dalam Kota',
  'Transportasi Antar Kota',
  'Transportasi Antar Negara PP',
  'Aplikasi Visa',
  'Asuransi Perjalanan',
  'Fiskal & Pajak Bandara',
  'Uang Harian',
  'Participant Key'
]);

const CC_HEADERS = Object.freeze(['Fakultas/Unit', 'Jabatan', 'Email']);
const SIGNATURE_HEADERS = Object.freeze(['Nama', 'Jabatan', 'NIK']);
const ACCESS_HEADERS = Object.freeze(['Email', 'Active', 'Role']);
const AUDIT_HEADERS = Object.freeze([
  'Log ID', 'Timestamp', 'User', 'Action', 'Entity ID', 'Success', 'Details JSON'
]);
const GENERATED_FILE_HEADERS = Object.freeze([
  'Request ID', 'Artifact Key', 'Revision', 'Type', 'File ID', 'URL',
  'Created At', 'Created By', 'Status', 'Metadata JSON'
]);

const TEMPLATE_CONFIG_HEADERS = Object.freeze([
  'Template Key',
  'Nama Template',
  'Jenis Surat',
  'Kode Surat',
  'Default To',
  'Default CC',
  'Default BCC',
  'Template Email',
  'Template Google Docs',
  'Status Aktif'
]);

const FALLBACK_ROUTING = Object.freeze({
  'Surat Balasan Campus Visit': {
    defaultTo: 'email mitra',
    defaultCc: ''
  },
  'Surat Rekomendasi Campus Visit - SU': {
    defaultTo: 'Sekretaris Universitas|Rektorat',
    defaultCc: ''
  },
  'Surat izin pimpinan - Campus Visit': {
    defaultTo: [
      'Sekretaris Universitas|Rektorat',
      'Dekan Fakultas Ekonomi|Fakultas Ekonomi',
      'Dekan Fakultas Hukum|Fakultas Hukum',
      'Dekan Fakultas Ilmu Sosial dan Ilmu Politik|Fakultas Ilmu Sosial dan Ilmu Politik',
      'Dekan Fakultas Teknik|Fakultas Teknik',
      'Dekan Fakultas Teknologi Rekayasa|Fakultas Teknologi Rekayasa',
      'Dekan Fakultas Sains|Fakultas Sains',
      'Dekan Fakultas Vokasi|Fakultas Vokasi',
      'Direktur Kemahasiswaan|Direktorat Kemahasiswaan',
      'Direktur Manajemen Aset, Keuangan, dan Sarana Prasarana|Direktorat Manajemen Aset, Keuangan, dan Sarana Prasarana',
      'Kepala Perpustakaan|Unit Perpustakaan'
    ].join('\n'),
    defaultCc: ''
  }
});
