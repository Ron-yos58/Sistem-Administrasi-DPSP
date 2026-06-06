function gabungkanPegawai_v5() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.getActiveSpreadsheet().toast("Sedang memproses data pegawai dan menghitung nama hari...", "Proses Dimulai", 5);

  var sheetMaster = ss.getSheetByName("Master Permohonan");
  var sheetPegawai = ss.getSheetByName("Database Pegawai");
  var sheetCC = ss.getSheetByName("Master CC");

  if (!sheetMaster || !sheetPegawai || !sheetCC) {
    SpreadsheetApp.getUi().alert("Sheet Master Permohonan, Database Pegawai, atau Master CC tidak ditemukan!");
    return;
  }

  // --- Ambil data dari Master (sampai kolom R untuk include tanggal) ---
  var dataMaster = sheetMaster.getRange(2, 1, sheetMaster.getLastRow() - 1, 18).getValues();

  // --- Ambil data pegawai ---
  var dataPegawai = sheetPegawai.getDataRange().getValues();
  dataPegawai.shift(); // hapus header

  var mapPegawai = {};
  dataPegawai.forEach(function(row) {
    var id = row[0];
    if (!id) return;
    var nama = row[1] || "";
    var nip = row[2] || "";
    var jabatan = row[3] || "";
    var prodi = row[4] || "";
    var email = row[5] || "";
    if (!mapPegawai[id]) {
      mapPegawai[id] = {nama: [], nip: [], jabatan: [], prodi: [], email: []};
    }
    mapPegawai[id].nama.push(nama);
    mapPegawai[id].nip.push(nip);
    mapPegawai[id].jabatan.push(jabatan);
    mapPegawai[id].prodi.push(prodi);
    if (email) mapPegawai[id].email.push(email);
  });

  // --- Ambil data CC ---
  var dataCC = sheetCC.getDataRange().getValues();
  dataCC.shift();

  var mapCC = {};
  dataCC.forEach(function(row) {
    var fakultas = row[0];
    var jabatan = row[1];
    var email = row[2];
    if (!mapCC[fakultas]) mapCC[fakultas] = {};
    mapCC[fakultas][jabatan] = email;
  });

  function cariEmail(mapCC, fakultas, keyword) {
    if (!mapCC[fakultas]) return { email: "", jabatan: "" };
    for (var jabatan in mapCC[fakultas]) {
      if (jabatan.indexOf(keyword) !== -1) {
        return { email: mapCC[fakultas][jabatan], jabatan: jabatan };
      }
    }
    return { email: "", jabatan: "" };
  }

  function cariBanyakEmail(mapCC, pairs) {
    var emails = [];
    var jabatans = [];
    for (var i = 0; i < pairs.length; i += 2) {
      var fakultas = pairs[i];
      var keyword = pairs[i + 1];
      var found = cariEmail(mapCC, fakultas, keyword);
      if (found.email) {
        emails.push(found.email);
        jabatans.push(found.jabatan);
      }
    }
    return {
      emails: uniqueArray(emails).join("\n"),
      jabatan: uniqueArray(jabatans).join("\n")
    };
  }

  function uniqueArray(arr) {
    return Array.from(new Set(arr));
  }

  // --- Helper Functions untuk Tanggal ---
  function parseTanggalString(str) {
    if (!str) return [null, null, "Empty input"];
    if (str instanceof Date) return [str, str, null];

    str = str.toString().trim();
    var bulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
                 "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    var daysInMonth = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

    function isLeapYear(year) {
      return (year % 4 === 0 && year % 100 !== 0) || (year % 400 === 0);
    }

    function isValidDate(day, monthIndex, year) {
      if (monthIndex < 0 || monthIndex > 11) return false;
      var maxDays = daysInMonth[monthIndex];
      if (monthIndex === 1 && isLeapYear(year)) maxDays = 29;
      return day >= 1 && day <= maxDays;
    }

    var rangeBedaBulanMatch = str.match(/^(\d{1,2}) (\w+) - (\d{1,2}) (\w+) (\d{4})$/i);
    if (rangeBedaBulanMatch) {
      var startDay = parseInt(rangeBedaBulanMatch[1]);
      var startMonth = rangeBedaBulanMatch[2];
      var endDay = parseInt(rangeBedaBulanMatch[3]);
      var endMonth = rangeBedaBulanMatch[4];
      var year = parseInt(rangeBedaBulanMatch[5]);
      var startMonthIndex = bulan.findIndex(m => m.toLowerCase() === startMonth.toLowerCase());
      var endMonthIndex = bulan.findIndex(m => m.toLowerCase() === endMonth.toLowerCase());

      if (startMonthIndex === -1) return [null, null, "Invalid start month: " + startMonth];
      if (endMonthIndex === -1) return [null, null, "Invalid end month: " + endMonth];
      if (!isValidDate(startDay, startMonthIndex, year)) return [null, null, "Invalid start date: " + startDay + " " + startMonth];
      if (!isValidDate(endDay, endMonthIndex, year)) return [null, null, "Invalid end date: " + endDay + " " + endMonth];

      var startDate = new Date(year, startMonthIndex, startDay);
      var endDate = new Date(year, endMonthIndex, endDay);
      if (startDate > endDate) return [null, null, "Start date after end date"];
      return [startDate, endDate, null];
    }

    var rangeMatch = str.match(/^(\d{1,2})\s*-\s*(\d{1,2}) (\w+) (\d{4})$/i);
    if (rangeMatch) {
      var startDay = parseInt(rangeMatch[1]);
      var endDay = parseInt(rangeMatch[2]);
      var month = rangeMatch[3];
      var year = parseInt(rangeMatch[4]);
      var monthIndex = bulan.findIndex(m => m.toLowerCase() === month.toLowerCase());

      if (monthIndex === -1) return [null, null, "Invalid month: " + month];
      if (!isValidDate(startDay, monthIndex, year)) return [null, null, "Invalid start date: " + startDay + " " + month];
      if (!isValidDate(endDay, monthIndex, year)) return [null, null, "Invalid end date: " + endDay + " " + month];

      var startDate = new Date(year, monthIndex, startDay);
      var endDate = new Date(year, monthIndex, endDay);
      if (startDate > endDate) return [null, null, "Start date after end date"];
      return [startDate, endDate, null];
    }

    var singleMatch = str.match(/^(\d{1,2}) (\w+) (\d{4})$/i);
    if (singleMatch) {
      var day = parseInt(singleMatch[1]);
      var month = singleMatch[2];
      var year = parseInt(singleMatch[3]);
      var monthIndex = bulan.findIndex(m => m.toLowerCase() === month.toLowerCase());

      if (monthIndex === -1) return [null, null, "Invalid month: " + month];
      if (!isValidDate(day, monthIndex, year)) return [null, null, "Invalid date: " + day + " " + month];

      var date = new Date(year, monthIndex, day);
      return [date, date, null];
    }

    return [null, null, "Invalid format: " + str];
  }

  function getDayName(date) {
    if (!date) return "";
    var days = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
    return days[date.getDay()];
  }

  function formatDayRange(tglMulai, tglAkhir, parseError) {
    if (parseError) return "Invalid date: " + parseError;
    if (!tglMulai) return "";
    var startDay = getDayName(tglMulai);
    if (!tglAkhir || tglMulai.getTime() === tglAkhir.getTime()) {
      return startDay;
    }
    var endDay = getDayName(tglAkhir);
    return startDay + " - " + endDay;
  }

  // --- Proses data Master ---
  var hasil = [];
  dataMaster.forEach(function(row) {
    var id = row[0];
    var jenisSurat = row[2];
    var fakultasUnit = row[5];
    var kolomJ = row[9] || "";
    var kolomM = row[12] || "";
    var emailMitra = row[14] || "";
    var tglPermohonan = row[17] || ""; // Kolom R

    var nomorList = "", namaList = "", nipList = "", jabatanList = "", prodiList = "", emailList = "";
    var toEmail = "", jabatanTo = "", ccEmail = "", jabatanCC = "", errorMsg = "";
    var dayRange = "";

    // Hitung nama hari dari kolom R hanya jika tidak kosong
    if (tglPermohonan) {
      var [tglMulai, tglAkhir, parseError] = parseTanggalString(tglPermohonan);
      dayRange = formatDayRange(tglMulai, tglAkhir, parseError);
    }

    if (mapPegawai[id]) {
      var jumlah = mapPegawai[id].nama.length;
      nomorList = Array.from({length: jumlah}, (_, i) => (i + 1)).join("\n");
      namaList = mapPegawai[id].nama.join("\n");
      nipList = mapPegawai[id].nip.join("\n");
      jabatanList = mapPegawai[id].jabatan.join("\n");
      prodiList = mapPegawai[id].prodi.join("\n");
      emailList = mapPegawai[id].email.join("\n");
    }

    // --- Aturan email ---
    if (jenisSurat == "Surat Balasan Campus Visit") {
      var cc = cariEmail(mapCC, "Rektorat", "Wakil Rektor Bidang Kerjasama");
      if (cc.email) {
        ccEmail = cc.email;
        jabatanCC = cc.jabatan;
      } else {
        errorMsg = "Alamat email WR Kerjasama tidak ditemukan di Master CC";
      }
      var jabatanToArray = [];
      if (kolomJ) jabatanToArray.push(kolomJ);
      if (kolomM) jabatanToArray.push(kolomM);
      jabatanTo = uniqueArray(jabatanToArray).join("\n");
    } else if (jenisSurat == "Surat Rekomendasi Campus Visit - SU") {
      var sekretaris = cariEmail(mapCC, "Rektorat", "Sekretaris Universitas");
      var wakilR = cariEmail(mapCC, "Rektorat", "Wakil Rektor Bidang Kerjasama");
      if (sekretaris.email && wakilR.email) {
        toEmail = sekretaris.email;
        jabatanTo = sekretaris.jabatan;
        ccEmail = wakilR.email;
        jabatanCC = wakilR.jabatan;
      } else {
        errorMsg = "Alamat email Sekretaris Universitas atau WR Kerjasama tidak ditemukan di Master CC";
      }
    } else if (jenisSurat == "Surat izin pimpinan - Campus Visit") {
      var pimpinan = cariBanyakEmail(mapCC, [
        "Rektorat", "Sekretaris Universitas",
        "Fakultas Ekonomi", "Dekan Fakultas Ekonomi",
        "Fakultas Hukum", "Dekan Fakultas Hukum",
        "Fakultas Ilmu Sosial dan Ilmu Politik", "Dekan Fakultas Ilmu Sosial dan Ilmu Politik",
        "Fakultas Teknik", "Dekan Fakultas Teknik",
        "Fakultas Teknologi Rekayasa", "Dekan Fakultas Teknologi Rekayasa",
        "Fakultas Sains", "Dekan Fakultas Sains",
        "Fakultas Vokasi", "Dekan Fakultas Vokasi",
        "Direktorat Kemahasiswaan", "Direktur Kemahasiswaan",
        "Direktorat Manajemen Aset, Keuangan, dan Sarana Prasarana", "Direktur Manajemen Aset, Keuangan, dan Sarana Prasarana",
        "Unit Perpustakaan", "Kepala Perpustakaan"
      ]);
      var penunjang = cariBanyakEmail(mapCC, [
        "Rektorat", "Wakil Rektor Bidang Kerjasama",
        "Direktorat Manajemen Aset, Keuangan, dan Sarana Prasarana", "Manajer Aset dan Sarana Prasarana",
        "Direktorat Kemahasiswaan", "Manajer Kemahasiswaan",
        "Direktorat Manajemen Aset, Keuangan, dan Sarana Prasarana", "Koordinator Kebersihan, Keamanan dan Ketertiban",
        "Direktorat Manajemen Aset, Keuangan, dan Sarana Prasarana", "Koordinator Kelas dan Fasilitas Pendukung",
        "Fakultas Ekonomi", "Koordinator Administrasi Fakultas Ekonomi",
        "Fakultas Hukum", "Koordinator Administrasi Fakultas Hukum",
        "Fakultas Ilmu Sosial dan Ilmu Politik", "Koordinator Administrasi Fakultas Ilmu Sosial dan Ilmu Politik",
        "Fakultas Sains", "Koordinator Administrasi Fakultas Sains",
        "Fakultas Teknik", "Koordinator Administrasi Fakultas Teknik",
        "Fakultas Teknologi Rekayasa", "Koordinator Administrasi Fakultas Teknologi Rekayasa",
        "Fakultas Vokasi", "Koordinator Administrasi Fakultas Vokasi"
      ]);
      if (pimpinan.emails && penunjang.emails) {
        toEmail = pimpinan.emails;
        jabatanTo = pimpinan.jabatan;
        ccEmail = penunjang.emails;
        jabatanCC = penunjang.jabatan;
      } else {
        errorMsg = "Alamat email pimpinan atau penunjang tidak ditemukan di Master CC";
      }
    } else if (jenisSurat == "Surat Permohonan Narasumber kepada Dekan") {
      var fakultasList = fakultasUnit ? fakultasUnit.split(",") : [];
      var toEmails = [], jabatanToList = [], ccEmails = [], jabatanCCList = [], notFoundDekan = [], notFoundAdmin = [];
      fakultasList.forEach(function(fak) {
        var f = fak.trim();
        var dekan = cariEmail(mapCC, f, "Dekan");
        if (dekan.email) {
          toEmails.push(dekan.email);
          jabatanToList.push(dekan.jabatan);
        } else {
          notFoundDekan.push(f);
        }
        var admin = cariEmail(mapCC, f, "Koordinator Administrasi");
        if (admin.email) {
          ccEmails.push(admin.email);
          jabatanCCList.push(admin.jabatan);
        } else {
          notFoundAdmin.push(f);
        }
      });
      if (toEmails.length > 0) {
        toEmail = uniqueArray(toEmails).join("\n");
        jabatanTo = uniqueArray(jabatanToList).join("\n");
      }
      if (ccEmails.length > 0) {
        ccEmail = uniqueArray(ccEmails).join("\n");
        jabatanCC = uniqueArray(jabatanCCList).join("\n");
      }
      var messages = [];
      if (notFoundDekan.length > 0) messages.push("Dekan tidak ditemukan: " + notFoundDekan.join("; "));
      if (notFoundAdmin.length > 0) messages.push("Koordinator Administrasi tidak ditemukan: " + notFoundAdmin.join("; "));
      if (messages.length > 0) errorMsg = messages.join(" | ");
    }

    if (emailMitra) {
      if (toEmail) {
        toEmail = toEmail + "\n" + emailMitra;
      } else {
        toEmail = emailMitra;
      }
    }

    hasil.push([
      nomorList, namaList, nipList, jabatanList, prodiList, emailList,
      toEmail, jabatanTo, ccEmail, jabatanCC, errorMsg, dayRange
    ]);
  });

  // Tulis ke kolom Z–AJ dan Q
  sheetMaster.getRange(2, 26, hasil.length, 11).setValues(hasil.map(row => row.slice(0, 11))); // Z–AJ
  sheetMaster.getRange(2, 17, hasil.length, 1).setValues(hasil.map(row => [row[11]])); // Q

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Proses selesai: Data Pegawai, Email Rules, dan Nama Hari digabungkan.",
    "Selesai",
    5
  );
}

// ----------------------------Generate Data Perjadin----------------------------------------

function generateDataPerjadin() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.getActiveSpreadsheet().toast("Sedang memproses Data Perjadin...", "Proses Dimulai", 5);
  var sheetMaster = ss.getSheetByName("Master Permohonan");
  var sheetPegawai = ss.getSheetByName("Database Pegawai");
  var sheetOutput = ss.getSheetByName("Data Perjadin");

  // --- Ambil data dari Master Permohonan ---
  var dataMaster = sheetMaster.getDataRange().getValues();
  var idxID = 0;       // Kolom A = ID Permohonan
  var idNoSurat = 6;   // Kolom F = Nomor Surat
  var idxTanggal = 17; // Kolom P = Tanggal Kegiatan
  var idxTempat = 19;  // Kolom R = Tempat Kegiatan
  var idxW = 24;       // Kolom W = status (Ya/Tidak)

  // --- Ambil data dari Database Pegawai ---
  var dataPegawai = sheetPegawai.getDataRange().getValues();
  var mapPegawai = {};

  // Simpan pegawai berdasarkan ID Permohonan
  for (var i = 1; i < dataPegawai.length; i++) {
    var row = dataPegawai[i];
    var id = row[0];
    if (!id) continue;
    if (!mapPegawai[id]) mapPegawai[id] = [];
    mapPegawai[id].push([
      row[0], // ID
      row[1], // Nama
      row[2], // NIP/NPM
      row[6], // Jabatan / Pangkat penunjang (kolom G)
      row[7]  // Kategori penerima tugas (kolom H)
    ]);
  }

  // --- Filter master: W = "Ya" ---
  var hasil = [];
  for (var j = 1; j < dataMaster.length; j++) {
    var rowM = dataMaster[j];
    if (rowM[idxW] === "Ya") {
      var idPermohonan = rowM[idxID];
      var noSurat = rowM[idNoSurat]; 
      var tglKegiatan = rowM[idxTanggal];
      var tempat = rowM[idxTempat];
      
      if (mapPegawai[idPermohonan]) {
        mapPegawai[idPermohonan].forEach(function(pegawai) {
          hasil.push([
            pegawai[0], // ID
            pegawai[1], // Nama
            pegawai[2], // NIP/NPM
            noSurat,    // Nomor Surat
            pegawai[3], // Jabatan
            pegawai[4], // Kategori
            tglKegiatan,// Tanggal Kegiatan
            tempat      // Tempat Kegiatan
          ]);
        });
      }
    }
  }

  // --- Tulis ke Data Perjadin mulai dari baris 2 (header tetap) ---
  sheetOutput.getRange(2, 1, sheetOutput.getLastRow(), 8).clearContent(); 
  if (hasil.length > 0) {
    sheetOutput.getRange(2, 1, hasil.length, hasil[0].length).setValues(hasil);
  }
    SpreadsheetApp.getActiveSpreadsheet().toast("Proses selesai: Data Perjadin sudah digenerate.", "Selesai", 5);
}