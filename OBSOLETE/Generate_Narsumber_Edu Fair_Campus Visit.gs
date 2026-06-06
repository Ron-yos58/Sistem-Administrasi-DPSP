// //---------------------------------------------buat sheet honor Edu Fair dan Campus Visit---------------------------------------------

function buatSheetPerPermohonan() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("Sedang membuat sheet honor per permohonan...", "Proses Dimulai", 5);

  // ====== DETEKSI LOCALE ======
  var locale = (typeof ss.getSpreadsheetLocale === 'function') ? ss.getSpreadsheetLocale() : Session.getActiveUserLocale();
  var argSep = (locale && locale.toLowerCase().indexOf('en') === 0) ? ',' : ';';

  var sheetMaster = ss.getSheetByName("Master Permohonan");
  var sheetPegawai = ss.getSheetByName("Database Pegawai");
  var sheetLog = ss.getSheetByName("Log_Sheet");

  if (!sheetMaster || !sheetPegawai) {
    SpreadsheetApp.getUi().alert("Sheet Master Permohonan atau Database Pegawai tidak ditemukan!");
    return;
  }

  if (!sheetLog) {
    sheetLog = ss.insertSheet("Log_Sheet");
    sheetLog.appendRow(["ID Permohonan", "Tanggal Dibuat"]);
  }

  var dataMaster = sheetMaster.getDataRange().getValues();
  var dataLog = sheetLog.getDataRange().getValues();
  var existingIds = dataLog.length > 1 ? dataLog.slice(1).map(function(r) { return r[0]; }) : [];

  var createdCount = 0;

  // ====== KOLOM INDEX MASTER ======
  var colId = 0;
  var colCampus = 1;
  var colJenisSurat = 2;
  var colJudul = 11;
  var colSekolah = 12;
  var colAlamat = 19;
  var colTanggalSurat = 15;
  var colTanggalKeg = 17;
  var colJamKeg = 18;
  var colPenandatangan = 20;
  var colJabatan = 22;
  var colYaTidak = 23;

  for (var i = 1; i < dataMaster.length; i++) {
    var idPermohonan = dataMaster[i][colId];
    var campus = dataMaster[i][colCampus];
    var jenisSurat = dataMaster[i][colJenisSurat];
    var status = dataMaster[i][colYaTidak];

    if (
      status === "Ya" &&
      (campus === "Edu Fair" || campus === "Campus Visit") &&
      jenisSurat === "Surat Tugas" &&
      idPermohonan &&
      existingIds.indexOf(idPermohonan) === -1
    ) {
      var newSheet = ss.insertSheet(String(idPermohonan), ss.getNumSheets());

      // ====== DATA HEADER ======
      var judul = dataMaster[i][colJudul] || "";
      var campus = dataMaster[i][colCampus] || "";
      var sekolah = dataMaster[i][colSekolah] || "";
      var alamat = dataMaster[i][colAlamat] || "";
      var tglSuratRaw = dataMaster[i][colTanggalSurat];
      var tglKegiatanRaw = dataMaster[i][colTanggalKeg];
      var jamKegiatan = dataMaster[i][colJamKeg] || "";
      var penandatangan = dataMaster[i][colPenandatangan] || "";
      var jabatan = dataMaster[i][colJabatan] || "";

      var tglSurat = safeFormatTanggal(tglSuratRaw);
      var tglKegiatanText = formatTanggalKegiatan(tglKegiatanRaw);
      var jamKegiatanText = (jamKegiatan && typeof jamKegiatan !== "object") ? String(jamKegiatan).trim() :
                            (jamKegiatan instanceof Date ? Utilities.formatDate(jamKegiatan, "Asia/Jakarta", "HH.mm-HH.mm") : "");

      var dataPegawai = sheetPegawai.getDataRange().getValues();
      var pegawaiDenganId = dataPegawai.filter(function(row, idx) {
        return idx > 0 && row[0] === idPermohonan;
      });

      // ====== HEADER ATAS ======
      setRichLabelEdunCamvis(newSheet, 1, "Daftar Honor: ", judul);
      setRichLabelEdunCamvis(newSheet, 2, "Tipe Kegiatan: ", campus);
      setRichLabelEdunCamvis(newSheet, 3, "Nama Mitra: ", sekolah);
      setRichLabelEdunCamvis(newSheet, 4, "Alamat Kegiatan: ", alamat);
      setRichLabelEdunCamvis(newSheet, 5, "Tanggal Kegiatan: ", tglKegiatanText);

      // ====== HEADER TABEL ======
      var headerRow = 8;
      var headerRange = newSheet.getRange(8, 1, 1, 7);
      var headerTabel = [["No.", "Nama Pegawai / Mahasiswa", "NIP/NPM", "Jam Mulai", "Jam Akhir", "Honor", "Transport", "Uang Makan", "Jumlah"]];
      newSheet.getRange(headerRow, 1, 1, headerTabel[0].length)
              .setValues(headerTabel)
              .setFontWeight("bold")
              .setBackground("#d9d9d9")
              .setBorder(true, true, true, true, true, true)
              .setWrap(true);

      // ====== ATUR ALIGNMENT HEADER & DATA RANGE ======
      headerRange.setHorizontalAlignment("center");
      headerRange.setVerticalAlignment("middle");

      // ====== UKURAN KOLOM ======
      newSheet.setColumnWidth(1, 40);   // A = NO
      newSheet.setColumnWidth(2, 220);  // B = Nama
      newSheet.setColumnWidth(3, 150);  // C = NIP/NPM
      newSheet.setColumnWidth(4, 100);  // D = Jam Mulai
      newSheet.setColumnWidth(5, 100);  // E = Jam Akhir
      newSheet.setColumnWidth(6, 120);  // F = Honor
      newSheet.setColumnWidth(7, 120);  // G = Transport
      newSheet.setColumnWidth(8, 120);  // H = Uang Makan
      newSheet.setColumnWidth(9, 120);  // I = Jumlah

      // ====== ISI DATA PEGAWAI ======
      var dataStart = headerRow + 1;
      var quoteDot = '"."';
      var quoteColon = '":"';

      for (var j = 0; j < pegawaiDenganId.length; j++) {
        var baris = pegawaiDenganId[j];
        var nomor = j + 1;
        var nama = baris[1] || "";
        var nip = baris[2] || "";

        var jamMulai = "";
        var jamAkhir = "";
        if (jamKegiatanText && String(jamKegiatanText).indexOf("-") !== -1) {
          var parts = String(jamKegiatanText).split("-");
          jamMulai = parts[0].trim().replace(".", ":");
          jamAkhir = (parts[1] || "").trim().replace(".", ":");
        }

        var r = dataStart + j;
        newSheet.getRange(r, 1, 1, 5).setValues([[nomor, nama, nip, jamMulai, jamAkhir]]);

        // ====== FORMULA HONOR (Locale aware) ======
        var subE = 'SUBSTITUTE(E' + r + argSep + quoteDot + argSep + quoteColon + ')';
        var subD = 'SUBSTITUTE(D' + r + argSep + quoteDot + argSep + quoteColon + ')';
        var timeDiff = 'MOD(TIMEVALUE(' + subE + ')-TIMEVALUE(' + subD + ')' + argSep + '1)*24*9000';
        var andPart = 'AND(LEN(D' + r + ')' + argSep + 'LEN(E' + r + '))';
        var honorFormula = '=IF(' + andPart + argSep + timeDiff + argSep + '0)';

        newSheet.getRange(r, 6).setFormula(honorFormula);
        newSheet.getRange(r, 7).setValue("");
        newSheet.getRange(r, 9).setFormula('=F' + r + '+G' + r + '+H' + r);

        // Styling
        newSheet.getRange(r, 1, 1, 9).setBorder(true, true, true, true, true, true).setWrap(true).setHorizontalAlignment("center").setVerticalAlignment("middle");
        newSheet.getRange(r, 2, 1, 1).setBorder(true, true, true, true, true, true).setWrap(true).setHorizontalAlignment("left").setVerticalAlignment("middle");
        newSheet.getRange(r, 6).setNumberFormat('"Rp"#,##0').setHorizontalAlignment("right");
        newSheet.getRange(r, 7).setNumberFormat('"Rp"#,##0').setHorizontalAlignment("right");
        newSheet.getRange(r, 8).setNumberFormat('"Rp"#,##0').setHorizontalAlignment("right");
        newSheet.getRange(r, 9).setNumberFormat('"Rp"#,##0').setHorizontalAlignment("right"); // Jumlah
      }

      // ====== TOTAL ======
      var lastRow = (pegawaiDenganId.length > 0) ? (dataStart + pegawaiDenganId.length - 1) : (dataStart - 1);
      var totalRow = lastRow + 2;

      newSheet.getRange(totalRow, 1, 1, 8).merge()
              .setValue("TOTAL")
              .setFontWeight("bold")
              .setBorder(true, true, true, false, true, true)
              .setHorizontalAlignment("right");

      if (pegawaiDenganId.length > 0) {
        newSheet.getRange(totalRow, 9)
                .setFormula("=SUM(I" + dataStart + ":I" + lastRow + ")")
                .setFontWeight("bold")
                .setBackground("#d9d9d9")
                .setBorder(true, true, true, true, true, true)
                .setNumberFormat('"Rp"#,##0')
                .setHorizontalAlignment("right");
      } else {
        newSheet.getRange(totalRow, 8)
                .setValue(0)
                .setFontWeight("bold")
                .setBackground("#d9d9d9")
                .setBorder(true, true, true, true, true, true)
                .setNumberFormat('"Rp"#,##0')
                .setHorizontalAlignment("right");
      }

      // ====== DROPDOWN JAM 24 JAM ======
      if (pegawaiDenganId.length > 0) {
        setDropdownJam(newSheet, 9, pegawaiDenganId.length);
      }

      // ====== TANDA TANGAN ======
      var tandaTanganRow = totalRow + 3;
      newSheet.getRange(tandaTanganRow, 7, 1, 3).merge().setValue("Bandung, " + tglSurat).setWrap(true);
      newSheet.getRange(tandaTanganRow + 5, 7, 1, 3).merge().setValue("Dr. Ir. Fransiscus Rian Pratikto, S.T., M.T., MIE.").setWrap(true).setFontWeight("bold");
      newSheet.getRange(tandaTanganRow + 6, 7, 1, 3).merge().setValue("Direktur Perencanaan Strategis dan Pemasaran").setWrap(true);

      // ====== FIX KOLOM A & B ======
      newSheet.setColumnWidth(1, 40);
      newSheet.setColumnWidth(2, 220);

      // Log
      sheetLog.appendRow([idPermohonan, new Date()]);
      existingIds.push(idPermohonan); // Update the in-memory list to prevent duplicates in the same run
      createdCount++;
    }
  }
  ss.toast("Proses selesai: " + createdCount + " sheet dibuat.", "Selesai", 5);
}

// ====== FUNGSI BUAT DROPDOWN JAM ======
function setDropdownJam(sheet, startRow, numRows) {
  var jamList = [];
  for (var h = 0; h < 24; h++) {
    var hh = (h < 10 ? "0" + h : "" + h) + ":00";
    jamList.push(hh);
  }

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(jamList, true)
    .setAllowInvalid(false)
    .build();

  // Terapkan ke kolom D (Jam Mulai)
  sheet.getRange(startRow, 4, numRows).setDataValidation(rule);

  // Terapkan ke kolom E (Jam Akhir)
  sheet.getRange(startRow, 5, numRows).setDataValidation(rule);
}

// ====== FORMAT TANGGAL KEGIATAN ======
function formatTanggalKegiatan(tglKegiatanRaw) {
  if (!tglKegiatanRaw) return "";
  if (typeof tglKegiatanRaw === "string") {
    var parts = tglKegiatanRaw.split("-");
    if (parts.length === 2) {
      var startDate = parseTanggal(parts[0].trim());
      var endDate = parseTanggal(parts[1].trim());
      if (startDate && endDate) {
        return formatTanggalRange(startDate, endDate);
      }
    }
    return safeFormatTanggal(parseTanggal(tglKegiatanRaw));
  } else if (tglKegiatanRaw instanceof Date) {
    return safeFormatTanggal(tglKegiatanRaw);
  }
  return "";
}

// ====== PARSE TANGGAL ======
function parseTanggal(str) {
  if (!str) return null;
  var bulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
               "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  var match = str.match(/^(\d{1,2}) (\w+) (\d{4})$/i);
  if (!match) return null;

  var day = parseInt(match[1]);
  var month = match[2];
  var year = parseInt(match[3]);
  var monthIndex = bulan.indexOf(month);

  if (monthIndex === -1) return null;
  var date = new Date(year, monthIndex, day);
  return isNaN(date.getTime()) ? null : date;
}

// ====== FORMAT TANGGAL RANGE ======
function formatTanggalRange(tglMulai, tglAkhir) {
  if (!tglMulai) return "";
  var bulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
               "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  
  if (!tglAkhir || tglMulai.getTime() === tglAkhir.getTime()) {
    return tglMulai.getDate() + " " + bulan[tglMulai.getMonth()] + " " + tglMulai.getFullYear();
  } else if (tglMulai.getMonth() === tglAkhir.getMonth()) {
    return tglMulai.getDate() + " - " + tglAkhir.getDate() + " " + bulan[tglMulai.getMonth()] + " " + tglMulai.getFullYear();
  } else {
    return tglMulai.getDate() + " " + bulan[tglMulai.getMonth()] + " - " + tglAkhir.getDate() + " " + bulan[tglAkhir.getMonth()] + " " + tglAkhir.getFullYear();
  }
}

// ====== FORMAT TANGGAL TUNGGAL ======
function safeFormatTanggal(val) {
  if (val instanceof Date) {
    var bulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
                 "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    return val.getDate() + " " + bulan[val.getMonth()] + " " + val.getFullYear();
  }
  return val || "";
}

// ====== SET RICH LABEL EDUN CAMVIS ======
function setRichLabelEdunCamvis(sheet, row, label, value) {
  var range = sheet.getRange(row, 1, 1, 6);
  range.merge();
  var richText = SpreadsheetApp.newRichTextValue()
    .setText(label + value)
    .setTextStyle(0, label.length, SpreadsheetApp.newTextStyle().setBold(true).build())
    .build();
  range.setRichTextValue(richText);
}

//---------------------------------------------buat sheet honor penugasan narasumber---------------------------------------------

function buatSheetPerPermohonan_narasumber() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("Sedang membuat sheet honor per permohonan...", "Proses Dimulai", 5);

  var sheetMaster = ss.getSheetByName("Master Permohonan");
  var sheetPegawai = ss.getSheetByName("Database Pegawai");
  var sheetLog = ss.getSheetByName("Log_Sheet");

  if (!sheetMaster || !sheetPegawai) {
    SpreadsheetApp.getUi().alert("Sheet Master Permohonan atau Database Pegawai tidak ditemukan!");
    return;
  }

  if (!sheetLog) {
    sheetLog = ss.insertSheet("Log_Sheet");
    sheetLog.appendRow(["ID Permohonan", "Tanggal Dibuat"]);
  }

  var dataMaster = sheetMaster.getDataRange().getValues();
  var dataLog = sheetLog.getDataRange().getValues();
  var existingIds = dataLog.length > 1 ? dataLog.slice(1).map(function(r) { return r[0]; }) : [];

  var createdCount = 0;

  // ====== KOLOM INDEX MASTER ======
  var colId = 0;
  var colCampus = 1;
  var colJenisSurat = 2;
  var colsubtipe = 3;
  var colJudul = 11;
  var colSekolah = 12;
  var colAlamat = 19;
  var colTanggalSurat = 15;
  var colTanggalKeg = 17;
  var colJamKeg = 18;
  var colPenandatangan = 20;
  var colJabatan = 22;
  var colYaTidak = 23;

  for (var i = 1; i < dataMaster.length; i++) {
    var idPermohonan = dataMaster[i][colId];
    var campus = dataMaster[i][colCampus];
    var jenisSurat = dataMaster[i][colJenisSurat];
    var status = dataMaster[i][colYaTidak];

    if (
      status === "Ya" &&
      jenisSurat === "Surat Tugas" &&
      campus === "Penugasan Narasumber" &&
      idPermohonan &&
      existingIds.indexOf(idPermohonan) === -1
    ) {
      var newSheet = ss.insertSheet(String(idPermohonan), ss.getNumSheets());

      // ====== DATA HEADER ======
      var judul = dataMaster[i][colJudul] || "";
      var campus = dataMaster[i][colCampus] || "";
      var subtipe = dataMaster[i][colsubtipe] || "";
      var sekolah = dataMaster[i][colSekolah] || "";
      var alamat = dataMaster[i][colAlamat] || "";
      var tglSuratRaw = dataMaster[i][colTanggalSurat];
      var tglKegiatanRaw = dataMaster[i][colTanggalKeg];
      var jamKegiatan = dataMaster[i][colJamKeg] || "";
      var penandatangan = dataMaster[i][colPenandatangan] || "";
      var jabatan = dataMaster[i][colJabatan] || "";

      var tglSurat = safeFormatTanggal(tglSuratRaw);
      var tglKegiatanText = formatTanggalKegiatan(tglKegiatanRaw);
      var jamKegiatanText = (jamKegiatan && typeof jamKegiatan !== "object") ? String(jamKegiatan).trim() :
                            (jamKegiatan instanceof Date ? Utilities.formatDate(jamKegiatan, "Asia/Jakarta", "HH.mm-HH.mm") : "");

      var dataPegawai = sheetPegawai.getDataRange().getValues();
      var pegawaiDenganId = dataPegawai.filter((row, idx) => idx > 0 && row[0] === idPermohonan);

      // ====== HEADER ATAS ======
      setRichLabelNarsum(newSheet, 1, "Daftar Honor: ", judul);
      setRichLabelNarsum(newSheet, 2, "Tipe Kegiatan: ", campus + (subtipe ? " - " + subtipe : ""));
      setRichLabelNarsum(newSheet, 3, "Nama Mitra: ", sekolah);
      setRichLabelNarsum(newSheet, 4, "Alamat Kegiatan: ", alamat);
      setRichLabelNarsum(newSheet, 5, "Tanggal Kegiatan: ", tglKegiatanText);

      // ====== HEADER TABEL ======
      var headerRow = 8;
      var headerRange = newSheet.getRange(8, 1, 1, 7);
      var headerTabel = [["No.", "Nama Pegawai", "NIP", "Makalah", "Honor", "Transport", "Jumlah"]];
      newSheet.getRange(headerRow, 1, 1, headerTabel[0].length)
              .setValues(headerTabel)
              .setFontWeight("bold")
              .setBackground("#d9d9d9")
              .setBorder(true, true, true, true, true, true)
              .setWrap(true);

      // ====== ATUR ALIGNMENT HEADER & DATA RANGE ======
      headerRange.setHorizontalAlignment("center");
      headerRange.setVerticalAlignment("middle");

      // ====== UKURAN KOLOM ======
      newSheet.setColumnWidth(1, 40);   // A = NO
      newSheet.setColumnWidth(2, 220);  // B = Nama
      newSheet.setColumnWidth(3, 150);  // C = NIP
      newSheet.setColumnWidth(4, 164);  // D = Makalah (Dropdown)
      newSheet.setColumnWidth(5, 120);  // E = Honor
      newSheet.setColumnWidth(6, 120);  // F = Transport
      newSheet.setColumnWidth(7, 120);  // G = Jumlah

      // ====== ISI DATA PEGAWAI ======
      var dataStart = headerRow + 1;

      for (var j = 0; j < pegawaiDenganId.length; j++) {
        var baris = pegawaiDenganId[j];
        var nomor = j + 1;
        var nama = baris[1] || "";
        var nip = baris[2] || "";

        var r = dataStart + j;
        newSheet.getRange(r, 1, 1, 7).setValues([[nomor, nama, nip, "", "", "", ""]]);

        // ====== DROPDOWN MAKALAH DENGAN DEFAULT ======
        var rule = SpreadsheetApp.newDataValidation()
          .requireValueInList(["Dengan Makalah", "Tidak Dengan Makalah"], true)
          .setAllowInvalid(false)
          .build();
        newSheet.getRange(r, 4).setDataValidation(rule);
        newSheet.getRange(r, 4).setValue("Dengan Makalah"); // default

        // ====== FORMULA HONOR BERDASARKAN PILIHAN MAKALAH ======
        newSheet.getRange(r, 5).setFormula('=IF(D' + r + '="Dengan Makalah";1000000;IF(D' + r + '="Tidak Dengan Makalah";750000;0))');

        // Jumlah = Honor + Transport
        newSheet.getRange(r, 7).setFormula('=E' + r + '+F' + r);

        // Styling
        newSheet.getRange(r, 1, 1, 7).setBorder(true, true, true, true, true, true).setWrap(true).setHorizontalAlignment("center").setVerticalAlignment("middle");
        newSheet.getRange(r, 2, 1, 1).setBorder(true, true, true, true, true, true).setWrap(true).setHorizontalAlignment("left").setVerticalAlignment("middle");
        newSheet.getRange(r, 5).setNumberFormat('"Rp"#,##0').setHorizontalAlignment("right");
        newSheet.getRange(r, 6).setNumberFormat('"Rp"#,##0').setHorizontalAlignment("right");
        newSheet.getRange(r, 7).setNumberFormat('"Rp"#,##0').setHorizontalAlignment("right");
      }

      // ====== TOTAL ======
      var lastRow = (pegawaiDenganId.length > 0) ? (dataStart + pegawaiDenganId.length - 1) : (dataStart - 1);
      var totalRow = lastRow + 2;

      newSheet.getRange(totalRow, 1, 1, 6).merge()
              .setValue("TOTAL")
              .setFontWeight("bold")
              .setBorder(true, true, true, false, true, true)
              .setHorizontalAlignment("right");

      if (pegawaiDenganId.length > 0) {
        newSheet.getRange(totalRow, 7)
                .setFormula("=SUM(G" + dataStart + ":G" + lastRow + ")")
                .setFontWeight("bold")
                .setBackground("#d9d9d9")
                .setBorder(true, true, true, true, true, true)
                .setNumberFormat('"Rp"#,##0')
                .setHorizontalAlignment("right");
      } else {
        newSheet.getRange(totalRow, 7)
                .setValue(0)
                .setFontWeight("bold")
                .setBackground("#d9d9d9")
                .setBorder(true, true, true, true, true, true)
                .setNumberFormat('"Rp"#,##0')
                .setHorizontalAlignment("right");
      }

      // ====== TANDA TANGAN ======
      var tandaTanganRow = totalRow + 3;
      newSheet.getRange(tandaTanganRow, 5, 1, 3).merge().setValue("Bandung, " + tglSurat).setWrap(true);
      newSheet.getRange(tandaTanganRow + 5, 5, 1, 3).merge().setValue("Dr. Ir. Fransiscus Rian Pratikto, S.T., M.T., MIE.").setWrap(true).setFontWeight("bold");
      newSheet.getRange(tandaTanganRow + 6, 5, 1, 3).merge().setValue("Direktur Perencanaan Strategis dan Pemasaran").setWrap(true);

      // Log
      sheetLog.appendRow([idPermohonan, new Date()]);
      existingIds.push(idPermohonan); // Update the in-memory list to prevent duplicates in the same run
      createdCount++;
    }
  }

  ss.toast("Proses selesai: " + createdCount + " sheet dibuat.", "Selesai", 5);
}

// -------- Helper untuk Permohonan Honor Narasumber dan Edu Fair/Campus Visit --------

function formatTanggalIndonesia(date) {
  var bulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
               "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  var hari = date.getDate();
  var bln = bulan[date.getMonth()];
  var tahun = date.getFullYear();
  return hari + " " + bln + " " + tahun;
}

function safeFormatTanggal(val) {
  if (val instanceof Date) {
    return formatTanggalIndonesia(val);
  }
  return val || "";
}

function formatTanggalKegiatan(val) {
  if (val instanceof Date) {
    return formatTanggalIndonesia(val);
  }
  if (typeof val === "string") {
    return val; // biarkan string (misalnya "11 - 12 Oktober 2025")
  }
  return "";
}

function setRichLabelNarsum(sheet, row, label, value) {
  var cell = sheet.getRange("A" + row + ":G" + row);
  cell.merge().setWrap(true);
  var richText = SpreadsheetApp.newRichTextValue()
    .setText(label + value)
    .setTextStyle(0, label.length, SpreadsheetApp.newTextStyle().setBold(true).build())
    .build();
  cell.setRichTextValue(richText);
}