function buatPerjadinTemplate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPerjadin = ss.getSheetByName("Data Perjadin");
  var sheetMaster = ss.getSheetByName("Master Permohonan");

  if (!sheetPerjadin || !sheetMaster) {
    SpreadsheetApp.getUi().alert("Sheet Data Perjadin atau Master Permohonan tidak ditemukan!");
    return;
  }

  var dataPerjadin = sheetPerjadin.getDataRange().getValues();
  var dataMaster = sheetMaster.getDataRange().getValues();

  // ====== LOG SHEET ======
  var sheetLog = ss.getSheetByName("Log_Sheet");
  if (!sheetLog) {
    sheetLog = ss.insertSheet("Log_Sheet");
    sheetLog.appendRow(["ID Sheet", "Tanggal Dibuat", "Debug Info"]);
  }

  // Retrieve Log_Sheet data and handle empty sheet case
  var dataLog = sheetLog.getDataRange().getValues();
  var existingSheetIds = dataLog.length > 1 ? dataLog.slice(1).map(function(row) { return row[0]; }) : [];

  var createdCount = 0;
  var idCounters = {};

  for (var i = 1; i < dataPerjadin.length; i++) { // skip header
    var row = dataPerjadin[i];
    var idPermohonan = row[0];
    var nama = row[1] || "";
    var nip = row[2] || "";
    var nomorSuratTugas = row[3] || "";
    var dosenPangkat = row[4] || "";
    var kategori = row[5] || "";
    var tanggalKegiatanRaw = row[6] || "";
    var tempatKegiatan = row[7] || "";

    // Initialize counter for ID if not exists
    if (!idCounters[idPermohonan]) {
      idCounters[idPermohonan] = 1;
    }
    var urut = idCounters[idPermohonan];
    var sheetId = idPermohonan + "_" + urut;

    // Skip if sheet already exists
    if (existingSheetIds.indexOf(sheetId) !== -1) {
      idCounters[idPermohonan]++;
      continue;
    }

    // Ambil data dari Master Permohonan
    var tglPermohonan = "";
    var untukKeperluan = ": ";
    var penandatangan = "";
    var jabatan = "";
    for (var m = 1; m < dataMaster.length; m++) {
      if (dataMaster[m][0] === idPermohonan) {
        tglPermohonan = dataMaster[m][15]; // Kolom R
        untukKeperluan = dataMaster[m][1] ? ": " + dataMaster[m][1].toString() : ": ";
        penandatangan = dataMaster[m][20] ? dataMaster[m][20].toString() : "";
        jabatan = dataMaster[m][22] ? dataMaster[m][22].toString() : "";
        break;
      }
    }
    var [tglPermohonanStart, tglPermohonanEnd] = parseTanggalString(tglPermohonan);
    tglPermohonan = formatTanggalRange(tglPermohonanStart, tglPermohonanEnd);

    // Parse tanggal kegiatan
    var [tglMulai, tglAkhir, parseError] = parseTanggalString(tanggalKegiatanRaw);
    var tanggalKegiatanStr = parseError ? ": Invalid date (" + parseError + ")" : formatTanggalRange(tglMulai, tglAkhir);
    var jmlHariKegiatan = parseError ? 0 : hitungJumlahHari(tglMulai, tglAkhir);

    // Debug logging
    var debugInfo = "ID: " + idPermohonan + ", Raw Tanggal Kegiatan: " + tanggalKegiatanRaw +
                    ", Parsed Start: " + (tglMulai ? tglMulai.toString() : "null") +
                    ", Parsed End: " + (tglAkhir ? tglAkhir.toString() : "null") +
                    ", Formatted: " + tanggalKegiatanStr +
                    ", Jml Hari: " + jmlHariKegiatan +
                    (parseError ? ", Error: " + parseError : "");

    // ====== BUAT SHEET BARU ======
    var newSheet = ss.insertSheet(sheetId, ss.getNumSheets());

    // HEADER TENGAH
    var headerRange = newSheet.getRange("A1:F1");
    headerRange.merge();
    headerRange.setValue("PERMOHONAN PENCAIRAN PERJALANAN DINAS")
               .setFontWeight("bold")
               .setHorizontalAlignment("center")
               .setVerticalAlignment("middle")
               .setFontSize(16);

    var tanggalRange = newSheet.getRange("A2:F2");
    tanggalRange.merge();
    tanggalRange.setValue("Tanggal: " + tglPermohonan)
                .setHorizontalAlignment("center")
                .setVerticalAlignment("middle");

    // DETAIL PENERIMA
    var detailLabels = [
      "Nama", "NIP/NPM", "Surat Tugas", "Dosen/Pangkat Penunjang",
      "Kategori penerima tugas", "Berangkat ke", "Tanggal Kegiatan",
      "Tanggal Berangkat", "Untuk keperluan", "Jml. Hari kegiatan resmi",
      "Tambahan hari lamanya perjalanan"
    ];

    var detailValues = [
      nama != null ? nama.toString() : "",
      nip != null ? nip.toString() : "",
      "III/DPSP/" + (nomorSuratTugas != null ? nomorSuratTugas.toString() : "") + "-ST",
      dosenPangkat != null ? dosenPangkat.toString() : "",
      kategori != null ? kategori.toString() : "",
      tempatKegiatan != null ? tempatKegiatan.toString() : "",
      tanggalKegiatanStr, // Use formatted string or error message
      ": ",
      untukKeperluan,
      jmlHariKegiatan > 0 ? jmlHariKegiatan.toString() + " hari" : ": ",
      ": "
    ];

    for (var d = 0; d < detailLabels.length; d++) {
      var r = 4 + d;
      newSheet.getRange(r, 1, 1, 2).merge().setValue(detailLabels[d]).setFontWeight("bold");

      var cellC = newSheet.getRange(r, 3, 1, 4).merge();
      cellC.setHorizontalAlignment("left").setVerticalAlignment("center");

      var value = detailValues[d];
      value = value != null ? value.toString() : "";
      if (!value.startsWith(": ")) value = ": " + value;
      cellC.setValue(value);

      if (d === 7 || d === 10) cellC.setBackground("#d9d9d9");
    }

    newSheet.setColumnWidth(2, 230);
    newSheet.setColumnWidth(3, 250);

    // ====== AMBIL DATA UANG DARI DATA PERJADIN ======
    var uangValues = [];
    for (var col = 8; col <= 19; col++) {
      uangValues.push(row[col] != null && row[col] !== "" ? row[col] : 0);
    }

    // TABEL KOMPONEN UANG PERJALANAN DINAS
    var tableStartRow = 16;
    var tableHeaders = ["No", "Komponen Uang Perjalanan Dinas", "Jumlah yang disetujui", "Keterangan"];
    var tableKomponen = [
      "Uang kegiatan", "Uang saku", "Uang makan", "Uang penginapan",
      "Uang transportasi dalam kota", "Uang transportasi antar kota",
      "Uang transportasi antar negara pp", "Uang aplikasi visa",
      "Uang asuransi perjalanan", "Uang fiskal & pajak bandara", "Uang harian"
    ];

    newSheet.getRange(tableStartRow, 1, 1, tableHeaders.length).setValues([tableHeaders])
            .setFontWeight("bold")
            .setHorizontalAlignment("center")
            .setBackground("#d9d9d9")
            .setBorder(true, true, true, true, true, true)
            .setWrap(true);

    newSheet.getRange(tableStartRow, 4, 1, 3).merge();

    for (var k = 0; k < tableKomponen.length; k++) {
      var rowIdx = tableStartRow + 1 + k;
      newSheet.getRange(rowIdx, 1, 1, 5).setBorder(true, true, true, true, true, true);
      newSheet.getRange(rowIdx, 1).setValue(k + 1).setHorizontalAlignment("center");
      newSheet.getRange(rowIdx, 2).setValue(tableKomponen[k]);
      newSheet.getRange(rowIdx, 3).setValue(uangValues[k])
                                 .setBackground("#d9d9d9")
                                 .setNumberFormat('"Rp"#,##0');
      newSheet.getRange(rowIdx, 4, 1, 3).merge().setBackground("#d9d9d9");
    }

    // TOTAL
    var totalRow = tableStartRow + 1 + tableKomponen.length;
    newSheet.getRange(totalRow, 1, 1, 2).merge()
            .setValue("TOTAL")
            .setFontWeight("bold")
            .setHorizontalAlignment("right");
    newSheet.getRange(totalRow, 3)
            .setFormula("=SUM(C" + (tableStartRow + 1) + ":C" + (totalRow - 1) + ")")
            .setFontWeight("bold")
            .setNumberFormat('"Rp"#,##0')
            .setHorizontalAlignment("right");
    newSheet.getRange(totalRow, 4, 1, 3).merge();
    newSheet.getRange(totalRow, 1, 1, 5)
            .setBackground("#d9d9d9")
            .setBorder(true, true, true, true, true, true);

    // ====== TANDA TANGAN ======
    var tandaTanganRow = totalRow + 3;
    newSheet.setColumnWidth(1, 40);
    newSheet.setColumnWidth(2, 270);
    newSheet.setColumnWidth(3, 165);
    newSheet.setColumnWidth(4, 100);
    newSheet.setColumnWidth(5, 100);
    newSheet.setColumnWidth(6, 100);

    newSheet.getRange(tandaTanganRow, 1, 1, 2).merge()
            .setValue("Bandung, " + tglPermohonan)
            .setHorizontalAlignment("left").setWrap(true);
    newSheet.getRange(tandaTanganRow, 3, 1, 2).merge()
            .setValue("Tanggal,").setHorizontalAlignment("left").setWrap(true);
    newSheet.getRange(tandaTanganRow, 5, 1, 2).merge()
            .setValue("Pemeriksa").setHorizontalAlignment("left").setWrap(true);

    newSheet.getRange(tandaTanganRow + 5, 1, 1, 2).merge()
            .setValue("Dr. Ir. Fransiscus Rian Pratikto, S.T., M.T., MIE.").setHorizontalAlignment("left").setWrap(true).setFontWeight("bold");
    newSheet.getRange(tandaTanganRow + 5, 3, 1, 2).merge()
            .setValue("Kepala BIKEU").setHorizontalAlignment("left").setWrap(true).setFontWeight("bold");
    newSheet.getRange(tandaTanganRow + 5, 5, 1, 2).merge()
            .setValue("Staff BIKEU").setHorizontalAlignment("left").setWrap(true).setFontWeight("bold");

    newSheet.getRange(tandaTanganRow + 6, 1, 1, 2).merge()
            .setValue("Direktur Perencanaan Strategis dan Pemasaran").setHorizontalAlignment("left").setWrap(true);
    newSheet.getRange(tandaTanganRow + 6, 3, 1, 2).merge()
            .setValue("Jabatan Kepala BIKEU").setHorizontalAlignment("left").setWrap(true);
    newSheet.getRange(tandaTanganRow + 6, 5, 1, 2).merge()
            .setValue("Jabatan Staff BIKEU").setHorizontalAlignment("left").setWrap(true);

    // TABEL KOMPONEN UANG PERJALANAN DINAS KE-2
    var table2StartRow = tandaTanganRow + 9;
    var table2Headers = ["No", "Komponen Uang Perjalanan Dinas", "Pencairan sebelum Pelaksanaan Tugas "];

    newSheet.getRange(table2StartRow, 1, 1, table2Headers.length).setValues([table2Headers])
            .setFontWeight("bold")
            .setHorizontalAlignment("center")
            .setBackground("#d9d9d9")
            .setBorder(true, true, true, true, true, true)
            .setWrap(true);

    for (var k = 0; k < tableKomponen.length; k++) {
      var rowIdx = table2StartRow + 1 + k;
      newSheet.getRange(rowIdx, 1, 1, 3).setBorder(true, true, true, true, true, true);
      newSheet.getRange(rowIdx, 1).setValue(k + 1).setHorizontalAlignment("center");
      newSheet.getRange(rowIdx, 2).setValue(tableKomponen[k]);
      newSheet.getRange(rowIdx, 3).setValue(uangValues[k])
                                .setBackground("#d9d9d9")
                                .setNumberFormat('"Rp"#,##0');
    }

    var totalRow2 = table2StartRow + 1 + tableKomponen.length;
    newSheet.getRange(totalRow2, 1, 1, 2).merge()
            .setValue("TOTAL")
            .setFontWeight("bold")
            .setHorizontalAlignment("right")
            .setBackground("#d9d9d9")
            .setBorder(true, true, true, true, true, true)
            .setWrap(true);
    newSheet.getRange(totalRow2, 3)
            .setFormula("=SUM(C" + (table2StartRow + 1) + ":C" + (totalRow2 - 1) + ")")
            .setFontWeight("bold")
            .setNumberFormat('"Rp"#,##0')
            .setHorizontalAlignment("right")
            .setBackground("#d9d9d9")
            .setBorder(true, true, true, true, true, true)
            .setWrap(true);

    // ====== TANDA TANGAN KOLOM E:F ======
    var tandaTangan2Row = 45;
    newSheet.getRange(tandaTangan2Row, 5, 1, 2).merge()
            .setValue("Tanggal, ")
            .setHorizontalAlignment("left")
            .setWrap(true);

    newSheet.getRange(50, 5, 1, 2).merge()
            .setValue("Penerima")
            .setHorizontalAlignment("left")
            .setWrap(true);

    newSheet.getRange(51, 5, 1, 2).merge()
            .setValue(nama)
            .setHorizontalAlignment("left")
            .setWrap(true)
            .setFontWeight("bold");

    // ====== TABEL KOMPONEN UANG PERJALANAN DINAS KE-3 ======
    var table3StartRow = totalRow2 + 3;
    var table3Headers = [
      "No",
      "Komponen Uang Perjalanan Dinas",
      "Pencairan Setelah Pelaksanaan Tugas",
      "Penggantian*)",
      "Reimbursement/Tanpa Penggantian*)"
    ];

    newSheet.getRange(table3StartRow, 1, 1, table3Headers.length).setValues([table3Headers])
            .setFontWeight("bold")
            .setHorizontalAlignment("center")
            .setBackground("#d9d9d9")
            .setBorder(true, true, true, true, true, true)
            .setWrap(true);

    newSheet.getRange(table3StartRow, 5, 1, 2).merge()
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle")
            .setFontWeight("bold")
            .setBackground("#d9d9d9");

    for (var k = 0; k < tableKomponen.length; k++) {
      var rowIdx = table3StartRow + 1 + k;
      newSheet.getRange(rowIdx, 1, 1, 6).setBorder(true, true, true, true, true, true);
      newSheet.getRange(rowIdx, 1).setValue(k + 1).setHorizontalAlignment("center");
      newSheet.getRange(rowIdx, 2).setValue(tableKomponen[k]);
      newSheet.getRange(rowIdx, 3)
              .setValue(0)
              .setBackground("#d9d9d9")
              .setNumberFormat('"Rp"#,##0');
      newSheet.getRange(rowIdx, 4).setBackground("#d9d9d9");
      newSheet.getRange(rowIdx, 5, 1, 2).merge().setBackground("#d9d9d9");
    }

    var totalRow3 = table3StartRow + 1 + tableKomponen.length;
    newSheet.getRange(totalRow3, 1, 1, 2).merge()
            .setValue("TOTAL")
            .setFontWeight("bold")
            .setHorizontalAlignment("right")
            .setBackground("#d9d9d9")
            .setBorder(true, true, true, true, true, true)
            .setWrap(true);
    newSheet.getRange(totalRow3, 3)
            .setFormula("=SUM(C" + (table3StartRow + 1) + ":C" + (totalRow3 - 1) + ")")
            .setFontWeight("bold")
            .setNumberFormat('"Rp"#,##0')
            .setHorizontalAlignment("right")
            .setBackground("#d9d9d9")
            .setBorder(true, true, true, true, true, true)
            .setWrap(true);
    newSheet.getRange(totalRow3, 4).setBackground("#d9d9d9").setBorder(true, true, true, true, true, true);
    newSheet.getRange(totalRow3, 5, 1, 2).merge().setBackground("#d9d9d9").setBorder(true, true, true, true, true, true);

    newSheet.getRange(totalRow3 + 1, 1, 1, 2).merge()
            .setValue("*) diisi setelah pelaksanaan tugas")
            .setHorizontalAlignment("left")
            .setWrap(true)
            .setFontSize(10)
            .setFontStyle("italic");

    newSheet.getRange(totalRow3 + 2, 1, 1, 2).merge()
            .setValue("**) untuk perjalanan dinas ke luar negeri")
            .setHorizontalAlignment("left")
            .setWrap(true)
            .setFontSize(10)
            .setFontStyle("italic");

    // Log
    sheetLog.appendRow([sheetId, new Date(), debugInfo]);
    existingSheetIds.push(sheetId); // Update the in-memory list to prevent duplicates in the same run
    idCounters[idPermohonan]++;
    createdCount++;
  }

  ss.toast("Selesai membuat " + createdCount + " sheet perjadin.", "Proses Selesai", 5);
}

// ======= Helper Functions =======

// Format tanggal tunggal menjadi string "dd MMMM yyyy"
function safeFormatTanggal(val) {
  if (val instanceof Date) {
    var bulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
                 "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    return val.getDate() + " " + bulan[val.getMonth()] + " " + val.getFullYear();
  }
  return val || "";
}

// Format tanggal range untuk sheet
function formatTanggalRange(tglMulai, tglAkhir) {
  if (!tglMulai) return ": ";
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

// Parsing string tanggal tunggal atau range, dengan validasi tanggal
function parseTanggalString(str) {
  if (!str) return [null, null, "Empty input"];
  if (str instanceof Date) return [str, str, null];

  str = str.toString().trim();
  var bulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
               "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  var daysInMonth = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]; // Non-leap year

  // Check if year is a leap year for February
  function isLeapYear(year) {
    return (year % 4 === 0 && year % 100 !== 0) || (year % 400 === 0);
  }

  // Validate day for a given month and year
  function isValidDate(day, monthIndex, year) {
    if (monthIndex < 0 || monthIndex > 11) return false;
    var maxDays = daysInMonth[monthIndex];
    if (monthIndex === 1 && isLeapYear(year)) maxDays = 29;
    return day >= 1 && day <= maxDays;
  }

  // Format range beda bulan "31 November - 5 Desember 2025"
  var rangeBedaBulanMatch = str.match(/^(\d{1,2}) (\w+) - (\d{1,2}) (\w+) (\d{4})$/i);
  if (rangeBedaBulanMatch) {
    var startDay = parseInt(rangeBedaBulanMatch[1]);
    var startMonth = rangeBedaBulanMatch[2];
    var endDay = parseInt(rangeBedaBulanMatch[3]);
    var endMonth = rangeBedaBulanMatch[4];
    var year = parseInt(rangeBedaBulanMatch[5]);
    var startMonthIndex = bulan.indexOf(startMonth);
    var endMonthIndex = bulan.indexOf(endMonth);

    if (startMonthIndex === -1) return [null, null, "Invalid start month: " + startMonth];
    if (endMonthIndex === -1) return [null, null, "Invalid end month: " + endMonth];
    if (!isValidDate(startDay, startMonthIndex, year)) return [null, null, "Invalid start date: " + startDay + " " + startMonth];
    if (!isValidDate(endDay, endMonthIndex, year)) return [null, null, "Invalid end date: " + endDay + " " + endMonth];

    var startDate = new Date(year, startMonthIndex, startDay);
    var endDate = new Date(year, endMonthIndex, endDay);
    if (startDate > endDate) return [null, null, "Start date after end date"];
    return [startDate, endDate, null];
  }

  // Format range sama bulan "30-31 November 2025"
  var rangeMatch = str.match(/^(\d{1,2})\s*-\s*(\d{1,2}) (\w+) (\d{4})$/i);
  if (rangeMatch) {
    var startDay = parseInt(rangeMatch[1]);
    var endDay = parseInt(rangeMatch[2]);
    var month = rangeMatch[3];
    var year = parseInt(rangeMatch[4]);
    var monthIndex = bulan.indexOf(month);

    if (monthIndex === -1) return [null, null, "Invalid month: " + month];
    if (!isValidDate(startDay, monthIndex, year)) return [null, null, "Invalid start date: " + startDay + " " + month];
    if (!isValidDate(endDay, monthIndex, year)) return [null, null, "Invalid end date: " + endDay + " " + month];

    var startDate = new Date(year, monthIndex, startDay);
    var endDate = new Date(year, monthIndex, endDay);
    if (startDate > endDate) return [null, null, "Start date after end date"];
    return [startDate, endDate, null];
  }

  // Format single date "30 November 2025"
  var singleMatch = str.match(/^(\d{1,2}) (\w+) (\d{4})$/i);
  if (singleMatch) {
    var day = parseInt(singleMatch[1]);
    var month = singleMatch[2];
    var year = parseInt(singleMatch[3]);
    var monthIndex = bulan.indexOf(month);

    if (monthIndex === -1) return [null, null, "Invalid month: " + month];
    if (!isValidDate(day, monthIndex, year)) return [null, null, "Invalid date: " + day + " " + month];

    var date = new Date(year, monthIndex, day);
    return [date, date, null];
  }

  return [null, null, "Invalid format: " + str];
}

// Hitung jumlah hari kegiatan dari tglMulai dan tglAkhir
function hitungJumlahHari(tglMulai, tglAkhir) {
  if (!tglMulai || !tglAkhir) return 0;
  var diff = tglAkhir.getTime() - tglMulai.getTime();
  return Math.max(1, Math.floor(diff / (1000 * 60 * 60 * 24)) + 1);
}