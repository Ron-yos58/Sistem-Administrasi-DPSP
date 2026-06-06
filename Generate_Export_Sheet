function exportSheets() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Pilih format export
  var formatResponse = ui.prompt(
    'Select Export Format',
    'Type "PDF" to export as PDF or "EXCEL" to export as Excel:',
    ui.ButtonSet.OK_CANCEL
  );
  if (formatResponse.getSelectedButton() === ui.Button.CANCEL) {
    ui.alert("Export canceled.");
    return;
  }

  var exportFormat = formatResponse.getResponseText().trim().toUpperCase();
  if (exportFormat !== "PDF" && exportFormat !== "EXCEL") {
    ui.alert("Invalid format! Please type \"PDF\" or \"EXCEL\".");
    return;
  }

  // Jika PDF, pilih ukuran kertas dan orientasi
  var paperSize = "A4";
  var orientation = "LANDSCAPE"; // default
  if (exportFormat === "PDF") {
    var sizeResponse = ui.prompt(
      'Select Paper Size',
      'Please type "Legal" or "A4" for the paper size:',
      ui.ButtonSet.OK_CANCEL
    );
    if (sizeResponse.getSelectedButton() === ui.Button.CANCEL) {
      ui.alert("Export canceled.");
      return;
    }
    paperSize = sizeResponse.getResponseText().trim().toUpperCase();
    if (paperSize !== "LEGAL" && paperSize !== "A4") {
      ui.alert("Invalid paper size! Defaulting to A4.");
      paperSize = "A4";
    }

    var orientResponse = ui.prompt(
      'Select Orientation',
      'Please type "Portrait" or "Landscape":',
      ui.ButtonSet.OK_CANCEL
    );
    if (orientResponse.getSelectedButton() === ui.Button.CANCEL) {
      ui.alert("Export canceled.");
      return;
    }
    orientation = orientResponse.getResponseText().trim().toUpperCase();
    if (orientation !== "PORTRAIT" && orientation !== "LANDSCAPE") {
      ui.alert("Invalid orientation! Defaulting to Landscape.");
      orientation = "LANDSCAPE";
    }
  }

  var BATCH_SIZE = 20;
  var scriptProperties = PropertiesService.getScriptProperties();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetExport_ID = spreadsheet.getSheetByName('Export_ID');
  var logSheet = getOrCreateLogSheet();

  var folderId = sheetExport_ID.getRange(2, 2).getValue().trim();
  if (!folderId) {
    logStatus(logSheet, "ERROR_FOLDER_ID", exportFormat);
    return;
  }

  var folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (e) {
    logStatus(logSheet, "INVALID_FOLDER_ID", exportFormat);
    return;
  }

  var sheets = spreadsheet.getSheets();
  var excludedSheets = [
    "Master Permohonan", "Database Pegawai", "Data Perjadin", "Export_ID", "Log_Sheet",
    "Dokumentasi", "NVScriptsProperties", "DO NOT DELETE - AutoCrat Job Settings", "Master CC", "Template"
  ];

  var exportableSheets = sheets.filter(sheet => !excludedSheets.includes(sheet.getName().trim()));
  if (exportableSheets.length === 0) {
    ui.alert("Tidak ada lembar yang dapat diekspor. Proses dibatalkan.");
    return;
  }

  // Ambil daftar log yang sudah ada
  var loggedData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 1).getValues().flat();

  var lastIndex = parseInt(scriptProperties.getProperty('lastExportIndex')) || 0;
  var exportCount = 0;
  var failedSheets = [];
  var startTime = new Date().getTime();
  var endIndex = Math.min(lastIndex + BATCH_SIZE, exportableSheets.length);

  ss.toast("Export process started (" + exportFormat + ")", "Export", 5);

  for (var i = lastIndex; i < endIndex; i++) {
    if (new Date().getTime() - startTime > 5 * 60 * 1000) {
      scriptProperties.setProperty('lastExportIndex', i);
      ss.toast("Ekspor sebagian disimpan. Silakan jalankan kembali untuk melanjutkan.", "Export Timeout", 8);
      return;
    }

    var sheet = exportableSheets[i];
    var sheetName = sheet.getName().trim();
    var logKey = sheetName + " - Export " + (exportFormat === "PDF" ? "PDF" : "Excel");

    // 🔎 Skip kalau sudah ada di log
    if (loggedData.includes(logKey)) {
      ss.toast("Skipped: " + logKey, "Already Exported", 3);
      continue;
    }

    var attempt = 0;
    var success = false;

    while (attempt < 3 && !success) {
      try {
        if (exportFormat === "PDF") {
          var sheetId = sheet.getSheetId();
          var url = 'https://docs.google.com/spreadsheets/d/' + spreadsheet.getId() + '/export?';
          var params = {
            format: 'pdf',
            gid: sheetId,
            size: paperSize,
            portrait: (orientation === "PORTRAIT"), // ✅ orientasi sesuai pilihan
            fitw: true,
            top_margin: 0.5,
            bottom_margin: 0.5,
            left_margin: 0.5,
            right_margin: 0.5,
            gridlines: 0,
            printtitle: 0
          };
          var queryString = Object.keys(params)
            .map(key => encodeURIComponent(key) + '=' + encodeURIComponent(params[key]))
            .join('&');
          var token = ScriptApp.getOAuthToken();
          var response = UrlFetchApp.fetch(url + queryString, {
            headers: { 'Authorization': 'Bearer ' + token },
            muteHttpExceptions: true
          });
          if (response.getResponseCode() === 200) {
            var pdfBlob = response.getBlob().setName(sheetName + '.pdf');
            folder.createFile(pdfBlob);
            logStatus(logSheet, sheetName, exportFormat);
            ss.toast("Exported: " + logKey, "Success", 3);
            success = true;
            exportCount++;
          }
        } else if (exportFormat === "EXCEL") {
          var url = 'https://docs.google.com/spreadsheets/d/' + spreadsheet.getId() + '/export?format=xlsx&gid=' + sheet.getSheetId();
          var token = ScriptApp.getOAuthToken();
          var response = UrlFetchApp.fetch(url, {
            headers: { 'Authorization': 'Bearer ' + token },
            muteHttpExceptions: true
          });
          if (response.getResponseCode() === 200) {
            var excelBlob = response.getBlob().setName(sheetName + '.xlsx');
            folder.createFile(excelBlob);
            logStatus(logSheet, sheetName, exportFormat);
            ss.toast("Exported: " + logKey, "Success", 3);
            success = true;
            exportCount++;
          }
        }

        Utilities.sleep(1500);
      } catch (err) {
        failedSheets.push(sheetName);
      }

      attempt++;
    }

    if (!success) {
      failedSheets.push(sheetName);
      ss.toast("Failed: " + sheetName, "Error", 5);
    }
  }

  if (endIndex >= exportableSheets.length) {
    scriptProperties.deleteProperty('lastExportIndex');
    ss.toast("Proses selesai. Total yang diekspor: " + exportCount, "Done", 8);
  } else {
    scriptProperties.setProperty('lastExportIndex', endIndex);
    ss.toast("Ekspor sebagian selesai. Lanjutkan dari indeks " + endIndex, "Paused", 8);
  }
}

function getOrCreateLogSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = spreadsheet.getSheetByName("Log_Sheet");
  if (!logSheet) {
    logSheet = spreadsheet.insertSheet("Log_Sheet");
    logSheet.appendRow(["ID Permohonan", "Tanggal Dibuat"]);
  }
  return logSheet;
}

function logStatus(sheet, idPermohonan, exportFormat) {
  var statusText = idPermohonan + " - Export " + (exportFormat === "PDF" ? "PDF" : "Excel");
  sheet.appendRow([statusText, new Date()]);
}
