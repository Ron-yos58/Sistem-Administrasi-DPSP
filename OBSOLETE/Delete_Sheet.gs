// --------Delete Sheets

function deleteUnnecessarySheets() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    "Confirmation of Deletion",
    "Are you sure you want to delete unnecessary sheets?",
    ui.ButtonSet.YES_NO
  );

  // Jika pengguna memilih "No", hentikan proses
  if (response == ui.Button.NO) {
    ui.alert("Deletion is canceled.");
    Logger.log("Deletion is canceled by the user.");
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var keepSheets = [
    "Master Permohonan", "Database Pegawai", "Data Perjadin", "Export_ID", "Log_Sheet",
    "Dokumentasi", "NVScriptsProperties", "DO NOT DELETE - AutoCrat Job Settings", "Master CC", "Template"
  ]; // Daftar sheet yang tidak boleh dihapus

  var deletedSheets = 0; // Counter jumlah sheet yang dihapus

  sheets.forEach(function(sheet) {
    var sheetName = sheet.getName();
    if (!keepSheets.includes(sheetName)) {
      ss.deleteSheet(sheet);
      Logger.log("Sheet removed: " + sheetName);
      deletedSheets++;
    }
  });

  // Tampilkan notifikasi setelah penghapusan selesai
  ui.alert("Process Completed", "Number of sheets deleted: " + deletedSheets, ui.ButtonSet.OK);
}
