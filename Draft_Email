// ------------------------------Draft Email Generator--------------------------------------

var EMAIL_DRAFTED = "EMAIL DRAFTED";

function draftMyEmails() {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadSheet.getSheetByName('Master Permohonan');
  var sheetTemplate = spreadSheet.getSheetByName('Template');
  
  SpreadsheetApp.getActiveSpreadsheet().toast("Sedang memproses pembuatan draft email...", "Proses Dimulai", 5);

  var startRow = 2;
  var numRows = sheet.getLastRow() - 1;
  var lastColumn = sheet.getLastColumn();
  var dataRange = sheet.getRange(startRow, 1, numRows, lastColumn);
  var data = dataRange.getValues();

  // --- Buat mapping template ---
  var templateData = sheetTemplate.getDataRange().getValues();
  var templateMap = {};
  templateData.forEach(function(row, index) {
    if (index === 0) return; // skip header
    
    var subTipe = row[0] ? row[0].toString().trim() : '';   // Kolom A
    var status = row[1] ? row[1].toString().trim() : '';    // Kolom B
    var subject = row[2] || ''; // Kolom C
    var body = row[3] || '';    // Kolom D

    if (subTipe) {
      var key = status ? (subTipe + "|" + status) : subTipe;
      templateMap[key] = { subject: subject, body: body };
    }
  });

  // --- Formatter tanggal Indonesia ---
  function formatTanggalIndo(date) {
    if (!(date instanceof Date)) return date || 'Tidak ada data';
    return new Intl.DateTimeFormat('id-ID', {
      day: 'numeric',
      month: 'long',
      year: 'numeric'
    }).format(date);
  }

  // --- Loop data Master Permohonan ---
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];

    var jenisSurat = row[1] || '';        
    var subTipe = row[2] ? row[2].toString().trim() : ''; 
    var statusNarasumber = row[4] ? row[4].toString().trim() : ''; // kolom E
    var nomorSurat = row[6] || '';       
    var nomorSuratMasuk = row[7] || '';  
    var tanggalSuratMasuk = row[8] || ''; 
    var pengirimSuratMasuk = row[9] || ''; 
    var perihalSuratMasuk = row[10] || ''; 
    var namaKegiatan = row[11] || '';    
    var namaMitra = row[12] || '';       
    var alamatMitra = row[13] || '';     
    var hari = row[16] || '';            
    var tanggal = row[17];               
    var waktuKegiatan = row[18] || '';   
    var tempatKegiatan = row[19] || '';  

    // --- Update sesuai perubahan kolom email ---
    var toEmailAE = row[30] || '';       // Kolom AE
    var toEmailAF = row[31] || '';       // Kolom AF
    var kepadaYth = row[32] || '';       // Kolom AG
    var ccEmailAG = row[33] || '';       // Kolom AH
    var ccEmailAH = row[34] || '';       // Kolom AI

    // Gabungkan email
    var toEmail = [toEmailAE, toEmailAF].filter(Boolean).join('\n');
    var emailStatus = row[lastColumn - 1]; 
    
    var tanggalFormatted = formatTanggalIndo(tanggal);
    var tanggalSuratMasukFormatted = formatTanggalIndo(tanggalSuratMasuk);

    // Handle nilai kosong
    var fields = [
      { value: jenisSurat },        
      { value: nomorSurat },        
      { value: nomorSuratMasuk },   
      { value: pengirimSuratMasuk },
      { value: perihalSuratMasuk }, 
      { value: namaKegiatan },     
      { value: namaMitra },        
      { value: alamatMitra },      
      { value: hari },             
      { value: waktuKegiatan },    
      { value: tempatKegiatan }    
    ];

    fields.forEach(function(field) {
      if (!field.value) field.value = 'Tidak ada data';
    });

    // update kembali
    jenisSurat = fields[0].value;
    nomorSurat = fields[1].value;
    nomorSuratMasuk = fields[2].value;
    pengirimSuratMasuk = fields[3].value;
    perihalSuratMasuk = fields[4].value;
    namaKegiatan = fields[5].value;
    namaMitra = fields[6].value;
    alamatMitra = fields[7].value;
    hari = fields[8].value;
    waktuKegiatan = fields[9].value;
    tempatKegiatan = fields[10].value;

    if (emailStatus !== EMAIL_DRAFTED && toEmail) {

      function formatList(str) {
        if (!str) return 'Tidak ada data';
        var items = str.split(/\r?\n/).map(function(item) { return item.trim(); });
        var html = '<ol>';
        items.forEach(function(it) { if (it) html += '<li>' + it + '</li>'; });
        html += '</ol>';
        return html;
      }

      // var kepadaYthFormatted = '';
      // if (subTipe === 'Surat Tugas') {
      //   var namaTugas = row[26] || 'Tidak ada data';
      //   kepadaYthFormatted = formatList(namaTugas);
      // } else {
      //   kepadaYthFormatted = formatList(kepadaYth);
      // }

      var kepadaYthFormatted = '';
      if (subTipe === 'Surat Tugas') {
        var namaTugas = row[26] || 'Tidak ada data';
        // Surat Tugas tetap tampilkan apa adanya, pisah dengan <br> bila ada banyak
        kepadaYthFormatted = namaTugas.toString().split(/\r?\n/).join('<br>');
      } else {
        // Selain Surat Tugas → juga tanpa nomor, hanya line break
        kepadaYthFormatted = (kepadaYth || 'Tidak ada data').toString().split(/\r?\n/).join('<br>');
      }


      var ccFormatted = formatList(ccEmailAH);

      // --- Ambil template dari sheet Template ---
      var key = statusNarasumber ? (subTipe + "|" + statusNarasumber) : subTipe;
      var template = templateMap[key];
      if (!template) {
        // Tidak ada template → skip baris ini
        continue;
      }

      // Format narasumber (jika ada)
      var narasumberFormatted = '';
      if (subTipe === 'Surat Permohonan Narasumber kepada Dekan' && statusNarasumber === 'Tidak Dicarikan') {
        var namaList = (row[26] || '').toString().split(/\r?\n/);
        var nipList = (row[27] || '').toString().split(/\r?\n/);
        var combined = [];
        for (var j = 0; j < namaList.length; j++) {
          if (namaList[j]) {
            var nip = nipList[j] ? nipList[j].trim() : '';
            combined.push((j+1) + '. ' + namaList[j].trim() + (nip ? ' (' + nip + ')' : ''));
          }
        }
        narasumberFormatted = combined.join('\n');
      }

      // Replace placeholder di subject
      var perihal = template.subject
        .replace(/{jenisSurat}/g, jenisSurat)
        .replace(/{subTipe}/g, subTipe)
        .replace(/{statusNarasumber}/g, statusNarasumber)
        .replace(/{namaKegiatan}/g, namaKegiatan)
        .replace(/{namaMitra}/g, namaMitra)
        .replace(/{hari}/g, hari)
        .replace(/{tanggal}/g, tanggalFormatted)
        .replace(/{tempat}/g, tempatKegiatan)
        .replace(/{nomorSurat}/g, nomorSurat)
        .replace(/{nomorSuratMasuk}/g, nomorSuratMasuk)
        .replace(/{tanggalSuratMasuk}/g, tanggalSuratMasukFormatted)
        .replace(/{pengirimSuratMasuk}/g, pengirimSuratMasuk)
        .replace(/{perihalSuratMasuk}/g, perihalSuratMasuk)
        .replace(/{alamatMitra}/g, alamatMitra);

      // Replace placeholder di body
      var bodyText = template.body
        .replace(/{kepadaYth}/g, kepadaYthFormatted)
        .replace(/{jenisSurat}/g, jenisSurat)
        .replace(/{subTipe}/g, subTipe)
        .replace(/{statusNarasumber}/g, statusNarasumber)
        .replace(/{namaKegiatan}/g, namaKegiatan)
        .replace(/{namaMitra}/g, namaMitra)
        .replace(/{hari}/g, hari)
        .replace(/{tanggal}/g, tanggalFormatted)
        .replace(/{tempat}/g, tempatKegiatan)
        .replace(/{tembusan}/g, ccFormatted)
        .replace(/{nomorSurat}/g, nomorSurat)
        .replace(/{nomorSuratMasuk}/g, nomorSuratMasuk)
        .replace(/{tanggalSuratMasuk}/g, tanggalSuratMasukFormatted)
        .replace(/{pengirimSuratMasuk}/g, pengirimSuratMasuk)
        .replace(/{perihalSuratMasuk}/g, perihalSuratMasuk)
        .replace(/{alamatMitra}/g, alamatMitra)
        .replace(/{waktuKegiatan}/g, waktuKegiatan)
        .replace(/{narasumber}/g, narasumberFormatted);

      // Format ke HTML
      var bodyFinal = bodyText
        .replace(/\n{2,}/g, '</p><p>')
        .replace(/\n/g, '<br>');
      bodyFinal = '<p>' + bodyFinal + '</p>';

      // Buat draft
      GmailApp.createDraft(
        toEmail,
        perihal,
        '',
        {
          htmlBody: bodyFinal,
          cc: ccEmailAG,
          bcc: '',
          markImportant: true
        }
      );

      // Tandai sudah diproses
      sheet.getRange(startRow + i, lastColumn).setValue(EMAIL_DRAFTED);
      SpreadsheetApp.flush();
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("Proses selesai: Draft email telah dibuat.", "Selesai", 5);
}
