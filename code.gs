// Code.gs

// Fungsi doGet 
function doGet() {
  const html = HtmlService.createHtmlOutputFromFile('index')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
  // Sebaiknya atur judul di HTML saja jika perlu
  // html.setTitle('Google Style Scanner');
  return html;
}

/**
 * Fungsi untuk mencari data peserta berdasarkan barcode (email)
 * dan mengupdate status kehadirannya di Kolom J (baru).
 * Disesuaikan untuk struktur sheet setelah Kolom F dihapus.
 *
 * @param {string} barcodeData Data dari QR code (diasumsikan email).
 * @return {object} Objek status: { status: 'new'/'duplicate'/'error', message: 'pesan error jika ada', email: 'email peserta' }
 */
function updateAttendanceByBarcode(barcodeData) {
  // Gunakan try...catch untuk penanganan error yang lebih baik
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1');
    // REVISI 1: Membaca data dari Kolom B sampai J (struktur baru)
    const dataRange = sheet.getRange("B2:J" + sheet.getLastRow());
    const data = dataRange.getValues();

    // Mencari baris berdasarkan email (barcodeData) di kolom pertama (Kolom B -> index 0)
    let rowIndexInData = -1;
    let rowDataFound = null; // Simpan data baris yg ditemukan

    for (let i = 0; i < data.length; i++) {
      // Pastikan ada data di kolom B dan cocok (case insensitive trim)
      if (data[i][0] && data[i][0].toString().trim().toLowerCase() === barcodeData.trim().toLowerCase()) {
        rowIndexInData = i;
        rowDataFound = data[i]; // Simpan datanya
        break;
      }
    }

    // Jika data tidak ditemukan di Kolom B
    if (rowIndexInData === -1) {
      Logger.log(`Error: Barcode data "${barcodeData}" tidak ditemukan di Kolom B.`);
      return { status: 'error', message: 'Data Peserta Tidak Ditemukan', email: barcodeData };
    }

    // Email ditemukan, ambil datanya
    const email = rowDataFound[0]; // Email ada di index 0

    // REVISI 2: Cek status kehadiran di Kolom J (baru)
    // Dalam data B:J, Kolom J ada di index ke-8 (B=0, C=1, ... J=8)
    const attendanceStatusIndex = 8;
    const currentStatus = rowDataFound[attendanceStatusIndex];

    // Jika status sudah 'Sudah Hadir'
    if (currentStatus === 'Sudah Hadir') {
       Logger.log(`Info: Email "${email}" (${barcodeData}) sudah tercatat hadir sebelumnya.`);
      return { status: 'duplicate', email: email };
    }

    // Jika belum hadir, update statusnya
    // Hitung nomor baris aktual di spreadsheet (index data + 2)
    const rowIndexInSheet = rowIndexInData + 2;

    // REVISI 3: Update Kolom J (baru), yang merupakan kolom ke-10 di sheet
    const attendanceColumnNumber = 10;
    sheet.getRange(rowIndexInSheet, attendanceColumnNumber).setValue('Sudah Hadir');
    SpreadsheetApp.flush(); // Pastikan perubahan segera tersimpan

    Logger.log(`Sukses: Email "${email}" (${barcodeData}) di baris ${rowIndexInSheet} diupdate menjadi "Sudah Hadir".`);
    return { status: 'new', email: email };

  } catch (e) {
    // Tangani error tak terduga
    Logger.log(`ERROR di updateAttendanceByBarcode: ${e.message}\nData: ${barcodeData}\nStack: ${e.stack}`);
    // Kembalikan error ke client-side
    return { status: 'error', message: `Terjadi Error Server: ${e.message}`, email: barcodeData };
  }
}