/**
 * File Script: CheckInSystem.gs (atau nama file .gs Anda)
 * Deskripsi: Mengelola pengiriman invoice/barcode otomatis dan manual
 * berdasarkan status pembayaran di Google Sheet.
 * Disesuaikan untuk struktur sheet TANPA Kolom F.
 */

//------------------------------------------------------------------
// FUNGSI UTAMA UNTUK TRIGGER OTOMATIS (On Edit)
//------------------------------------------------------------------

/**
 * Fungsi ini akan dijalankan oleh Installable Trigger 'On Edit'.
 * Memeriksa apakah Kolom G diedit menjadi 'sudah valid' di sheet yang benar,
 * lalu memanggil fungsi pengiriman email untuk baris tersebut.
 *
 * @param {Object} e Event object yang disediakan oleh trigger OnEdit.
 */
function handleEdit(e) {
  try { // Tambahkan try-catch di level atas handleEdit
    const editedRange = e.range; // Range yang diedit
    const sheet = editedRange.getSheet();
    const sheetName = 'Form Responses 1'; // PASTIKAN NAMA SHEET SUDAH BENAR
    const triggerColumnIndex = 7; // Kolom G adalah kolom ke-7
    const requiredValue = "sudah valid";

    // 1. Cek apakah pengeditan terjadi di sheet yang benar dan hanya satu sel
    if (sheet.getName() !== sheetName || editedRange.getNumRows() !== 1 || editedRange.getNumColumns() !== 1) {
      // Logger.log(`handleEdit: Edit di luar sheet/range/kolom target. Sheet: ${sheet.getName()}, Col: ${editedRange.getColumn()}`);
      return; // Keluar jika bukan sheet/edit yang diinginkan
    }

    // 2. Cek apakah kolom yang diedit adalah Kolom G
    if (editedRange.getColumn() !== triggerColumnIndex) {
      // Logger.log(`handleEdit: Bukan kolom G yang diedit (Kolom: ${editedRange.getColumn()})`);
      return; // Keluar jika bukan kolom G yang diedit
    }

    // 3. Cek apakah nilai baru adalah "sudah valid" (case insensitive)
    const newValue = e.value; // Nilai baru di sel yang diedit
    if (!newValue || newValue.toString().toLowerCase().trim() !== requiredValue) {
      // Logger.log(`handleEdit: Nilai "${newValue}" di kolom G baris ${editedRange.getRow()} tidak sama dengan "${requiredValue}".`);
      return; // Keluar jika nilai tidak sesuai
    }

    // --- Jika semua kondisi terpenuhi ---
    const editedRowIndex = editedRange.getRow();
    // Hindari memproses baris header jika teredit
    if (editedRowIndex < 2) {
        Logger.log(`handleEdit: Baris header (baris ${editedRowIndex}) teredit, diabaikan.`);
        return;
    }

    Logger.log(`handleEdit: Kolom G baris ${editedRowIndex} diubah menjadi "${requiredValue}". Memulai proses pengiriman.`);

    // 4. Ambil data yang diperlukan dari baris yang diedit
    // Kolom: B(Email), C(Nama), E(Qty), H(Inv Sent), I(Barcode Sent)
    // Relatif thd Kolom A: B=2, C=3, D=4, E=5, F(dihapus), G=7, H=8, I=9
    // Perlu baca range B:I (8 kolom)
    const rowDataRange = sheet.getRange(editedRowIndex, 2, 1, 8); // Baca Kolom B sampai I
    const rowData = rowDataRange.getValues()[0];

    const email = rowData[0];      // Index 0 => Kolom B
    const name = rowData[1];       // Index 1 => Kolom C
    // Kolom D di Index 2
    const ticketQty = rowData[3];  // Index 3 => Kolom E
    // Kolom G (Status Bayar) di Index 5 (tdk perlu dibaca lagi krn trigger value)
    const invoiceSent = rowData[6]; // Index 6 => Kolom H (Invoice Sent)
    const barcodeSent = rowData[7]; // Index 7 => Kolom I (Barcode Sent)

    // 5. Pastikan email ada dan belum pernah dikirim sebelumnya
    if (!email) {
      Logger.log(`WARN handleEdit: Tidak ada email di baris ${editedRowIndex}. Pengiriman dibatalkan.`);
      SpreadsheetApp.getActiveSpreadsheet().toast(`Baris ${editedRowIndex}: Gagal kirim otomatis, email kosong.`); // Feedback visual
      return;
    }
    if (invoiceSent === 'Ya' && barcodeSent === 'Ya') {
      Logger.log(`INFO handleEdit: Invoice/Barcode untuk ${email} di baris ${editedRowIndex} sudah terkirim sebelumnya.`);
      // SpreadsheetApp.getActiveSpreadsheet().toast(`Baris ${editedRowIndex}: Email sudah pernah dikirim.`); // Opsional: feedback jika sudah terkirim
      return;
    }

    // 6. Panggil fungsi pengiriman email
    Logger.log(`Memanggil sendInvoiceAndBarcode untuk baris ${editedRowIndex}, email: ${email}`);
    // Fungsi sendInvoiceAndBarcode sudah punya try-catch sendiri
    sendInvoiceAndBarcode(email, name, ticketQty, editedRowIndex);
    // Tidak perlu log sukses disini krn sdh ada di dlm sendInvoiceAndBarcode
    // SpreadsheetApp.getActiveSpreadsheet().toast(`Baris ${editedRowIndex}: Proses kirim email untuk ${email} dimulai.`); // Feedback visual

  } catch (error) { // Menangkap error tak terduga di handleEdit
      Logger.log(`FATAL ERROR handleEdit saat memproses baris ${e.range ? e.range.getRow() : 'tidak diketahui'}: ${error.message}\nStack: ${error.stack}`);
      // Mungkin beri notifikasi ke diri sendiri jika trigger gagal total
      // MailApp.sendEmail(Session.getEffectiveUser().getEmail(), "Error Trigger handleEdit", `Error: ${error.message}\nRange: ${e.range ? e.range.getA1Notation() : 'N/A'}`);
  }
}


//------------------------------------------------------------------
// FUNGSI PENDUKUNG (Pengiriman Email, Generate PDF, Barcode)
//------------------------------------------------------------------

/**
 * Fungsi untuk mengirim invoice dan barcode via email.
 * (Versi Revisi untuk Struktur Kolom Baru: Tanpa Kolom F)
 * @param {string} email Alamat email penerima.
 * @param {string} name Nama penerima.
 * @param {number} ticketQty Jumlah tiket.
 * @param {number} rowIndex Nomor baris di spreadsheet untuk update status.
 */
function sendInvoiceAndBarcode(email, name, ticketQty, rowIndex) {
  try {
    const barcodeBlob = generateBarcode(email); // Fungsi ini tidak berubah
    if (!barcodeBlob) {
        Logger.log(`Gagal membuat barcode untuk ${email} (baris ${rowIndex})`);
        // Mungkin tandai error di sheet
        // SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1').getRange(rowIndex, 10).setValue('Error Barcode');
        return;
    }

    const invoicePdf = generateInvoice(name, ticketQty, barcodeBlob); // Fungsi ini tidak berubah
    if (!invoicePdf) {
        Logger.log(`Gagal membuat PDF invoice untuk ${email} (baris ${rowIndex})`);
        // Mungkin tandai error di sheet
        // SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1').getRange(rowIndex, 10).setValue('Error PDF');
        return;
    }

    const subject = 'Invoice dan Barcode Tiket Growbalization PIB 2025';
    const body = `Halo ${name},\n\nTerima kasih telah melakukan pembayaran untuk acara Growbalization PIB 2025.\n\nBerikut terlampir Invoice dan Barcode (QR Code) untuk ${ticketQty} tiket Anda.\n\nMohon simpan email ini dan tunjukkan Barcode (QR Code) di lampiran saat registrasi ulang di lokasi acara.\n\nSampai jumpa di acara!\n\nSalam,\nPanitia Growbalization PIB 2025`;

    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: body,
      attachments: [invoicePdf],
      name: 'Panitia Growbalization PIB 2025'
    });

    Logger.log(`Email berhasil dikirim ke: ${email} (baris ${rowIndex})`);

    // Update kolom H dan I (baru) setelah berhasil mengirim email
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1');
    sheet.getRange(rowIndex, 8).setValue('Ya');  // Kolom H: Invoice yang Terkirim (Kolom ke-8)
    sheet.getRange(rowIndex, 9).setValue('Ya');  // Kolom I: Barcode yang Terkirim (Kolom ke-9)
    // Tambahkan timestamp pengiriman (opsional, misal di kolom J (baru) / kolom ke-10)
    // const timestampColIndex = 10;
    // sheet.getRange(rowIndex, timestampColIndex).setValue(new Date());
    SpreadsheetApp.getActiveSpreadsheet().toast(`Baris ${rowIndex}: Email invoice/barcode berhasil dikirim ke ${email}.`); // Feedback visual

  } catch (error) {
      Logger.log(`ERROR saat mengirim ke ${email} (baris ${rowIndex}): ${error.message}\nStack: ${error.stack}`);
      // Mungkin tandai error di sheet juga
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1');
      const errorColIndex = 10; // Kolom J (baru) untuk error
      try { // Gunakan try-catch lagi untuk setValue jika sheet error
          sheet.getRange(rowIndex, errorColIndex).setValue(`Error Kirim: ${error.message.substring(0,100)}`); // Batasi panjang pesan error
      } catch (e) {
           Logger.log(`Gagal menulis error ke sheet baris ${rowIndex}: ${e.message}`);
      }
      SpreadsheetApp.getActiveSpreadsheet().toast(`Baris ${rowIndex}: GAGAL kirim email ke ${email}. Cek log.`); // Feedback visual
  }
}

/**
 * Fungsi untuk generate QR code sebagai Blob. (TIDAK BERUBAH)
 * @param {string} data Teks yang akan di-encode dalam QR code (misal: email).
 * @return {Blob|null} Blob gambar PNG QR code atau null jika gagal.
 */
function generateBarcode(data) {
  if (!data) {
      Logger.log("Gagal generate barcode: Data (email) kosong.");
      return null;
  }
  try {
    const barcodeUrl = 'https://api.qrserver.com/v1/create-qr-code/?data=' + encodeURIComponent(data) + '&size=150x150&format=png';
    const response = UrlFetchApp.fetch(barcodeUrl, { muteHttpExceptions: true }); // muteHttpExceptions agar bisa cek status code

    if (response.getResponseCode() == 200) {
      const barcodeBlob = response.getBlob().setName(`barcode_${data}.png`);
      Logger.log(`Barcode berhasil dibuat untuk: ${data}`);
      return barcodeBlob;
    } else {
      Logger.log(`Gagal fetch barcode dari API untuk ${data}. Kode status: ${response.getResponseCode()}, Respon: ${response.getContentText()}`);
      return null;
    }
  } catch (error) {
    Logger.log(`Error saat generate barcode untuk ${data}: ${error.message}`);
    return null;
  }
}

/**
 * Fungsi untuk generate invoice + barcode dalam 1 file PDF. (TIDAK BERUBAH)
 * @param {string} name Nama peserta.
 * @param {number} ticketQty Jumlah tiket.
 * @param {Blob} barcodeBlob Blob gambar QR code.
 * @return {Blob|null} Blob file PDF atau null jika gagal.
 */
function generateInvoice(name, ticketQty, barcodeBlob) {
  let doc;
  try {
    // Validasi input dasar
    if (!name || !ticketQty || !barcodeBlob) {
        Logger.log(`Gagal generate invoice: Input tidak lengkap (Nama: ${name}, Jml Tiket: ${ticketQty}, Barcode: ${barcodeBlob ? 'Ada' : 'Tidak Ada'})`);
        return null;
    }

    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
    const docName = `Invoice_${name.replace(/[^a-zA-Z0-9]/g, '_')}_${timestamp}`;
    doc = DocumentApp.create(docName);
    const body = doc.getBody();
    const invoiceId = "INV-" + timestamp + "-" + Math.floor(Math.random() * 1000);

    // --- Styling ---
    const headingStyle = {};
    headingStyle[DocumentApp.Attribute.BOLD] = true;
    headingStyle[DocumentApp.Attribute.FONT_SIZE] = 14;
    headingStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
    const normalStyle = {};
    normalStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
    const boldStyle = {};
    boldStyle[DocumentApp.Attribute.BOLD] = true;
    boldStyle[DocumentApp.Attribute.FONT_SIZE] = 11;

    // --- Header Invoice ---
    body.appendParagraph('INVOICE').setHeading(DocumentApp.ParagraphHeading.HEADING1).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('Growbalization PIB 2025').setAttributes(headingStyle);
    body.appendParagraph('------------------------------------').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph(`No. Invoice: ${invoiceId}`).setAttributes(normalStyle);
    body.appendParagraph(`Tanggal: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMMM yyyy")}`).setAttributes(normalStyle); // Format tanggal diperbaiki
    body.appendParagraph(`Nama Pemesan: ${name}`).setAttributes(normalStyle);
    body.appendParagraph('\n');

    // --- Detail Tiket ---
    body.appendParagraph('Detail Pemesanan:').setAttributes(boldStyle);
    const hargaPerTiket = 50000;
    const totalHarga = ticketQty * hargaPerTiket;
    const tableData = [
        ['Deskripsi', 'Jumlah', 'Harga Satuan', 'Subtotal'],
        ['Tiket Growbalization PIB 2025', ticketQty.toString(), `Rp ${hargaPerTiket.toLocaleString('id-ID')}`, `Rp ${totalHarga.toLocaleString('id-ID')}`]
    ];
    const table = body.appendTable(tableData);
    const headerRow = table.getRow(0);
    for (let i = 0; i < headerRow.getNumCells(); i++) {
        headerRow.getCell(i).setBold(true).setBackgroundColor('#DDDDDD');
    }
    // Perbaiki alignment dan style total
    body.appendParagraph(`\nTotal Pembayaran:`).setAttributes(boldStyle).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    body.appendParagraph(`Rp ${totalHarga.toLocaleString('id-ID')}`).setAttributes(boldStyle).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    body.appendParagraph('Status: LUNAS').setAttributes(boldStyle).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    body.appendParagraph('------------------------------------').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('\n');

    // --- Barcode ---
    body.appendParagraph('Barcode Tiket (Tunjukkan saat Registrasi):').setAttributes(boldStyle).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    const paraBarcode = body.appendParagraph('');
    // Validasi barcodeBlob sebelum append
    if (barcodeBlob.getBytes().length > 0) { // Cek jika blob tidak kosong
        const img = paraBarcode.appendInlineImage(barcodeBlob);
        img.setWidth(150).setHeight(150);
        paraBarcode.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        Logger.log(`Barcode ditambahkan ke Dokumen untuk ${name}.`);
    } else {
        Logger.log(`WARN: Barcode blob kosong untuk ${name}, tidak ditambahkan ke dokumen.`);
        paraBarcode.appendText("Error: Gagal memuat barcode.").setAttributes(normalStyle).setForegroundColor("#FF0000"); // Pesan error di Doc
        paraBarcode.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    }
    body.appendParagraph('\n');

    // --- Footer ---
    body.appendParagraph('Terima kasih atas partisipasi Anda!').setAttributes(normalStyle).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('Panitia Growbalization PIB 2025').setAttributes(normalStyle).setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    // Simpan dan tutup dokumen
    doc.saveAndClose();
    const docId = doc.getId(); // Ambil ID sebelum diakses lagi
    Logger.log(`Dokumen ${docName} (ID: ${docId}) disimpan dan ditutup.`);

    // Konversi ke PDF
    const pdfBlob = DriveApp.getFileById(docId).getAs('application/pdf');
    pdfBlob.setName(docName + '.pdf');

    Logger.log(`PDF ${pdfBlob.getName()} berhasil dibuat.`);
    return pdfBlob;

  } catch (error) {
    Logger.log(`Error saat generate invoice untuk ${name}: ${error.message}\nStack: ${error.stack}`);
    return null;
  } finally {
    // Pembersihan file Google Docs sementara
    if (doc && doc.getId()) { // Cek jika doc dan ID nya ada
        const docId = doc.getId();
        try {
            Utilities.sleep(1000); // Beri jeda 1 detik sebelum hapus, jaga2 proses belum selesai
            const file = DriveApp.getFileById(docId);
             // Pastikan file tidak null dan bisa diakses (belum terhapus oleh proses lain)
            if (file && typeof file.setTrashed === 'function') {
                file.setTrashed(true);
                Logger.log(`Dokumen sementara (ID: ${docId}) dipindahkan ke sampah.`);
            } else {
                 Logger.log(`Dokumen sementara (ID: ${docId}) tidak ditemukan atau tidak bisa dihapus.`);
            }
        } catch (e) {
             Logger.log(`Gagal memindahkan dokumen sementara (ID: ${docId}) ke sampah: ${e.message}`);
        }
    }
  }
}


//------------------------------------------------------------------
// FUNGSI UNTUK EKSEKUSI MANUAL (Opsional)
//------------------------------------------------------------------
// Fungsi checkAndSendInvoice masih ada jika Anda perlu menjalankannya manual
// untuk memproses semua baris sekaligus (misalnya saat awal setup atau jika trigger gagal)
/**
 * Memeriksa status pembayaran dan mengirimkan invoice + barcode jika valid.
 * (Versi Revisi untuk Struktur Kolom Baru: Tanpa Kolom F)
 * Ini bisa dijalankan manual dari editor script.
 */
function checkAndSendInvoice() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1');
  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  const data = dataRange.getValues();

  Logger.log(`Memulai checkAndSendInvoice Manual. Jumlah baris data: ${data.length}`);
  let processedCount = 0;
  let skippedSentCount = 0;
  let skippedInvalidCount = 0;
  let skippedNoEmailCount = 0;
  let errorCount = 0;

  for (let i = 0; i < data.length; i++) {
    const rowIndex = i + 2;
    try {
      const row = data[i];
      const email = row[1];         // Kolom B: Email peserta (Index 1)
      const name = row[2];          // Kolom C: Nama peserta (Index 2)
      const ticketQty = row[4];     // Kolom E: Jumlah tiket (Index 4)
      const paymentStatus = row[6]; // REVISI -> Kolom G: Status pembayaran (Index 6)
      const invoiceSent = row[7];   // REVISI -> Kolom H: Invoice Terkirim? (Index 7)
      const barcodeSent = row[8];   // REVISI -> Kolom I: Barcode Terkirim? (Index 8)

      if (!email) {
        // Logger.log(`WARN Baris ${rowIndex}: Dilewati, tidak ada alamat email.`);
        skippedNoEmailCount++;
        continue; // Lanjut ke baris berikutnya
      }

      if (paymentStatus && paymentStatus.toString().toLowerCase().trim() === "sudah valid") {
         if (invoiceSent !== 'Ya' || barcodeSent !== 'Ya') {
            Logger.log(`Memproses baris ${rowIndex}: ${name} (${email}) - Status Valid, Belum terkirim.`);
            sendInvoiceAndBarcode(email, name, ticketQty, rowIndex);
            processedCount++;
            // Beri jeda sedikit antar email untuk menghindari limit MailApp
            if (processedCount % 10 === 0) { // Jeda setiap 10 email
               Utilities.sleep(1000);
            }
         } else {
            // Logger.log(`INFO Baris ${rowIndex}: Dilewati, Invoice/Barcode untuk ${email} sudah terkirim.`);
            skippedSentCount++;
         }
      } else {
        // Logger.log(`INFO Baris ${rowIndex}: Dilewati, Status pembayaran untuk ${email} bukan 'sudah valid' ("${paymentStatus}").`);
        skippedInvalidCount++;
      }

    } catch (error) {
      errorCount++;
      Logger.log(`ERROR di baris ${rowIndex} saat checkAndSendInvoice: ${error.message}\nStack: ${error.stack}`);
      // Mungkin tandai error di sheet juga
      // const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1');
      // const errorColIndex = 10; // Kolom J (baru)
      // sheet.getRange(rowIndex, errorColIndex).setValue(`Error Check: ${error.message.substring(0,100)}`);
    }
  }
  Logger.log(`CheckAndSendInvoice Manual Selesai. Diproses: ${processedCount}, Dilewati (Sudah Kirim): ${skippedSentCount}, Dilewati (Status Tdk Valid): ${skippedInvalidCount}, Dilewati (Tanpa Email): ${skippedNoEmailCount}, Error: ${errorCount}`);
  SpreadsheetApp.getActiveSpreadsheet().toast(`Proses manual selesai. ${processedCount} email diproses.`);
}