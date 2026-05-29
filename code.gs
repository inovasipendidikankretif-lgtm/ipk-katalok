const SS_ID = "1kCVxbwRvFSjVXV60efxXaaV1ECuf6GoVoE-BBHSOeD8"; 
const FOLDER_COVER_ID = "1SOogj1x50sDAkXQ-XODrpkxOB7XKiT-T"; 
const ADMIN_PASS = "46035";

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Katalog Buku - IPK')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Fungsi Otentikasi Admin di sisi Server
function verifyAdminPassword(inputPass) {
  return inputPass === ADMIN_PASS;
}

// Menarik data dari Google Sheets dengan mapping header asli
function getSheetData() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return [];

    const headers = data[0].map(h => h.toString().trim());
    const rows = data.slice(1);

    // Menyisipkan rowNum asli untuk kebutuhan tracking Edit/Hapus data
    return rows.map((row, index) => {
      let obj = { rowNum: index + 2 }; 
      headers.forEach((header, i) => {
        obj[header] = row[i] !== undefined ? row[i] : "";
      });
      return obj;
    }).reverse(); // Buku terbaru otomatis tampil di urutan paling atas katalog
  } catch (e) { throw new Error(e.message); }
}

// Fungsi Menyimpan Data (Menangani Input Buku Baru & Perubahan/Edit Buku)
function saveBookData(formData) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheets()[0];
    let fileId = formData.idCover || "";

    // Memproses konversi file gambar cover ke Google Drive jika diunggah berkas baru
    if (formData.fileData && formData.fileData.includes(",")) {
      const folder = DriveApp.getFolderById(FOLDER_COVER_ID);
      const bytes = Utilities.base64Decode(formData.fileData.split(',')[1]);
      const blob = Utilities.newBlob(bytes, formData.fileType, formData.fileName);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileId = file.getId();
    }

    const rowData = [
      formData.judul, 
      formData.penulis, 
      formData.penerbit, 
      formData.tahun, 
      formData.halaman, 
      formData.isbn, 
      formData.kategori, 
      formData.deskripsi, 
      fileId, 
      formData.harga
    ];

    if (formData.rowNum && formData.rowNum !== "") {
      // JIKA MODE EDIT: Perbarui baris yang sudah ada
      const range = sheet.getRange(Number(formData.rowNum), 1, 1, rowData.length);
      range.setValues([rowData]);
    } else {
      // JIKA MODE TAMBAH BUKU BARU: Tambah baris baru otomatis di Google Sheets
      sheet.appendRow(rowData);
    }
    return { success: true };
  } catch (e) { return { success: false, message: e.toString() }; }
}

// Fungsi Menghapus Baris Data Buku
function deleteBookData(rowNum) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheets()[0];
    sheet.deleteRow(Number(rowNum));
    return { success: true };
  } catch (e) { return { success: false, message: e.toString() }; }
}