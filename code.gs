const SS_ID = "1kCVxbwRvFSjVXV60efxXaaV1ECuf6GoVoE-BBHSOeD8"; 
const FOLDER_COVER_ID = "1SOogj1x50sDAkXQ-XODrpkxOB7XKiT-T"; 
const ADMIN_PASS = "46035";

// Tetap dipertahankan jika Anda ingin membuka langsung dari link Google Web App
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Katalog Buku - IPK')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Pintu masuk utama untuk request dari GitHub Pages via Fetch API
function doPost(e) {
  try {
    const requestData = JSON.parse(e.postData.contents);
    const action = requestData.action;
    let result = {};

    if (action === "getSheetData") {
      result = getSheetData();
    } else if (action === "verifyAdminPassword") {
      result = verifyAdminPassword(requestData.pass);
    } else if (action === "saveBookData") {
      result = saveBookData(requestData.formData);
    } else if (action === "deleteBookData") {
      result = deleteBookData(requestData.rowNum);
    } else {
      result = { success: false, message: "Aksi tidak dikenali." };
    }

    // Mengembalikan output dalam bentuk JSON agar diizinkan oleh CORS Browser di GitHub
    return ContentService.createTextOutput(JSON.stringify(result))
                         .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: error.toString() }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

function verifyAdminPassword(inputPass) {
  return inputPass === ADMIN_PASS;
}

function getSheetData() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return [];

    const headers = data[0].map(h => h.toString().trim());
    const rows = data.slice(1);

    return rows.map((row, index) => {
      let obj = { rowNum: index + 2 }; 
      headers.forEach((header, i) => {
        obj[header] = row[i] !== undefined ? row[i] : "";
      });
      return obj;
    }).reverse();
  } catch (e) { 
    return { success: false, message: e.message };
  }
}

function saveBookData(formData) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheets()[0];
    let fileId = formData.idCover || "";

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
      const range = sheet.getRange(Number(formData.rowNum), 1, 1, rowData.length);
      range.setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
    return { success: true };
  } catch (e) { 
    return { success: false, message: e.toString() };
  }
}

function deleteBookData(rowNum) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheets()[0];
    sheet.deleteRow(Number(rowNum));
    return { success: true };
  } catch (e) { 
    return { success: false, message: e.toString() }; 
  }
}
