// ==========================================
// CONFIGURATION (Sesuaikan dengan akun Anda)
// ==========================================
const SPREADSHEET_ID = "11JzRAf3TAXulLYk8Sckq8I8SBgiQjOWWKUd0no_5iq8"; 
const DATA_SHEET_NAME = "LaporanPenukaran";
const SETTINGS_SHEET_NAME = "Pengaturan";
const ADMIN_USERS_SHEET_NAME = "AdminUsers";   
const FOLDER_VOUCHER_ID = "1TkvEludNNS9TUotgeWqrxTBOFviAvtOx";   
const FOLDER_PENERIMA_ID = "1XzkCnyN9xDZoN3tV-webx_MtaSHhFXf0"; 

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Sistem Penukaran Voucher Mall')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==========================================
// LOGIKA AUTENTIKASI LOGIN TERINTEGRASI
// ==========================================
function checkUserLogin(email, password) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(ADMIN_USERS_SHEET_NAME);
    
    if (!sheet) {
      return { success: false, message: "Sistem Error: Sheet 'AdminUsers' tidak ditemukan!" };
    }
    
    const data = sheet.getDataRange().getValues();
    const cleanEmail = email.toLowerCase().trim();
    
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0] || !data[i][1]) continue; 
      
      const sheetEmail = data[i][0].toString().trim().toLowerCase();
      const sheetPassword = data[i][1].toString().trim();
      const sheetRole = data[i][2] ? data[i][2].toString().trim().toLowerCase() : "cs";
      const sheetNama = data[i][3] ? data[i][3].toString().trim() : sheetEmail;
      
      if (cleanEmail === sheetEmail && password === sheetPassword) {
        return { 
          success: true, 
          message: "Login Berhasil", 
          role: sheetRole, 
          userEmail: sheetNama 
        };
      }
    }
    
    return { success: false, message: "Email atau Password salah!" };
    
  } catch (error) {
    return { success: false, message: "Gagal terhubung ke database: " + error.toString() };
  }
}

// ==========================================
// CORE LOGIC: INPUT DATA LAPORAN & UPLOAD (OPTIMIZED WITH LOCK)
// ==========================================
function submitCSReport(payload) {
  // Menggunakan LockService mencegah tabrakan kuota voucher jika disubmit bersamaan
  const lock = LockService.getScriptLock();
  try {
    // Tunggu maksimal 30 detik untuk mendapatkan giliran jika server sibuk
    lock.waitLock(30000); 

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(DATA_SHEET_NAME);
    
    if (!sheet) {
      sheet = ss.insertSheet(DATA_SHEET_NAME);
      sheet.appendRow(["ID Transaksi", "Timestamp", "Nama CS", "Nama Pelanggan", "No HP/Kontak", "Jenis Voucher", "Kode Voucher", "ID Foto Voucher", "ID Foto Penerima", "Catatan"]);
    } else if (sheet.getLastRow() === 0) {
      sheet.appendRow(["ID Transaksi", "Timestamp", "Nama CS", "Nama Pelanggan", "No HP/Kontak", "Jenis Voucher", "Kode Voucher", "ID Foto Voucher", "ID Foto Penerima", "Catatan"]);
    }

    const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    if (settingsSheet) {
      const settingsData = settingsSheet.getDataRange().getValues();
      let voucherKetemu = false;
      let sisaKuota = 0;
      let barisVoucher = -1;

      for (let i = 1; i < settingsData.length; i++) {
        if (settingsData[i][1] && settingsData[i][1].toString().trim() === payload.jenisVoucher.trim()) {
          voucherKetemu = true;
          sisaKuota = Number(settingsData[i][2] || 0);
          barisVoucher = i + 1; 
          break;
        }
      }

      if (voucherKetemu && sisaKuota <= 0) {
        return { success: false, message: `Gagal! Kuota untuk voucher "${payload.jenisVoucher}" sudah habis.` };
      }

      if (voucherKetemu && barisVoucher !== -1) {
        settingsSheet.getRange(barisVoucher, 3).setValue(sisaKuota - 1);
      }
    }

    const trxId = "TX-" + Array.from({length: 8}, () => Math.floor(Math.random()*16).toString(16)).join('').toUpperCase();
    const timestamp = new Date();

    let fileVoucherId = "";
    if (payload.fotoVoucherBase64) {
      fileVoucherId = uploadToDrive(payload.fotoVoucherBase64, `VCH-${payload.kodeVoucher}-${trxId}.jpg`, FOLDER_VOUCHER_ID);
    }

    let filePenerimaId = "";
    if (payload.fotoPenerimaBase64) {
      filePenerimaId = uploadToDrive(payload.fotoPenerimaBase64, `PNR-${payload.namaPelanggan.replace(/\s+/g, '_')}-${trxId}.jpg`, FOLDER_PENERIMA_ID);
    }

    sheet.appendRow([
      trxId,
      timestamp,
      payload.namaCs, 
      payload.namaPelanggan,
      "'" + payload.kontak, 
      payload.jenisVoucher,
      payload.kodeVoucher,
      fileVoucherId,
      filePenerimaId,
      payload.catatan || "-"
    ]);

    return { success: true, message: `Sukses! Laporan berhasil disimpan dengan ID: ${trxId}` };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    // Selalu lepaskan kunci lock setelah proses selesai / error terjadi
    lock.releaseLock();
  }
}

function uploadToDrive(base64Data, fileName, folderId) {
  const splitData = base64Data.split(',');
  const contentType = splitData[0].match(/:(.*?);/)[1];
  const bytes = Utilities.base64Decode(splitData[1]);
  const blob = Utilities.newBlob(bytes, contentType, fileName);
  
  const folder = DriveApp.getFolderById(folderId);
  const file = folder.createFile(blob);
  
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch(e) {
    Logger.log("Gagal mengubah izin akses file: " + e.toString());
  }
  
  return file.getId();
}

function getFormDataOptions() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    if (!sheet) {
      return { success: true, options: { voucherList: [] } };
    }
    
    const data = sheet.getDataRange().getValues();
    let voucherList = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1]) {
        voucherList.push({
          nama: data[i][1].toString(),
          qty: data[i][2] ? Number(data[i][2]) : 0
        });
      }
    }
    
    return { success: true, options: { voucherList: voucherList } };
  } catch(e) {
    return { success: false, options: { voucherList: [] } };
  }
}

// ==========================================
// DASHBOARD DATA FOR ADMIN MONITOR
// ==========================================
function getAdminDashboardData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(DATA_SHEET_NAME);
    const userSheet = ss.getSheetByName(ADMIN_USERS_SHEET_NAME);
    const settings = getFormDataOptions().options;
    
    let userList = [];
    if (userSheet && userSheet.getLastRow() > 1) {
      const uData = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 4).getValues();
      uData.forEach(r => {
        if(r[0]) {
          userList.push({ email: r[0], role: r[2] || "cs", nama: r[3] || "" });
        }
      });
    }

    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: true, data: [], stats: { total: 0, topCs: "-", topVoucher: "-" }, settings, userList };
    }

    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
    const today = new Date();
    today.setHours(0,0,0,0);

    let filteredData = [];
    let csCounts = {};
    let voucherCounts = {};

    rows.forEach(row => {
      if (!row[1]) return;
      const rowDate = new Date(row[1]);
      rowDate.setHours(0,0,0,0);

      if (rowDate.getTime() === today.getTime()) {
        const timeStr = Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), "HH:mm");
        
        filteredData.unshift({
          id: row[0],
          timestamp: timeStr,
          namaCs: row[2],
          namaPelanggan: row[3],
          kontak: row[4],
          jenisVoucher: row[5],
          kodeVoucher: row[6],
          fotoVoucherId: row[7] ? cleanDriveId(row[7].toString()) : "", 
          fotoPenerimaId: row[8] ? cleanDriveId(row[8].toString()) : "",
          catatan: row[9]
        });

        if (row[2]) csCounts[row[2]] = (csCounts[row[2]] || 0) + 1;
        if (row[5]) voucherCounts[row[5]] = (voucherCounts[row[5]] || 0) + 1;
      }
    });

    let topCs = "-"; let maxCs = 0;
    for (let k in csCounts) { if(csCounts[k] > maxCs) { maxCs = csCounts[k]; topCs = k; } }

    let topVoucher = "-"; let maxVch = 0;
    for (let k in voucherCounts) { if(voucherCounts[k] > maxVch) { maxVch = voucherCounts[k]; topVoucher = k; } }

    return {
      success: true,
      data: filteredData,
      stats: { total: filteredData.length, topCs: topCs + (maxCs > 0 ? ` (${maxCs} Trx)` : ""), topVoucher: topVoucher },
      settings: settings,
      userList: userList
    };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function getFormImageBase64(rowId, columnType) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(DATA_SHEET_NAME);
    const rows = sheet.getDataRange().getValues();
    
    let fileId = "";
    const colIdx = (columnType === "voucher") ? 7 : 8;

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === rowId) { fileId = cleanDriveId(rows[i][colIdx].toString()); break; }
    }

    if (!fileId) return "";
    const file = DriveApp.getFileById(fileId);
    return "data:" + file.getMimeType() + ";base64," + Utilities.base64Encode(file.getBlob().getBytes());
  } catch (e) {
    return "ERROR: " + e.toString();
  }
}

function cleanDriveId(inputStr) {
  if (!inputStr) return "";
  let txt = inputStr.trim();
  if (txt.includes("id=")) {
    let match = txt.match(/id=([^&]+)/);
    if (match) return match[1];
  }
  if (txt.includes("/d/")) {
    let match = txt.match(/\/d\/([^/]+)/);
    if (match) return match[1];
  }
  return txt.replace(/['"\[\]]/g, ""); 
}

// ==========================================
// MASTER EXPORT: MURNI DATA TEKS
// ==========================================
function exportLogToExcelWithImages() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const logSheet = ss.getSheetByName(DATA_SHEET_NAME);
    
    if (!logSheet || logSheet.getLastRow() <= 1) {
      return "ERROR: Tidak ada data transaksi untuk diekspor.";
    }
    
    const data = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 10).getValues();
    
    const reportSs = SpreadsheetApp.create("Laporan_Voucher_Murni_" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd"));
    const sheet = reportSs.getSheets()[0];
    sheet.setName("Data Penukaran");
    sheet.setHiddenGridlines(false); 
    
    const titleRange = sheet.getRange("A1:G1");
    titleRange.merge();
    titleRange.setValue("LAPORAN DATA PENUKARAN VOUCHER MALL");
    titleRange.setFontSize(14);
    titleRange.setFontWeight("bold");
    titleRange.setHorizontalAlignment("center");
    titleRange.setVerticalAlignment("middle");
    titleRange.setBackground("#0f172a");
    titleRange.setFontColor("#ffffff");
    
    const subtitleRange = sheet.getRange("A2:G2");
    subtitleRange.merge();
    subtitleRange.setValue("Waktu Cetak Laporan: " + Utilities.formatDate(new Date(), "GMT+7", "dd-MM-yyyy HH:mm") + " WIB");
    subtitleRange.setFontSize(10);
    subtitleRange.setFontStyle("italic");
    subtitleRange.setHorizontalAlignment("center");
    subtitleRange.setVerticalAlignment("middle");
    subtitleRange.setBackground("#f1f5f9");
    subtitleRange.setFontColor("#475569");
    
    sheet.setRowHeight(1, 40);
    sheet.setRowHeight(2, 25);
    
    const headers = ["ID Transaksi", "Tanggal & Waktu", "Nama CS", "Nama Pelanggan", "No HP / Kontak", "Jenis Voucher", "Kode Voucher"];
    const headerRange = sheet.getRange(4, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#cbd5e1");
    headerRange.setHorizontalAlignment("center");
    headerRange.setVerticalAlignment("middle");
    sheet.setRowHeight(4, 30);
    
    sheet.setColumnWidth(1, 120); 
    sheet.setColumnWidth(2, 150); 
    sheet.setColumnWidth(3, 120); 
    sheet.setColumnWidth(4, 150); 
    sheet.setColumnWidth(5, 130); 
    sheet.setColumnWidth(6, 180); 
    sheet.setColumnWidth(7, 130); 
    
    let targetRow = 5;
    
    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      if (!r[0]) continue; 
      
      sheet.getRange(targetRow, 1).setValue(r[0]); 
      
      if (r[1] instanceof Date) {
        sheet.getRange(targetRow, 2).setValue(Utilities.formatDate(r[1], "GMT+7", "dd-MM-yyyy HH:mm"));
      } else {
        sheet.getRange(targetRow, 2).setValue(r[1]);
      }
      
      sheet.getRange(targetRow, 3).setValue(r[2]); 
      sheet.getRange(targetRow, 4).setValue(r[3]); 
      sheet.getRange(targetRow, 5).setNumberFormat("@").setValue(r[4]); 
      sheet.getRange(targetRow, 6).setValue(r[5]); 
      sheet.getRange(targetRow, 7).setValue(r[6]); 
      
      sheet.setRowHeight(targetRow, 22); 
      targetRow++;
    }
    
    if (targetRow > 5) {
      const totalRangeData = sheet.getRange(4, 1, targetRow - 4, 7);
      totalRangeData.setBorder(true, true, true, true, true, true, "#94a3b8", SpreadsheetApp.BorderStyle.SOLID);
      totalRangeData.setVerticalAlignment("middle");
      
      sheet.getRange(5, 1, targetRow - 5, 2).setHorizontalAlignment("center");
      sheet.getRange(5, 3, targetRow - 5, 2).setHorizontalAlignment("left");   
      sheet.getRange(5, 5, targetRow - 5, 3).setHorizontalAlignment("center"); 
    }
    
    SpreadsheetApp.flush();
    return "https://docs.google.com/spreadsheets/d/" + reportSs.getId() + "/export?format=xlsx";
    
  } catch(e) {
    return "ERROR backend: " + e.toString();
  }
}

// ==========================================
// PENGATURAN PANEL CONFIGURATION & RESTOCK
// ==========================================
function addSettingItem(type, payloadObj) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SETTINGS_SHEET_NAME) || ss.insertSheet(SETTINGS_SHEET_NAME);
  if(sheet.getLastRow() === 0) sheet.appendRow(["Daftar Nama CS", "Daftar Jenis Voucher", "Jumlah Kuota"]);
  
  if (type === 'voucher') {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 2).setValue(payloadObj.nama);
    sheet.getRange(lastRow + 1, 3).setValue(payloadObj.qty);
  } 
  else if (type === 'voucher_restock') {
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: "Sheet kosong!" };
    const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 2).getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0].toString().trim() === payloadObj.nama.trim()) {
        const kuotaLama = Number(data[i][1] || 0);
        sheet.getRange(i + 2, 3).setValue(kuotaLama + Number(payloadObj.qty));
        break;
      }
    }
  }
  return { success: true };
}

function deleteSettingItem(type, value) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() <= 1) return { success: false };
  
  const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0].toString() === value.toString()) {
      sheet.deleteRow(i + 2);
      break;
    }
  }
  return { success: true };
}

function adminAddUserAccount(email, password, role, nama) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(ADMIN_USERS_SHEET_NAME) || ss.insertSheet(ADMIN_USERS_SHEET_NAME);
    if(sheet.getLastRow() === 0) sheet.appendRow(["Email", "Password", "Role", "Nama Lengkap"]);
    
    sheet.appendRow([email.toLowerCase().trim(), password, role, nama.trim()]);
    return { success: true, message: "Akun baru berhasil didaftarkan!" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function adminDeleteUserAccount(email) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(ADMIN_USERS_SHEET_NAME);
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: "Sheet kosong" };
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase() === email.toLowerCase()) {
        sheet.deleteRow(i + 2);
        break;
      }
    }
    return { success: true, message: "Akun berhasil dihapus." };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}