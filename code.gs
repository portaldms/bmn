/**
 * ====================================================================
 * GAS WebApp Expert - Portal Dashboard & Link Manager (Production Backend)
 * File: Code.gs
 * ====================================================================
 */

const SHEET_DATA = "DashboardData";
const SHEET_SETTINGS = "Settings";

/**
 * Endpoint Utama Web App (Rendering HTML)
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("Portal Dashboard Terpadu")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Mengambil Seluruh Data Dashboard & Settings (Batch Operation Fast Read)
 */
function getDashboardPayload() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Inisialisasi Sheet Settings jika belum ada
    let settingSheet = ss.getSheetByName(SHEET_SETTINGS);
    if (!settingSheet) {
      settingSheet = ss.insertSheet(SHEET_SETTINGS);
      settingSheet.appendRow(["Key", "Value"]);
      settingSheet.appendRow(["ADMIN_PIN", "123456"]);
      settingSheet.appendRow(["APP_TITLE", "Portal Terpadu Apps"]);
      settingSheet.appendRow(["APP_SUBTITLE", "Akses Cepat Layanan & Informasi Perusahaan"]);
    }
    const settingValues = settingSheet.getDataRange().getValues();
    const settings = {};
    for (let i = 1; i < settingValues.length; i++) {
      if (settingValues[i][0]) {
        settings[settingValues[i][0]] = settingValues[i][1];
      }
    }

    // Inisialisasi Sheet DashboardData jika belum ada
    let dataSheet = ss.getSheetByName(SHEET_DATA);
    if (!dataSheet) {
      dataSheet = ss.insertSheet(SHEET_DATA);
      dataSheet.appendRow(["ID", "Category", "Title", "Subtitle", "URL", "Icon", "Badge", "Order", "Status"]);
      dataSheet.appendRow(["MENU_1", "Operasional", "Form Permohonan Cuti", "Pengajuan cuti tahunan, sakit, dan izin pegawai", "https://docs.google.com", "fa-calendar-check", "Populer", 1, "Active"]);
      dataSheet.appendRow(["MENU_2", "Keuangan", "Klaim Reimbursment", "Pengajuan klaim dana operasional & perjalanan dinas", "https://sheets.google.com", "fa-receipt", "Keuangan", 2, "Active"]);
    }
    
    const dataValues = dataSheet.getDataRange().getValues();
    const items = [];
    
    for (let i = 1; i < dataValues.length; i++) {
      const row = dataValues[i];
      if (row[0]) {
        items.push({
          id: row[0].toString(),
          category: row[1] ? row[1].toString() : 'Umum',
          title: row[2] ? row[2].toString() : '',
          subtitle: row[3] ? row[3].toString() : '',
          url: row[4] ? row[4].toString() : '#',
          icon: row[5] ? row[5].toString() : 'fa-link',
          badge: row[6] ? row[6].toString() : '',
          order: Number(row[7]) || 0,
          status: row[8] ? row[8].toString() : 'Active'
        });
      }
    }

    items.sort((a, b) => a.order - b.order);

    return { success: true, settings: settings, items: items };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * Verifikasi PIN Admin (Server-side Validation)
 */
function verifyAdminPin(inputPin) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingSheet = ss.getSheetByName(SHEET_SETTINGS);
    if (!settingSheet) return { success: false, message: "Sheet Settings belum siap." };
    const data = settingSheet.getDataRange().getValues();
    
    let storedPin = "123456";
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === "ADMIN_PIN") {
        storedPin = data[i][1].toString();
        break;
      }
    }

    if (inputPin.toString() === storedPin) {
      return { success: true, token: "ADMIN_AUTH_SUCCESS_" + new Date().getTime() };
    } else {
      return { success: false, message: "PIN Admin tidak valid!" };
    }
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * Menyimpan / Mengubah Item Menu (Admin Only Validation)
 */
function saveMenuItem(pinInput, itemData) {
  const auth = verifyAdminPin(pinInput);
  if (!auth.success) return { success: false, message: "Akses ditolak: PIN tidak valid!" };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_DATA);
    const data = sheet.getDataRange().getValues();

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === itemData.id.toString()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex > -1) {
      sheet.getRange(rowIndex, 1, 1, 9).setValues([[
        itemData.id, itemData.category, itemData.title, itemData.subtitle,
        itemData.url, itemData.icon, itemData.badge || "", itemData.order || 1, itemData.status || "Active"
      ]]);
    } else {
      const newId = itemData.id || ("MENU_" + new Date().getTime());
      sheet.appendRow([
        newId, itemData.category, itemData.title, itemData.subtitle,
        itemData.url, itemData.icon, itemData.badge || "", itemData.order || (data.length), itemData.status || "Active"
      ]);
    }
    return { success: true, message: "Menu berhasil disimpan ke Google Sheets!" };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * Mengubah Nama Kategori Secara Masal (Batch Write - Super Fast)
 */
function renameCategory(pinInput, oldCategory, newCategory) {
  const auth = verifyAdminPin(pinInput);
  if (!auth.success) return { success: false, message: "Akses ditolak." };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_DATA);
    const range = sheet.getDataRange();
    const data = range.getValues();

    let count = 0;
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === oldCategory) {
        data[i][1] = newCategory;
        count++;
      }
    }
    
    // Tulis seluruh sheet sekaligus dalam 1 kali API call (Batch Operation)
    if (count > 0) {
      range.setValues(data);
    }
    return { success: true, message: `Berhasil memperbarui ${count} menu!` };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * Menghapus Item Menu
 */
function deleteMenuItem(pinInput, itemId) {
  const auth = verifyAdminPin(pinInput);
  if (!auth.success) return { success: false, message: "Akses ditolak." };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_DATA);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === itemId.toString()) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "Menu berhasil dihapus dari Google Sheets!" };
      }
    }
    return { success: false, message: "Item tidak ditemukan." };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}
