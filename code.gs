/**
 * ==============================================================================
 * PORTAL DMS BACKEND ENGINE (Google Apps Script)
 * Versi: 2.0
 * Deskripsi: Script ini bertindak sebagai API Server untuk Portal DMS Blogger/Frontend.
 * ==============================================================================
 */

// Nama-nama Tab Sheet di Google Sheets
const SHEET_MENUS = "Menus";
const SHEET_USERS = "Users";

// Header untuk Tab Menus
const MENUS_HEADERS = [
  "id",
  "parentId",
  "title",
  "category",
  "visibility",
  "type",
  "iconClass",
  "targetUrl",
  "description",
  "ownerUserId"
];

// Header untuk Tab Users
const USERS_HEADERS = [
  "userId",
  "username",
  "password",
  "fullName",
  "role"
];

/**
 * Handler Request HTTP GET
 * Menangani pengambilan data menu dan daftar pengguna
 */
function doGet(e) {
  try {
    ensureSheetsExist();
    
    const params = e.parameter || {};
    const action = params.action || "getMenus";
    
    let result = { status: "error", message: "Aksi GET tidak dikenal" };

    if (action === "getMenus") {
      const userId = params.userId || "PUBLIC";
      const role = params.role || "UMUM";
      const menus = fetchFilteredMenuTree(userId, role);
      result = { status: "success", data: menus };
    } 
    else if (action === "getUsers") {
      const users = fetchAllUsers();
      result = { status: "success", data: users };
    }

    return createJsonResponse(result);
  } catch (error) {
    return createJsonResponse({ status: "error", message: error.toString() });
  }
}

/**
 * Handler Request HTTP POST
 * Menangani Login, Simpan/Edit Menu, Hapus Menu, Simpan User, Hapus User
 */
function doPost(e) {
  try {
    ensureSheetsExist();

    let postData = {};
    if (e.postData && e.postData.contents) {
      postData = JSON.parse(e.postData.contents);
    }

    const action = postData.action || "";
    let result = { status: "error", message: "Aksi POST tidak dikenal" };

    switch (action) {
      case "login":
        result = processLogin(postData.username, postData.password);
        break;

      case "saveMenu":
        result = processSaveMenu(postData);
        break;

      case "deleteMenu":
        result = processDeleteMenu(postData.id, postData.userId, postData.role);
        break;

      case "saveUser":
        result = processSaveUser(postData);
        break;

      case "deleteUser":
        result = processDeleteUser(postData.userId);
        break;

      default:
        result = { status: "error", message: `Aksi '${action}' tidak didukung` };
        break;
    }

    return createJsonResponse(result);
  } catch (error) {
    return createJsonResponse({ status: "error", message: error.toString() });
  }
}

/**
 * Mengambil dan menyusun hirarki menu bertingkat (Tree) dengan filter visibilitas
 */
function fetchFilteredMenuTree(userId, role) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MENUS);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; // Hanya ada header

  const headers = data[0];
  const rawMenus = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue; // Lewati jika ID kosong

    const menu = {
      id: String(row[0]),
      parentId: String(row[1] || ""),
      title: String(row[2] || ""),
      category: String(row[3] || "General"),
      visibility: String(row[4] || "UMUM"),
      type: String(row[5] || "link"),
      iconClass: String(row[6] || "fas fa-link"),
      targetUrl: String(row[7] || "#"),
      description: String(row[8] || ""),
      ownerUserId: String(row[9] || "PUBLIC")
    };

    // Filter Visibilitas Akses
    if (canUserAccessMenu(menu, userId, role)) {
      rawMenus.push(menu);
    }
  }

  return buildMenuTree(rawMenus, "");
}

/**
 * Mengecek apakah pengguna berhak melihat menu tertentu
 */
function canUserAccessMenu(menu, userId, role) {
  // Admin selalu bisa melihat semua menu
  if (role === "ADMIN") return true;

  // Visibilitas UMUM dapat dilihat siapa saja
  if (menu.visibility === "UMUM") return true;

  // Visibilitas PRIVASI hanya untuk pemilik menu
  if (menu.visibility === "PRIVASI") {
    return userId !== "PUBLIC" && menu.ownerUserId === userId;
  }

  // Visibilitas PRIVASI_ADMIN untuk pemilik menu atau Admin
  if (menu.visibility === "PRIVASI_ADMIN") {
    return (userId !== "PUBLIC" && menu.ownerUserId === userId) || role === "ADMIN";
  }

  return false;
}

/**
 * Menyusun data array datar menjadi struktur pohon (Parent-Submenu)
 */
function buildMenuTree(items, parentId) {
  const branch = [];

  items.forEach(item => {
    if (item.parentId === parentId || (parentId === "" && (!item.parentId || item.parentId === "ROOT"))) {
      const children = buildMenuTree(items, item.id);
      if (children.length > 0) {
        item.submenus = children;
      } else {
        item.submenus = [];
      }
      branch.push(item);
    }
  });

  return branch;
}

/**
 * Menyimpan atau memperbarui data menu
 */
function processSaveMenu(payload) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MENUS);
  const data = sheet.getDataRange().getValues();

  const id = payload.id ? String(payload.id) : "MNU-" + Date.now();
  const parentId = payload.parentId ? String(payload.parentId) : "";
  const title = payload.title || "Menu Baru";
  const category = payload.category || "General";
  const visibility = payload.visibility || "UMUM";
  const type = payload.type || "link";
  const iconClass = payload.iconClass || "fas fa-link";
  const targetUrl = payload.targetUrl || "#";
  const description = payload.description || "";
  const ownerUserId = payload.ownerUserId || "PUBLIC";

  let rowIndex = -1;

  // Cari baris jika ID sudah ada (Update)
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === id) {
      rowIndex = i + 1;
      break;
    }
  }

  const rowData = [id, parentId, title, category, visibility, type, iconClass, targetUrl, description, ownerUserId];

  if (rowIndex > 0) {
    // Update Baris
    sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
  } else {
    // Tambah Baris Baru
    sheet.appendRow(rowData);
  }

  return { status: "success", message: "Menu berhasil disimpan", id: id };
}

/**
 * Menghapus menu beserta sub-menu bawaannya secara rekursif
 */
function processDeleteMenu(id, userId, role) {
  if (!id) return { status: "error", message: "ID Menu diperlukan" };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MENUS);
  let data = sheet.getDataRange().getValues();

  // Kumpulkan semua ID yang akan dihapus (termasuk anak/sub-menu)
  const idsToDelete = [String(id)];
  collectChildMenuIds(data, String(id), idsToDelete);

  let deletedCount = 0;
  // Hapus dari baris paling bawah agar indeks tidak bergeser
  for (let i = data.length - 1; i >= 1; i--) {
    const rowId = String(data[i][0]);
    if (idsToDelete.includes(rowId)) {
      sheet.deleteRow(i + 1);
      deletedCount++;
    }
  }

  return { status: "success", message: `${deletedCount} item menu berhasil dihapus` };
}

function collectChildMenuIds(data, parentId, accumulator) {
  for (let i = 1; i < data.length; i++) {
    const rowId = String(data[i][0]);
    const rowParentId = String(data[i][1]);

    if (rowParentId === parentId) {
      accumulator.push(rowId);
      collectChildMenuIds(data, rowId, accumulator);
    }
  }
}

/**
 * Proses Autentikasi Pengguna (Login)
 */
function processLogin(username, password) {
  if (!username || !password) {
    return { status: "error", message: "Username dan password wajib diisi." };
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const u = String(data[i][1] || "").trim();
    const p = String(data[i][2] || "").trim();

    if (u.toLowerCase() === username.trim().toLowerCase() && p === password.trim()) {
      return {
        status: "success",
        user: {
          userId: String(data[i][0]),
          username: String(data[i][1]),
          fullName: String(data[i][3]),
          role: String(data[i][4])
        }
      };
    }
  }

  return { status: "error", message: "Username atau password salah." };
}

/**
 * Mengambil seluruh daftar pengguna
 */
function fetchAllUsers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  const users = [];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    users.push({
      userId: String(data[i][0]),
      username: String(data[i][1]),
      fullName: String(data[i][3]),
      role: String(data[i][4])
    });
  }

  return users;
}

/**
 * Menyimpan atau memperbarui data pengguna
 */
function processSaveUser(payload) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  const data = sheet.getDataRange().getValues();

  const userId = payload.userId ? String(payload.userId).trim() : "USR-" + Date.now();
  const username = payload.username ? String(payload.username).trim() : "";
  const password = payload.password ? String(payload.password).trim() : "";
  const fullName = payload.fullName ? String(payload.fullName).trim() : "";
  const role = payload.role || "USER_KHUSUS";

  let rowIndex = -1;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === userId) {
      rowIndex = i + 1;
      break;
    }
  }

  const rowData = [userId, username, password, fullName, role];

  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }

  return { status: "success", message: "Data user berhasil disimpan", userId: userId };
}

/**
 * Menghapus pengguna berdasarkan userId
 */
function processDeleteUser(userId) {
  if (!userId) return { status: "error", message: "User ID diperlukan" };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  const data = sheet.getDataRange().getValues();

  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === userId) {
      sheet.deleteRow(i + 1);
      return { status: "success", message: "User berhasil dihapus" };
    }
  }

  return { status: "error", message: "User ID tidak ditemukan" };
}

/**
 * Memastikan Tab 'Menus' dan 'Users' tersedia di Google Sheets
 */
function ensureSheetsExist() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Cek & Inisialisasi Sheet Menus
  let menuSheet = ss.getSheetByName(SHEET_MENUS);
  if (!menuSheet) {
    menuSheet = ss.insertSheet(SHEET_MENUS);
    menuSheet.appendRow(MENUS_HEADERS);
    menuSheet.getRange(1, 1, 1, MENUS_HEADERS.length).setFontWeight("bold").setBackground("#1f2937").setFontColor("#ffffff");
    
    // Data contoh awal
    menuSheet.appendRow(["MNU-001", "", "Dashboard Utama", "General", "UMUM", "link", "fas fa-gauge", "https://google.com", "Halaman Utama Dashboard", "ADM-01"]);
  }

  // 2. Cek & Inisialisasi Sheet Users
  let userSheet = ss.getSheetByName(SHEET_USERS);
  if (!userSheet) {
    userSheet = ss.insertSheet(SHEET_USERS);
    userSheet.appendRow(USERS_HEADERS);
    userSheet.getRange(1, 1, 1, USERS_HEADERS.length).setFontWeight("bold").setBackground("#1f2937").setFontColor("#ffffff");
    
    // Akun default admin jika sheet baru dibuat
    userSheet.appendRow(["ADM-01", "admin", "admin123", "Administrator Portal", "ADMIN"]);
  }
}

/**
 * Format Response JSON standar untuk API
 */
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
