/**
 * HIGH-PERFORMANCE BACKEND ENGINE WITH MANUAL USER ID - GAS-Bloger V2 II
 */

const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const SHEET_MENUS = "DB_MENUS";
const SHEET_USERS = "DB_USERS";

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const action = e.parameter.action || "getMenus";
  const userId = e.parameter.userId || "PUBLIC";
  const role = e.parameter.role || "UMUM";

  try {
    if (action === "getMenus") {
      const data = getMenusHierarchicalOptimized(userId, role);
      return createJsonResponse({ status: "success", data: data });
    }
    else if (action === "getUsers") {
      if (role !== "ADMIN") throw new Error("Akses Ditolak: Hanya ADMIN yang berhak mengelola user.");
      const users = getUsersList();
      return createJsonResponse({ status: "success", data: users });
    }
    return createJsonResponse({ status: "error", message: "Action doGet tidak dikenal." });
  } catch (err) {
    return createJsonResponse({ status: "error", message: err.toString() });
  }
}

function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;

    if (action === "login") {
      const auth = authenticateUser(postData.username, postData.password);
      return createJsonResponse(auth);
    }

    const authCheck = validateSession(postData.sessionUserId);
    if (!authCheck.valid) {
      return createJsonResponse({ status: "error", message: "Sesi kadaluarsa. Silakan login kembali." });
    }

    const sessionUser = authCheck.user;

    // --- FITUR MENU MANAGEMENT ---
    if (action === "saveMenu") {
      if (sessionUser.role === "USER_KHUSUS") {
        if (postData.payload.id) checkOwnership(postData.payload.id, sessionUser.userId);
        postData.payload.ownerUserId = sessionUser.userId;
      } else if (sessionUser.role === "ADMIN") {
        postData.payload.ownerUserId = postData.payload.ownerUserId || sessionUser.userId;
      }

      const result = saveOrUpdateMenu(postData.payload);
      return createJsonResponse({ status: "success", data: result });
    } 
    else if (action === "deleteMenu") {
      if (sessionUser.role === "USER_KHUSUS") checkOwnership(postData.id, sessionUser.userId);
      deleteMenuRecursive(postData.id);
      return createJsonResponse({ status: "success", message: "Menu berhasil dihapus." });
    }

    // --- FITUR USER MANAGEMENT (KHUSUS ADMIN) ---
    else if (action === "saveUser") {
      if (sessionUser.role !== "ADMIN") throw new Error("Akses Ditolak: Fitur ini khusus ADMIN.");
      const result = saveOrUpdateUser(postData.payload);
      return createJsonResponse({ status: "success", data: result });
    }
    else if (action === "deleteUser") {
      if (sessionUser.role !== "ADMIN") throw new Error("Akses Ditolak: Fitur ini khusus ADMIN.");
      if (sessionUser.userId === postData.targetUserId) throw new Error("Ditolak: Anda tidak dapat menghapus akun Anda sendiri yang sedang aktif.");
      deleteUserRecord(postData.targetUserId);
      return createJsonResponse({ status: "success", message: "User berhasil dihapus." });
    }

    return createJsonResponse({ status: "error", message: "Action doPost tidak dikenal." });
  } catch (err) {
    return createJsonResponse({ status: "error", message: err.toString() });
  }
}

// In-Memory Fast Reading for Menus
function getMenusHierarchicalOptimized(userId, role) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const rows = ss.getSheetByName(SHEET_MENUS).getDataRange().getValues();
  if (rows.length <= 1) return [];

  const rawList = [];
  const len = rows.length;

  for (let i = 1; i < len; i++) {
    const r = rows[i];
    if (!r[0]) continue;

    const ownerUserId = String(r[8] || "PUBLIC");
    const visibility = String(r[9] || "UMUM");

    let visible = false;
    if (visibility === "UMUM") {
      visible = true;
    } else if (visibility === "PRIVASI") {
      if (userId !== "PUBLIC" && ownerUserId === userId) visible = true;
    } else if (visibility === "PRIVASI_ADMIN") {
      if (role === "ADMIN" || (userId !== "PUBLIC" && ownerUserId === userId)) visible = true;
    }

    if (visible) {
      rawList.push({
        id: String(r[0]), parentId: String(r[1] || ""), title: String(r[2] || ""),
        category: String(r[3] || "Umum"), type: String(r[4] || "link"),
        iconClass: String(r[5] || "fas fa-link"), targetUrl: String(r[6] || "#"),
        description: String(r[7] || ""), ownerUserId: ownerUserId, visibility: visibility, submenus: []
      });
    }
  }

  const map = {};
  const tree = [];
  for (let i = 0; i < rawList.length; i++) map[rawList[i].id] = rawList[i];
  for (let i = 0; i < rawList.length; i++) {
    const item = rawList[i];
    if (item.parentId && map[item.parentId]) {
      map[item.parentId].submenus.push(item);
      if (map[item.parentId].type !== "folder") map[item.parentId].type = "folder";
    } else {
      tree.push(item);
    }
  }
  return tree;
}

function authenticateUser(username, password) {
  const rows = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_USERS).getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) === String(username) && String(rows[i][2]) === String(password)) {
      return {
        status: "success",
        user: { userId: String(rows[i][0]), username: String(rows[i][1]), fullName: String(rows[i][3]), role: String(rows[i][4]) }
      };
    }
  }
  return { status: "error", message: "Username atau Password tidak valid." };
}

function validateSession(userId) {
  if (!userId) return { valid: false };
  const rows = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_USERS).getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(userId)) {
      return {
        valid: true,
        user: { userId: String(rows[i][0]), username: String(rows[i][1]), fullName: String(rows[i][3]), role: String(rows[i][4]) }
      };
    }
  }
  return { valid: false };
}

function checkOwnership(menuId, userId) {
  const rows = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_MENUS).getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(menuId)) {
      if (String(rows[i][8] || "PUBLIC") !== userId) throw new Error("Akses Ditolak: Anda bukan pemilik menu ini.");
      return;
    }
  }
}

function saveOrUpdateMenu(payload) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_MENUS);
  const rows = sheet.getDataRange().getValues();
  let targetRowIndex = -1;

  if (payload.id) {
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(payload.id)) {
        targetRowIndex = i + 1;
        break;
      }
    }
  }

  const recordId = payload.id ? payload.id : "MENU-" + Date.now();
  const rowData = [
    recordId, payload.parentId || "", payload.title, payload.category || "Umum",
    payload.type || "link", payload.iconClass || "fas fa-link", payload.targetUrl || "#",
    payload.description || "", payload.ownerUserId || "PUBLIC", payload.visibility || "UMUM",
    new Date().toISOString()
  ];

  if (targetRowIndex > 0) {
    sheet.getRange(targetRowIndex, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  return { id: recordId };
}

function deleteMenuRecursive(targetId) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_MENUS);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return;

  const idsToDelete = new Set([String(targetId)]);
  let added = true;

  while (added) {
    added = false;
    for (let i = 1; i < rows.length; i++) {
      const id = String(rows[i][0]);
      const pId = String(rows[i][1]);
      if (idsToDelete.has(pId) && !idsToDelete.has(id)) {
        idsToDelete.add(id);
        added = true;
      }
    }
  }

  for (let i = rows.length - 1; i >= 1; i--) {
    if (idsToDelete.has(String(rows[i][0]))) sheet.deleteRow(i + 1);
  }
}

function getUsersList() {
  const rows = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_USERS).getDataRange().getValues();
  const users = [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) {
      users.push({
        userId: String(rows[i][0]),
        username: String(rows[i][1]),
        password: String(rows[i][2]),
        fullName: String(rows[i][3]),
        role: String(rows[i][4])
      });
    }
  }
  return users;
}

// LOGIKA INPUT MANUAL USER_ID & UPDATE USER
function saveOrUpdateUser(payload) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_USERS);
  const rows = sheet.getDataRange().getValues();
  let targetRowIndex = -1;

  if (payload.isEditMode) {
    // Mode Update: Cari berdasarkan User_ID
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).toLowerCase() === String(payload.userId).toLowerCase()) {
        targetRowIndex = i + 1;
        break;
      }
    }
    if (targetRowIndex === -1) throw new Error("User ID tidak ditemukan untuk diperbarui.");
  } else {
    // Mode Create: Validasi Unique User_ID dan Unique Username
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).toLowerCase() === String(payload.userId).toLowerCase()) {
        throw new Error("User ID '" + payload.userId + "' sudah digunakan. Gunakan User ID lain.");
      }
      if (String(rows[i][1]).toLowerCase() === String(payload.username).toLowerCase()) {
        throw new Error("Username '" + payload.username + "' sudah terdaftar. Gunakan Username lain.");
      }
    }
  }

  const rowData = [
    payload.userId, payload.username, payload.password, payload.fullName, payload.role || "USER_KHUSUS"
  ];

  if (targetRowIndex > 0) {
    sheet.getRange(targetRowIndex, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  return { userId: payload.userId };
}

function deleteUserRecord(targetUserId) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_USERS);
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(targetUserId)) {
      sheet.deleteRow(i + 1);
      return;
    }
  }
}
