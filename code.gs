/**
 * PORTAL DMS ENGINE V2 - BACKEND API (Google Apps Script)
 */

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  var action = e.parameter.action;
  var postData = {};
  
  if (e.postData && e.postData.contents) {
    try {
      postData = JSON.parse(e.postData.contents);
      if (!action) action = postData.action;
    } catch (err) {}
  }

  var response = { status: "error", message: "Invalid Action" };

  try {
    if (action === "getMenus") {
      response = getMenus(e.parameter.userId || postData.userId, e.parameter.role || postData.role);
    } else if (action === "saveMenu") {
      response = saveMenu(postData);
    } else if (action === "deleteMenu") {
      response = deleteMenu(postData.id);
    } else if (action === "login") {
      response = processLogin(postData.username, postData.password);
    } else if (action === "getUsers") {
      response = getUsers();
    } else if (action === "saveUser") {
      response = saveUser(postData);
    } else if (action === "deleteUser") {
      response = deleteUser(postData.userId);
    }
  } catch (err) {
    response = { status: "error", message: err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

// ------------------- HELPER SPREADSHEET ------------------- //
function getSheet(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error("Tab Sheet '" + sheetName + "' tidak ditemukan. Harap buat tab tersebut!");
  }
  return sheet;
}

// ------------------- LOGIKA MENUS ------------------- //
function getMenus(userId, role) {
  var sheet = getSheet("Menus");
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { status: "success", data: [] };

  var rawMenus = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var menuObj = {
      id: String(row[0]),
      parentId: String(row[1] || ""),
      title: String(row[2] || ""),
      category: String(row[3] || ""),
      type: String(row[4] || "link"),
      iconClass: String(row[5] || "fas fa-link"),
      targetUrl: String(row[6] || "#"),
      description: String(row[7] || ""),
      ownerUserId: String(row[8] || ""),
      visibility: String(row[9] || "UMUM"),
      createdAt: String(row[10] || "")
    };

    // Filter Hak Akses (Visibility)
    var isVisible = false;
    if (menuObj.visibility === "UMUM") {
      isVisible = true;
    } else if (role === "ADMIN") {
      isVisible = true;
    } else if (menuObj.visibility === "PRIVASI" && userId && userId === menuObj.ownerUserId) {
      isVisible = true;
    } else if (menuObj.visibility === "PRIVASI_ADMIN" && (role === "ADMIN" || userId === menuObj.ownerUserId)) {
      isVisible = true;
    }

    if (isVisible) {
      rawMenus.push(menuObj);
    }
  }

  // Susun Hirarki (Parent - Submenu)
  var tree = [];
  var menuMap = {};

  rawMenus.forEach(function(item) {
    item.submenus = [];
    menuMap[item.id] = item;
  });

  rawMenus.forEach(function(item) {
    if (item.parentId && menuMap[item.parentId]) {
      menuMap[item.parentId].submenus.push(item);
    } else {
      tree.push(item);
    }
  });

  return { status: "success", data: tree };
}

function saveMenu(p) {
  var sheet = getSheet("Menus");
  var data = sheet.getDataRange().getValues();
  var menuId = p.id;
  var now = new Date().toISOString();

  if (menuId) {
    // UPDATE DATA EKSISTING
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(menuId)) {
        sheet.getRange(i + 1, 2).setValue(p.parentId || "");
        sheet.getRange(i + 1, 3).setValue(p.title);
        sheet.getRange(i + 1, 4).setValue(p.category || "");
        sheet.getRange(i + 1, 5).setValue(p.type || "link");
        sheet.getRange(i + 1, 6).setValue(p.iconClass || "fas fa-link");
        sheet.getRange(i + 1, 7).setValue(p.targetUrl || "#");
        sheet.getRange(i + 1, 8).setValue(p.description || "");
        sheet.getRange(i + 1, 10).setValue(p.visibility || "UMUM");
        return { status: "success", message: "Data menu berhasil diperbarui." };
      }
    }
  }

  // INSERT DATA BARU
  var newId = "MENU-" + new Date().getTime();
  sheet.appendRow([
    newId,
    p.parentId || "",
    p.title,
    p.category || "",
    p.type || "link",
    p.iconClass || "fas fa-link",
    p.targetUrl || "#",
    p.description || "",
    p.ownerUserId || "PUBLIC",
    p.visibility || "UMUM",
    now
  ]);

  return { status: "success", message: "Data menu baru berhasil ditambahkan." };
}

function deleteMenu(id) {
  var sheet = getSheet("Menus");
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { status: "success", message: "Menu berhasil dihapus." };
    }
  }
  return { status: "error", message: "ID Menu tidak ditemukan." };
}

// ------------------- LOGIKA USERS & AUTH ------------------- //
function processLogin(username, password) {
  var sheet = getSheet("Users");
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(username).trim() && String(data[i][2]).trim() === String(password).trim()) {
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
  return { status: "error", message: "Username atau Password salah." };
}

function getUsers() {
  var sheet = getSheet("Users");
  var data = sheet.getDataRange().getValues();
  var users = [];

  for (var i = 1; i < data.length; i++) {
    users.push({
      userId: String(data[i][0]),
      username: String(data[i][1]),
      password: String(data[i][2]),
      fullName: String(data[i][3]),
      role: String(data[i][4])
    });
  }
  return { status: "success", data: users };
}

function saveUser(p) {
  var sheet = getSheet("Users");
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(p.userId)) {
      sheet.getRange(i + 1, 2).setValue(p.username);
      sheet.getRange(i + 1, 3).setValue(p.password);
      sheet.getRange(i + 1, 4).setValue(p.fullName);
      sheet.getRange(i + 1, 5).setValue(p.role);
      return { status: "success", message: "Data User berhasil diperbarui." };
    }
  }

  sheet.appendRow([p.userId, p.username, p.password, p.fullName, p.role]);
  return { status: "success", message: "User baru berhasil dibuat." };
}

function deleteUser(userId) {
  var sheet = getSheet("Users");
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(userId)) {
      sheet.deleteRow(i + 1);
      return { status: "success", message: "User berhasil dihapus." };
    }
  }
  return { status: "error", message: "User ID tidak ditemukan." };
}
