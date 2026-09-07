/************************************************************
 MOBILE SHOP MANAGEMENT SYSTEM
 COMPLETE UPGRADED Code.gs
 Version 2.0
************************************************************/

const SHEETS = {
  STOCK: "Stock",
  USERS: "Users",
  SALES: "Sales",
  SALE_ORDERS: "SaleOrders",
  PURCHASES: "Purchases",
  VOUCHERS: "Vouchers",
  LEDGER: "Ledger",
  PARTIES: "Parties",
  MASTERS: "Masters",
  EXPENSES: "Expenses",
  RETURNS: "Returns",
  TRANSFERS: "Transfers",
  CASHBANK: "CashBank",
  CLOSING: "DailyClosing",
  JOURNAL: "JournalVouchers"
};

const SESSION_SECONDS = 21600;


/************************************************************
 DO GET
************************************************************/
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("Index")
    .setTitle("Mobile Shop Management System")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/************************************************************
 SETUP
************************************************************/
function setupSystem() {

  createSheet_(SHEETS.STOCK, [
    "IMEI",
    "Brand",
    "Model",
    "Memory",
    "Color",
    "Purchase Price",
    "Sale Price",
    "Activation Status",
    "Supplier",
    "Shop",
    "Stock Status",
    "Stock In Date",
    "Sold Date",
    "Invoice No"
  ]);

  createSheet_(SHEETS.USERS, [
    "Username",
    "Password",
    "Name",
    "Role",
    "Status",
    "Failed Attempts",
    "View",
    "Add",
    "Edit",
    "Delete"
  ]);

  createSheet_(SHEETS.SALES, [
    "Invoice No",
    "Date",
    "Customer Name",
    "Customer Mobile",
    "Customer Type",
    "IMEI",
    "Brand",
    "Model",
    "Memory",
    "Color",
    "Purchase Price",
    "Sale Price",
    "Discount",
    "Final Amount",
    "Payment Amount",
    "Payment Method",
    "Profit",
    "Sold By"
  ]);

  createSheet_(SHEETS.SALE_ORDERS, [
    "Order No",
    "Date",
    "Customer Name",
    "Customer Mobile",
    "Customer Type",
    "IMEI",
    "Brand",
    "Model",
    "Memory",
    "Color",
    "Amount",
    "Status",
    "User"
  ]);

  createSheet_(SHEETS.PURCHASES, [
    "Purchase No",
    "Date",
    "Supplier Name",
    "Supplier Mobile",
    "IMEI",
    "Brand",
    "Model",
    "Memory",
    "Color",
    "Purchase Price",
    "Sale Price",
    "Qty",
    "Total Amount",
    "Payment Amount",
    "Payment Method",
    "Shop",
    "Purchased By"
  ]);

  createSheet_(SHEETS.VOUCHERS, [
    "Voucher No",
    "Date",
    "Type",
    "Party",
    "Description",
    "Amount",
    "Payment Method",
    "User"
  ]);

  createSheet_(SHEETS.LEDGER, [
    "Date",
    "Party",
    "Type",
    "Reference",
    "Description",
    "Debit",
    "Credit",
    "Balance",
    "User"
  ]);

  createSheet_(SHEETS.PARTIES, [
    "ID",
    "Type",
    "Name",
    "Mobile",
    "Address",
    "Customer Type",
    "Opening Balance",
    "Status",
    "Created Date"
  ]);

  createSheet_(SHEETS.MASTERS, [
    "Type",
    "Name",
    "Status",
    "Created Date"
  ]);

  createSheet_(SHEETS.EXPENSES, [
    "Expense No",
    "Date",
    "Category",
    "Description",
    "Amount",
    "Payment Method",
    "Shop",
    "User"
  ]);

  createSheet_(SHEETS.RETURNS, [
    "Return No",
    "Date",
    "Type",
    "Party",
    "IMEI",
    "Brand",
    "Model",
    "Memory",
    "Color",
    "Amount",
    "Reason",
    "Payment Method",
    "User"
  ]);

  createSheet_(SHEETS.TRANSFERS, [
    "Transfer No",
    "Date",
    "IMEI",
    "Brand",
    "Model",
    "Memory",
    "Color",
    "From Shop",
    "To Shop",
    "User"
  ]);

  createSheet_(SHEETS.JOURNAL, [
    "JV No", "Date", "Debit Account", "Credit Account", "Amount", "Description", "User"
  ]);

  createSheet_(SHEETS.CASHBANK, [
    "Date",
    "Type",
    "Reference",
    "Description",
    "Cash",
    "Bank",
    "User"
  ]);

  createSheet_(SHEETS.CLOSING, [
    "Date",
    "Opening Cash",
    "Cash In",
    "Cash Out",
    "Expected Cash",
    "Actual Cash",
    "Difference",
    "Notes",
    "User"
  ]);

  seedMasters_();
  seedCashCustomer_();
  createDefaultAdmin_();

  return "System Setup Complete";
}


/************************************************************
 CREATE SHEET
************************************************************/
function createSheet_(name, headers) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);

  if (!sh) {
    sh = ss.insertSheet(name);
  }

  if (sh.getLastRow() === 0) {

    sh.getRange(
      1,
      1,
      1,
      headers.length
    ).setValues([headers]);

    sh.getRange(
      1,
      1,
      1,
      headers.length
    ).setFontWeight("bold");

    sh.setFrozenRows(1);
  }
}


/************************************************************
 SAFE SHEET
************************************************************/
function getSheet_(name) {

  const sh =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(name);

  if (!sh) {
    throw new Error(
      "Sheet '" + name +
      "' not found. Please run setupSystem()."
    );
  }

  return sh;
}


/************************************************************
 UPGRADE EXISTING SYSTEM
************************************************************/
function upgradeSystem(token) {

  adminCheck_(token);

  setupSystem();

  return "System upgraded successfully.";
}


/************************************************************
 DEFAULT MASTERS
************************************************************/
function seedMasters_() {

  const sh = getSheet_(SHEETS.MASTERS);

  const defaults = [
    ["Brand", "Oppo"],
    ["Brand", "Vivo"],
    ["Brand", "Samsung"],
    ["Brand", "Infinix"],
    ["Brand", "Tecno"],
    ["Brand", "Realme"],
    ["Brand", "Xiaomi"],
    ["Brand", "Nokia"],
    ["Brand", "Honor"],
    ["Brand", "OnePlus"],

    ["Memory", "2/32"],
    ["Memory", "3/64"],
    ["Memory", "4/64"],
    ["Memory", "4/128"],
    ["Memory", "6/128"],
    ["Memory", "8/128"],
    ["Memory", "8/256"],
    ["Memory", "12/256"]
  ];

  const existing = {};

  if (sh.getLastRow() >= 2) {

    const data =
      sh.getRange(
        2,
        1,
        sh.getLastRow() - 1,
        4
      ).getValues();

    data.forEach(r => {

      existing[
        String(r[0]).toLowerCase() +
        "|" +
        String(r[1]).toLowerCase()
      ] = true;

    });
  }

  const rows = [];

  defaults.forEach(x => {

    const key =
      x[0].toLowerCase() +
      "|" +
      x[1].toLowerCase();

    if (!existing[key]) {

      rows.push([
        x[0],
        x[1],
        "Active",
        new Date()
      ]);
    }
  });

  if (rows.length) {

    sh.getRange(
      sh.getLastRow() + 1,
      1,
      rows.length,
      4
    ).setValues(rows);
  }
}


/************************************************************
 CASH CUSTOMER
************************************************************/
function seedCashCustomer_() {

  const sh = getSheet_(SHEETS.PARTIES);

  if (sh.getLastRow() >= 2) {

    const data =
      sh.getRange(
        2,
        1,
        sh.getLastRow() - 1,
        9
      ).getValues();

    for (let i = 0; i < data.length; i++) {

      if (
        String(data[i][2])
          .trim()
          .toLowerCase() ===
        "cash customer"
      ) {

        if (
          String(data[i][7])
            .toLowerCase() !==
          "active"
        ) {
          sh.getRange(i + 2, 8)
            .setValue("Active");
        }

        return;
      }
    }
  }

  sh.appendRow([
    "CASH-001",
    "Customer",
    "Cash Customer",
    "",
    "",
    "Cash",
    0,
    "Active",
    new Date()
  ]);
}


/************************************************************
 DEFAULT ADMIN
************************************************************/
function createDefaultAdmin_() {

  const sh = getSheet_(SHEETS.USERS);

  if (sh.getLastRow() >= 2) {

    const data =
      sh.getRange(
        2,
        1,
        sh.getLastRow() - 1,
        10
      ).getValues();

    for (let i = 0; i < data.length; i++) {

      if (
        String(data[i][0])
          .trim()
          .toLowerCase() === "admin"
      ) {
        return;
      }
    }
  }

  sh.appendRow([
    "admin",
    "12345",
    "Administrator",
    "Admin",
    "Active",
    0,
    true,
    true,
    true,
    true
  ]);
}


/************************************************************
 LOGIN
************************************************************/
function loginUser(username, password) {

  const sh = getSheet_(SHEETS.USERS);

  if (sh.getLastRow() < 2) {
    throw new Error(
      "No users found. Run setupSystem()."
    );
  }

  const data =
    sh.getRange(
      2,
      1,
      sh.getLastRow() - 1,
      10
    ).getValues();

  for (let i = 0; i < data.length; i++) {

    const r = data[i];

    if (
      String(r[0]).toLowerCase() ===
      String(username).toLowerCase()
    ) {

      if (
        String(r[4]).toLowerCase() ===
        "blocked"
      ) {
        return {
          success: false,
          message: "This user is blocked."
        };
      }

      if (
        String(r[1]) !==
        String(password)
      ) {

        let attempts =
          Number(r[5]) || 0;

        attempts++;

        sh.getRange(i + 2, 6)
          .setValue(attempts);

        if (attempts >= 5) {

          sh.getRange(i + 2, 5)
            .setValue("Blocked");

          return {
            success: false,
            message:
              "5 wrong attempts. User blocked."
          };
        }

        return {
          success: false,
          message:
            "Wrong password. Attempt " +
            attempts + " of 5."
        };
      }

      sh.getRange(i + 2, 6)
        .setValue(0);

      const token =
        Utilities.getUuid();

      CacheService
        .getScriptCache()
        .put(
          "SESSION_" + token,
          String(r[0]),
          SESSION_SECONDS
        );

      return {
        success: true,
        token: token,
        username: r[0],
        name: r[2],
        role: r[3],
        view: r[6],
        add: r[7],
        edit: r[8],
        delete: r[9]
      };
    }
  }

  return {
    success: false,
    message: "Username not found."
  };
}


/************************************************************
 SESSION
************************************************************/
function validateSession(token) {

  if (!token) {
    throw new Error(
      "Session expired. Please login again."
    );
  }

  const username =
    CacheService
      .getScriptCache()
      .get("SESSION_" + token);

  if (!username) {
    throw new Error(
      "Session expired. Please login again."
    );
  }

  return username;
}


function logoutUser(token) {

  if (token) {
    CacheService
      .getScriptCache()
      .remove("SESSION_" + token);
  }

  return true;
}


function getCurrentUser(token) {

  const username =
    validateSession(token);

  const user =
    findUser_(username);

  if (!user) {
    throw new Error("User not found.");
  }

  return {
    success: true,
    username: user.username,
    name: user.name,
    role: user.role,
    view: user.view,
    add: user.add,
    edit: user.edit,
    delete: user.delete
  };
}


function findUser_(username) {

  const sh = getSheet_(SHEETS.USERS);

  if (sh.getLastRow() < 2) {
    return null;
  }

  const data =
    sh.getRange(
      2,
      1,
      sh.getLastRow() - 1,
      10
    ).getValues();

  for (let i = 0; i < data.length; i++) {

    if (
      String(data[i][0])
        .toLowerCase() ===
      String(username)
        .toLowerCase()
    ) {

      return {
        row: i + 2,
        username: data[i][0],
        password: data[i][1],
        name: data[i][2],
        role: data[i][3],
        status: data[i][4],
        failed: data[i][5],
        view: data[i][6],
        add: data[i][7],
        edit: data[i][8],
        delete: data[i][9]
      };
    }
  }

  return null;
}


function adminCheck_(token) {

  const username =
    validateSession(token);

  const user =
    findUser_(username);

  if (
    !user ||
    String(user.role).toLowerCase() !==
    "admin"
  ) {
    throw new Error(
      "Admin permission required."
    );
  }

  return user;
}


function adminOrPermission_(token, permission) {

  const username =
    validateSession(token);

  const user =
    findUser_(username);

  if (!user) {
    throw new Error("User not found.");
  }

  if (
    String(user.role).toLowerCase() ===
    "admin"
  ) {
    return true;
  }

  let allowed = false;

  if (permission === "view") {
    allowed = user.view;
  } else if (permission === "add") {
    allowed = user.add;
  } else if (permission === "edit") {
    allowed = user.edit;
  } else if (permission === "delete") {
    allowed = user.delete;
  }

  if (
    allowed === false ||
    String(allowed).toLowerCase() === "false"
  ) {
    throw new Error(
      "You do not have permission for this action."
    );
  }

  return true;
}


/************************************************************
 USERS
************************************************************/
function getAllUsers(token) {

  adminCheck_(token);

  const sh = getSheet_(SHEETS.USERS);

  if (sh.getLastRow() < 2) {
    return [];
  }

  const data =
    sh.getRange(
      2,
      1,
      sh.getLastRow() - 1,
      10
    ).getValues();

  return data.map(r => ({
    username: r[0],
    name: r[2],
    role: r[3],
    status: r[4],
    view: r[6],
    add: r[7],
    edit: r[8],
    delete: r[9]
  }));
}


function createUser(token, obj) {

  adminCheck_(token);

  obj = obj || {};

  if (!obj.username || !obj.password) {
    throw new Error(
      "Username and password required."
    );
  }

  if (findUser_(obj.username)) {
    throw new Error(
      "Username already exists."
    );
  }

  getSheet_(SHEETS.USERS).appendRow([
    obj.username,
    obj.password,
    obj.name || "",
    obj.role || "User",
    "Active",
    0,
    obj.view !== false,
    obj.add !== false,
    obj.edit !== false,
    obj.delete !== false
  ]);

  return "User created successfully.";
}


function editUser(token, username, obj) {

  adminCheck_(token);

  const user = findUser_(username);

  if (!user) {
    throw new Error("User not found.");
  }

  obj = obj || {};

  const sh = getSheet_(SHEETS.USERS);

  sh.getRange(user.row, 1, 1, 10)
    .setValues([[
      user.username,
      obj.password != null
        ? obj.password
        : user.password,
      obj.name != null
        ? obj.name
        : user.name,
      obj.role != null
        ? obj.role
        : user.role,
      obj.status != null
        ? obj.status
        : user.status,
      0,
      obj.view !== undefined
        ? obj.view
        : user.view,
      obj.add !== undefined
        ? obj.add
        : user.add,
      obj.edit !== undefined
        ? obj.edit
        : user.edit,
      obj.delete !== undefined
        ? obj.delete
        : user.delete
    ]]);

  return "User updated successfully.";
}


function blockUser(token, username) {

  adminCheck_(token);

  if (
    String(username).toLowerCase() ===
    "admin"
  ) {
    throw new Error(
      "Main admin cannot be blocked."
    );
  }

  const user = findUser_(username);

  if (!user) {
    throw new Error("User not found.");
  }

  getSheet_(SHEETS.USERS)
    .getRange(user.row, 5)
    .setValue("Blocked");

  return "User blocked.";
}


function unblockUser(token, username) {

  adminCheck_(token);

  const user = findUser_(username);

  if (!user) {
    throw new Error("User not found.");
  }

  getSheet_(SHEETS.USERS)
    .getRange(user.row, 5)
    .setValue("Active");

  getSheet_(SHEETS.USERS)
    .getRange(user.row, 6)
    .setValue(0);

  return "User unblocked.";
}


function deleteUser(token, username) {

  adminCheck_(token);

  if (
    String(username).toLowerCase() ===
    "admin"
  ) {
    throw new Error(
      "Main admin cannot be deleted."
    );
  }

  const user = findUser_(username);

  if (!user) {
    throw new Error("User not found.");
  }

  getSheet_(SHEETS.USERS)
    .deleteRow(user.row);

  return "User deleted.";
}


/************************************************************
 MASTERS
************************************************************/
function getMasters(token, type) {

  validateSession(token);

  type = String(type || "").trim();

  const sh = getSheet_(SHEETS.MASTERS);

  if (sh.getLastRow() < 2) {
    return [];
  }

  const data =
    sh.getRange(
      2,
      1,
      sh.getLastRow() - 1,
      4
    ).getValues();

  return data
    .filter(r =>
      String(r[0]).trim().toLowerCase() ===
      type.toLowerCase() &&
      String(r[2]).trim().toLowerCase() ===
      "active"
    )
    .map(r => String(r[1]).trim())
    .filter(Boolean)
    .sort();
}


function getAllMasters(token) {

  adminCheck_(token);

  const sh = getSheet_(SHEETS.MASTERS);

  if (sh.getLastRow() < 2) {
    return [];
  }

  const data =
    sh.getRange(
      2,
      1,
      sh.getLastRow() - 1,
      4
    ).getValues();

  return data.map((r, i) => ({
    row: i + 2,
    type: r[0],
    name: r[1],
    status: r[2],
    createdDate:
      r[3] ? formatDate_(r[3]) : ""
  }));
}


function addMaster(token, type, name) {

  adminCheck_(token);

  type = String(type || "").trim();
  name = String(name || "").trim();

  const allowed = [
    "Brand",
    "Model",
    "Memory",
    "Color",
    "Shop"
  ];

  if (allowed.indexOf(type) === -1) {
    throw new Error("Invalid master type.");
  }

  if (!name) {
    throw new Error("Name required.");
  }

  const sh = getSheet_(SHEETS.MASTERS);

  if (sh.getLastRow() >= 2) {

    const data =
      sh.getRange(
        2,
        1,
        sh.getLastRow() - 1,
        4
      ).getValues();

    for (let i = 0; i < data.length; i++) {

      if (
        String(data[i][0]).toLowerCase() ===
        type.toLowerCase() &&
        String(data[i][1]).toLowerCase() ===
        name.toLowerCase()
      ) {

        if (
          String(data[i][2]).toLowerCase() ===
          "deleted"
        ) {

          sh.getRange(i + 2, 2)
            .setValue(name);

          sh.getRange(i + 2, 3)
            .setValue("Active");

          return "Master restored.";
        }

        throw new Error(
          type + " already exists."
        );
      }
    }
  }

  sh.appendRow([
    type,
    name,
    "Active",
    new Date()
  ]);

  return "Added successfully.";
}


function editMaster(token, row, name) {

  adminCheck_(token);

  row = Number(row);
  name = String(name || "").trim();

  if (!row || row < 2) {
    throw new Error("Invalid row.");
  }

  if (!name) {
    throw new Error("Name required.");
  }

  const sh = getSheet_(SHEETS.MASTERS);

  if (row > sh.getLastRow()) {
    throw new Error("Master not found.");
  }

  const type =
    String(
      sh.getRange(row, 1).getValue()
    ).trim();

  sh.getRange(row, 2)
    .setValue(name);

  sh.getRange(row, 3)
    .setValue("Active");

  return "Master updated.";
}


function deleteMaster(token, row) {

  adminCheck_(token);

  row = Number(row);

  if (!row || row < 2) {
    throw new Error("Invalid row.");
  }

  const sh = getSheet_(SHEETS.MASTERS);

  if (row > sh.getLastRow()) {
    throw new Error("Master not found.");
  }

  sh.deleteRow(row);

  return "Master deleted permanently.";
}


/************************************************************
 PARTIES
************************************************************/
function getParties(token, type) {

  validateSession(token);

  const sh = getSheet_(SHEETS.PARTIES);

  if (sh.getLastRow() < 2) {
    return [];
  }

  const data =
    sh.getRange(
      2,
      1,
      sh.getLastRow() - 1,
      9
    ).getValues();

  return data
    .filter(r =>
      String(r[1]).toLowerCase() ===
      String(type).toLowerCase() &&
      String(r[7]).toLowerCase() ===
      "active"
    )
    .map(r => ({
      id: r[0],
      name: r[2],
      mobile: r[3],
      address: r[4],
      customerType: r[5],
      openingBalance:
        Number(r[6]) || 0
    }));
}


function getSuppliers(token) {
  return getParties(token, "Supplier");
}


function getCustomersList(token) {
  return getParties(token, "Customer");
}


function getSupplierNames(token) {
  return getSuppliers(token)
    .map(x => x.name)
    .filter(Boolean);
}


function getCustomerNames(token) {
  return getCustomersList(token)
    .map(x => x.name)
    .filter(Boolean);
}


function getCustomers(token) {
  return getCustomerNames(token);
}


function getPurchaseParties(token) {
  return getSupplierNames(token);
}


function getAllParties(token) {

  adminCheck_(token);

  const sh = getSheet_(SHEETS.PARTIES);

  if (sh.getLastRow() < 2) {
    return [];
  }

  const data =
    sh.getRange(
      2,
      1,
      sh.getLastRow() - 1,
      9
    ).getValues();

  return data.map((r, i) => ({
    row: i + 2,
    id: r[0],
    type: r[1],
    name: r[2],
    mobile: r[3],
    address: r[4],
    customerType: r[5],
    openingBalance:
      Number(r[6]) || 0,
    status: r[7],
    createdDate:
      r[8] ? formatDate_(r[8]) : ""
  }));
}


function addParty(token, obj) {

  adminOrPermission_(token, "add");

  obj = obj || {};

  const type =
    String(obj.type || "").trim();

  const name =
    String(obj.name || "").trim();

  if (
    type !== "Supplier" &&
    type !== "Customer"
  ) {
    throw new Error(
      "Party type must be Supplier or Customer."
    );
  }

  if (!name) {
    throw new Error("Name required.");
  }

  const sh = getSheet_(SHEETS.PARTIES);

  if (sh.getLastRow() >= 2) {

    const data =
      sh.getRange(
        2,
        1,
        sh.getLastRow() - 1,
        9
      ).getValues();

    for (let i = 0; i < data.length; i++) {

      if (
        String(data[i][1]).toLowerCase() ===
        type.toLowerCase() &&
        String(data[i][2]).toLowerCase() ===
        name.toLowerCase()
      ) {

        throw new Error(
          type + " already exists."
        );
      }
    }
  }

  const id =
    type.substring(0, 3).toUpperCase() +
    "-" +
    new Date().getTime();

  const opening =
    Number(obj.openingBalance) || 0;

  sh.appendRow([
    id,
    type,
    name,
    obj.mobile || "",
    obj.address || "",
    obj.customerType || "",
    opening,
    "Active",
    new Date()
  ]);

  if (opening !== 0) {

    addLedgerInternal_(
      new Date(),
      name,
      "Opening",
      id,
      "Opening Balance",
      opening > 0 ? opening : 0,
      opening < 0 ? Math.abs(opening) : 0,
      getCurrentUsername_(token)
    );
  }

  return {
    success: true,
    id: id,
    message: "Party added successfully."
  };
}


function editParty(token, row, obj) {

  adminOrPermission_(token, "edit");

  row = Number(row);
  obj = obj || {};

  const sh = getSheet_(SHEETS.PARTIES);

  if (
    !row ||
    row < 2 ||
    row > sh.getLastRow()
  ) {
    throw new Error("Party not found.");
  }

  const name =
    String(obj.name || "").trim();

  if (!name) {
    throw new Error("Name required.");
  }

  sh.getRange(row, 3, 1, 5)
    .setValues([[
      name,
      obj.mobile || "",
      obj.address || "",
      obj.customerType || "",
      Number(obj.openingBalance) || 0
    ]]);

  sh.getRange(row, 8)
    .setValue("Active");

  return "Party updated successfully.";
}


function deleteParty(token, row) {

  adminOrPermission_(token, "delete");

  row = Number(row);

  const sh = getSheet_(SHEETS.PARTIES);

  if (
    !row ||
    row < 2 ||
    row > sh.getLastRow()
  ) {
    throw new Error("Party not found.");
  }

  const name =
    String(
      sh.getRange(row, 3).getValue()
    );

  if (
    name.toLowerCase() ===
    "cash customer"
  ) {
    throw new Error(
      "Cash Customer cannot be deleted."
    );
  }

  sh.deleteRow(row);

  return "Party deleted permanently.";
}


/************************************************************
 NUMBER GENERATORS
************************************************************/
function nextNumber_(sheetName, prefix, col) {

  const sh = getSheet_(sheetName);

  let maxNo = 0;

  if (sh.getLastRow() >= 2) {

    const values =
      sh.getRange(
        2,
        col || 1,
        sh.getLastRow() - 1,
        1
      ).getValues();

    values.forEach(r => {

      const m =
        String(r[0] || "")
          .match(
            new RegExp("^" + prefix + "-(\\d+)$", "i")
          );

      if (m) {
        maxNo =
          Math.max(
            maxNo,
            Number(m[1]) || 0
          );
      }
    });
  }

  return prefix + "-" +
    Utilities.formatString(
      "%05d",
      maxNo + 1
    );
}


function generatePurchaseNo(token) {
  validateSession(token);
  return nextNumber_(
    SHEETS.PURCHASES,
    "PUR",
    1
  );
}


function generateInvoiceNo(token) {
  validateSession(token);
  return nextNumber_(
    SHEETS.SALES,
    "INV",
    1
  );
}


function generateVoucherNo(token) {
  validateSession(token);
  return nextNumber_(
    SHEETS.VOUCHERS,
    "VCH",
    1
  );
}


/************************************************************
 PURCHASE
************************************************************/
function createPurchase(token, obj) {

  adminOrPermission_(token, "add");

  obj = obj || {};

  if (
    !obj.items ||
    !obj.items.length
  ) {
    throw new Error(
      "Please add at least one item."
    );
  }

  const stock = getSheet_(SHEETS.STOCK);
  const purchases = getSheet_(SHEETS.PURCHASES);

  const lock =
    LockService.getScriptLock();

  lock.waitLock(10000);

  try {

    const purchaseNo =
      nextNumber_(
        SHEETS.PURCHASES,
        "PUR",
        1
      );

    const now = new Date();

    let total = 0;

    obj.items.forEach(item => {

      const imei =
        String(item.imei || "").trim();

      if (!imei) {
        throw new Error("IMEI required.");
      }

      if (findStockRow_(imei)) {
        throw new Error(
          "IMEI already exists: " + imei
        );
      }

      total +=
        Number(item.purchasePrice) || 0;
    });

    const payment =
      Math.min(
        total,
        Math.max(
          0,
          Number(obj.paymentAmount) || 0
        )
      );

    obj.items.forEach(item => {

      const purchasePrice =
        Number(item.purchasePrice) || 0;

      stock.appendRow([
        item.imei,
        item.brand || "",
        item.model || "",
        item.memory || "",
        item.color || "",
        purchasePrice,
        Number(item.salePrice) || 0,
        item.activationStatus ||
          "Not Activated",
        obj.supplierName || "",
        item.shop || obj.shop || "",
        "Available",
        now,
        "",
        ""
      ]);

      purchases.appendRow([
        purchaseNo,
        now,
        obj.supplierName || "",
        obj.supplierMobile || "",
        item.imei,
        item.brand || "",
        item.model || "",
        item.memory || "",
        item.color || "",
        purchasePrice,
        Number(item.salePrice) || 0,
        1,
        purchasePrice,
        payment,
        obj.paymentMethod || "Cash",
        item.shop || obj.shop || "",
        getCurrentUsername_(token)
      ]);
    });

    if (obj.supplierName) {

      addLedgerInternal_(
        now,
        obj.supplierName,
        "Purchase",
        purchaseNo,
        "Purchase Stock",
        0,
        total,
        getCurrentUsername_(token)
      );

      if (payment > 0) {

        addLedgerInternal_(
          now,
          obj.supplierName,
          "Supplier Payment",
          purchaseNo,
          "Purchase Payment",
          payment,
          0,
          getCurrentUsername_(token)
        );
      }
    }

    addCashBank_(
      now,
      "Purchase",
      purchaseNo,
      "Purchase Payment",
      obj.paymentMethod || "Cash",
      payment,
      token
    );

    return {
      success: true,
      purchaseNo: purchaseNo,
      qty: obj.items.length,
      total: total,
      payment: payment,
      balance: total - payment
    };

  } finally {
    lock.releaseLock();
  }
}



/************************************************************
 PURCHASE HISTORY
************************************************************/
function getAllPurchases(token, filters) {
  validateSession(token);
  filters = filters || {};

  const sh = getSheet_(SHEETS.PURCHASES);
  if (sh.getLastRow() < 2) return [];

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 18).getValues();
  const groups = {};

  data.forEach(r => {
    const no = String(r[0] || "").trim();
    if (!no) return;

    if (!groups[no]) {
      groups[no] = {
        purchaseNo: r[0],
        date: formatDate_(r[1]),
        supplier: r[2],
        supplierMobile: r[3],
        shop: r[16],
        paymentMethod: r[14] || "Cash",
        payment: Number(r[13]) || 0,
        total: 0,
        qty: 0,
        models: []
      };
    }

    groups[no].qty += Number(r[11]) || 1;
    groups[no].total += Number(r[12]) || Number(r[9]) || 0;

    const model = String(r[6] || "").trim();
    if (model && groups[no].models.indexOf(model) === -1) {
      groups[no].models.push(model);
    }

    if (!groups[no].payment && Number(r[13])) {
      groups[no].payment = Number(r[13]) || 0;
    }
  });

  const search = String(filters.search || "").trim().toLowerCase();

  return Object.keys(groups)
    .map(k => groups[k])
    .filter(x => {
      if (!search) return true;
      const hay = [
        x.purchaseNo, x.supplier, x.supplierMobile,
        x.shop, x.models.join(" ")
      ].join(" ").toLowerCase();
      return hay.indexOf(search) !== -1;
    })
    .map(x => {
      x.balance = Math.max(0, x.total - x.payment);
      return x;
    });
}


/************************************************************
 GET PURCHASE
************************************************************/
function getPurchaseByNo(token, purchaseNo) {
  validateSession(token);

  const sh = getSheet_(SHEETS.PURCHASES);
  if (sh.getLastRow() < 2) return null;

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 18).getValues();
  const items = [];
  let first = null;

  data.forEach((r, i) => {
    if (
      String(r[0]).trim().toLowerCase() ===
      String(purchaseNo).trim().toLowerCase()
    ) {
      if (!first) {
        first = {
          row: i + 2,
          purchaseNo: r[0],
          date: formatDate_(r[1]),
          supplierName: r[2],
          supplierMobile: r[3],
          paymentAmount: Number(r[13]) || 0,
          paymentMethod: r[14] || "Cash",
          shop: r[16] || ""
        };
      }

      items.push({
        imei: r[4],
        brand: r[5],
        model: r[6],
        memory: r[7],
        color: r[8],
        purchasePrice: Number(r[9]) || 0,
        salePrice: Number(r[10]) || 0,
        qty: Number(r[11]) || 1,
        shop: r[16] || ""
      });
    }
  });

  if (!first) return null;

  return {
    row: first.row,
    purchaseNo: first.purchaseNo,
    date: first.date,
    supplierName: first.supplierName,
    supplierMobile: first.supplierMobile,
    paymentAmount: first.paymentAmount,
    paymentMethod: first.paymentMethod,
    shop: first.shop,
    items: items
  };
}


/************************************************************
 EDIT PURCHASE
************************************************************/
function editPurchase(token, purchaseNo, obj) {
  adminOrPermission_(token, "edit");

  const old = getPurchaseByNo(token, purchaseNo);
  if (!old) throw new Error("Purchase not found.");

  obj = obj || {};
  if (!obj.items || !obj.items.length) {
    throw new Error("Please add at least one item.");
  }

  const stock = getSheet_(SHEETS.STOCK);
  const purchases = getSheet_(SHEETS.PURCHASES);

  // Sold IMEIs cannot be edited because they are linked to a sale.
  old.items.forEach(item => {
    const s = findStockRow_(item.imei);
    if (
      s &&
      String(s.values[10]).trim().toLowerCase() === "sold"
    ) {
      throw new Error(
        "Cannot edit purchase. IMEI already sold: " + item.imei
      );
    }
  });

  // Validate new IMEIs.
  const seen = {};
  obj.items.forEach(item => {
    const imei = String(item.imei || "").trim();
    if (!imei) throw new Error("IMEI required.");

    const key = imei.toLowerCase();
    if (seen[key]) {
      throw new Error("Duplicate IMEI: " + imei);
    }
    seen[key] = true;

    const existing = findStockRow_(imei);
    const belongsToOld = old.items.some(x =>
      String(x.imei).trim().toLowerCase() === key
    );

    if (existing && !belongsToOld) {
      throw new Error("IMEI already exists: " + imei);
    }
  });

  const supplierName =
    obj.supplierName != null
      ? String(obj.supplierName).trim()
      : old.supplierName;

  const supplierMobile =
    obj.supplierMobile != null
      ? String(obj.supplierMobile).trim()
      : old.supplierMobile;

  const shop =
    obj.shop != null
      ? String(obj.shop).trim()
      : old.shop;

  const paymentMethod =
    obj.paymentMethod ||
    old.paymentMethod ||
    "Cash";

  let total = 0;
  obj.items.forEach(item => {
    const qty = Number(item.qty) || 1;
    const price = Number(item.purchasePrice) || 0;
    total += price * qty;
  });

  const payment = Math.min(
    total,
    Math.max(
      0,
      Number(
        obj.paymentAmount != null
          ? obj.paymentAmount
          : old.paymentAmount
      ) || 0
    )
  );

  const now = new Date();
  const user = getCurrentUsername_(token);

  // Remove old accounting entries.
  removeRowsByReference_(SHEETS.CASHBANK, old.purchaseNo, 3);
  removeRowsByReference_(SHEETS.LEDGER, old.purchaseNo, 4);

  // Remove old purchase rows.
  const pData = purchases.getRange(
    2, 1, purchases.getLastRow() - 1, 18
  ).getValues();

  const pRows = [];
  pData.forEach((r, i) => {
    if (
      String(r[0]).trim().toLowerCase() ===
      String(old.purchaseNo).trim().toLowerCase()
    ) {
      pRows.push(i + 2);
    }
  });

  for (let i = pRows.length - 1; i >= 0; i--) {
    purchases.deleteRow(pRows[i]);
  }

  // Remove old available stock rows.
  old.items.forEach(item => {
    const s = findStockRow_(item.imei);
    if (s) stock.deleteRow(s.row);
  });

  // Write updated purchase.
  obj.items.forEach(item => {
    const imei = String(item.imei || "").trim();
    const purchasePrice = Number(item.purchasePrice) || 0;
    const salePrice = Number(item.salePrice) || 0;

    stock.appendRow([
      imei,
      item.brand || "",
      item.model || "",
      item.memory || "",
      item.color || "",
      purchasePrice,
      salePrice,
      item.activationStatus || "Not Activated",
      supplierName,
      item.shop || shop || "",
      "Available",
      now,
      "",
      ""
    ]);

    purchases.appendRow([
      old.purchaseNo,
      now,
      supplierName,
      supplierMobile,
      imei,
      item.brand || "",
      item.model || "",
      item.memory || "",
      item.color || "",
      purchasePrice,
      salePrice,
      1,
      purchasePrice,
      payment,
      paymentMethod,
      item.shop || shop || "",
      user
    ]);
  });

  // Rebuild accounting.
  if (supplierName) {
    addLedgerInternal_(
      now, supplierName, "Purchase", old.purchaseNo,
      "Purchase Stock", 0, total, user
    );

    if (payment > 0) {
      addLedgerInternal_(
        now, supplierName, "Supplier Payment", old.purchaseNo,
        "Purchase Payment", payment, 0, user
      );
    }
  }

  addCashBank_(
    now, "Purchase", old.purchaseNo,
    "Purchase Payment", paymentMethod, payment, token
  );

  return {
    success: true,
    purchaseNo: old.purchaseNo,
    total: total,
    payment: payment,
    balance: total - payment,
    message: "Purchase updated successfully."
  };
}


/************************************************************
 DELETE PURCHASE
************************************************************/
function deletePurchase(token, purchaseNo) {
  adminOrPermission_(token, "delete");

  const old = getPurchaseByNo(token, purchaseNo);
  if (!old) throw new Error("Purchase not found.");

  old.items.forEach(item => {
    const s = findStockRow_(item.imei);
    if (
      s &&
      String(s.values[10]).trim().toLowerCase() === "sold"
    ) {
      throw new Error(
        "Cannot delete purchase. IMEI already sold: " + item.imei
      );
    }
  });

  const purchases = getSheet_(SHEETS.PURCHASES);
  const stock = getSheet_(SHEETS.STOCK);

  const pData = purchases.getRange(
    2, 1, purchases.getLastRow() - 1, 18
  ).getValues();

  const pRows = [];
  pData.forEach((r, i) => {
    if (
      String(r[0]).trim().toLowerCase() ===
      String(old.purchaseNo).trim().toLowerCase()
    ) {
      pRows.push(i + 2);
    }
  });

  for (let i = pRows.length - 1; i >= 0; i--) {
    purchases.deleteRow(pRows[i]);
  }

  old.items.forEach(item => {
    const s = findStockRow_(item.imei);
    if (s) stock.deleteRow(s.row);
  });

  removeRowsByReference_(SHEETS.CASHBANK, old.purchaseNo, 3);
  removeRowsByReference_(SHEETS.LEDGER, old.purchaseNo, 4);

  return {
    success: true,
    message: "Purchase deleted and accounting/stock updated."
  };
}


/************************************************************
 PURCHASE REPORT
************************************************************/
function getPurchaseReport(token, filters) {

  validateSession(token);

  filters = filters || {};

  const sh = getSheet_(SHEETS.PURCHASES);

  if (sh.getLastRow() < 2) {
    return {
      rows: [],
      totalQty: 0,
      totalAmount: 0
    };
  }

  const data =
    sh.getRange(
      2,
      1,
      sh.getLastRow() - 1,
      18
    ).getValues();

  const from =
    parseReportDate_(
      filters.fromDate,
      false
    );

  const to =
    parseReportDate_(
      filters.toDate,
      true
    );

  let rows = [];
  let totalQty = 0;
  let totalAmount = 0;

  data.forEach(r => {

    const date = new Date(r[1]);

    if (from && date < from) return;
    if (to && date > to) return;

    if (
      filters.supplier &&
      filters.supplier !== "all" &&
      String(r[2]) !==
      String(filters.supplier)
    ) {
      return;
    }

    const qty = Number(r[12]) || 0;
    const amount = Number(r[13]) || 0;

    totalQty += qty;
    totalAmount += amount;

    rows.push({
      purchaseNo: r[0],
      date: formatDate_(r[1]),
      supplier: r[2],
      imei: r[4],
      item: r[6],
      memory: r[7],
      color: r[8],
      qty: qty,
      price: Number(r[9]) || 0,
      amount: amount
    });
  });

  return {
    rows: rows,
    totalQty: totalQty,
    totalAmount: totalAmount
  };
}


/************************************************************
 STOCK
************************************************************/
function getAllStock(token) {

  validateSession(token);

  const sh = getSheet_(SHEETS.STOCK);

  if (sh.getLastRow() < 2) {
    return [];
  }

  const data =
    sh.getRange(
      2,
      1,
      sh.getLastRow() - 1,
      14
    ).getValues();

  return data.map((r, i) => ({
    row: i + 2,
    imei: r[0],
    brand: r[1],
    model: r[2],
    memory: r[3],
    color: r[4],
    purchasePrice: Number(r[5]) || 0,
    salePrice: Number(r[6]) || 0,
    activationStatus: r[7],
    supplier: r[8],
    shop: r[9],
    stockStatus: r[10],
    stockInDate: formatDate_(r[11]),
    soldDate: formatDate_(r[12]),
    invoiceNo: r[13]
  }));
}


function getStock(token) {

  return getAllStock(token)
    .filter(r =>
      String(r.stockStatus)
        .toLowerCase() ===
      "available"
    );
}


function searchIMEI(token, imei) {

  validateSession(token);

  const result =
    findStockRow_(imei);

  if (!result) {
    return null;
  }

  const r = result.values;

  return {
    row: result.row,
    imei: r[0],
    brand: r[1],
    model: r[2],
    memory: r[3],
    color: r[4],
    purchasePrice: Number(r[5]) || 0,
    salePrice: Number(r[6]) || 0,
    activationStatus: r[7],
    supplier: r[8],
    shop: r[9],
    stockStatus: r[10],
    stockInDate: formatDate_(r[11]),
    soldDate: formatDate_(r[12]),
    invoiceNo: r[13]
  };
}


function editStock(token, row, obj) {

  adminOrPermission_(token, "edit");

  row = Number(row);
  obj = obj || {};

  const sh = getSheet_(SHEETS.STOCK);

  if (
    !row ||
    row < 2 ||
    row > sh.getLastRow()
  ) {
    throw new Error("Stock row not found.");
  }

  const current =
    sh.getRange(row, 1, 1, 14)
      .getValues()[0];

  sh.getRange(row, 1, 1, 10)
    .setValues([[
      obj.imei || current[0],
      obj.brand || "",
      obj.model || "",
      obj.memory || "",
      obj.color || "",
      Number(obj.purchasePrice) || 0,
      Number(obj.salePrice) || 0,
      obj.activationStatus ||
        "Not Activated",
      obj.supplier || "",
      obj.shop || ""
    ]]);

  return "Stock updated.";
}


function deleteStock(token, row) {

  adminOrPermission_(token, "delete");

  row = Number(row);

  const sh = getSheet_(SHEETS.STOCK);

  if (
    !row ||
    row < 2 ||
    row > sh.getLastRow()
  ) {
    throw new Error("Stock row not found.");
  }

  const status =
    String(
      sh.getRange(row, 11).getValue()
    ).toLowerCase();

  if (status === "sold") {
    throw new Error(
      "Sold stock cannot be deleted."
    );
  }

  sh.deleteRow(row);

  return "Stock deleted.";
}


/************************************************************
 SALES
************************************************************/
function createSale(token, obj) {
  adminOrPermission_(token, "add"); obj=obj||{};
  const raw=Array.isArray(obj.imeis)?obj.imeis:[obj.imei], imeis=[], seen={};
  raw.map(x=>String(x||"").trim()).filter(Boolean).forEach(x=>{const k=x.toLowerCase();if(!seen[k]){seen[k]=true;imeis.push(x);}});
  if(!imeis.length)throw new Error("At least one IMEI is required.");
  const lock=LockService.getScriptLock(); lock.waitLock(10000);
  try{
    const stock=getSheet_(SHEETS.STOCK), sales=getSheet_(SHEETS.SALES), items=[]; let purchaseTotal=0;
    imeis.forEach(imei=>{const info=findStockRow_(imei);if(!info)throw new Error("IMEI not found in stock: "+imei);if(String(info.values[10]).toLowerCase()==="sold")throw new Error("This IMEI is already sold: "+imei);items.push({imei,info});purchaseTotal+=Number(info.values[5])||0;});
    const invoiceNo=nextNumber_(SHEETS.SALES,"INV",1), qty=items.length, unitPrice=Number(obj.salePrice)||Number(items[0].info.values[6])||0;
    const discount=Math.max(0,Number(obj.discount)||0), gross=unitPrice*qty, finalAmount=Math.max(0,gross-discount), payment=Math.min(finalAmount,Math.max(0,Number(obj.paymentAmount)||0));
    const now=new Date(), customer=obj.customerName||"Cash Customer", customerType=obj.customerType||"Cash", method=obj.paymentMethod||"Cash", user=getCurrentUsername_(token), perDiscount=qty?discount/qty:0;
    items.forEach(x=>{const r=x.info.values, itemPrice=Number(obj.salePrice)||Number(r[6])||unitPrice, itemFinal=Math.max(0,itemPrice-perDiscount);stock.getRange(x.info.row,11,1,4).setValues([["Sold",r[11],now,invoiceNo]]);sales.appendRow([invoiceNo,now,customer,obj.customerMobile||"",customerType,x.imei,r[1],r[2],r[3],r[4],Number(r[5])||0,itemPrice,perDiscount,itemFinal,payment,method,itemFinal-(Number(r[5])||0),user]);});
    addLedgerInternal_(now,customer,"Sale",invoiceNo,"Mobile Sale",finalAmount,0,user);
    if(payment>0)addLedgerInternal_(now,customer,"Customer Payment",invoiceNo,"Payment Received",0,payment,user);
    addCashBank_(now,"Sale",invoiceNo,"Sale Payment",method,payment,token);
    return {success:true,invoiceNo,customer,qty,finalAmount,payment,balance:finalAmount-payment,profit:finalAmount-purchaseTotal};
  }finally{lock.releaseLock();}
}

/************************************************************
 SALE REPORT
************************************************************/
function getSaleReport(token, filters) {

  validateSession(token);

  filters = filters || {};

  const sh = getSheet_(SHEETS.SALES);

  if (sh.getLastRow() < 2) {
    return {
      rows: [],
      totalQty: 0,
      totalAmount: 0
    };
  }

  const data =
    sh.getRange(
      2,
      1,
      sh.getLastRow() - 1,
      18
    ).getValues();

  const from =
    parseReportDate_(
      filters.fromDate,
      false
    );

  const to =
    parseReportDate_(
      filters.toDate,
      true
    );

  let rows = [];
  let totalQty = 0;
  let totalAmount = 0;

  data.forEach(r => {

    const date = new Date(r[1]);

    if (from && date < from) return;
    if (to && date > to) return;

    if (
      filters.customer &&
      filters.customer !== "all" &&
      String(r[2]) !==
      String(filters.customer)
    ) {
      return;
    }

    if (
      filters.customerType &&
      filters.customerType !== "all" &&
      String(r[4]) !==
      String(filters.customerType)
    ) {
      return;
    }

    totalQty++;
    totalAmount += Number(r[13]) || 0;

    rows.push({
      invoiceNo: r[0],
      date: formatDate_(r[1]),
      customer: r[2],
      customerType: r[4],
      imei: r[5],
      brand: r[6],
      item: r[7],
      memory: r[8],
      color: r[9],
      price: Number(r[11]) || 0,
      discount: Number(r[12]) || 0,
      amount: Number(r[13]) || 0,
      payment: Number(r[14]) || 0,
      balance:
        (Number(r[13]) || 0) -
        (Number(r[14]) || 0),
      profit: Number(r[16]) || 0
    });
  });

  return {
    rows: rows,
    totalQty: totalQty,
    totalAmount: totalAmount
  };
}


function getAllSales(token, filters) {

  validateSession(token);

  filters = filters || {};

  const sh = getSheet_(SHEETS.SALES);

  if (sh.getLastRow() < 2) {
    return [];
  }

  const data =
    sh.getRange(
      2,
      1,
      sh.getLastRow() - 1,
      18
    ).getValues();

  const from =
    parseReportDate_(
      filters.fromDate,
      false
    );

  const to =
    parseReportDate_(
      filters.toDate,
      true
    );

  const search =
    String(filters.search || "")
      .trim()
      .toLowerCase();

  return data.map((r, i) => ({
    row: i + 2,
    invoiceNo: r[0],
    date: formatDate_(r[1]),
    customerName: r[2],
    customerMobile: r[3],
    customerType: r[4],
    imei: r[5],
    brand: r[6],
    model: r[7],
    memory: r[8],
    color: r[9],
    purchasePrice: Number(r[10]) || 0,
    salePrice: Number(r[11]) || 0,
    discount: Number(r[12]) || 0,
    finalAmount: Number(r[13]) || 0,
    paymentAmount: Number(r[14]) || 0,
    paymentMethod: r[15],
    profit: Number(r[16]) || 0,
    soldBy: r[17],
    rawDate: r[1]
  })).filter(x => {

    if (
      from &&
      new Date(x.rawDate) < from
    ) return false;

    if (
      to &&
      new Date(x.rawDate) > to
    ) return false;

    if (
      filters.customerType &&
      filters.customerType !== "all" &&
      String(x.customerType) !==
      String(filters.customerType)
    ) return false;

    if (search) {

      const hay = [
        x.invoiceNo,
        x.customerName,
        x.customerMobile,
        x.imei,
        x.brand,
        x.model
      ]
        .join(" ")
        .toLowerCase();

      if (
        hay.indexOf(search) === -1
      ) return false;
    }

    return true;

  }).map(x => {

    delete x.rawDate;
    return x;

  });
}


/************************************************************
 GET SALE
************************************************************/
function getSaleByInvoice(token, invoiceNo) {
  validateSession(token);
  const rows = getSaleRowsByInvoice(token, invoiceNo);
  if (!rows.length) return null;
  const r = rows[0];
  return Object.assign({}, r, {
    items: rows,
    qty: rows.length,
    totalFinalAmount: rows.reduce((a,x)=>a+(Number(x.finalAmount)||0),0),
    totalPurchasePrice: rows.reduce((a,x)=>a+(Number(x.purchasePrice)||0),0),
    totalProfit: rows.reduce((a,x)=>a+(Number(x.profit)||0),0)
  });
}


function getSaleRowsByInvoice(token, invoiceNo){
  validateSession(token); const sh=getSheet_(SHEETS.SALES); if(sh.getLastRow()<2)return [];
  const d=sh.getRange(2,1,sh.getLastRow()-1,18).getValues();
  return d.map((r,i)=>({row:i+2,invoiceNo:r[0],date:formatDate_(r[1]),customerName:r[2],customerMobile:r[3],customerType:r[4],imei:r[5],brand:r[6],model:r[7],memory:r[8],color:r[9],purchasePrice:Number(r[10])||0,salePrice:Number(r[11])||0,discount:Number(r[12])||0,finalAmount:Number(r[13])||0,paymentAmount:Number(r[14])||0,paymentMethod:r[15],profit:Number(r[16])||0,soldBy:r[17]})).filter(x=>String(x.invoiceNo).trim().toLowerCase()===String(invoiceNo).trim().toLowerCase());
}

/************************************************************
 EDIT SALE
************************************************************/
function editSale(token, invoiceNo, obj) {
  adminOrPermission_(token, "edit");
  obj = obj || {};

  const oldRows = getSaleRowsByInvoice(token, invoiceNo);
  if (!oldRows.length) throw new Error("Sale invoice not found.");

  let imeis = Array.isArray(obj.imeis)
    ? obj.imeis.map(x=>String(x||"").trim()).filter(Boolean)
    : (obj.imei ? [String(obj.imei).trim()] : oldRows.map(x=>String(x.imei).trim()));

  const seen = {};
  imeis = imeis.filter(x=>{ const k=x.toLowerCase(); if(seen[k]) return false; seen[k]=true; return true; });
  if (!imeis.length) throw new Error("At least one IMEI is required.");

  const oldImeiSet = {};
  oldRows.forEach(x=>oldImeiSet[String(x.imei).trim().toLowerCase()]=true);

  const stock = getSheet_(SHEETS.STOCK);
  const sales = getSheet_(SHEETS.SALES);
  const newItems = [];

  // Validate every new IMEI before changing anything.
  imeis.forEach(imei=>{
    const info=findStockRow_(imei);
    if(!info) throw new Error("IMEI not found in stock: "+imei);
    const status=String(info.values[10]).trim().toLowerCase();
    if(status==="sold" && !oldImeiSet[imei.toLowerCase()]) {
      throw new Error("IMEI already sold: "+imei);
    }
    newItems.push({imei,info});
  });

  const oldCustomer=oldRows[0].customerName;
  const customer=obj.customerName!=null ? (String(obj.customerName).trim()||"Cash Customer") : oldCustomer;
  const customerMobile=obj.customerMobile!=null ? obj.customerMobile : oldRows[0].customerMobile;
  const customerType=obj.customerType!=null ? (obj.customerType||"Cash") : oldRows[0].customerType;
  const method=obj.paymentMethod || oldRows[0].paymentMethod || "Cash";
  const unitPriceInput=Number(obj.salePrice)!=0 ? Number(obj.salePrice) : 0;
  const unitPriceDefault=unitPriceInput || Number(newItems[0].info.values[6]) || 0;
  const qty=newItems.length;
  const discount=Math.max(0,Number(obj.discount)||0);
  const gross=unitPriceDefault*qty;
  const finalAmount=Math.max(0,gross-discount);
  const payment=Math.min(finalAmount,Math.max(0,Number(obj.paymentAmount!=null?obj.paymentAmount:oldRows[0].paymentAmount)||0));
  const totalPurchase=newItems.reduce((a,x)=>a+(Number(x.info.values[5])||0),0);
  const now=new Date();
  const user=getCurrentUsername_(token);
  const perDiscount=qty ? discount/qty : 0;

  // Restore all old stock first, but keep the original stock row data intact.
  oldRows.forEach(old=>{
    const s=findStockRow_(old.imei);
    if(s){
      stock.getRange(s.row,11,1,4).setValues([["Available",s.values[11],"",""]]);
    }
  });

  // Remove old invoice rows and old accounting.
  const allSales=sales.getRange(2,1,Math.max(1,sales.getLastRow()-1),18).getValues();
  const oldSaleRows=[];
  allSales.forEach((r,i)=>{ if(String(r[0]).trim().toLowerCase()===String(invoiceNo).trim().toLowerCase()) oldSaleRows.push(i+2); });
  for(let i=oldSaleRows.length-1;i>=0;i--) sales.deleteRow(oldSaleRows[i]);
  removeRowsByReference_(SHEETS.CASHBANK, invoiceNo, 3);
  removeRowsByReference_(SHEETS.LEDGER, invoiceNo, 4);

  // Write the edited invoice as one row per physical phone/IMEI.
  newItems.forEach((x,idx)=>{
    const r=x.info.values;
    const itemUnitPrice=unitPriceInput || Number(r[6]) || unitPriceDefault;
    const itemDiscount=perDiscount;
    const itemFinal=Math.max(0,itemUnitPrice-itemDiscount);
    const purchasePrice=Number(r[5])||0;
    const itemProfit=itemFinal-purchasePrice;
    if(obj.purchasePrice!=null) stock.getRange(x.info.row,6).setValue(Number(obj.purchasePrice)||0);
    stock.getRange(x.info.row,11,1,4).setValues([["Sold",r[11],now,invoiceNo]]);
    sales.appendRow([
      invoiceNo,now,customer,customerMobile,customerType,x.imei,
      r[1],r[2],r[3],r[4],obj.purchasePrice!=null?(Number(obj.purchasePrice)||0):purchasePrice,
      itemUnitPrice,itemDiscount,itemFinal,idx===0?payment:0,method,itemProfit,user
    ]);
  });

  addLedgerInternal_(now,customer,"Sale",invoiceNo,"Mobile Sale",finalAmount,0,user);
  if(payment>0) addLedgerInternal_(now,customer,"Customer Payment",invoiceNo,"Payment Received",0,payment,user);
  addCashBank_(now,"Sale",invoiceNo,"Sale Payment",method,payment,token);

  return {success:true,invoiceNo:invoiceNo,qty:qty,finalAmount:finalAmount,payment:payment,balance:finalAmount-payment,profit:finalAmount-totalPurchase};
}


/************************************************************
 DELETE SALE
************************************************************/
function deleteSale(token, invoiceNo) {

  adminOrPermission_(token, "delete");

  const old =
    getSaleByInvoice(
      token,
      invoiceNo
    );

  if (!old) {
    throw new Error(
      "Sale invoice not found."
    );
  }

  const stockInfo =
    findStockRow_(old.imei);

  if (
    stockInfo &&
    String(stockInfo.values[13]).trim() ===
    String(old.invoiceNo).trim()
  ) {

    getSheet_(SHEETS.STOCK)
      .getRange(
        stockInfo.row,
        11,
        1,
        4
      )
      .setValues([[
        "Available",
        stockInfo.values[11],
        "",
        ""
      ]]);
  }

  getSheet_(SHEETS.SALES)
    .deleteRow(old.row);

  removeRowsByReference_(
    SHEETS.CASHBANK,
    old.invoiceNo,
    3
  );

  removeRowsByReference_(
    SHEETS.LEDGER,
    old.invoiceNo,
    4
  );

  return {
    success: true,
    message:
      "Sale deleted and stock restored."
  };
}


/************************************************************
 SALE ORDERS
************************************************************/
function createSaleOrder(token, obj) {

  adminOrPermission_(token, "add");

  obj = obj || {};

  const no =
    nextNumber_(
      SHEETS.SALE_ORDERS,
      "ORD",
      1
    );

  getSheet_(SHEETS.SALE_ORDERS)
    .appendRow([
      no,
      new Date(),
      obj.customerName ||
        "Cash Customer",
      obj.customerMobile || "",
      obj.customerType || "Retail",
      obj.imei || "",
      obj.brand || "",
      obj.model || "",
      obj.memory || "",
      obj.color || "",
      Number(obj.amount) || 0,
      obj.status || "Pending",
      getCurrentUsername_(token)
    ]);

  return {
    success: true,
    orderNo: no
  };
}


function getSaleOrders(token) {

  validateSession(token);

  const sh =
    getSheet_(SHEETS.SALE_ORDERS);

  if (sh.getLastRow() < 2) {
    return [];
  }

  const data =
    sh.getRange(
      2,
      1,
      sh.getLastRow() - 1,
      13
    ).getValues();

  return data.map((r, i) => ({
    row: i + 2,
    orderNo: r[0],
    date: formatDate_(r[1]),
    customer: r[2],
    mobile: r[3],
    customerType: r[4],
    imei: r[5],
    brand: r[6],
    model: r[7],
    memory: r[8],
    color: r[9],
    amount: Number(r[10]) || 0,
    status: r[11],
    user: r[12]
  }));
}


/************************************************************
 LEDGER
************************************************************/
function addLedger(
  token,
  party,
  type,
  reference,
  description,
  debit,
  credit
) {

  adminOrPermission_(token, "add");

  addLedgerInternal_(
    new Date(),
    party,
    type,
    reference,
    description,
    Number(debit) || 0,
    Number(credit) || 0,
    getCurrentUsername_(token)
  );

  return "Ledger entry added.";
}


function addLedgerInternal_(
  date,
  party,
  type,
  reference,
  description,
  debit,
  credit,
  user
) {

  getSheet_(SHEETS.LEDGER)
    .appendRow([
      date,
      party,
      type,
      reference,
      description,
      debit,
      credit,
      "",
      user
    ]);
}


function getLedgerParties(token) {

  validateSession(token);

  const sh =
    getSheet_(SHEETS.LEDGER);

  if (sh.getLastRow() < 2) {
    return [];
  }

  const data =
    sh.getRange(
      2,
      1,
      sh.getLastRow() - 1,
      9
    ).getValues();

  const set = {};

  data.forEach(r => {
    if (r[1]) {
      set[String(r[1])] = true;
    }
  });

  return Object.keys(set).sort();
}


function getLedger(token, filters) {

  validateSession(token);

  filters = filters || {};

  const sh =
    getSheet_(SHEETS.LEDGER);

  if (sh.getLastRow() < 2) {
    return {
      rows: [],
      debit: 0,
      credit: 0,
      balance: 0
    };
  }

  const data =
    sh.getRange(
      2,
      1,
      sh.getLastRow() - 1,
      9
    ).getValues();

  const from =
    parseReportDate_(
      filters.fromDate,
      false
    );

  const to =
    parseReportDate_(
      filters.toDate,
      true
    );

  let rows = [];
  let debit = 0;
  let credit = 0;
  let balance = 0;

  data.forEach(r => {

    if (
      filters.party &&
      filters.party !== "all" &&
      String(r[1]) !==
      String(filters.party)
    ) {
      return;
    }

    const date = new Date(r[0]);

    if (from && date < from) return;
    if (to && date > to) return;

    const d = Number(r[5]) || 0;
    const c = Number(r[6]) || 0;

    debit += d;
    credit += c;
    balance += d - c;

    rows.push({
      date: formatDate_(r[0]),
      party: r[1],
      type: r[2],
      reference: r[3],
      description: r[4],
      debit: d,
      credit: c,
      balance: balance
    });
  });

  return {
    rows: rows,
    debit: debit,
    credit: credit,
    balance: balance
  };
}


/************************************************************
 VOUCHERS
************************************************************/
function createVoucher(token, obj) {

  adminOrPermission_(token, "add");

  obj = obj || {};

  const type =
    String(obj.type || "").trim();

  if (
    type !== "Cash Received" &&
    type !== "Cash Payment" &&
    type !== "Bank Received" &&
    type !== "Bank Payment"
  ) {
    throw new Error(
      "Invalid voucher type."
    );
  }

  const amount =
    Number(obj.amount) || 0;

  if (amount <= 0) {
    throw new Error("Amount required.");
  }

  const no =
    nextNumber_(
      SHEETS.VOUCHERS,
      "VCH",
      1
    );

  const now = new Date();

  const party =
    String(obj.party || "").trim();

  const method =
    obj.paymentMethod ||
    (
      type.indexOf("Bank") === 0
        ? "Bank"
        : "Cash"
    );

  const description =
    obj.description ||
    type;

  const user =
    getCurrentUsername_(token);

  getSheet_(SHEETS.VOUCHERS)
    .appendRow([
      no,
      now,
      type,
      party,
      description,
      amount,
      method,
      user
    ]);

  if (party) {

    if (
      type === "Cash Received" ||
      type === "Bank Received"
    ) {

      addLedgerInternal_(
        now,
        party,
        type,
        no,
        description,
        0,
        amount,
        user
      );

    } else {

      addLedgerInternal_(
        now,
        party,
        type,
        no,
        description,
        amount,
        0,
        user
      );
    }
  }

  addCashBank_(
    now,
    type,
    no,
    description,
    method,
    amount,
    token
  );

  return {
    success: true,
    voucherNo: no,
    amount: amount,
    type: type,
    party: party
  };
}


function getVouchers(token, filters) {

  validateSession(token);

  filters = filters || {};

  const sh =
    getSheet_(SHEETS.VOUCHERS);

  if (sh.getLastRow() < 2) {
    return [];
  }

  const data =
    sh.getRange(
      2,
      1,
      sh.getLastRow() - 1,
      8
    ).getValues();

  const from =
    parseReportDate_(
      filters.fromDate,
      false
    );

  const to =
    parseReportDate_(
      filters.toDate,
      true
    );

  const search =
    String(filters.search || "")
      .trim()
      .toLowerCase();

  return data
    .map((r, i) => ({
      row: i + 2,
      voucherNo: r[0],
      date: formatDate_(r[1]),
      type: r[2],
      party: r[3],
      description: r[4],
      amount: Number(r[5]) || 0,
      paymentMethod: r[6],
      user: r[7],
      rawDate: r[1]
    }))
    .filter(x => {

      if (
        from &&
        new Date(x.rawDate) < from
      ) return false;

      if (
        to &&
        new Date(x.rawDate) > to
      ) return false;

      if (
        filters.type &&
        filters.type !== "all" &&
        String(x.type) !==
        String(filters.type)
      ) return false;

      if (search) {

        const hay = [
          x.voucherNo,
          x.type,
          x.party,
          x.description,
          x.paymentMethod
        ]
          .join(" ")
          .toLowerCase();

        if (
          hay.indexOf(search) === -1
        ) return false;
      }

      return true;

    })
    .map(x => {

      delete x.rawDate;
      return x;

    });
}


function getVoucherByNo(token, voucherNo) {

  validateSession(token);

  const sh =
    getSheet_(SHEETS.VOUCHERS);

  if (sh.getLastRow() < 2) {
    return null;
  }

  const data =
    sh.getRange(
      2,
      1,
      sh.getLastRow() - 1,
      8
    ).getValues();

  for (let i = 0; i < data.length; i++) {

    if (
      String(data[i][0])
        .trim()
        .toLowerCase() ===
      String(voucherNo)
        .trim()
        .toLowerCase()
    ) {

      return {
        row: i + 2,
        voucherNo: data[i][0],
        date: formatDate_(data[i][1]),
        type: data[i][2],
        party: data[i][3],
        description: data[i][4],
        amount: Number(data[i][5]) || 0,
        paymentMethod: data[i][6],
        user: data[i][7]
      };
    }
  }

  return null;
}


function editVoucher(
  token,
  voucherNo,
  obj
) {

  adminOrPermission_(token, "edit");

  const old =
    getVoucherByNo(
      token,
      voucherNo
    );

  if (!old) {
    throw new Error(
      "Voucher not found."
    );
  }

  obj = obj || {};

  const type =
    obj.type || old.type;

  const amount =
    Number(obj.amount) || 0;

  if (amount <= 0) {
    throw new Error("Amount required.");
  }

  const party =
    obj.party != null
      ? String(obj.party).trim()
      : old.party;

  const description =
    obj.description != null
      ? obj.description
      : old.description;

  const method =
    obj.paymentMethod ||
    old.paymentMethod ||
    "Cash";

  const now = new Date();

  const user =
    getCurrentUsername_(token);

  removeRowsByReference_(
    SHEETS.CASHBANK,
    old.voucherNo,
    3
  );

  removeRowsByReference_(
    SHEETS.LEDGER,
    old.voucherNo,
    4
  );

  getSheet_(SHEETS.VOUCHERS)
    .getRange(
      old.row,
      1,
      1,
      8
    )
    .setValues([[
      old.voucherNo,
      now,
      type,
      party,
      description,
      amount,
      method,
      user
    ]]);

  if (party) {

    if (
      type === "Cash Received" ||
      type === "Bank Received"
    ) {

      addLedgerInternal_(
        now,
        party,
        type,
        old.voucherNo,
        description,
        0,
        amount,
        user
      );

    } else {

      addLedgerInternal_(
        now,
        party,
        type,
        old.voucherNo,
        description,
        amount,
        0,
        user
      );
    }
  }

  addCashBank_(
    now,
    type,
    old.voucherNo,
    description,
    method,
    amount,
    token
  );

  return {
    success: true,
    voucherNo: old.voucherNo,
    message:
      "Voucher updated successfully."
  };
}


function deleteVoucher(
  token,
  voucherNo
) {

  adminOrPermission_(token, "delete");

  const old =
    getVoucherByNo(
      token,
      voucherNo
    );

  if (!old) {
    throw new Error(
      "Voucher not found."
    );
  }

  getSheet_(SHEETS.VOUCHERS)
    .deleteRow(old.row);

  removeRowsByReference_(
    SHEETS.CASHBANK,
    old.voucherNo,
    3
  );

  removeRowsByReference_(
    SHEETS.LEDGER,
    old.voucherNo,
    4
  );

  return {
    success: true,
    message:
      "Voucher deleted successfully."
  };
}


/************************************************************
 EXPENSES
************************************************************/
function addExpense(token, obj) {

  adminOrPermission_(token, "add");

  obj = obj || {};

  const sh =
    getSheet_(SHEETS.EXPENSES);

  const no =
    nextNumber_(
      SHEETS.EXPENSES,
      "EXP",
      1
    );

  const amount =
    Number(obj.amount) || 0;

  if (amount <= 0) {
    throw new Error("Amount required.");
  }

  const now = new Date();

  sh.appendRow([
    no,
    now,
    obj.category || "",
    obj.description || "",
    amount,
    obj.paymentMethod || "Cash",
    obj.shop || "",
    getCurrentUsername_(token)
  ]);

  addCashBank_(
    now,
    "Expense",
    no,
    obj.description || "Expense",
    obj.paymentMethod || "Cash",
    amount,
    token
  );

  return {
    success: true,
    expenseNo: no
  };
}


function getExpenses(token, filters) {

  validateSession(token);

  filters = filters || {};

  const sh =
    getSheet_(SHEETS.EXPENSES);

  if (sh.getLastRow() < 2) {
    return [];
  }

  const data =
    sh.getRange(
      2,
      1,
      sh.getLastRow() - 1,
      8
    ).getValues();

  const from =
    parseReportDate_(
      filters.fromDate,
      false
    );

  const to =
    parseReportDate_(
      filters.toDate,
      true
    );

  return data
    .map((r, i) => ({
      row: i + 2,
      no: r[0],
      date: formatDate_(r[1]),
      category: r[2],
      description: r[3],
      amount: Number(r[4]) || 0,
      paymentMethod: r[5],
      shop: r[6],
      rawDate: r[1]
    }))
    .filter(x => {

      if (
        from &&
        new Date(x.rawDate) < from
      ) return false;

      if (
        to &&
        new Date(x.rawDate) > to
      ) return false;

      return true;

    })
    .map(x => {

      delete x.rawDate;
      return x;

    });
}



function getExpenseByNo(token, expenseNo) {
  validateSession(token);
  const sh=getSheet_(SHEETS.EXPENSES);
  if(sh.getLastRow()<2) return null;
  const data=sh.getRange(2,1,sh.getLastRow()-1,8).getValues();
  for(let i=0;i<data.length;i++) if(String(data[i][0]).trim().toLowerCase()===String(expenseNo).trim().toLowerCase())
    return {row:i+2,no:data[i][0],date:formatDate_(data[i][1]),category:data[i][2],description:data[i][3],amount:Number(data[i][4])||0,paymentMethod:data[i][5],shop:data[i][6],user:data[i][7]};
  return null;
}

function editExpense(token, expenseNo, obj) {
  adminOrPermission_(token,"edit");
  const old=getExpenseByNo(token,expenseNo); if(!old) throw new Error("Expense not found.");
  obj=obj||{}; const amount=Number(obj.amount)||0; if(amount<=0) throw new Error("Amount required.");
  const method=obj.paymentMethod||old.paymentMethod||"Cash", desc=obj.description||"Expense", now=new Date(), user=getCurrentUsername_(token);
  removeRowsByReference_(SHEETS.CASHBANK,old.no,3);
  getSheet_(SHEETS.EXPENSES).getRange(old.row,1,1,8).setValues([[old.no,now,obj.category||old.category,desc,amount,method,obj.shop||old.shop,user]]);
  addCashBank_(now,"Expense",old.no,desc,method,amount,token);
  return {success:true,expenseNo:old.no};
}

function deleteExpense(token, row) {

  adminOrPermission_(token, "delete");

  row = Number(row);

  const sh =
    getSheet_(SHEETS.EXPENSES);

  if (
    !row ||
    row < 2 ||
    row > sh.getLastRow()
  ) {
    throw new Error("Expense not found.");
  }

  const no =
    sh.getRange(row, 1)
      .getValue();

  sh.deleteRow(row);

  removeRowsByReference_(
    SHEETS.CASHBANK,
    no,
    3
  );

  return "Expense deleted.";
}


/************************************************************
 CASH / BANK
************************************************************/
function addCashBank_(
  date,
  type,
  reference,
  description,
  method,
  amount,
  token
) {

  amount =
    Number(amount) || 0;

  if (!amount) {
    return;
  }

  method =
    String(method || "Cash")
      .toLowerCase();

  let cash = 0;
  let bank = 0;

  if (method === "cash") {
    cash = amount;
  } else {
    bank = amount;
  }

  let sign = 1;

  if (
    type === "Purchase" ||
    type === "Expense" ||
    type === "Cash Payment" ||
    type === "Bank Payment"
  ) {
    sign = -1;
  }

  getSheet_(SHEETS.CASHBANK)
    .appendRow([
      date,
      sign > 0 ? "In" : "Out",
      reference,
      description,
      cash * sign,
      bank * sign,
      getCurrentUsername_(token)
    ]);
}


function getCashBankBalance(token) {

  validateSession(token);

  const sh =
    getSheet_(SHEETS.CASHBANK);

  if (sh.getLastRow() < 2) {

    return {
      cash: 0,
      bank: 0,
      total: 0
    };
  }

  const data =
    sh.getRange(
      2,
      1,
      sh.getLastRow() - 1,
      7
    ).getValues();

  let cash = 0;
  let bank = 0;

  data.forEach(r => {

    cash += Number(r[4]) || 0;
    bank += Number(r[5]) || 0;

  });

  return {
    cash: cash,
    bank: bank,
    total: cash + bank
  };
}


function getCashBankReport(token) {

  validateSession(token);

  const sh =
    getSheet_(SHEETS.CASHBANK);

  if (sh.getLastRow() < 2) {
    return [];
  }

  const data =
    sh.getRange(
      2,
      1,
      sh.getLastRow() - 1,
      7
    ).getValues();

  return data.map((r, i) => ({
    row: i + 2,
    date: formatDate_(r[0]),
    type: r[1],
    reference: r[2],
    description: r[3],
    cash: Number(r[4]) || 0,
    bank: Number(r[5]) || 0,
    user: r[6]
  }));
}


/************************************************************
 DAILY CLOSING
************************************************************/
function dailyClosing(token, obj) {

  adminOrPermission_(token, "add");

  obj = obj || {};

  const balance =
    getCashBankBalance(token);

  const actual =
    Number(obj.actualCash) || 0;

  const difference =
    actual - balance.cash;

  getSheet_(SHEETS.CLOSING)
    .appendRow([
      new Date(),
      balance.cash,
      Number(obj.cashIn) || 0,
      Number(obj.cashOut) || 0,
      balance.cash,
      actual,
      difference,
      obj.notes || "",
      getCurrentUsername_(token)
    ]);

  return {
    success: true,
    expectedCash: balance.cash,
    actualCash: actual,
    difference: difference
  };
}


function getDailyClosing(token) {

  validateSession(token);

  const sh =
    getSheet_(SHEETS.CLOSING);

  if (sh.getLastRow() < 2) {
    return [];
  }

  const data =
    sh.getRange(
      2,
      1,
      sh.getLastRow() - 1,
      9
    ).getValues();

  return data.map((r, i) => ({
    row: i + 2,
    date: formatDate_(r[0]),
    openingCash: Number(r[1]) || 0,
    cashIn: Number(r[2]) || 0,
    cashOut: Number(r[3]) || 0,
    expectedCash: Number(r[4]) || 0,
    actualCash: Number(r[5]) || 0,
    difference: Number(r[6]) || 0,
    notes: r[7]
  }));
}


/************************************************************
 RETURNS
************************************************************/
function createReturn(token, obj) {

  adminOrPermission_(token, "add");

  obj = obj || {};

  const imei =
    String(obj.imei || "").trim();

  if (!imei) {
    throw new Error("IMEI required.");
  }

  const result =
    findStockRow_(imei);

  if (!result) {
    throw new Error("IMEI not found.");
  }

  const r = result.values;

  const type =
    obj.type || "Sale Return";

  const no =
    nextNumber_(
      SHEETS.RETURNS,
      "RET",
      1
    );

  const amount =
    Number(obj.amount) || 0;

  const now = new Date();

  getSheet_(SHEETS.RETURNS)
    .appendRow([
      no,
      now,
      type,
      obj.party || "",
      imei,
      r[1],
      r[2],
      r[3],
      r[4],
      amount,
      obj.reason || "",
      obj.paymentMethod || "Cash",
      getCurrentUsername_(token)
    ]);

  const stock =
    getSheet_(SHEETS.STOCK);

  /*
   Sale Return:
   Sold -> Available
  */
  if (
    type.toLowerCase() ===
    "sale return"
  ) {

    if (
      String(r[10]).toLowerCase() ===
      "sold"
    ) {

      stock.getRange(
        result.row,
        11,
        1,
        4
      ).setValues([[
        "Available",
        r[11],
        "",
        ""
      ]]);
    }

    if (obj.party && amount > 0) {

      addLedgerInternal_(
        now,
        obj.party,
        "Sale Return",
        no,
        "Sale Return",
        0,
        amount,
        getCurrentUsername_(token)
      );
    }

    if (amount > 0) {

      addCashBank_(
        now,
        "Sale Return",
        no,
        "Sale Return Refund",
        obj.paymentMethod || "Cash",
        amount,
        token
      );
    }
  }

  /*
   Purchase Return:
   Available -> Returned
  */
  if (
    type.toLowerCase() ===
    "purchase return"
  ) {

    stock.getRange(
      result.row,
      11
    ).setValue("Returned");

    if (obj.party && amount > 0) {

      addLedgerInternal_(
        now,
        obj.party,
        "Purchase Return",
        no,
        "Purchase Return",
        amount,
        0,
        getCurrentUsername_(token)
      );
    }

    if (amount > 0) {

      addCashBank_(
        now,
        "Purchase Return",
        no,
        "Purchase Return Received",
        obj.paymentMethod || "Cash",
        amount,
        token
      );
    }
  }

  return {
    success: true,
    returnNo: no
  };
}


function getReturns(token) {

  validateSession(token);

  const sh =
    getSheet_(SHEETS.RETURNS);

  if (sh.getLastRow() < 2) {
    return [];
  }

  const data =
    sh.getRange(
      2,
      1,
      sh.getLastRow() - 1,
      13
    ).getValues();

  return data.map((r, i) => ({
    row: i + 2,
    no: r[0],
    date: formatDate_(r[1]),
    type: r[2],
    party: r[3],
    imei: r[4],
    brand: r[5],
    model: r[6],
    memory: r[7],
    color: r[8],
    amount: Number(r[9]) || 0,
    reason: r[10],
    paymentMethod: r[11],
    user: r[12]
  }));
}



function getReturnByNo(token, returnNo) {
  validateSession(token); const sh=getSheet_(SHEETS.RETURNS); if(sh.getLastRow()<2)return null;
  const data=sh.getRange(2,1,sh.getLastRow()-1,13).getValues();
  for(let i=0;i<data.length;i++) if(String(data[i][0]).trim().toLowerCase()===String(returnNo).trim().toLowerCase())
    return {row:i+2,no:data[i][0],date:formatDate_(data[i][1]),type:data[i][2],party:data[i][3],imei:data[i][4],brand:data[i][5],model:data[i][6],memory:data[i][7],color:data[i][8],amount:Number(data[i][9])||0,reason:data[i][10],paymentMethod:data[i][11],user:data[i][12]};
  return null;
}

function deleteReturn(token, returnNo) {
  adminOrPermission_(token,"delete"); const old=getReturnByNo(token,returnNo); if(!old)throw new Error("Return not found.");
  const stock=findStockRow_(old.imei); const type=String(old.type).toLowerCase();
  if(stock){ if(type==="sale return") getSheet_(SHEETS.STOCK).getRange(stock.row,11,1,4).setValues([["Sold",stock.values[11],new Date(),stock.values[13]||""]]);
    if(type==="purchase return") getSheet_(SHEETS.STOCK).getRange(stock.row,11).setValue("Available"); }
  removeRowsByReference_(SHEETS.CASHBANK,old.no,3); removeRowsByReference_(SHEETS.LEDGER,old.no,4); getSheet_(SHEETS.RETURNS).deleteRow(old.row); return {success:true};
}

function editReturn(token, returnNo, obj) {
  adminOrPermission_(token,"edit");
  const old=getReturnByNo(token,returnNo); if(!old)throw new Error("Return not found.");
  obj=obj||{};
  const oldStock=findStockRow_(old.imei);
  if(oldStock){
    if(String(old.type).toLowerCase()==="sale return") getSheet_(SHEETS.STOCK).getRange(oldStock.row,11,1,4).setValues([["Sold",oldStock.values[11],oldStock.values[12]||"",oldStock.values[13]||""]]);
    if(String(old.type).toLowerCase()==="purchase return") getSheet_(SHEETS.STOCK).getRange(oldStock.row,11).setValue("Available");
  }
  removeRowsByReference_(SHEETS.CASHBANK,old.no,3); removeRowsByReference_(SHEETS.LEDGER,old.no,4);
  const newImei=String(obj.imei||old.imei).trim(), result=findStockRow_(newImei); if(!result)throw new Error("IMEI not found."); const r=result.values;
  const type=obj.type||old.type, amount=Number(obj.amount!=null?obj.amount:old.amount)||0, party=obj.party!=null?String(obj.party):old.party, method=obj.paymentMethod||old.paymentMethod||"Cash", reason=obj.reason!=null?obj.reason:old.reason, now=new Date(), user=getCurrentUsername_(token);
  getSheet_(SHEETS.RETURNS).getRange(old.row,1,1,13).setValues([[old.no,now,type,party,newImei,r[1],r[2],r[3],r[4],amount,reason,method,user]]);
  if(String(type).toLowerCase()==="sale return") { getSheet_(SHEETS.STOCK).getRange(result.row,11,1,4).setValues([["Available",r[11],"",""]]); if(party&&amount)addLedgerInternal_(now,party,type,old.no,reason,0,amount,user); if(amount)addCashBank_(now,type,old.no,reason,method,amount,token); }
  else if(String(type).toLowerCase()==="purchase return") { getSheet_(SHEETS.STOCK).getRange(result.row,11).setValue("Returned"); if(party&&amount)addLedgerInternal_(now,party,type,old.no,reason,amount,0,user); if(amount)addCashBank_(now,type,old.no,reason,method,amount,token); }
  return {success:true,returnNo:old.no};
}

/************************************************************
 STOCK TRANSFER
************************************************************/
function stockTransfer(token, obj) {

  adminOrPermission_(token, "edit");

  obj = obj || {};

  const imei =
    String(obj.imei || "").trim();

  if (!imei) {
    throw new Error("IMEI required.");
  }

  const result =
    findStockRow_(imei);

  if (!result) {
    throw new Error("IMEI not found.");
  }

  if (
    String(result.values[10]).toLowerCase() ===
    "sold"
  ) {
    throw new Error(
      "Sold stock cannot be transferred."
    );
  }

  const fromShop =
    obj.fromShop ||
    result.values[9] ||
    "";

  const toShop =
    obj.toShop || "";

  if (!toShop) {
    throw new Error("To Shop required.");
  }

  const no =
    nextNumber_(
      SHEETS.TRANSFERS,
      "TR",
      1
    );

  getSheet_(SHEETS.STOCK)
    .getRange(result.row, 10)
    .setValue(toShop);

  getSheet_(SHEETS.TRANSFERS)
    .appendRow([
      no,
      new Date(),
      imei,
      result.values[1],
      result.values[2],
      result.values[3],
      result.values[4],
      fromShop,
      toShop,
      getCurrentUsername_(token)
    ]);

  return {
    success: true,
    transferNo: no
  };
}


function getTransfers(token) {

  validateSession(token);

  const sh =
    getSheet_(SHEETS.TRANSFERS);

  if (sh.getLastRow() < 2) {
    return [];
  }

  const data =
    sh.getRange(
      2,
      1,
      sh.getLastRow() - 1,
      10
    ).getValues();

  return data.map((r, i) => ({
    row: i + 2,
    transferNo: r[0],
    date: formatDate_(r[1]),
    imei: r[2],
    brand: r[3],
    model: r[4],
    memory: r[5],
    color: r[6],
    fromShop: r[7],
    toShop: r[8],
    user: r[9]
  }));
}



function getTransferByNo(token, no) { validateSession(token); const sh=getSheet_(SHEETS.TRANSFERS); if(sh.getLastRow()<2)return null; const d=sh.getRange(2,1,sh.getLastRow()-1,10).getValues(); for(let i=0;i<d.length;i++) if(String(d[i][0]).toLowerCase()===String(no).toLowerCase()) return {row:i+2,transferNo:d[i][0],date:formatDate_(d[i][1]),imei:d[i][2],brand:d[i][3],model:d[i][4],memory:d[i][5],color:d[i][6],fromShop:d[i][7],toShop:d[i][8],user:d[i][9]}; return null; }
function deleteTransfer(token,no){ adminOrPermission_(token,"delete"); const old=getTransferByNo(token,no); if(!old)throw new Error("Transfer not found."); const stock=findStockRow_(old.imei); if(stock)getSheet_(SHEETS.STOCK).getRange(stock.row,10).setValue(old.fromShop); getSheet_(SHEETS.TRANSFERS).deleteRow(old.row); return {success:true}; }
function editTransfer(token,no,obj){ adminOrPermission_(token,"edit"); const old=getTransferByNo(token,no); if(!old)throw new Error("Transfer not found."); obj=obj||{}; const stock=findStockRow_(old.imei); if(!stock)throw new Error("IMEI not found."); const toShop=obj.toShop||old.toShop; if(!toShop)throw new Error("To Shop required."); getSheet_(SHEETS.STOCK).getRange(stock.row,10).setValue(toShop); getSheet_(SHEETS.TRANSFERS).getRange(old.row,1,1,10).setValues([[old.transferNo,new Date(),old.imei,old.brand,old.model,old.memory,old.color,obj.fromShop||old.fromShop,toShop,getCurrentUsername_(token)]]); return {success:true}; }

/************************************************************
 DASHBOARD - FULL
************************************************************/
function getDashboardStats(token) {

  validateSession(token);

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const now = new Date();

  const start =
    new Date(
      now.getFullYear(),
      now.getMonth(),
      now.getDate(),
      0, 0, 0, 0
    );

  const end =
    new Date(
      now.getFullYear(),
      now.getMonth(),
      now.getDate(),
      23, 59, 59, 999
    );

  /**************** STOCK ****************/

  let totalStock = 0;
  let availableStock = 0;
  let soldStock = 0;
  let stockValue = 0;
  let stockSaleValue = 0;

  const stockSh =
    getSheet_(SHEETS.STOCK);

  if (stockSh.getLastRow() >= 2) {

    const data =
      stockSh.getRange(
        2,
        1,
        stockSh.getLastRow() - 1,
        14
      ).getValues();

    data.forEach(r => {

      totalStock++;

      const status =
        String(r[10])
          .toLowerCase();

      if (status === "available") {

        availableStock++;

        stockValue +=
          Number(r[5]) || 0;

        stockSaleValue +=
          Number(r[6]) || 0;

      } else if (
        status === "sold"
      ) {

        soldStock++;
      }
    });
  }


  /**************** SALES ****************/

  let totalSales = 0;
  let todaySale = 0;
  let todaySaleQty = 0;

  let grossProfit = 0;
  let todayGrossProfit = 0;

  let totalPayments = 0;
  let todayPayments = 0;

  const salesSh =
    getSheet_(SHEETS.SALES);

  if (salesSh.getLastRow() >= 2) {

    const data =
      salesSh.getRange(
        2,
        1,
        salesSh.getLastRow() - 1,
        18
      ).getValues();

    data.forEach(r => {

      const date =
        new Date(r[1]);

      const amount =
        Number(r[13]) || 0;

      const payment =
        Number(r[14]) || 0;

      const profit =
        Number(r[16]) || 0;

      totalSales += amount;
      grossProfit += profit;
      totalPayments += payment;

      if (
        date >= start &&
        date <= end
      ) {

        todaySale += amount;
        todaySaleQty++;
        todayGrossProfit += profit;
        todayPayments += payment;
      }
    });
  }


  /**************** EXPENSES ****************/

  let totalExpenses = 0;
  let todayExpenses = 0;

  const expenseSh =
    getSheet_(SHEETS.EXPENSES);

  if (expenseSh.getLastRow() >= 2) {

    const data =
      expenseSh.getRange(
        2,
        1,
        expenseSh.getLastRow() - 1,
        8
      ).getValues();

    data.forEach(r => {

      const date =
        new Date(r[1]);

      const amount =
        Number(r[4]) || 0;

      totalExpenses += amount;

      if (
        date >= start &&
        date <= end
      ) {
        todayExpenses += amount;
      }
    });
  }


  /**************** CASH BANK ****************/

  const cashBank =
    getCashBankBalance(token);


  /**************** RECEIVABLE / PAYABLE ****************/

  let receivables = 0;
  let payables = 0;

  const ledgerSh =
    getSheet_(SHEETS.LEDGER);

  const partyBalances = {};

  if (ledgerSh.getLastRow() >= 2) {

    const data =
      ledgerSh.getRange(
        2,
        1,
        ledgerSh.getLastRow() - 1,
        9
      ).getValues();

    data.forEach(r => {

      const party =
        String(r[1] || "").trim();

      if (!party) return;

      const debit =
        Number(r[5]) || 0;

      const credit =
        Number(r[6]) || 0;

      if (!partyBalances[party]) {
        partyBalances[party] = 0;
      }

      partyBalances[party] +=
        debit - credit;
    });
  }

  Object.keys(partyBalances)
    .forEach(party => {

      const balance =
        partyBalances[party];

      if (balance > 0) {
        receivables += balance;
      } else if (balance < 0) {
        payables += Math.abs(balance);
      }

    });


  const netProfit =
    grossProfit -
    totalExpenses;

  const todayNetProfit =
    todayGrossProfit -
    todayExpenses;

  return {

    /* Stock */
    totalStock: totalStock,
    availableStock: availableStock,
    soldStock: soldStock,
    stockValue: stockValue,
    stockSaleValue: stockSaleValue,

    /* Sales */
    totalSales: totalSales,
    todaySale: todaySale,
    todaySaleQty: todaySaleQty,

    /* Profit */
    grossProfit: grossProfit,
    todayGrossProfit: todayGrossProfit,
    netProfit: netProfit,
    todayNetProfit: todayNetProfit,

    /* Payments */
    totalPayments: totalPayments,
    todayPayments: todayPayments,

    /* Expenses */
    totalExpenses: totalExpenses,
    todayExpenses: todayExpenses,

    /* Cash Bank */
    cashIn: cashBank.cash > 0
      ? cashBank.cash
      : 0,

    cashBalance: cashBank.cash,
    bankBalance: cashBank.bank,
    totalCashBank: cashBank.total,

    /* Party */
    receivables: receivables,
    payables: payables,

    /* Dashboard compatibility */
    available: availableStock,
    total: totalStock,
    sold: soldStock,
    salesTotal: totalSales,
    profitTotal: grossProfit,

    /* Orders */
    confirmOrders:
      getOrderCount_(token, "Confirmed"),

    pendingOrders:
      getOrderCount_(token, "Pending"),

    /* In Hand */
    inHand: cashBank.cash

  };
}


function getOrderCount_(token, status) {

  validateSession(token);

  const sh =
    getSheet_(SHEETS.SALE_ORDERS);

  if (sh.getLastRow() < 2) {
    return 0;
  }

  const data =
    sh.getRange(
      2,
      1,
      sh.getLastRow() - 1,
      13
    ).getValues();

  let count = 0;

  data.forEach(r => {

    if (
      String(r[11]).toLowerCase() ===
      String(status).toLowerCase()
    ) {
      count++;
    }
  });

  return count;
}


/************************************************************
 PROFIT LOSS
************************************************************/
function getProfitLoss(token, filters) {

  validateSession(token);

  filters = filters || {};

  const sales =
    getSaleReport(
      token,
      filters
    );

  const purchases =
    getPurchaseReport(
      token,
      filters
    );

  let salesTotal = 0;
  let grossProfit = 0;

  sales.rows.forEach(r => {

    salesTotal +=
      Number(r.amount) || 0;

    grossProfit +=
      Number(r.profit) || 0;
  });

  const costOfSold =
    salesTotal -
    grossProfit;

  let expenses = 0;

  const exp =
    getExpenses(
      token,
      filters
    );

  exp.forEach(r => {
    expenses +=
      Number(r.amount) || 0;
  });

  const netProfit =
    grossProfit -
    expenses;

  return {
    purchaseTotal:
      purchases.totalAmount,

    sales:
      salesTotal,

    costOfSold:
      costOfSold,

    grossProfit:
      grossProfit,

    expenses:
      expenses,

    netProfit:
      netProfit,

    salesQty:
      sales.totalQty
  };
}


/************************************************************
 RECEIVABLE / PAYABLE
************************************************************/
function getReceivablePayable(token) {

  validateSession(token);

  const sh =
    getSheet_(SHEETS.LEDGER);

  const balances = {};

  if (sh.getLastRow() >= 2) {

    const data =
      sh.getRange(
        2,
        1,
        sh.getLastRow() - 1,
        9
      ).getValues();

    data.forEach(r => {

      const party =
        String(r[1] || "").trim();

      if (!party) return;

      const debit =
        Number(r[5]) || 0;

      const credit =
        Number(r[6]) || 0;

      if (!balances[party]) {
        balances[party] = 0;
      }

      balances[party] +=
        debit - credit;
    });
  }

  const rows = [];

  Object.keys(balances)
    .forEach(party => {

      const balance =
        balances[party];

      rows.push({
        party: party,
        balance: balance,
        receivable:
          balance > 0
            ? balance
            : 0,
        payable:
          balance < 0
            ? Math.abs(balance)
            : 0
      });
    });

  return rows.sort(
    (a, b) =>
      Math.abs(b.balance) -
      Math.abs(a.balance)
  );
}


/************************************************************
 SEARCH EVERYTHING
************************************************************/
function globalSearch(token, search) {

  validateSession(token);

  search =
    String(search || "")
      .trim()
      .toLowerCase();

  if (!search) {
    return [];
  }

  const results = [];

  /* Stock */
  getAllStock(token)
    .forEach(r => {

      const hay = [
        r.imei,
        r.brand,
        r.model,
        r.memory,
        r.color,
        r.supplier,
        r.shop,
        r.invoiceNo
      ]
        .join(" ")
        .toLowerCase();

      if (hay.indexOf(search) !== -1) {

        results.push({
          type: "Stock",
          title: r.imei,
          detail:
            r.brand + " " +
            r.model + " " +
            r.color,
          status: r.stockStatus
        });
      }
    });

  /* Sales */
  getAllSales(token, {
    search: search
  }).forEach(r => {

    results.push({
      type: "Sale",
      title: r.invoiceNo,
      detail:
        r.customerName +
        " - " +
        r.imei,
      status: "Sale"
    });
  });

  return results.slice(0, 100);
}


/************************************************************
 INVOICE HTML
************************************************************/


/************************************************************
 JOURNAL VOUCHERS
************************************************************/
function createJournalVoucher(token,obj){
  adminOrPermission_(token,"add"); obj=obj||{}; const debit=String(obj.debitAccount||"").trim(), credit=String(obj.creditAccount||"").trim(), amount=Number(obj.amount)||0;
  if(!debit||!credit)throw new Error("Debit and Credit accounts required."); if(debit===credit)throw new Error("Debit and Credit cannot be same."); if(amount<=0)throw new Error("Amount required.");
  const no=nextNumber_(SHEETS.JOURNAL,"JV",1), now=new Date(), user=getCurrentUsername_(token); getSheet_(SHEETS.JOURNAL).appendRow([no,now,debit,credit,amount,obj.description||"",user]);
  addLedgerInternal_(now,debit,"Journal Voucher",no,obj.description||"JV",amount,0,user); addLedgerInternal_(now,credit,"Journal Voucher",no,obj.description||"JV",0,amount,user);
  if(/^cash$/i.test(debit)||/^bank$/i.test(debit)) addCashBank_(now,"JV",no,obj.description||"JV",debit,amount,token);
  if(/^cash$/i.test(credit)||/^bank$/i.test(credit)) addCashBank_(now,"Cash Payment",no,obj.description||"JV",credit,amount,token);
  return {success:true,jvNo:no};
}
function getJournalVouchers(token){validateSession(token); const sh=getSheet_(SHEETS.JOURNAL); if(sh.getLastRow()<2)return []; return sh.getRange(2,1,sh.getLastRow()-1,7).getValues().map((r,i)=>({row:i+2,jvNo:r[0],date:formatDate_(r[1]),debitAccount:r[2],creditAccount:r[3],amount:Number(r[4])||0,description:r[5],user:r[6]}));}
function getJournalVoucherByNo(token,no){validateSession(token); const rows=getJournalVouchers(token); return rows.find(x=>String(x.jvNo).toLowerCase()===String(no).toLowerCase())||null;}
function editJournalVoucher(token,no,obj){adminOrPermission_(token,"edit"); const old=getJournalVoucherByNo(token,no); if(!old)throw new Error("JV not found."); obj=obj||{}; const debit=String(obj.debitAccount||old.debitAccount).trim(), credit=String(obj.creditAccount||old.creditAccount).trim(), amount=Number(obj.amount!=null?obj.amount:old.amount)||0; if(!debit||!credit||debit===credit||amount<=0)throw new Error("Valid debit, credit and amount required."); removeRowsByReference_(SHEETS.LEDGER,old.jvNo,4); getSheet_(SHEETS.JOURNAL).getRange(old.row,1,1,7).setValues([[old.jvNo,new Date(),debit,credit,amount,obj.description!=null?obj.description:old.description,getCurrentUsername_(token)]]); const user=getCurrentUsername_(token),desc=obj.description!=null?obj.description:old.description; addLedgerInternal_(new Date(),debit,"Journal Voucher",old.jvNo,desc,amount,0,user); addLedgerInternal_(new Date(),credit,"Journal Voucher",old.jvNo,desc,0,amount,user); return {success:true,jvNo:old.jvNo};}
function deleteJournalVoucher(token,no){adminOrPermission_(token,"delete"); const old=getJournalVoucherByNo(token,no); if(!old)throw new Error("JV not found."); getSheet_(SHEETS.JOURNAL).deleteRow(old.row); removeRowsByReference_(SHEETS.LEDGER,old.jvNo,4); return {success:true};}

function getSaleInvoiceHtml(token, invoiceNo){
  const rows=getSaleRowsByInvoice(token,invoiceNo); if(!rows.length)throw new Error("Invoice not found."); const s=rows[0];
  const esc=v=>String(v==null?"":v).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"), total=rows.reduce((a,r)=>a+Number(r.finalAmount||0),0), discount=rows.reduce((a,r)=>a+Number(r.discount||0),0), paid=Number(s.paymentAmount||0), profit=rows.reduce((a,r)=>a+Number(r.profit||0),0);
  const itemRows=rows.map(r=>`<tr><td>${esc(r.imei)}</td><td>${esc(r.brand)}</td><td>${esc(r.model)}</td><td>${esc(r.memory)}</td><td>${esc(r.color)}</td><td>${Number(r.salePrice||0).toFixed(2)}</td></tr>`).join('');
  return `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>${esc(s.invoiceNo)}</title><style>body{font-family:Arial,sans-serif;margin:0;padding:25px;background:#f4f6f8;color:#111}.invoice{max-width:900px;margin:auto;background:#fff;border:1px solid #ddd;padding:30px}.header{display:flex;justify-content:space-between;border-bottom:3px solid #0878d1;padding-bottom:18px}.shop{font-size:25px;font-weight:bold;color:#0878d1}.title{font-size:22px;font-weight:bold}.grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin:22px 0}.info{border:1px solid #eee;padding:10px}table{width:100%;border-collapse:collapse;margin-top:15px}th,td{border:1px solid #ddd;padding:10px;text-align:left}th{background:#f1f5f9}.footer{margin-top:40px;border-top:1px solid #ddd;padding-top:15px;font-size:13px}@media print{body{background:#fff;padding:0}.invoice{border:0}}</style></head><body><div class="invoice"><div class="header"><div><div class="shop">MOBILE SHOP</div><div>Mobile Phone Management</div></div><div><div class="title">SALE INVOICE</div><b>${esc(s.invoiceNo)}</b><br>${esc(s.date)}</div></div><div class="grid"><div class="info"><b>Customer:</b><br>${esc(s.customerName)}</div><div class="info"><b>Mobile:</b><br>${esc(s.customerMobile)}</div><div class="info"><b>Customer Type:</b><br>${esc(s.customerType)}</div><div class="info"><b>Payment Method:</b><br>${esc(s.paymentMethod)}</div></div><table><tr><th>IMEI</th><th>Brand</th><th>Model</th><th>Memory</th><th>Color</th><th>Unit Sale Price</th></tr>${itemRows}</table><table><tr><td>Quantity</td><td>${rows.length}</td></tr><tr><td>Discount</td><td>${discount.toFixed(2)}</td></tr><tr><td><b>Final Amount</b></td><td><b>${total.toFixed(2)}</b></td></tr><tr><td>Paid</td><td>${paid.toFixed(2)}</td></tr><tr><td>Balance</td><td>${(total-paid).toFixed(2)}</td></tr></table><div class="footer">Sold By: ${esc(s.soldBy)} | Profit: ${profit.toFixed(2)}</div></div></body></html>`;
}
function createSaleInvoicePdf(
  token,
  invoiceNo
) {

  const html =
    getSaleInvoiceHtml(
      token,
      invoiceNo
    );

  const blob =
    Utilities
      .newBlob(
        html,
        "text/html",
        invoiceNo + ".html"
      )
      .getAs("application/pdf")
      .setName(
        invoiceNo + ".pdf"
      );

  const file =
    DriveApp.createFile(blob);

  return {
    url: file.getUrl(),
    name: file.getName()
  };
}


function createVoucherPdf(
  token,
  voucherNo
) {

  const html =
    getVoucherHtml(
      token,
      voucherNo
    );

  const blob =
    Utilities
      .newBlob(
        html,
        "text/html",
        voucherNo + ".html"
      )
      .getAs("application/pdf")
      .setName(
        voucherNo + ".pdf"
      );

  const file =
    DriveApp.createFile(blob);

  return {
    url: file.getUrl(),
    name: file.getName()
  };
}


/************************************************************
 EXCEL / CSV EXPORT
************************************************************/
function createSaleInvoiceExcel(
  token,
  invoiceNo
) {

  const s =
    getSaleByInvoice(
      token,
      invoiceNo
    );

  if (!s) {
    throw new Error(
      "Invoice not found."
    );
  }

  const rows = [
    [
      "Invoice No",
      "Date",
      "Customer",
      "Mobile",
      "Type",
      "IMEI",
      "Brand",
      "Model",
      "Memory",
      "Color",
      "Purchase Price",
      "Sale Price",
      "Discount",
      "Final Amount",
      "Payment",
      "Payment Method",
      "Profit",
      "Sold By"
    ],

    [
      s.invoiceNo,
      s.date,
      s.customerName,
      s.customerMobile,
      s.customerType,
      s.imei,
      s.brand,
      s.model,
      s.memory,
      s.color,
      s.purchasePrice,
      s.salePrice,
      s.discount,
      s.finalAmount,
      s.paymentAmount,
      s.paymentMethod,
      s.profit,
      s.soldBy
    ]
  ];

  const csv =
    rows.map(row =>
      row.map(v =>
        '"' +
        String(
          v == null ? "" : v
        ).replace(/"/g, '""') +
        '"'
      ).join(",")
    ).join("\n");

  const file =
    DriveApp.createFile(
      Utilities.newBlob(
        csv,
        "text/csv",
        invoiceNo + ".csv"
      )
    );

  return {
    url: file.getUrl(),
    name: file.getName()
  };
}


function createVoucherExcel(
  token,
  voucherNo
) {

  const v =
    getVoucherByNo(
      token,
      voucherNo
    );

  if (!v) {
    throw new Error(
      "Voucher not found."
    );
  }

  const rows = [
    [
      "Voucher No",
      "Date",
      "Type",
      "Party",
      "Description",
      "Amount",
      "Payment Method",
      "User"
    ],

    [
      v.voucherNo,
      v.date,
      v.type,
      v.party,
      v.description,
      v.amount,
      v.paymentMethod,
      v.user
    ]
  ];

  const csv =
    rows.map(row =>
      row.map(v =>
        '"' +
        String(
          v == null ? "" : v
        ).replace(/"/g, '""') +
        '"'
      ).join(",")
    ).join("\n");

  const file =
    DriveApp.createFile(
      Utilities.newBlob(
        csv,
        "text/csv",
        voucherNo + ".csv"
      )
    );

  return {
    url: file.getUrl(),
    name: file.getName()
  };
}


/************************************************************
 REMOVE ACCOUNTING REFERENCES
************************************************************/
function removeRowsByReference_(
  sheetName,
  reference,
  refCol
) {

  const sh =
    getSheet_(sheetName);

  if (sh.getLastRow() < 2) {
    return 0;
  }

  const col =
    Number(refCol) || 1;

  const values =
    sh.getRange(
      2,
      col,
      sh.getLastRow() - 1,
      1
    ).getValues();

  const rows = [];

  values.forEach((r, i) => {

    if (
      String(r[0]).trim().toLowerCase() ===
      String(reference).trim().toLowerCase()
    ) {
      rows.push(i + 2);
    }
  });

  for (
    let i = rows.length - 1;
    i >= 0;
    i--
  ) {
    sh.deleteRow(rows[i]);
  }

  return rows.length;
}


/************************************************************
 FIND STOCK
************************************************************/
function findStockRow_(imei) {

  const sh =
    getSheet_(SHEETS.STOCK);

  if (sh.getLastRow() < 2) {
    return null;
  }

  const data =
    sh.getRange(
      2,
      1,
      sh.getLastRow() - 1,
      14
    ).getValues();

  const search =
    String(imei)
      .trim()
      .toLowerCase();

  for (let i = 0; i < data.length; i++) {

    if (
      String(data[i][0])
        .trim()
        .toLowerCase() ===
      search
    ) {

      return {
        row: i + 2,
        values: data[i]
      };
    }
  }

  return null;
}


/************************************************************
 CURRENT USER
************************************************************/
function getCurrentUsername_(token) {

  return CacheService
    .getScriptCache()
    .get(
      "SESSION_" + token
    ) || "";
}


/************************************************************
 DATE
************************************************************/
function formatDate_(date) {

  if (!date) {
    return "";
  }

  try {

    return Utilities.formatDate(
      new Date(date),
      Session.getScriptTimeZone(),
      "dd-MM-yyyy HH:mm"
    );

  } catch (e) {

    return String(date);
  }
}


function parseReportDate_(
  value,
  endOfDay
) {

  if (!value) {
    return null;
  }

  const parts =
    String(value).split("-");

  if (parts.length !== 3) {
    return null;
  }

  const y =
    Number(parts[0]);

  const m =
    Number(parts[1]) - 1;

  const d =
    Number(parts[2]);

  if (endOfDay) {

    return new Date(
      y,
      m,
      d,
      23,
      59,
      59,
      999
    );
  }

  return new Date(
    y,
    m,
    d,
    0,
    0,
    0,
    0
  );
}
