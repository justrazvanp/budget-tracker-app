/**
 * Personal finance backend — Google Sheet + Apps Script Web App.
 * Run setupSheets() once from the editor to create/seed the tabs, then deploy as a Web App.
 */

var SHEETS = {
  ACCOUNTS: 'Accounts',
  TRANSACTIONS: 'Transactions',
  TRANSFERS: 'Transfers',
  CATEGORIES: 'Categories',
  RECURRING: 'Recurring',
  BUDGETS: 'Budgets'
};

function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var accounts = getOrCreateSheet_(ss, SHEETS.ACCOUNTS);
  accounts.clear();
  accounts.getRange(1, 1, 1, 4).setValues([['id', 'name', 'currency', 'opening_balance']]);
  accounts.getRange(2, 1, 4, 4).setValues([
    [1, 'Checking', 'RON', 0],
    [2, 'Fond de urgență', 'EUR', 0],
    [3, 'Economii pe termen lung / Neprevăzute', 'RON', 0],
    [4, 'Signal Iduna Investment', 'EUR', 0]
  ]);

  var transactions = getOrCreateSheet_(ss, SHEETS.TRANSACTIONS);
  transactions.clear();
  transactions.getRange(1, 1, 1, 7).setValues([
    ['id', 'account_id', 'date', 'type', 'category', 'description', 'amount']
  ]);

  var transfers = getOrCreateSheet_(ss, SHEETS.TRANSFERS);
  transfers.clear();
  transfers.getRange(1, 1, 1, 8).setValues([
    ['id', 'date', 'source_account_id', 'dest_account_id', 'source_amount', 'dest_amount', 'fx_rate', 'description']
  ]);

  var categories = getOrCreateSheet_(ss, SHEETS.CATEGORIES);
  categories.clear();
  categories.getRange(1, 1, 1, 2).setValues([['type', 'name']]);
  var incomeCategories = ['Salariu (net)', 'Side hustle (net)', 'Cadou', 'Dobândă', 'Alte surse', 'Datorii/Împrumut'];
  var expenseCategories = [
    'Cheltuieli uzuale/Necesități', 'Cheltuieli pe termen lung/Neprevăzute', 'Educație',
    'Divertisment/Vacanțe', 'Donații/Cadouri', 'Corecții', 'Datorii/Împrumut'
  ];
  var rows = incomeCategories.map(function (name) { return ['Income', name]; })
    .concat(expenseCategories.map(function (name) { return ['Expense', name]; }));
  categories.getRange(2, 1, rows.length, 2).setValues(rows);

  var defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) ss.deleteSheet(defaultSheet);
}

// Phase 4: adds the Recurring/Budgets tabs without touching any existing sheet or data.
// Safe to run on an already-live spreadsheet — run this once instead of setupSheets().
function setupPhase4Sheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var recurring = getOrCreateSheet_(ss, SHEETS.RECURRING);
  if (recurring.getLastRow() < 1) {
    recurring.getRange(1, 1, 1, 7).setValues([
      ['id', 'account_id', 'type', 'category', 'description', 'amount', 'active']
    ]);
  }

  var budgets = getOrCreateSheet_(ss, SHEETS.BUDGETS);
  if (budgets.getLastRow() < 1) {
    budgets.getRange(1, 1, 1, 2).setValues([['category', 'monthly_limit']]);
  }
}

// Manual-only reset: wipes Transactions/Transfers and zeroes every opening_balance.
// Not exposed via doGet/doPost — run it directly from the Apps Script editor when you want
// to blank the ledger back to the initial seed state without touching account/category rows.
function clearTestData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  clearDataRows_(ss.getSheetByName(SHEETS.TRANSACTIONS));
  clearDataRows_(ss.getSheetByName(SHEETS.TRANSFERS));

  var accounts = ss.getSheetByName(SHEETS.ACCOUNTS);
  var lastRow = accounts.getLastRow();
  if (lastRow > 1) {
    var zeros = accounts.getRange(2, 1, lastRow - 1, 1).getValues().map(function () { return [0]; });
    accounts.getRange(2, 4, lastRow - 1, 1).setValues(zeros);
  }
}

function clearDataRows_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
}

function getOrCreateSheet_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function getSheet_(name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('Missing sheet: ' + name + '. Run setupSheets() first.');
  return sheet;
}

function nextId_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 1;
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues()
    .map(function (r) { return Number(r[0]); })
    .filter(function (v) { return !isNaN(v); });
  return ids.length ? Math.max.apply(null, ids) + 1 : 1;
}

function findRowIndexById_(sheet, id) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function requireFields_(params, fields) {
  fields.forEach(function (f) {
    if (params[f] === undefined || params[f] === null || params[f] === '') {
      throw new Error('Missing required field: ' + f);
    }
  });
}

function getAccountBalance_(accountId) {
  var accountsSheet = getSheet_(SHEETS.ACCOUNTS);
  var rowIndex = findRowIndexById_(accountsSheet, accountId);
  if (rowIndex === -1) throw new Error('Account not found: ' + accountId);
  var openingBalance = Number(accountsSheet.getRange(rowIndex, 4).getValue());

  var txSheet = getSheet_(SHEETS.TRANSACTIONS);
  var txLastRow = txSheet.getLastRow();
  var incomeSum = 0, expenseSum = 0;
  if (txLastRow > 1) {
    txSheet.getRange(2, 1, txLastRow - 1, 7).getValues().forEach(function (row) {
      if (String(row[1]) === String(accountId)) {
        if (row[3] === 'Income') incomeSum += Number(row[6]);
        else if (row[3] === 'Expense') expenseSum += Number(row[6]);
      }
    });
  }

  var trSheet = getSheet_(SHEETS.TRANSFERS);
  var trLastRow = trSheet.getLastRow();
  var destSum = 0, sourceSum = 0;
  if (trLastRow > 1) {
    trSheet.getRange(2, 1, trLastRow - 1, 8).getValues().forEach(function (row) {
      if (String(row[3]) === String(accountId)) destSum += Number(row[5]);
      if (String(row[2]) === String(accountId)) sourceSum += Number(row[4]);
    });
  }

  return openingBalance + incomeSum - expenseSum + destSum - sourceSum;
}

function addTransaction(p) {
  requireFields_(p, ['account_id', 'date', 'type', 'category', 'amount']);
  if (['Income', 'Expense'].indexOf(p.type) === -1) throw new Error('type must be Income or Expense');
  var amount = Number(p.amount);
  if (!(amount > 0)) throw new Error('amount must be a positive number');
  var accountsSheet = getSheet_(SHEETS.ACCOUNTS);
  if (findRowIndexById_(accountsSheet, p.account_id) === -1) throw new Error('Unknown account_id: ' + p.account_id);

  var id;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_(SHEETS.TRANSACTIONS);
    id = nextId_(sheet);
    sheet.appendRow([id, Number(p.account_id), new Date(p.date), p.type, p.category, p.description || '', amount]);
  } finally {
    lock.releaseLock();
  }

  return { id: id, account_id: Number(p.account_id), balance: getAccountBalance_(p.account_id) };
}

function addTransfer(p) {
  requireFields_(p, ['date', 'source_account_id', 'dest_account_id', 'source_amount', 'dest_amount']);
  var sourceAmount = Number(p.source_amount);
  var destAmount = Number(p.dest_amount);
  if (!(sourceAmount > 0) || !(destAmount > 0)) throw new Error('source_amount and dest_amount must be positive numbers');
  var accountsSheet = getSheet_(SHEETS.ACCOUNTS);
  if (findRowIndexById_(accountsSheet, p.source_account_id) === -1) throw new Error('Unknown source_account_id: ' + p.source_account_id);
  if (findRowIndexById_(accountsSheet, p.dest_account_id) === -1) throw new Error('Unknown dest_account_id: ' + p.dest_account_id);
  var fxRate = destAmount / sourceAmount;

  var id;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_(SHEETS.TRANSFERS);
    id = nextId_(sheet);
    sheet.appendRow([id, new Date(p.date), Number(p.source_account_id), Number(p.dest_account_id), sourceAmount, destAmount, fxRate, p.description || '']);
  } finally {
    lock.releaseLock();
  }

  return {
    id: id,
    fx_rate: fxRate,
    source_account: { account_id: Number(p.source_account_id), balance: getAccountBalance_(p.source_account_id) },
    dest_account: { account_id: Number(p.dest_account_id), balance: getAccountBalance_(p.dest_account_id) }
  };
}

function setOpeningBalance(p) {
  requireFields_(p, ['account_id', 'opening_balance']);
  var sheet = getSheet_(SHEETS.ACCOUNTS);
  var rowIndex = findRowIndexById_(sheet, p.account_id);
  if (rowIndex === -1) throw new Error('Unknown account_id: ' + p.account_id);

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    sheet.getRange(rowIndex, 4).setValue(Number(p.opening_balance));
  } finally {
    lock.releaseLock();
  }

  return { account_id: Number(p.account_id), opening_balance: Number(p.opening_balance), balance: getAccountBalance_(p.account_id) };
}

function getFxRate(p) {
  requireFields_(p, ['from_currency', 'to_currency']);
  var from = String(p.from_currency).toUpperCase();
  var to = String(p.to_currency).toUpperCase();
  if (from === to) return { from_currency: from, to_currency: to, rate: 1, date: new Date().toISOString().slice(0, 10) };

  var url = 'https://api.frankfurter.app/latest?amount=1&from=' + encodeURIComponent(from) + '&to=' + encodeURIComponent(to);
  var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) throw new Error('FX API error: ' + response.getContentText());
  var data = JSON.parse(response.getContentText());
  var rate = data.rates && data.rates[to];
  if (!rate) throw new Error('No rate returned for ' + from + ' -> ' + to);

  return { from_currency: from, to_currency: to, rate: rate, date: data.date };
}

function getAccounts() {
  var sheet = getSheet_(SHEETS.ACCOUNTS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 4).getValues().map(function (row) {
    return { id: row[0], name: row[1], currency: row[2], opening_balance: row[3] };
  });
}

function getCategories() {
  var sheet = getSheet_(SHEETS.CATEGORIES);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 2).getValues().map(function (row) {
    return { type: row[0], name: row[1] };
  });
}

function getTransactions(p) {
  var sheet = getSheet_(SHEETS.TRANSACTIONS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var filterId = (p && p.account_id !== undefined && p.account_id !== null && p.account_id !== '')
    ? String(p.account_id) : null;
  return sheet.getRange(2, 1, lastRow - 1, 7).getValues()
    .filter(function (row) { return !filterId || String(row[1]) === filterId; })
    .map(function (row) {
      return {
        id: row[0],
        account_id: row[1],
        date: formatDate_(row[2]),
        type: row[3],
        category: row[4],
        description: row[5],
        amount: row[6]
      };
    });
}

function getTransfers() {
  var sheet = getSheet_(SHEETS.TRANSFERS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 8).getValues().map(function (row) {
    return {
      id: row[0],
      date: formatDate_(row[1]),
      source_account_id: row[2],
      dest_account_id: row[3],
      source_amount: row[4],
      dest_amount: row[5],
      fx_rate: row[6],
      description: row[7]
    };
  });
}

function getRecurring() {
  var sheet = getSheet_(SHEETS.RECURRING);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 7).getValues().map(function (row) {
    return {
      id: row[0],
      account_id: row[1],
      type: row[2],
      category: row[3],
      description: row[4],
      amount: row[5],
      active: row[6] === true || String(row[6]).toUpperCase() === 'TRUE'
    };
  });
}

function addRecurring(p) {
  requireFields_(p, ['account_id', 'type', 'category', 'amount']);
  if (['Income', 'Expense'].indexOf(p.type) === -1) throw new Error('type must be Income or Expense');
  var amount = Number(p.amount);
  if (!(amount > 0)) throw new Error('amount must be a positive number');
  var accountsSheet = getSheet_(SHEETS.ACCOUNTS);
  if (findRowIndexById_(accountsSheet, p.account_id) === -1) throw new Error('Unknown account_id: ' + p.account_id);

  var id;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_(SHEETS.RECURRING);
    id = nextId_(sheet);
    sheet.appendRow([id, Number(p.account_id), p.type, p.category, p.description || '', amount, true]);
  } finally {
    lock.releaseLock();
  }

  return { id: id };
}

function updateRecurring(p) {
  requireFields_(p, ['id', 'account_id', 'type', 'category', 'amount', 'active']);
  if (['Income', 'Expense'].indexOf(p.type) === -1) throw new Error('type must be Income or Expense');
  var amount = Number(p.amount);
  if (!(amount > 0)) throw new Error('amount must be a positive number');
  var accountsSheet = getSheet_(SHEETS.ACCOUNTS);
  if (findRowIndexById_(accountsSheet, p.account_id) === -1) throw new Error('Unknown account_id: ' + p.account_id);
  var active = (p.active === true || String(p.active).toUpperCase() === 'TRUE');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_(SHEETS.RECURRING);
    var rowIndex = findRowIndexById_(sheet, p.id);
    if (rowIndex === -1) throw new Error('Unknown recurring id: ' + p.id);
    sheet.getRange(rowIndex, 1, 1, 7).setValues([[
      Number(p.id), Number(p.account_id), p.type, p.category, p.description || '', amount, active
    ]]);
  } finally {
    lock.releaseLock();
  }

  return { id: Number(p.id) };
}

function getBudgets() {
  var sheet = getSheet_(SHEETS.BUDGETS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 2).getValues()
    .filter(function (row) { return row[0] !== ''; })
    .map(function (row) { return { category: row[0], monthly_limit: row[1] }; });
}

function setBudget(p) {
  requireFields_(p, ['category', 'monthly_limit']);
  var limit = Number(p.monthly_limit);
  if (!(limit >= 0)) throw new Error('monthly_limit must be a non-negative number');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_(SHEETS.BUDGETS);
    var rowIndex = findRowIndexById_(sheet, p.category);
    if (rowIndex === -1) {
      sheet.appendRow([p.category, limit]);
    } else {
      sheet.getRange(rowIndex, 2).setValue(limit);
    }
  } finally {
    lock.releaseLock();
  }

  return { category: p.category, monthly_limit: limit };
}

function formatDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return value;
}

var ACTIONS_ = {
  addTransaction: addTransaction,
  addTransfer: addTransfer,
  setOpeningBalance: setOpeningBalance,
  getFxRate: getFxRate,
  getAccounts: getAccounts,
  getCategories: getCategories,
  getTransactions: getTransactions,
  getTransfers: getTransfers,
  getRecurring: getRecurring,
  addRecurring: addRecurring,
  updateRecurring: updateRecurring,
  getBudgets: getBudgets,
  setBudget: setBudget
};

function doGet(e) { return handleRequest_(e); }
function doPost(e) { return handleRequest_(e); }

function handleRequest_(e) {
  try {
    var params;
    if (e.postData && e.postData.contents) {
      params = JSON.parse(e.postData.contents);
    } else {
      params = (e && e.parameter) || {};
    }

    checkAuth_(params);

    var action = params.action;
    if (!action || !ACTIONS_.hasOwnProperty(action)) throw new Error('Unknown or missing action: ' + action);

    var result = ACTIONS_[action](params);
    return jsonOutput_({ success: true, data: result });
  } catch (err) {
    return jsonOutput_({ success: false, error: err.message });
  }
}

// Optional shared-secret gate: set an API_TOKEN script property (Project Settings > Script
// Properties) before deploying with "Anyone" access, since that setting otherwise makes this
// a public write API onto your financial data. Leave the property unset to skip auth entirely.
function checkAuth_(params) {
  var required = PropertiesService.getScriptProperties().getProperty('API_TOKEN');
  if (required && params.token !== required) throw new Error('Unauthorized: invalid or missing token');
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
