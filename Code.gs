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
  RECURRING_TRANSFERS: 'RecurringTransfers',
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

// Phase 4b: adds day_of_month/last_generated_month to Recurring and recurring_id/confirmed
// to Transactions, without touching any existing column or row data. Safe to re-run.
function setupPhase4bSheets() {
  var recurring = getSheet_(SHEETS.RECURRING);
  if (recurring.getRange(1, 8).getValue() !== 'day_of_month') {
    recurring.getRange(1, 8).setValue('day_of_month');
    var recLastRow = recurring.getLastRow();
    if (recLastRow > 1) {
      var dayDefaults = [];
      for (var i = 0; i < recLastRow - 1; i++) dayDefaults.push([1]);
      recurring.getRange(2, 8, recLastRow - 1, 1).setValues(dayDefaults);
    }
  }
  if (recurring.getRange(1, 9).getValue() !== 'last_generated_month') {
    recurring.getRange(1, 9).setValue('last_generated_month');
  }

  var tx = getSheet_(SHEETS.TRANSACTIONS);
  if (tx.getRange(1, 8).getValue() !== 'recurring_id') {
    tx.getRange(1, 8).setValue('recurring_id');
  }
  if (tx.getRange(1, 9).getValue() !== 'confirmed') {
    tx.getRange(1, 9).setValue('confirmed');
    var txLastRow = tx.getLastRow();
    if (txLastRow > 1) {
      var confirmedDefaults = [];
      for (var j = 0; j < txLastRow - 1; j++) confirmedDefaults.push([true]);
      tx.getRange(2, 9, txLastRow - 1, 1).setValues(confirmedDefaults);
    }
  }
}

// Budgets.monthly_limit (lei) becomes Budgets.percent (0-100, % of income). Old monetary
// values aren't valid percentages, so this clears existing budget rows' values once during
// the rename — re-enter targets as percentages from Setări afterwards. Safe to re-run: it
// no-ops once the header already reads 'percent'.
function setupBudgetPercentMigration() {
  var sheet = getSheet_(SHEETS.BUDGETS);
  if (sheet.getRange(1, 2).getValue() === 'percent') return;
  sheet.getRange(1, 2).setValue('percent');
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var blanks = [];
    for (var i = 0; i < lastRow - 1; i++) blanks.push(['']);
    sheet.getRange(2, 2, lastRow - 1, 1).setValues(blanks);
  }
}

// Adds restricted_account_id to Categories (empty by default), then sets it to 3 (the
// "Economii pe termen lung / Neprevăzute" account) for "Cheltuieli pe termen lung/Neprevăzute"
// specifically — a one-time manual restriction so that category can only be logged against
// that account. Every other category keeps the field empty. Safe to re-run.
function setupCategoryRestrictionMigration() {
  var sheet = getSheet_(SHEETS.CATEGORIES);
  if (sheet.getRange(1, 3).getValue() !== 'restricted_account_id') {
    sheet.getRange(1, 3).setValue('restricted_account_id');
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var rows = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][1] === 'Cheltuieli pe termen lung/Neprevăzute') {
      sheet.getRange(2 + i, 3).setValue(3);
      break;
    }
  }
}

// Adds frequency/month_of_year to Recurring and renames column 9 from
// last_generated_month to last_generated_period (same column, now dual-purpose: "YYYY-MM"
// for Lunar items, "YYYY" for Anual items — unaffected either way, since generation just
// compares it against whichever key it computes for that item's own frequency). Existing
// rows default to frequency = "Lunar" (the only kind that existed before this). Safe to
// re-run.
function setupRecurringFrequencyMigration() {
  var sheet = getSheet_(SHEETS.RECURRING);
  if (sheet.getRange(1, 9).getValue() === 'last_generated_month') {
    sheet.getRange(1, 9).setValue('last_generated_period');
  }
  if (sheet.getRange(1, 10).getValue() !== 'frequency') {
    sheet.getRange(1, 10).setValue('frequency');
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var freqDefaults = [];
      for (var i = 0; i < lastRow - 1; i++) freqDefaults.push(['Lunar']);
      sheet.getRange(2, 10, lastRow - 1, 1).setValues(freqDefaults);
    }
  }
  if (sheet.getRange(1, 11).getValue() !== 'month_of_year') {
    sheet.getRange(1, 11).setValue('month_of_year');
  }
}

// Creates the RecurringTransfers tab (mirrors Recurring's shape, minus the
// type/category/description/account_id-single fields that don't apply to a transfer).
// Safe to re-run — no-ops once the sheet already has a header row.
function setupRecurringTransfersSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet_(ss, SHEETS.RECURRING_TRANSFERS);
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, 10).setValues([[
      'id', 'source_account_id', 'dest_account_id', 'source_amount', 'dest_amount',
      'frequency', 'day_of_month', 'month_of_year', 'active', 'last_generated_period'
    ]]);
  }
}

// Adds recurring_transfer_id/confirmed to Transfers, mirroring the equivalent Transactions
// migration exactly (columns land at 9/10 here since Transfers already has 8 base columns
// vs. Transactions' 7). Existing rows default to confirmed = TRUE, since they were all
// entered manually before this feature existed. Safe to re-run.
function setupTransferConfirmationMigration() {
  var sheet = getSheet_(SHEETS.TRANSFERS);
  if (sheet.getRange(1, 9).getValue() !== 'recurring_transfer_id') {
    sheet.getRange(1, 9).setValue('recurring_transfer_id');
  }
  if (sheet.getRange(1, 10).getValue() !== 'confirmed') {
    sheet.getRange(1, 10).setValue('confirmed');
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var confirmedDefaults = [];
      for (var i = 0; i < lastRow - 1; i++) confirmedDefaults.push([true]);
      sheet.getRange(2, 10, lastRow - 1, 1).setValues(confirmedDefaults);
    }
  }
}

// RecurringTransfers.source_amount/dest_amount (both "fixed" at creation) become
// fixed_side/fixed_amount — only one side is ever actually known exactly, the other gets
// estimated via FX at generation time. Existing rows default to fixed_side = "source",
// carrying over their old source_amount value as fixed_amount (their dest_amount is
// discarded — it was never the exact one for a cross-currency transfer anyway). Safe to
// re-run: no-ops once column 4 already reads "fixed_side".
function setupRecurringTransferFixedSideMigration() {
  var sheet = getSheet_(SHEETS.RECURRING_TRANSFERS);
  if (sheet.getRange(1, 4).getValue() === 'fixed_side') return;

  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var oldSourceAmounts = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
    var fixedSideCol = [], fixedAmountCol = [];
    for (var i = 0; i < oldSourceAmounts.length; i++) {
      fixedSideCol.push(['source']);
      fixedAmountCol.push([oldSourceAmounts[i][0]]);
    }
    sheet.getRange(2, 4, oldSourceAmounts.length, 1).setValues(fixedSideCol);
    sheet.getRange(2, 5, oldSourceAmounts.length, 1).setValues(fixedAmountCol);
  }
  sheet.getRange(1, 4).setValue('fixed_side');
  sheet.getRange(1, 5).setValue('fixed_amount');
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
    sheet.appendRow([id, Number(p.account_id), new Date(p.date), p.type, p.category, p.description || '', amount, '', true]);
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
    sheet.appendRow([
      id, new Date(p.date), Number(p.source_account_id), Number(p.dest_account_id), sourceAmount,
      destAmount, fxRate, p.description || '', '', true
    ]);
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
  return sheet.getRange(2, 1, lastRow - 1, 3).getValues().map(function (row) {
    return { type: row[0], name: row[1], restricted_account_id: row[2] || null };
  });
}

function getTransactions(p) {
  var sheet = getSheet_(SHEETS.TRANSACTIONS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var filterId = (p && p.account_id !== undefined && p.account_id !== null && p.account_id !== '')
    ? String(p.account_id) : null;
  var dateFrom = (p && p.date_from) ? String(p.date_from) : null;
  var dateTo = (p && p.date_to) ? String(p.date_to) : null;
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
    })
    .filter(function (t) {
      if (dateFrom && t.date < dateFrom) return false;
      if (dateTo && t.date > dateTo) return false;
      return true;
    });
}

function getTransfers(p) {
  var sheet = getSheet_(SHEETS.TRANSFERS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var dateFrom = (p && p.date_from) ? String(p.date_from) : null;
  var dateTo = (p && p.date_to) ? String(p.date_to) : null;
  return sheet.getRange(2, 1, lastRow - 1, 8).getValues()
    .map(function (row) {
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
    })
    .filter(function (t) {
      if (dateFrom && t.date < dateFrom) return false;
      if (dateTo && t.date > dateTo) return false;
      return true;
    });
}

function getRecurring() {
  var sheet = getSheet_(SHEETS.RECURRING);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 11).getValues().map(function (row) {
    return {
      id: row[0],
      account_id: row[1],
      type: row[2],
      category: row[3],
      description: row[4],
      amount: row[5],
      active: row[6] === true || String(row[6]).toUpperCase() === 'TRUE',
      day_of_month: row[7],
      last_generated_period: row[8],
      frequency: row[9] || 'Lunar',
      month_of_year: row[10] || null
    };
  });
}

function validateDayOfMonth_(value) {
  var dayOfMonth = Number(value);
  if (!(dayOfMonth >= 1 && dayOfMonth <= 28 && Math.floor(dayOfMonth) === dayOfMonth)) {
    throw new Error('day_of_month must be an integer between 1 and 28');
  }
  return dayOfMonth;
}

function validateFrequency_(value) {
  if (value !== 'Lunar' && value !== 'Anual') throw new Error('frequency must be "Lunar" or "Anual"');
  return value;
}

// month_of_year only means something for Anual items — Lunar ones store it blank.
function validateMonthOfYear_(value, frequency) {
  if (frequency !== 'Anual') return '';
  var month = Number(value);
  if (!(month >= 1 && month <= 12 && Math.floor(month) === month)) {
    throw new Error('month_of_year must be an integer between 1 and 12 for Anual frequency');
  }
  return month;
}

// Defaults to "source" when omitted — the frontend doesn't even show a choice when the two
// accounts share a currency, since which side is "fixed" is moot there.
function validateFixedSide_(value) {
  var side = value || 'source';
  if (side !== 'source' && side !== 'dest') throw new Error('fixed_side must be "source" or "dest"');
  return side;
}

function addRecurring(p) {
  requireFields_(p, ['account_id', 'type', 'category', 'amount', 'day_of_month', 'frequency']);
  if (['Income', 'Expense'].indexOf(p.type) === -1) throw new Error('type must be Income or Expense');
  var amount = Number(p.amount);
  if (!(amount > 0)) throw new Error('amount must be a positive number');
  var dayOfMonth = validateDayOfMonth_(p.day_of_month);
  var frequency = validateFrequency_(p.frequency);
  var monthOfYear = validateMonthOfYear_(p.month_of_year, frequency);
  var accountsSheet = getSheet_(SHEETS.ACCOUNTS);
  if (findRowIndexById_(accountsSheet, p.account_id) === -1) throw new Error('Unknown account_id: ' + p.account_id);

  var id;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_(SHEETS.RECURRING);
    id = nextId_(sheet);
    sheet.appendRow([
      id, Number(p.account_id), p.type, p.category, p.description || '', amount, true, dayOfMonth,
      '', frequency, monthOfYear
    ]);
  } finally {
    lock.releaseLock();
  }

  return { id: id };
}

// Deliberately leaves column 9 (last_generated_period) untouched — that's internal
// bookkeeping for generateRecurringTransactions_, not something an edit should reset. A
// stale "YYYY-MM" left behind by a frequency change to/from Anual just never matches the
// new check's key format, which correctly behaves as "not generated yet under this scheme".
function updateRecurring(p) {
  requireFields_(p, ['id', 'account_id', 'type', 'category', 'amount', 'active', 'day_of_month', 'frequency']);
  if (['Income', 'Expense'].indexOf(p.type) === -1) throw new Error('type must be Income or Expense');
  var amount = Number(p.amount);
  if (!(amount > 0)) throw new Error('amount must be a positive number');
  var dayOfMonth = validateDayOfMonth_(p.day_of_month);
  var frequency = validateFrequency_(p.frequency);
  var monthOfYear = validateMonthOfYear_(p.month_of_year, frequency);
  var accountsSheet = getSheet_(SHEETS.ACCOUNTS);
  if (findRowIndexById_(accountsSheet, p.account_id) === -1) throw new Error('Unknown account_id: ' + p.account_id);
  var active = (p.active === true || String(p.active).toUpperCase() === 'TRUE');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_(SHEETS.RECURRING);
    var rowIndex = findRowIndexById_(sheet, p.id);
    if (rowIndex === -1) throw new Error('Unknown recurring id: ' + p.id);
    sheet.getRange(rowIndex, 1, 1, 8).setValues([[
      Number(p.id), Number(p.account_id), p.type, p.category, p.description || '', amount, active, dayOfMonth
    ]]);
    sheet.getRange(rowIndex, 10, 1, 2).setValues([[frequency, monthOfYear]]);
  } finally {
    lock.releaseLock();
  }

  return { id: Number(p.id) };
}

function getRecurringTransfers() {
  var sheet = getSheet_(SHEETS.RECURRING_TRANSFERS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 10).getValues().map(function (row) {
    return {
      id: row[0],
      source_account_id: row[1],
      dest_account_id: row[2],
      fixed_side: row[3] || 'source',
      fixed_amount: row[4],
      frequency: row[5] || 'Lunar',
      day_of_month: row[6],
      month_of_year: row[7] || null,
      active: row[8] === true || String(row[8]).toUpperCase() === 'TRUE',
      last_generated_period: row[9]
    };
  });
}

function addRecurringTransfer(p) {
  requireFields_(p, ['source_account_id', 'dest_account_id', 'fixed_amount', 'day_of_month', 'frequency']);
  var fixedAmount = Number(p.fixed_amount);
  if (!(fixedAmount > 0)) throw new Error('fixed_amount must be a positive number');
  var fixedSide = validateFixedSide_(p.fixed_side);
  var dayOfMonth = validateDayOfMonth_(p.day_of_month);
  var frequency = validateFrequency_(p.frequency);
  var monthOfYear = validateMonthOfYear_(p.month_of_year, frequency);
  var accountsSheet = getSheet_(SHEETS.ACCOUNTS);
  if (findRowIndexById_(accountsSheet, p.source_account_id) === -1) throw new Error('Unknown source_account_id: ' + p.source_account_id);
  if (findRowIndexById_(accountsSheet, p.dest_account_id) === -1) throw new Error('Unknown dest_account_id: ' + p.dest_account_id);

  var id;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_(SHEETS.RECURRING_TRANSFERS);
    id = nextId_(sheet);
    sheet.appendRow([
      id, Number(p.source_account_id), Number(p.dest_account_id), fixedSide, fixedAmount,
      frequency, dayOfMonth, monthOfYear, true, ''
    ]);
  } finally {
    lock.releaseLock();
  }

  return { id: id };
}

// Deliberately leaves column 10 (last_generated_period) untouched, same reasoning as
// updateRecurring: internal bookkeeping for generateRecurringTransfers_, not something an
// edit should reset.
function updateRecurringTransfer(p) {
  requireFields_(p, ['id', 'source_account_id', 'dest_account_id', 'fixed_amount', 'day_of_month', 'frequency', 'active']);
  var fixedAmount = Number(p.fixed_amount);
  if (!(fixedAmount > 0)) throw new Error('fixed_amount must be a positive number');
  var fixedSide = validateFixedSide_(p.fixed_side);
  var dayOfMonth = validateDayOfMonth_(p.day_of_month);
  var frequency = validateFrequency_(p.frequency);
  var monthOfYear = validateMonthOfYear_(p.month_of_year, frequency);
  var accountsSheet = getSheet_(SHEETS.ACCOUNTS);
  if (findRowIndexById_(accountsSheet, p.source_account_id) === -1) throw new Error('Unknown source_account_id: ' + p.source_account_id);
  if (findRowIndexById_(accountsSheet, p.dest_account_id) === -1) throw new Error('Unknown dest_account_id: ' + p.dest_account_id);
  var active = (p.active === true || String(p.active).toUpperCase() === 'TRUE');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_(SHEETS.RECURRING_TRANSFERS);
    var rowIndex = findRowIndexById_(sheet, p.id);
    if (rowIndex === -1) throw new Error('Unknown recurring transfer id: ' + p.id);
    sheet.getRange(rowIndex, 1, 1, 9).setValues([[
      Number(p.id), Number(p.source_account_id), Number(p.dest_account_id), fixedSide, fixedAmount,
      frequency, dayOfMonth, monthOfYear, active
    ]]);
  } finally {
    lock.releaseLock();
  }

  return { id: Number(p.id) };
}

var BUDGET_EXCLUDED_CATEGORIES_ = ['Corecții', 'Datorii/Împrumut'];

function getBudgets() {
  var sheet = getSheet_(SHEETS.BUDGETS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 2).getValues()
    .filter(function (row) { return row[0] !== ''; })
    .map(function (row) { return { category: row[0], percent: row[1] }; });
}

function setBudget(p) {
  requireFields_(p, ['category', 'percent']);
  if (BUDGET_EXCLUDED_CATEGORIES_.indexOf(p.category) !== -1) {
    throw new Error('Category cannot have a budget target: ' + p.category);
  }
  var percent = Number(p.percent);
  if (!(percent >= 0 && percent <= 100)) throw new Error('percent must be a number between 0 and 100');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_(SHEETS.BUDGETS);
    var rowIndex = findRowIndexById_(sheet, p.category);
    if (rowIndex === -1) {
      sheet.appendRow([p.category, percent]);
    } else {
      sheet.getRange(rowIndex, 2).setValue(percent);
    }
  } finally {
    lock.releaseLock();
  }

  return { category: p.category, percent: percent };
}

function getPendingConfirmations() {
  var sheet = getSheet_(SHEETS.TRANSACTIONS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 9).getValues()
    .filter(function (row) { return row[8] === false || String(row[8]).toUpperCase() === 'FALSE'; })
    .map(function (row) {
      return {
        id: row[0],
        account_id: row[1],
        date: formatDate_(row[2]),
        type: row[3],
        category: row[4],
        description: row[5],
        amount: row[6],
        recurring_id: row[7]
      };
    });
}

function confirmTransaction(p) {
  requireFields_(p, ['id', 'amount']);
  var amount = Number(p.amount);
  if (!(amount > 0)) throw new Error('amount must be a positive number');

  var accountId;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_(SHEETS.TRANSACTIONS);
    var rowIndex = findRowIndexById_(sheet, p.id);
    if (rowIndex === -1) throw new Error('Unknown transaction id: ' + p.id);
    accountId = sheet.getRange(rowIndex, 2).getValue();
    sheet.getRange(rowIndex, 7).setValue(amount);
    sheet.getRange(rowIndex, 9).setValue(true);
  } finally {
    lock.releaseLock();
  }

  return { id: Number(p.id), account_id: Number(accountId), balance: getAccountBalance_(accountId) };
}

function getPendingTransferConfirmations() {
  var sheet = getSheet_(SHEETS.TRANSFERS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 10).getValues()
    .filter(function (row) { return row[9] === false || String(row[9]).toUpperCase() === 'FALSE'; })
    .map(function (row) {
      return {
        id: row[0],
        date: formatDate_(row[1]),
        source_account_id: row[2],
        dest_account_id: row[3],
        source_amount: row[4],
        dest_amount: row[5],
        fx_rate: row[6],
        description: row[7],
        recurring_transfer_id: row[8]
      };
    });
}

function confirmTransfer(p) {
  requireFields_(p, ['id', 'source_amount', 'dest_amount']);
  var sourceAmount = Number(p.source_amount);
  var destAmount = Number(p.dest_amount);
  if (!(sourceAmount > 0) || !(destAmount > 0)) throw new Error('source_amount and dest_amount must be positive numbers');
  var fxRate = destAmount / sourceAmount;

  var sourceAccountId, destAccountId;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_(SHEETS.TRANSFERS);
    var rowIndex = findRowIndexById_(sheet, p.id);
    if (rowIndex === -1) throw new Error('Unknown transfer id: ' + p.id);
    sourceAccountId = sheet.getRange(rowIndex, 3).getValue();
    destAccountId = sheet.getRange(rowIndex, 4).getValue();
    sheet.getRange(rowIndex, 5, 1, 3).setValues([[sourceAmount, destAmount, fxRate]]);
    sheet.getRange(rowIndex, 10).setValue(true);
  } finally {
    lock.releaseLock();
  }

  return {
    id: Number(p.id),
    fx_rate: fxRate,
    source_account: { account_id: Number(sourceAccountId), balance: getAccountBalance_(sourceAccountId) },
    dest_account: { account_id: Number(destAccountId), balance: getAccountBalance_(destAccountId) }
  };
}

// Deliberately preserves column 8 (recurring_id) — editing a transaction shouldn't sever its
// link back to the Recurring row that generated it. Editing always implies confirmed = TRUE.
function updateTransaction(p) {
  requireFields_(p, ['id', 'account_id', 'date', 'type', 'category', 'amount']);
  if (['Income', 'Expense'].indexOf(p.type) === -1) throw new Error('type must be Income or Expense');
  var amount = Number(p.amount);
  if (!(amount > 0)) throw new Error('amount must be a positive number');
  var accountsSheet = getSheet_(SHEETS.ACCOUNTS);
  if (findRowIndexById_(accountsSheet, p.account_id) === -1) throw new Error('Unknown account_id: ' + p.account_id);

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_(SHEETS.TRANSACTIONS);
    var rowIndex = findRowIndexById_(sheet, p.id);
    if (rowIndex === -1) throw new Error('Unknown transaction id: ' + p.id);
    var recurringId = sheet.getRange(rowIndex, 8).getValue();
    sheet.getRange(rowIndex, 1, 1, 9).setValues([[
      Number(p.id), Number(p.account_id), new Date(p.date), p.type, p.category, p.description || '',
      amount, recurringId, true
    ]]);
  } finally {
    lock.releaseLock();
  }

  return { id: Number(p.id), account_id: Number(p.account_id), balance: getAccountBalance_(p.account_id) };
}

function deleteTransaction(p) {
  requireFields_(p, ['id']);
  var accountId;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_(SHEETS.TRANSACTIONS);
    var rowIndex = findRowIndexById_(sheet, p.id);
    if (rowIndex === -1) throw new Error('Unknown transaction id: ' + p.id);
    accountId = sheet.getRange(rowIndex, 2).getValue();
    sheet.deleteRow(rowIndex);
  } finally {
    lock.releaseLock();
  }

  return { id: Number(p.id), account_id: Number(accountId), balance: getAccountBalance_(accountId) };
}

function updateTransfer(p) {
  requireFields_(p, ['id', 'date', 'source_account_id', 'dest_account_id', 'source_amount', 'dest_amount']);
  var sourceAmount = Number(p.source_amount);
  var destAmount = Number(p.dest_amount);
  if (!(sourceAmount > 0) || !(destAmount > 0)) throw new Error('source_amount and dest_amount must be positive numbers');
  var accountsSheet = getSheet_(SHEETS.ACCOUNTS);
  if (findRowIndexById_(accountsSheet, p.source_account_id) === -1) throw new Error('Unknown source_account_id: ' + p.source_account_id);
  if (findRowIndexById_(accountsSheet, p.dest_account_id) === -1) throw new Error('Unknown dest_account_id: ' + p.dest_account_id);
  var fxRate = destAmount / sourceAmount;

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_(SHEETS.TRANSFERS);
    var rowIndex = findRowIndexById_(sheet, p.id);
    if (rowIndex === -1) throw new Error('Unknown transfer id: ' + p.id);
    sheet.getRange(rowIndex, 1, 1, 8).setValues([[
      Number(p.id), new Date(p.date), Number(p.source_account_id), Number(p.dest_account_id),
      sourceAmount, destAmount, fxRate, p.description || ''
    ]]);
    sheet.getRange(rowIndex, 10).setValue(true);
  } finally {
    lock.releaseLock();
  }

  return {
    id: Number(p.id),
    fx_rate: fxRate,
    source_account: { account_id: Number(p.source_account_id), balance: getAccountBalance_(p.source_account_id) },
    dest_account: { account_id: Number(p.dest_account_id), balance: getAccountBalance_(p.dest_account_id) }
  };
}

function deleteTransfer(p) {
  requireFields_(p, ['id']);
  var sourceAccountId, destAccountId;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_(SHEETS.TRANSFERS);
    var rowIndex = findRowIndexById_(sheet, p.id);
    if (rowIndex === -1) throw new Error('Unknown transfer id: ' + p.id);
    var row = sheet.getRange(rowIndex, 1, 1, 8).getValues()[0];
    sourceAccountId = row[2];
    destAccountId = row[3];
    sheet.deleteRow(rowIndex);
  } finally {
    lock.releaseLock();
  }

  return {
    id: Number(p.id),
    source_account: { account_id: Number(sourceAccountId), balance: getAccountBalance_(sourceAccountId) },
    dest_account: { account_id: Number(destAccountId), balance: getAccountBalance_(destAccountId) }
  };
}

// Meant to run on a daily time-driven trigger (Apps Script editor > Triggers), not via the
// web app — it is intentionally not registered in ACTIONS_. For each active Recurring row
// whose day_of_month has been reached and hasn't already been generated this month, appends
// an unconfirmed Transaction and stamps last_generated_month — that stamp is what makes
// re-running the same day (or later the same month) a no-op instead of a duplicate.
function generateRecurringTransactions() {
  var today = new Date();
  var currentDay = today.getDate();
  var currentMonthNum = today.getMonth() + 1;
  var currentMonthKey = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM');
  var currentYearKey = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy');

  var recurringSheet = getSheet_(SHEETS.RECURRING);
  var lastRow = recurringSheet.getLastRow();
  if (lastRow < 2) return { generated: 0 };

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  var generated = 0;
  try {
    var rows = recurringSheet.getRange(2, 1, lastRow - 1, 11).getValues();
    var txSheet = getSheet_(SHEETS.TRANSACTIONS);

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      var active = row[6] === true || String(row[6]).toUpperCase() === 'TRUE';
      if (!active) continue;

      var dayOfMonth = Number(row[7]);
      var lastGeneratedPeriod = row[8];
      var frequency = row[9] || 'Lunar';
      var monthOfYear = row[10];

      var periodKey;
      if (frequency === 'Anual') {
        if (currentMonthNum !== Number(monthOfYear)) continue;
        if (currentDay < dayOfMonth) continue;
        if (lastGeneratedPeriod === currentYearKey) continue;
        periodKey = currentYearKey;
      } else {
        if (currentDay < dayOfMonth) continue;
        if (lastGeneratedPeriod === currentMonthKey) continue;
        periodKey = currentMonthKey;
      }

      var recurringId = row[0], accountId = row[1], type = row[2], category = row[3],
        description = row[4], amount = row[5];
      var newId = nextId_(txSheet);
      txSheet.appendRow([
        newId, Number(accountId), today, type, category, description || '', Number(amount),
        Number(recurringId), false
      ]);
      recurringSheet.getRange(2 + i, 9).setValue(periodKey);
      generated++;
    }
  } finally {
    lock.releaseLock();
  }

  // Recurring transfers ride the same daily trigger as this function — no separate
  // trigger to configure — since Apps Script time-driven triggers call one function each.
  var transfersResult = generateRecurringTransfers();

  return { generated: generated, transfersGenerated: transfersResult.generated };
}

// Mirrors generateRecurringTransactions' Lunar/Anual gating exactly, but appends to
// Transfers instead of Transactions. Called from generateRecurringTransactions() above so
// both run under the one existing daily trigger; also safe to run directly (e.g. to test
// transfer generation in isolation from the Apps Script editor).
function generateRecurringTransfers() {
  var today = new Date();
  var currentDay = today.getDate();
  var currentMonthNum = today.getMonth() + 1;
  var currentMonthKey = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM');
  var currentYearKey = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy');

  var recurringSheet = getSheet_(SHEETS.RECURRING_TRANSFERS);
  var lastRow = recurringSheet.getLastRow();
  if (lastRow < 2) return { generated: 0 };

  var accountsSheet = getSheet_(SHEETS.ACCOUNTS);

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  var generated = 0;
  try {
    var rows = recurringSheet.getRange(2, 1, lastRow - 1, 10).getValues();
    var trSheet = getSheet_(SHEETS.TRANSFERS);

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      var active = row[8] === true || String(row[8]).toUpperCase() === 'TRUE';
      if (!active) continue;

      var dayOfMonth = Number(row[6]);
      var lastGeneratedPeriod = row[9];
      var frequency = row[5] || 'Lunar';
      var monthOfYear = row[7];

      var periodKey;
      if (frequency === 'Anual') {
        if (currentMonthNum !== Number(monthOfYear)) continue;
        if (currentDay < dayOfMonth) continue;
        if (lastGeneratedPeriod === currentYearKey) continue;
        periodKey = currentYearKey;
      } else {
        if (currentDay < dayOfMonth) continue;
        if (lastGeneratedPeriod === currentMonthKey) continue;
        periodKey = currentMonthKey;
      }

      var recurringTransferId = row[0], sourceAccountId = row[1], destAccountId = row[2],
        fixedSide = row[3] || 'source', fixedAmount = Number(row[4]);

      var sourceRowIndex = findRowIndexById_(accountsSheet, sourceAccountId);
      var destRowIndex = findRowIndexById_(accountsSheet, destAccountId);
      var sourceCurrency = String(accountsSheet.getRange(sourceRowIndex, 3).getValue()).toUpperCase();
      var destCurrency = String(accountsSheet.getRange(destRowIndex, 3).getValue()).toUpperCase();

      var sourceAmount, destAmount;
      if (sourceCurrency === destCurrency) {
        sourceAmount = fixedAmount;
        destAmount = fixedAmount;
      } else {
        // Only the fixed side is known exactly — the other is an FX estimate the user
        // corrects with the real bank amount when they confirm. A failed FX lookup skips
        // this item for today rather than aborting the whole batch; it retries tomorrow
        // since last_generated_period is only stamped after a successful generation.
        try {
          if (fixedSide === 'dest') {
            destAmount = fixedAmount;
            sourceAmount = fixedAmount * getFxRate({ from_currency: destCurrency, to_currency: sourceCurrency }).rate;
          } else {
            sourceAmount = fixedAmount;
            destAmount = fixedAmount * getFxRate({ from_currency: sourceCurrency, to_currency: destCurrency }).rate;
          }
        } catch (fxErr) {
          continue;
        }
      }
      var fxRate = destAmount / sourceAmount;

      var newId = nextId_(trSheet);
      trSheet.appendRow([
        newId, today, Number(sourceAccountId), Number(destAccountId), sourceAmount,
        destAmount, fxRate, '', Number(recurringTransferId), false
      ]);
      recurringSheet.getRange(2 + i, 10).setValue(periodKey);
      generated++;
    }
  } finally {
    lock.releaseLock();
  }

  return { generated: generated };
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
  setBudget: setBudget,
  getPendingConfirmations: getPendingConfirmations,
  confirmTransaction: confirmTransaction,
  updateTransaction: updateTransaction,
  deleteTransaction: deleteTransaction,
  updateTransfer: updateTransfer,
  deleteTransfer: deleteTransfer,
  getRecurringTransfers: getRecurringTransfers,
  addRecurringTransfer: addRecurringTransfer,
  updateRecurringTransfer: updateRecurringTransfer,
  confirmTransfer: confirmTransfer,
  getPendingTransferConfirmations: getPendingTransferConfirmations
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
