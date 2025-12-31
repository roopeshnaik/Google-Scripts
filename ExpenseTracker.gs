function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Expense Tracker')
    .addItem('Add Expense', 'showSidebar')
    .addToUi();

  showSidebar();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('ExpenseForm')
    .setTitle('Add Expense Record');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Dropdown options for the form
 * Expense_Master sheet with columns ordered
 * (Expense Name, CategoryRange, Sub-Category, DefaultPayment, Expense Nature)
 */
function getDropdownOptions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName('Expense_Master');

  const lastRow = master.getLastRow();
  if (lastRow < 2) {
    return { expenseNames: [], categories: [], payment: [] };
  }

  const data = master.getRange(2, 1, lastRow - 1, 5).getValues();

  const categoriesSet = new Set();
  const paymentsSet = new Set();
  const expenseNames = [];

  data.forEach(row => {
    const expenseName = row[0];
    const category = row[1];
    const payment = row[3];

    if (expenseName) expenseNames.push(expenseName);
    if (category) categoriesSet.add(category);
    if (payment) paymentsSet.add(payment);
  });

  // 🔑 Alphabetical sort (case-insensitive)
  expenseNames.sort((a, b) =>
    a.toString().localeCompare(b.toString(), undefined, { sensitivity: 'base' })
  );

  return {
    expenseNames,
    categories: [...categoriesSet],
    payment: [...paymentsSet]
  };
}


/**
 * Lookup expense details by Expense Name
 */
function getExpenseDetails(expenseName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName('Expense_Master');
  const data = master.getDataRange().getValues();

  const header = data.shift();
  const idx = {
    expense: header.indexOf('Expense Name'),
    category: header.indexOf('CategoryRange'),
    sub: header.indexOf('Sub-Category'),
    payment: header.indexOf('DefaultPayment'),
    nature: header.indexOf('Expense Nature')
  };

  const row = data.find(r => r[idx.expense] === expenseName);
  if (!row) return {};

  return {
    category: row[idx.category],
    subCategory: row[idx.sub],
    payment: row[idx.payment],
    nature: row[idx.nature]
  };
}

/**
 * Find insert row so newest dates stay on top
 */
function findInsertRowByDate(sheet, newDate) {
  const lastRow = sheet.getLastRow();
  const testDate = new Date(newDate);

  for (let i = 2; i <= lastRow; i++) {
    const cellDate = sheet.getRange(i, 1).getValue();
    if (!cellDate) continue;

    if (testDate >= new Date(cellDate)) {
      return i;
    }
  }
  return lastRow + 1;
}

/**
 * Save expense — ALWAYS re-derive from Expense_Master
 */
function saveExpense(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = getMonthSheetName(data.date);
  const sheet = getOrCreateMonthSheet(sheetName);
  const master = ss.getSheetByName('Expense_Master');

  const masterData = master.getDataRange().getValues();
  const header = masterData.shift();

  const idx = {
    expense: header.indexOf('Expense Name'),
    category: header.indexOf('CategoryRange'),
    sub: header.indexOf('Sub-Category'),
    payment: header.indexOf('DefaultPayment'),
    nature: header.indexOf('Expense Nature')
  };

  const row = masterData.find(r => r[idx.expense] === data.expenseName);
  if (!row) throw new Error('Invalid Expense Name');

  const category = row[idx.category];
  const subCategory = row[idx.sub];
  const payment = data.paymentMethod || row[idx.payment];
  const nature = row[idx.nature];

  let spent = Number(data.amount) || 0;
  sheet.activate();
  const insertRow = findInsertRowByDate(sheet, data.date);
  sheet.insertRowBefore(insertRow);

  const rowData = [
    data.date,
    data.expenseName,
    spent,
    category,
    payment,
    subCategory,
    nature,
    data.notes || ''
  ];

  sheet.getRange(insertRow, 1, 1, rowData.length).setValues([rowData]);

  return `Added: ${data.expenseName} – ₹${spent || investmentValue}`;
}

/**
 * Find the month sheet name
 */
function getMonthSheetName(dateStr) {
  const d = new Date(dateStr);
  return Utilities.formatDate(
    d,
    Session.getScriptTimeZone(),
    'MMM yy'
  );
}

/**
 * Creates or gets month sheet from _Month_Template
 */
function getOrCreateMonthSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    const template = ss.getSheetByName('_Month_Template');
    if (!template) {
      throw new Error('Template sheet "_Month_Template" not found');
    }

    // Step 1: copy template (new sheet appears at end)
    sheet = template.copyTo(ss);
    sheet.setName(sheetName);
    sheet.showSheet();

    // Step 4: clear leftover data below header
    const lastRow = sheet.getLastRow();
    if (lastRow > 2) {
      sheet
        .getRange(3, 1, lastRow - 2, sheet.getLastColumn())
        .clearContent();
    }
  }

  return sheet;
}
