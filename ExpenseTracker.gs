function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Expense Tracker')
    .addItem('Add Expense', 'showSidebar')
    .addToUi();

  // Open sidebar form automatically on sheet open
  showSidebar();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('ExpenseForm')
    .setTitle('Add Expense Record');
  SpreadsheetApp.getUi().showSidebar(html);
}

// Get Category and Payment Method options from named ranges
function getDropdownOptions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const categoryRange = sheet.getRangeByName('CategoryRange');
  const paymentRange = sheet.getRangeByName('PaymentMethodRange');
  
  const uniqueCategories = [];
  const uniquePayments = [];
  
  if (categoryRange) {
    const categories = categoryRange.getValues().flat();
    categories.forEach(cat => {
      if (cat && !uniqueCategories.includes(cat)) {
        uniqueCategories.push(cat);
      }
    });
  }
  
  if (paymentRange) {
    const payments = paymentRange.getValues().flat();
    payments.forEach(pay => {
      if (pay && !uniquePayments.includes(pay)) {
        uniquePayments.push(pay);
      }
    });
  }
  
  return {
    categories: uniqueCategories,
    payments: uniquePayments
  };
}

function getFirstRecordRow(sheet) {
  // If header is at row 1, records start at row 2
  const startRow = 2;
  const lastRow = sheet.getLastRow();

  // Scan from startRow down for first record
  for (let i = startRow; i <= lastRow; i++) {
    const val = sheet.getRange(i, 1).getValue();
    if (val !== "" && val !== null) {
      return i;
    }
  }
  // If no records found, return where first record should go
  return startRow;
}

function saveExpense(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Find top of records block
  const insertRow = getFirstRecordRow(sheet);

  // Insert a new row above first record row (shifts all records down)
  sheet.insertRowBefore(insertRow);

  // Write the expense at the inserted row
  const rowData = [
    data.date,
    data.category,
    data.description,
    parseFloat(data.amount),
    data.investment,
    data.paymentMethod,
    data.notes
  ];
  sheet.getRange(insertRow, 1, 1, rowData.length).setValues([rowData]);

  return 'New expense added at row ' + insertRow + ', for '+ data.date + '.';
}


