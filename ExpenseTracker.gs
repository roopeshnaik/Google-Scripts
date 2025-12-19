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

function findInsertRowByDate(sheet, newDate) {
  // Assuming "Date" is in column 1, data starts at row 2
  const lastRow = sheet.getLastRow();
  for (let i = 2; i <= lastRow; i++) {
    const cellDateStr = sheet.getRange(i, 1).getValue();
    if (!cellDateStr) continue;

    // Parse both Excel and ISO date formats
    const cellDate = new Date(cellDateStr);
    const testDate = new Date(newDate);

    // If the new date is greater or equal, insert above this row
    // For descending order: newer dates higher
    if (testDate >= cellDate) {
      return i;
    }
  }
  // If not found, insert at the bottom
  return lastRow + 1;
}

function saveExpense(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Find where to insert by date
  const insertRow = findInsertRowByDate(sheet, data.date);

  // Insert a new row above found position
  sheet.insertRowBefore(insertRow);

  // Correct expense/investment column logic
  let amountValue = data.amount || 0;
  let investmentValue = 0;
  if (data.category.trim().toLowerCase() === 'investments') {
    investmentValue = amountValue;
    amountValue = 0;
  }

  // Write the expense at the inserted row
  const rowData = [
    data.date,
    data.category,
    data.description,
    amountValue,
    investmentValue,
    data.paymentMethod,
    data.notes
  ];
  sheet.getRange(insertRow, 1, 1, rowData.length).setValues([rowData]);
  return 'New expense added at ' + insertRow + ' for ' + amountValue + ', against ' + data.category + '.';
}
