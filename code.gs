/**
 * Configuration
 */
const SHEET_NAME = 'FinanceData';
const TRANSACTIONS_SHEET = 'Transactions';
const CATEGORIES_SHEET = 'Categories';
const BUDGETS_SHEET = 'Budgets';
const ACCOUNTS_SHEET = 'Accounts'; // ใหม่: ชีตบัญชี
const TRANSFERS_SHEET = 'Transfers'; // ใหม่: ชีตการโอน
const TAX_SHEET = 'Tax'; // ใหม่: ชีตภาษี

function doGet() {
  let html = HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('ระบบบันทึกรายรับ-รายจ่าย')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}

function getURL(){
  return ScriptApp.getService().getUrl();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


/**
 * Initialize sheets if they don't exist
 */
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Transactions Sheet ---
  let transactionsSheet = ss.getSheetByName(TRANSACTIONS_SHEET);
  if (!transactionsSheet) {
    transactionsSheet = ss.insertSheet(TRANSACTIONS_SHEET);
    // เพิ่ม AccountID เพื่อผูกรายการกับบัญชี
    const headers = [['ID', 'Date', 'Category', 'Type', 'Amount', 'Description', 'AccountID']];
    transactionsSheet.getRange(1, 1, 1, headers[0].length).setValues(headers).setFontWeight('bold');
  }

  // --- Categories Sheet ---
  let categoriesSheet = ss.getSheetByName(CATEGORIES_SHEET);
  if (!categoriesSheet) {
    categoriesSheet = ss.insertSheet(CATEGORIES_SHEET);
    categoriesSheet.getRange(1, 1, 1, 4).setValues([['ID', 'Name', 'Type', 'Color']]).setFontWeight('bold');
  }

  // --- Budgets Sheet ---
  let budgetsSheet = ss.getSheetByName(BUDGETS_SHEET);
  if (!budgetsSheet) {
    budgetsSheet = ss.insertSheet(BUDGETS_SHEET);
    budgetsSheet.getRange(1, 1, 1, 3).setValues([['MonthYear', 'CategoryID', 'Amount']]).setFontWeight('bold');
  }

  // --- Accounts Sheet (ใหม่) ---
  let accountsSheet = ss.getSheetByName(ACCOUNTS_SHEET);
  if (!accountsSheet) {
    accountsSheet = ss.insertSheet(ACCOUNTS_SHEET);
    const headers = [['ID', 'Name', 'Type', 'InitialBalance', 'Icon']];
    accountsSheet.getRange(1, 1, 1, headers[0].length).setValues(headers).setFontWeight('bold');
  }

  // --- Transfers Sheet (ใหม่) ---
  let transfersSheet = ss.getSheetByName(TRANSFERS_SHEET);
  if (!transfersSheet) {
    transfersSheet = ss.insertSheet(TRANSFERS_SHEET);
    const headers = [['ID', 'Date', 'FromAccountID', 'ToAccountID', 'Amount', 'Description']];
    transfersSheet.getRange(1, 1, 1, headers[0].length).setValues(headers).setFontWeight('bold');
  }

  // --- Tax Sheet (ใหม่) ---
  let taxSheet = ss.getSheetByName(TAX_SHEET);
  if (!taxSheet) {
    taxSheet = ss.insertSheet(TAX_SHEET);
    const headers = [['Year', 'TotalIncome', 'DeductibleExpenses', 'TaxPaid']];
    taxSheet.getRange(1, 1, 1, headers[0].length).setValues(headers).setFontWeight('bold');
  }

  return 'Sheets initialized successfully';
}


// ==================================================================
//  ACCOUNT FUNCTIONS
// ==================================================================
function saveAccount(account) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(ACCOUNTS_SHEET);
    const id = account.id || new Date().getTime();
    const newRow = [id, account.name, account.type, account.initialBalance, account.icon];
    sheet.appendRow(newRow);
    return { success: true, id: id };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function getAllAccounts() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(ACCOUNTS_SHEET);
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const accounts = [];
    for (let i = 1; i < data.length; i++) {
      accounts.push({
        id: data[i][0],
        name: data[i][1],
        type: data[i][2],
        initialBalance: parseFloat(data[i][3]) || 0,
        icon: data[i][4]
      });
    }
    return accounts;
  } catch (error) {
    return [];
  }
}

function deleteAccount(accountId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(ACCOUNTS_SHEET);
    if (!sheet) return { success: false, error: 'Sheet not found' };

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == accountId) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Account not found' };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function updateAccount(account) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(ACCOUNTS_SHEET);
    if (!sheet) return { success: false, error: 'Sheet not found' };

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == account.id) { // ค้นหาแถวด้วย ID
        // อัปเดตข้อมูลในแถวที่เจอ
        sheet.getRange(i + 1, 2, 1, 3).setValues([
          [account.name, account.type, account.initialBalance]
        ]);
        return { success: true };
      }
    }
    return { success: false, error: 'Account not found' };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// ==================================================================
//  TRANSFER FUNCTIONS
// ==================================================================
function saveTransfer(transfer) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TRANSFERS_SHEET);
    const id = transfer.id || new Date().getTime();

    const newRow = [
      id,
      transfer.date,
      transfer.from,
      transfer.to,
      transfer.amount,
      transfer.description || ''
    ];
    sheet.appendRow(newRow);
    return { success: true, id: id };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function getAllTransfers() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TRANSFERS_SHEET);
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const transfers = [];
    for (let i = 1; i < data.length; i++) {
      transfers.push({
        id: data[i][0],
        date: Utilities.formatDate(new Date(data[i][1]), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        from: data[i][2],
        to: data[i][3],
        amount: parseFloat(data[i][4]) || 0,
        description: data[i][5]
      });
    }
    return transfers;
  } catch (error) {
    return [];
  }
}

// ==================================================================
//  TAX FUNCTIONS
// ==================================================================
function saveTaxData(taxData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TAX_SHEET);
    const data = sheet.getDataRange().getValues();
    let recordFound = false;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == taxData.year) {
        sheet.getRange(i + 1, 2, 1, 3).setValues([[
          taxData.totalIncome,
          taxData.deductibleExpenses,
          taxData.taxPaid
        ]]);
        recordFound = true;
        break;
      }
    }

    if (!recordFound) {
      sheet.appendRow([
        taxData.year,
        taxData.totalIncome,
        taxData.deductibleExpenses,
        taxData.taxPaid
      ]);
    }
    return { success: true };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function getAllTaxData() {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(TAX_SHEET);
        if (!sheet) return {};

        const data = sheet.getDataRange().getValues();
        if (data.length <= 1) return {};

        const taxRecords = {};
        for (let i = 1; i < data.length; i++) {
            taxRecords[data[i][0]] = {
                year: data[i][0],
                totalIncome: parseFloat(data[i][1]) || 0,
                deductibleExpenses: parseFloat(data[i][2]) || 0,
                taxPaid: parseFloat(data[i][3]) || 0
            };
        }
        return taxRecords;
    } catch (error) {
        Logger.log('Error getting tax data: ' + error);
        return {};
    }
}

/**
 * Transaction functions
 */
function saveTransaction(transaction) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(TRANSACTIONS_SHEET);
    
    if (!sheet) {
      initializeSheets();
      sheet = ss.getSheetByName(TRANSACTIONS_SHEET);
    }
    
    const id = transaction.id || new Date().getTime();
    const newRow = [
      id,
      transaction.date,
      transaction.category,
      transaction.type,
      transaction.amount,
      transaction.description || '',
      transaction.accountId || null
    ];
    
    sheet.appendRow(newRow);
    return { success: true, id: id };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function saveMultipleTransactions(transactions) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(TRANSACTIONS_SHEET);
    if (!sheet) {
      initializeSheets();
      sheet = ss.getSheetByName(TRANSACTIONS_SHEET);
    }

    const rows = transactions.map(t => [
      t.id || new Date().getTime() + Math.random(),
      new Date(t.date),
      t.category,
      t.type,
      t.amount,
      t.description || '',
      t.accountId || null
    ]);

    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }

    return { success: true, count: rows.length };
  } catch (error) {
    Logger.log(`Error in saveMultipleTransactions: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}

function getAllTransactions() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TRANSACTIONS_SHEET);
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const transactions = [];
    for (let i = 1; i < data.length; i++) {
      const dateValue = data[i][1];
      let dateStr = '';

      if (dateValue instanceof Date) {
        dateStr = Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        dateStr = dateValue || '';
      }

      transactions.push({
        id: data[i][0],
        date: dateStr,
        category: data[i][2],
        type: data[i][3],
        amount: data[i][4],
        description: data[i][5],
        accountId: data[i][6]
      });
    }

    return transactions;
  } catch (error) {
    Logger.log('Error getting transactions: ' + error);
    return [];
  }
}


function deleteTransaction(transactionId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TRANSACTIONS_SHEET);
    if (!sheet) return { success: false, error: 'Sheet not found' };
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == transactionId) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    
    return { success: false, error: 'Transaction not found' };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Category functions
 */
function saveCategory(category) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CATEGORIES_SHEET);
    
    if (!sheet) {
      initializeSheets();
      sheet = ss.getSheetByName(CATEGORIES_SHEET);
    }
    
    const id = category.id || new Date().getTime();
    const newRow = [
      id,
      category.name,
      category.type,
      category.color
    ];
    sheet.appendRow(newRow);
    return { success: true, id: id };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function getAllCategories() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CATEGORIES_SHEET);
    if (!sheet) {
      initializeSheets();
      sheet = ss.getSheetByName(CATEGORIES_SHEET);
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    const categories = [];
    for (let i = 1; i < data.length; i++) {
      categories.push({
        id: data[i][0],
        name: data[i][1],
        type: data[i][2],
        color: data[i][3]
      });
    }
    
    return categories;
  } catch (error) {
    console.error('Error getting categories:', error);
    return [];
  }
}

function deleteCategory(categoryId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CATEGORIES_SHEET);
    if (!sheet) return { success: false, error: 'Sheet not found' };
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == categoryId) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    
    return { success: false, error: 'Category not found' };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Budget functions
 */
function saveBudget(monthYear, budgets) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(BUDGETS_SHEET);

    if (!sheet) {
      initializeSheets();
      sheet = ss.getSheetByName(BUDGETS_SHEET);
    }

    deleteBudgetsByMonth(monthYear);

    if (budgets.length > 0) {
      const rows = budgets.map(b => [monthYear, b.categoryId, b.amount]);
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
    }

    return { success: true };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function getBudgetsByMonth(monthYear) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(BUDGETS_SHEET);
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const budgets = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === monthYear) {
        budgets.push({
          categoryId: data[i][1],
          amount: data[i][2]
        });
      }
    }

    return budgets;
  } catch (error) {
    console.error('Error getting budgets by month:', error);
    return [];
  }
}

function getAllBudgets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(BUDGETS_SHEET);
    if (!sheet) return {};

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return {};

    const budgets = {};
    for (let i = 1; i < data.length; i++) {
      let rawDate = data[i][0];
      let dateObj = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
      if (isNaN(dateObj.getTime())) continue;

      const month = String(dateObj.getMonth() + 1).padStart(2, '0');
      const year = dateObj.getFullYear();
      const monthYear = `${year}-${month}`;

      if (!budgets[monthYear]) budgets[monthYear] = [];
      budgets[monthYear].push({
        categoryId: data[i][1],
        amount: data[i][2]
      });
    }

    return budgets;
  } catch (error) {
    console.error('Error getting all budgets:', error);
    return {};
  }
}

function deleteBudgetsByMonth(monthYear) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(BUDGETS_SHEET);
    if (!sheet) return { success: false, error: 'Sheet not found' };

    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      let rawDate = data[i][0];
      let dateObj = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
      if (isNaN(dateObj.getTime())) continue;
      const rowMonth = String(dateObj.getMonth() + 1).padStart(2, '0');
      const rowYear = dateObj.getFullYear();
      const rowMonthYear = `${rowYear}-${rowMonth}`;
      if (rowMonthYear === monthYear) {
        sheet.deleteRow(i + 1);
      }
    }

    return { success: true };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Utility functions
 */
function clearAllData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const sheetNames = [TRANSACTIONS_SHEET, BUDGETS_SHEET, CATEGORIES_SHEET, ACCOUNTS_SHEET, TRANSFERS_SHEET, TAX_SHEET];
    
    sheetNames.forEach(name => {
      const sheet = ss.getSheetByName(name);
      if (sheet && sheet.getLastRow() > 1) {
        sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear();
      }
    });
    
    return { success: true };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function getAppData() {
  try {
    const transactions = getAllTransactions();
    const categories = getAllCategories();
    const budgets = getAllBudgets();
    const accounts = getAllAccounts();
    const transfers = getAllTransfers();
    const taxData = getAllTaxData();

    const data = {
      transactions: transactions || [],
      categories: categories || [],
      budgets: budgets || {},
      accounts: accounts || [],
      transfers: transfers || [],
      taxData: taxData || {}
    };

    return data;

  } catch (error) {
    Logger.log('❌ Error in getAppData: ' + error);
    return {
      transactions: [],
      categories: [],
      budgets: {},
      accounts: [],
      transfers: [],
      taxData: {}
    };
  }
}

/**
 * Test functions
 */
function testConnection() {
  return 'Google Apps Script connection successful!';
}
