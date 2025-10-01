

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
 * Configuration
 */
const SHEET_NAME = 'FinanceData';
const TRANSACTIONS_SHEET = 'Transactions';
const CATEGORIES_SHEET = 'Categories';
const BUDGETS_SHEET = 'Budgets'; 

/**
 * Initialize sheets if they don't exist
 */
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create Transactions sheet
  let transactionsSheet = ss.getSheetByName(TRANSACTIONS_SHEET);
  if (!transactionsSheet) {
    transactionsSheet = ss.insertSheet(TRANSACTIONS_SHEET);
    const headers = [['ID', 'Date', 'Category', 'Type', 'Amount', 'Description']];
    transactionsSheet.getRange(1, 1, 1, headers[0].length).setValues(headers).setFontWeight('bold');
  }
  
  // Create Categories sheet
  let categoriesSheet = ss.getSheetByName(CATEGORIES_SHEET);
  if (!categoriesSheet) {
    categoriesSheet = ss.insertSheet(CATEGORIES_SHEET);
    categoriesSheet.getRange(1, 1, 1, 4).setValues([
      ['ID', 'Name', 'Type', 'Color']
    ]);
    categoriesSheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    
    // Add default categories
    const defaultCategories = [
      [1, 'เงินเดือน', 'income', '#4caf50'],
      [2, 'รายได้เสริม', 'income', '#8bc34a'],
      [3, 'อาหาร', 'expense', '#ff9800'],
      [4, 'เดินทาง', 'expense', '#f44336'],
      [5, 'ช้อปปิ้ง', 'expense', '#e91e63'],
      [6, 'บิล/ค่าใช้จ่าย', 'expense', '#9c27b0']
    ];
    categoriesSheet.getRange(2, 1, defaultCategories.length, 4).setValues(defaultCategories);
  }
  
  // Create Budgets sheet
  let budgetsSheet = ss.getSheetByName(BUDGETS_SHEET);
  if (!budgetsSheet) {
    budgetsSheet = ss.insertSheet(BUDGETS_SHEET);
    budgetsSheet.getRange(1, 1, 1, 3).setValues([
      ['MonthYear', 'CategoryID', 'Amount']
    ]);
    budgetsSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  }
  
  return 'Sheets initialized successfully';
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
      transaction.description || ''
    ];
    
    sheet.appendRow(newRow);
    return { success: true, id: id };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function saveMultipleTransactions(transactions) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadpreheet();
    let sheet = ss.getSheetByName(TRANSACTIONS_SHEET);
    if (!sheet) {
      initializeSheets(); // ถ้ายังไม่มีชีต ให้สร้างก่อน
      sheet = ss.getSheetByName(TRANSACTIONS_SHEET);
    }
    
    // แปลง array of objects ให้เป็น array of arrays (rows) สำหรับการบันทึก
    const rows = transactions.map(t => [
      t.id || new Date().getTime() + Math.random(), // สร้าง ID ถ้าไม่มี
      new Date(t.date), // แปลง string date เป็น Date object
      t.category,
      t.type,
      t.amount,
      t.description || '' // ใส่ค่าว่างถ้าไม่มี description
    ]);
    
    if (rows.length > 0) {
      // บันทึกข้อมูลทั้งหมดลงชีตในครั้งเดียวเพื่อประสิทธิภาพที่ดีกว่าการ appendRow ทีละแถว
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
        description: data[i][5]
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
      if (isNaN(dateObj.getTime())) continue; // ข้ามถ้าไม่ใช่วันถูกต้อง

      const month = String(dateObj.getMonth() + 1).padStart(2, '0');
      const year = dateObj.getFullYear();
      const monthYear = `${year}-${month}`; // ใช้ key แบบ YYYY-MM

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
    
    // Clear transactions
    const transactionsSheet = ss.getSheetByName(TRANSACTIONS_SHEET);
    if (transactionsSheet && transactionsSheet.getLastRow() > 1) {
      transactionsSheet.getRange(2, 1, transactionsSheet.getLastRow() - 1, 6).clear();
    }
    
    // Clear budgets
    const budgetsSheet = ss.getSheetByName(BUDGETS_SHEET);
    if (budgetsSheet && budgetsSheet.getLastRow() > 1) {
      budgetsSheet.getRange(2, 1, budgetsSheet.getLastRow() - 1, 3).clear();
    }
    
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

    Logger.log('Transactions: ' + JSON.stringify(transactions));
    Logger.log('Categories: ' + JSON.stringify(categories));
    Logger.log('Budgets: ' + JSON.stringify(budgets));

    const data = {
      transactions: transactions || [],
      categories: categories || [],
      budgets: budgets || {}
    };

    return data;

  } catch (error) {
    Logger.log('❌ Error in getAppData: ' + error);
    return {
      transactions: [],
      categories: [],
      budgets: {}
    };
  }
}


/**
 * Test functions
 */
function testConnection() {
  return 'Google Apps Script connection successful!';
}
