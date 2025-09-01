// 請替換為您的 Google Sheets ID
const SHEET_ID = '1tcGHNiKmedqv6A_1a3zH95sxkJxhBoKdtXqwQ-jnoS0';
const SHEET_NAME = '預算表';

/**
 * 載入並顯示 index.html 網頁。
 * 這是使用者存取網頁時執行的函式。
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate();
}

/**
 * 處理來自前端的資料同步請求。
 * 由 google.script.run.syncData() 呼叫。
 * @param {Array<object>} data 包含要同步的資料物件陣列。
 * @returns {object} 包含同步結果的狀態和訊息。
 */
function syncData(data) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) {
      throw new Error(`Sheet with name "${SHEET_NAME}" not found.`);
    }

    if (!Array.isArray(data) || data.length === 0) {
      throw new Error('Invalid data format. Expected an array of objects.');
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headerMap = {
      'StoreName': headers.indexOf('StoreName'),
      'Year': headers.indexOf('Year'),
      'Month': headers.indexOf('Month'),
      'Amount': headers.indexOf('Amount')
    };

    for (const key in headerMap) {
      if (headerMap[key] === -1) {
        throw new Error(`Required header "${key}" not found.`);
      }
    }

    // 讀取現有資料
    const lastRow = sheet.getLastRow();
    const existingData = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, headers.length).getValues() : [];
    const existingMap = new Map();
    existingData.forEach((row, index) => {
      const key = `${row[headerMap.StoreName]}-${row[headerMap.Year]}-${row[headerMap.Month]}`;
      existingMap.set(key, index + 2); // 儲存行號
    });

    const newRows = [];
    data.forEach(item => {
      const storeName = String(item.store_name || '').trim();
      const year = parseInt(item.year);
      const month = parseInt(item.month);
      const amount = parseFloat(item.amount);

      if (!storeName || isNaN(year) || isNaN(month) || isNaN(amount)) {
        console.warn(`Skipping invalid data: ${JSON.stringify(item)}`);
        return;
      }

      const key = `${storeName}-${year}-${month}`;
      if (existingMap.has(key)) {
        const rowIndex = existingMap.get(key);
        sheet.getRange(rowIndex, headerMap.Amount + 1).setValue(amount);
      } else {
        newRows.push([storeName, year, month, amount]);
      }
    });

    if (newRows.length > 0) {
      sheet.getRange(lastRow + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    }

    return { status: 'success', message: `資料已成功同步。新增 ${newRows.length} 筆，更新 ${data.length - newRows.length} 筆。` };
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}


/**
 * 處理來自前端的讀取資料請求。
 * 由 google.script.run.getSheetData() 呼叫。
 * @returns {object} 包含試算表資料的狀態和訊息。
 */
function getSheetData() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) {
      throw new Error(`Sheet with name "${SHEET_NAME}" not found.`);
    }

    const lastRow = sheet.getLastRow();
    const values = lastRow <= 1 ? [] : sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

    return { status: 'success', data: values };
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}