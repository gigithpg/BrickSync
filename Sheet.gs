const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

const Utils = {
  clearRange: (sheet, row, col, numRows, numCols) => {
    const startTime = new Date().getTime();
    if (numRows <= 0) {
      logMessage('WARN', `Utils.clearRange skipped for ${sheet.getName()}!R${row}C${col}:R${row + numRows - 1}C${col + numCols - 1} due to invalid numRows: ${numRows}`);
      return;
    }
    logMessage('INFO', `Utils.clearRange started for ${sheet.getName()}!R${row}C${col}:R${row + numRows - 1}C${col + numCols - 1}`);
    if (!sheet) throw new Error('Sheet not found');
    sheet.getRange(row, col, numRows, numCols).clearContent();
    logMessage('INFO', `Utils.clearRange completed, took ${new Date().getTime() - startTime}ms`);
  }
};

function logMessage(level, message) {
  try {
    const logSheet = spreadsheet.getSheetByName('Logs');
    if (!logSheet) {
      Logger.log(`${level}: ${message}`);
      return;
    }
    if (logSheet.getLastRow() > 10000) {
      const archiveSheet = spreadsheet.insertSheet(`Logs_Archive_${new Date().toISOString().slice(0, 10)}`);
      archiveSheet.getRange(1, 1, logSheet.getLastRow(), 2).setValues(logSheet.getRange(1, 1, logSheet.getLastRow(), 2).getValues());
      logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 2).clearContent();
      logMessage('INFO', `Archived logs to Logs_Archive_${new Date().toISOString().slice(0, 10)}`);
    }
    const timestamp = new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' });
    logSheet.appendRow([timestamp, `${level}: ${message}`]);
  } catch (e) {
    Logger.log(`Failed to log: ${e.message}`);
  }
}

function getLastContentRow(sheet, columns) {
  const startTime = new Date().getTime();
  logMessage('INFO', `getLastContentRow started for ${sheet.getName()}, columns: ${columns}`);
  const cacheKey = `LastRow_${sheet.getName()}_${columns.join(',')}`;
  const props = PropertiesService.getScriptProperties();
  const cached = props.getProperty(cacheKey);
  const sheetLastRow = sheet.getLastRow();
  if (cached && Number(cached) >= sheetLastRow && Number(cached) > 1) {
    logMessage('INFO', `getLastContentRow completed, lastRow: ${cached}, took ${new Date().getTime() - startTime}ms (cached)`);
    return Number(cached);
  }
  if (sheetLastRow <= 1) {
    props.setProperty(cacheKey, '1');
    logMessage('INFO', `getLastContentRow completed, lastRow: 1, took ${new Date().getTime() - startTime}ms`);
    return 1;
  }
  const minCol = Math.min(...columns);
  const maxCol = Math.max(...columns);
  const startRow = Math.max(2, sheetLastRow - 5);
  const range = sheet.getRange(startRow, minCol, sheetLastRow - startRow + 1, maxCol - minCol + 1);
  const values = range.getValues();
  let maxRow = 1;
  columns.forEach(col => {
    const colIdx = col - minCol;
    for (let row = values.length - 1; row >= 0; row--) {
      if (values[row][colIdx] !== '' && values[row][colIdx] != null) {
        maxRow = Math.max(maxRow, row + startRow);
        break;
      }
    }
  });
  props.setProperty(cacheKey, maxRow.toString());
  logMessage('INFO', `getLastContentRow completed, lastRow: ${maxRow}, took ${new Date().getTime() - startTime}ms`);
  return maxRow;
}

function batchCacheLastRows(sheetLastRows) {
  const props = PropertiesService.getScriptProperties();
  const cache = {};
  Object.entries(sheetLastRows).forEach(([sheetName, { lastRow, columns }]) => {
    const cacheKey = `LastRow_${sheetName}_${columns.join(',')}`;
    cache[cacheKey] = lastRow.toString();
  });
  props.setProperties(cache);
  logMessage('INFO', `Batch cached last rows for ${Object.keys(sheetLastRows).length} sheets`);
}

function invalidateLastRowCache(sheetName, columns) {
  const cacheKey = `LastRow_${sheetName}_${columns.join(',')}`;
  PropertiesService.getScriptProperties().deleteProperty(cacheKey);
  logMessage('INFO', `Invalidated LastRow cache for ${sheetName}, columns: ${columns}`);
}

function getCachedCustomers() {
  const cacheKey = 'CustomerNames';
  const props = PropertiesService.getScriptProperties();
  let customers = JSON.parse(props.getProperty(cacheKey) || '[]');
  if (customers.length > 0) {
    logMessage('INFO', 'Retrieved customer names from cache');
    return customers;
  }
  const customersSheet = spreadsheet.getSheetByName('Customers');
  const lastRow = getLastContentRow(customersSheet, [2]);
  customers = lastRow > 1 ? customersSheet.getRange(`B2:B${lastRow}`).getValues().flat().filter(name => name) : [];
  props.setProperty(cacheKey, JSON.stringify(customers));
  logMessage('INFO', `Cached ${customers.length} customer names`);
  return customers;
}

function invalidateCustomerCache() {
  PropertiesService.getScriptProperties().deleteProperty('CustomerNames');
  logMessage('INFO', 'Invalidated customer names cache');
}

function validateCustomers(sheetName, customerCol, range, startRow, numRows) {
  const startTime = new Date().getTime();
  logMessage('INFO', `validateCustomers started for ${sheetName}!${customerCol}${startRow}:${customerCol}${startRow + numRows - 1}`);
  const customerNames = getCachedCustomers();
  const values = range.getValues().slice(startRow - range.getRow(), startRow - range.getRow() + numRows);
  const errors = [];
  values.forEach((row, idx) => {
    const customer = row[0];
    if (customer && !customerNames.includes(customer)) {
      errors.push(`Invalid customer '${customer}' in ${sheetName}!${customerCol}${startRow + idx}`);
    }
  });
  if (errors.length > 0) {
    logMessage('ERROR', `Customer validation failed:\n${errors.join('\n')}`);
    throw new Error(`Customer validation failed:\n${errors.join('\n')}`);
  }
  logMessage('INFO', `validateCustomers completed for ${sheetName}!${customerCol}${startRow}:${customerCol}${startRow + numRows - 1}, took ${new Date().getTime() - startTime}ms`);
}

function formatDateColumns() {
  const startTime = new Date().getTime();
  logMessage('INFO', 'formatDateColumns started');
  const sheetsToFormat = [
    { name: 'Sales', dateColumns: ['A'] },
    { name: 'Payments', dateColumns: ['A'] },
    { name: 'Transactions', dateColumns: ['B'] }
  ];
  const sheetLastRows = {};
  const sheets = {};
  sheetsToFormat.forEach(sheetInfo => {
    const sheet = sheets[sheetInfo.name] || spreadsheet.getSheetByName(sheetInfo.name);
    sheets[sheetInfo.name] = sheet;
    if (!sheet) return;
    const colIndices = sheetInfo.dateColumns.map(col => col.charCodeAt(0) - 64);
    const lastRow = getLastContentRow(sheet, colIndices);
    if (lastRow <= 1) return;
    sheetLastRows[sheetInfo.name] = { lastRow, columns: colIndices };
    sheetInfo.dateColumns.forEach(col => {
      logMessage('INFO', `Formatting ${sheetInfo.name}!${col}2:${col}${lastRow}`);
      sheet.getRange(`${col}2:${col}${lastRow}`).setNumberFormat(CONFIG.DATE_FORMAT);
      logMessage('INFO', `Formatted ${sheetInfo.name}!${col}2:${col}${lastRow} as ${CONFIG.DATE_FORMAT}`);
      Utilities.sleep(50);
    });
  });
  batchCacheLastRows(sheetLastRows);
  logMessage('INFO', `formatDateColumns completed, took ${new Date().getTime() - startTime}ms`);
}

function normalizeDate(date) {
  if (!date) return '';
  if (typeof date === 'string' && date.includes('T')) {
    try {
      const parsedDate = new Date(date);
      if (!isNaN(parsedDate)) {
        const today = new Date();
        if (parsedDate > today) {
          throw new Error(`Date ${Utilities.formatDate(parsedDate, 'Asia/Kolkata', 'dd/MM/yyyy')} is in the future`);
        }
        return Utilities.formatDate(parsedDate, 'Asia/Kolkata', 'dd/MM/yyyy');
      }
    } catch (e) {
      logMessage('ERROR', `normalizeDate failed: ${e.message}`);
      throw e;
    }
  }
  if (date instanceof Date && !isNaN(date)) {
    const today = new Date();
    if (date > today) {
      throw new Error(`Date ${Utilities.formatDate(date, 'Asia/Kolkata', 'dd/MM/yyyy')} is in the future`);
    }
    return Utilities.formatDate(date, 'Asia/Kolkata', 'dd/MM/yyyy');
  }
  if (typeof date === 'string' && date.match(/^\d{2}\/\d{2}\/\d{4}$/)) {
    return date;
  }
  return date;
}

function normalizePaymentMethod(method) {
  if (!method || typeof method !== 'string') return method || '';
  const validMethods = ['Bank Transfer', 'Cheque', 'Cash', 'UPI', 'Others'];
  const methodMap = {
    'bank transfer': 'Bank Transfer',
    'cheque': 'Cheque',
    'cash': 'Cash',
    'upi': 'UPI',
    'others': 'Others'
  };
  const lowerMethod = method.toLowerCase();
  return methodMap[lowerMethod] || (validMethods.includes(method) ? method : 'Others');
}

function updateTransactions(incrementalRange = null) {
  const startTime = new Date().getTime();
  logMessage('INFO', `updateTransactions started, incrementalRange: ${incrementalRange ? incrementalRange.getA1Notation() : 'full'}`);
  try {
    const sheets = {
      transactions: spreadsheet.getSheetByName('Transactions'),
      sales: spreadsheet.getSheetByName('Sales'),
      payments: spreadsheet.getSheetByName('Payments')
    };
    if (!sheets.transactions || !sheets.sales || !sheets.payments) throw new Error('Required sheets not found');

    const sheetLastRows = {};
    const lastRows = {
      sales: getLastContentRow(sheets.sales, [3]),
      payments: getLastContentRow(sheets.payments, [3, 5]),
      transactions: getLastContentRow(sheets.transactions, [3])
    };
    sheetLastRows['Sales'] = { lastRow: lastRows.sales, columns: [3] };
    sheetLastRows['Payments'] = { lastRow: lastRows.payments, columns: [3, 5] };
    sheetLastRows['Transactions'] = { lastRow: lastRows.transactions, columns: [3] };

    let salesData = [], paymentsData = [];
    let existingIds = new Set();

    if (lastRows.transactions > 1) {
      const transIds = sheets.transactions.getRange(`C2:C${lastRows.transactions}`).getValues().flat().filter(id => id);
      existingIds = new Set(transIds);
      logMessage('INFO', `Fetched ${transIds.length} transaction IDs from Transactions!C2:C${lastRows.transactions}`);
    }

    if (!incrementalRange) {
      if (lastRows.sales > 1) {
        const rawSalesData = sheets.sales.getRange(`A2:J${lastRows.sales}`).getValues();
        logMessage('INFO', `Raw Sales!A2:J${lastRows.sales}, ${rawSalesData.length} rows, sample: ${JSON.stringify(rawSalesData.slice(0, 5))}`);
        salesData = rawSalesData
          .map(row => [
            'Sale', normalizeDate(row[0]), row[1], row[2], Number(row[3]) || 0, Number(row[4]) || 0, Number(row[5]) || 0,
            Number(row[6]) || 0, normalizePaymentMethod(row[7]), Number(row[8]) || 0, row[9] || ''
          ])
          .filter(row => row[1] && row[2] && row[3] && row[7] && !existingIds.has(row[2]));
        logMessage('INFO', `Fetched Sales!A2:J${lastRows.sales}, ${salesData.length} valid rows out of ${rawSalesData.length}, sample: ${JSON.stringify(salesData.slice(0, 5))}`);
        if (salesData.length > 0) {
          validateCustomers('Sales', 'C', sheets.sales.getRange(`C2:C${lastRows.sales}`), 2, lastRows.sales - 1);
        }
      }

      if (lastRows.payments > 1) {
        const rawPaymentsData = sheets.payments.getRange(`A2:F${lastRows.payments}`).getValues();
        logMessage('INFO', `Raw Payments!A2:F${lastRows.payments}, ${rawPaymentsData.length} rows, sample: ${JSON.stringify(rawPaymentsData.slice(0, 5))}`);
        paymentsData = rawPaymentsData
          .map(row => [
            'Payment', normalizeDate(row[0]), row[1], row[2], 0, 0, 0, 0, normalizePaymentMethod(row[3]), Number(row[4]) || 0, row[5] || ''
          ])
          .filter(row => row[1] && row[2] && row[3] && row[9] && !existingIds.has(row[2]));
        logMessage('INFO', `Fetched Payments!A2:F${lastRows.payments}, ${paymentsData.length} valid rows out of ${rawPaymentsData.length}, sample: ${JSON.stringify(paymentsData.slice(0, 5))}`);
        if (paymentsData.length > 0) {
          validateCustomers('Payments', 'C', sheets.payments.getRange(`C2:C${lastRows.payments}`), 2, lastRows.payments - 1);
        }
      }

      const combinedData = [...salesData, ...paymentsData];
      if (combinedData.length === 0) {
        logMessage('INFO', `updateTransactions completed, no new data, took ${new Date().getTime() - startTime}ms`);
        batchCacheLastRows(sheetLastRows);
        return;
      }

      const nextRow = lastRows.transactions > 1 ? lastRows.transactions + 1 : 2;
      sheets.transactions.getRange(nextRow, 1, combinedData.length, 11).setValues(combinedData);
      logMessage('INFO', `Appended ${combinedData.length} rows to Transactions!A${nextRow}:K, sample: ${JSON.stringify(combinedData.slice(0, 5))}`);
    } else {
      const sheetName = incrementalRange.getSheet().getName();
      const rowStart = incrementalRange.getRow();
      const rowEnd = rowStart + incrementalRange.getNumRows() - 1;
      let newData = [];

      if (sheetName === 'Sales') {
        const range = sheets.sales.getRange(`A${rowStart}:J${rowEnd}`);
        newData = range.getValues()
          .map(row => [
            'Sale', normalizeDate(row[0]), row[1], row[2], Number(row[3]) || 0, Number(row[4]) || 0, Number(row[5]) || 0,
            Number(row[6]) || 0, normalizePaymentMethod(row[7]), Number(row[8]) || 0, row[9] || ''
          ])
          .filter(row => row[1] && row[2] && row[3] && row[7] && !existingIds.has(row[2]));
        logMessage('INFO', `Fetched Sales!A${rowStart}:J${rowEnd}, ${newData.length} valid rows out of ${rowEnd - rowStart + 1}`);
        if (newData.length > 0) {
          validateCustomers('Sales', 'C', range, rowStart, rowEnd - rowStart + 1);
        }
      } else if (sheetName === 'Payments') {
        const range = sheets.payments.getRange(`A${rowStart}:F${rowEnd}`);
        newData = range.getValues()
          .map(row => [
            'Payment', normalizeDate(row[0]), row[1], row[2], 0, 0, 0, 0, normalizePaymentMethod(row[3]), Number(row[4]) || 0, row[5] || ''
          ])
          .filter(row => row[1] && row[2] && row[3] && row[9] && !existingIds.has(row[2]));
        logMessage('INFO', `Fetched Payments!A${rowStart}:F${rowEnd}, ${newData.length} valid rows out of ${rowEnd - rowStart + 1}`);
        if (newData.length > 0) {
          validateCustomers('Payments', 'C', range, rowStart, rowEnd - rowStart + 1);
        }
      }

      if (newData.length === 0) {
        logMessage('INFO', `updateTransactions completed, no valid new data, took ${new Date().getTime() - startTime}ms`);
        batchCacheLastRows(sheetLastRows);
        return;
      }

      const nextRow = lastRows.transactions > 1 ? lastRows.transactions + 1 : 2;
      sheets.transactions.getRange(nextRow, 1, newData.length, 11).setValues(newData);
      logMessage('INFO', `Appended ${newData.length} rows to Transactions!A${nextRow}:K, sample: ${JSON.stringify(newData.slice(0, 5))}`);
    }

    batchCacheLastRows(sheetLastRows);
    logMessage('INFO', `updateTransactions completed, took ${new Date().getTime() - startTime}ms`);
  } catch (e) {
    logMessage('ERROR', `updateTransactions failed: ${e.message}`);
    throw e;
  }
}

function updateBalances(incrementalCustomers = null) {
  const startTime = new Date().getTime();
  logMessage('INFO', `updateBalances started, incrementalCustomers: ${incrementalCustomers ? incrementalCustomers.join(', ') : 'full'}`);
  try {
    const sheets = {
      balances: spreadsheet.getSheetByName('Balances'),
      customers: spreadsheet.getSheetByName('Customers'),
      sales: spreadsheet.getSheetByName('Sales'),
      payments: spreadsheet.getSheetByName('Payments')
    };
    if (!sheets.balances || !sheets.customers || !sheets.sales || !sheets.payments) throw new Error('Required sheets not found');

    const sheetLastRows = {};
    const lastRows = {
      customers: getLastContentRow(sheets.customers, [2]),
      sales: getLastContentRow(sheets.sales, [3]),
      payments: getLastContentRow(sheets.payments, [3, 5]),
      balances: getLastContentRow(sheets.balances, [1])
    };
    sheetLastRows['Customers'] = { lastRow: lastRows.customers, columns: [2] };
    sheetLastRows['Sales'] = { lastRow: lastRows.sales, columns: [3] };
    sheetLastRows['Payments'] = { lastRow: lastRows.payments, columns: [3, 5] };
    sheetLastRows['Balances'] = { lastRow: lastRows.balances, columns: [1] };

    if (lastRows.customers <= 1) {
      logMessage('WARN', 'No customer data found in Customers!B2:B');
      if (lastRows.balances > 1) {
        Utils.clearRange(sheets.balances, 2, 1, lastRows.balances - 1, 4);
      }
      batchCacheLastRows(sheetLastRows);
      return;
    }
    let customers = sheets.customers.getRange(`B2:B${lastRows.customers}`).getValues()
      .map((name, idx) => ({ name: name[0], row: idx + 2 }))
      .filter(c => c.name);
    if (incrementalCustomers) {
      customers = customers.filter(c => incrementalCustomers.includes(c.name));
    }
    if (customers.length === 0) {
      logMessage('INFO', `updateBalances completed, no customers to process, took ${new Date().getTime() - startTime}ms`);
      batchCacheLastRows(sheetLastRows);
      return;
    }

    const salesData = lastRows.sales > 1 ? sheets.sales.getRange(`C2:J${lastRows.sales}`).getValues() : [];
    const paymentsData = lastRows.payments > 1 ? sheets.payments.getRange(`C2:F${lastRows.payments}`).getValues() : [];

    const balances = customers.map(customer => {
      const validSales = salesData.filter(row => row[0] === customer.name && Number(row[4]) > 0 && row[1]);
      const validPayments = paymentsData.filter(row => row[0] === customer.name && Number(row[2]) > 0 && row[1]);
      const totalSales = validSales.reduce((sum, row) => sum + Number(row[4]), 0);
      const salesPayments = validSales.reduce((sum, row) => sum + Number(row[6]), 0);
      const paymentReceived = validPayments.reduce((sum, row) => sum + Number(row[2]), 0);
      const totalPayments = salesPayments + paymentReceived;
      const pendingBalance = totalSales - totalPayments;
      logMessage('INFO', `Balance for ${customer.name}: Sales=${totalSales}, Payments=${totalPayments}, Balance=${pendingBalance}`);
      return [customer.name, totalSales, totalPayments, pendingBalance];
    });

    if (incrementalCustomers) {
      balances.forEach(([name, sales, payments, balance]) => {
        const balancesRow = sheets.balances.getRange(`A2:A${lastRows.balances}`).getValues()
          .findIndex(row => row[0] === name) + 2;
        if (balancesRow >= 2) {
          sheets.balances.getRange(balancesRow, 2, 1, 3).setValues([[sales, payments, balance]]);
        } else {
          sheets.balances.getRange(lastRows.balances + 1, 1, 1, 4).setValues([[name, sales, payments, balance]]);
          lastRows.balances++;
        }
      });
    } else {
      if (lastRows.balances > 1) {
        Utils.clearRange(sheets.balances, 2, 1, lastRows.balances - 1, 4);
      }
      if (balances.length > 0) {
        sheets.balances.getRange(2, 1, balances.length, 4).setValues(balances);
        lastRows.balances = 2 + balances.length - 1;
      }
    }

    if (lastRows.balances > 1) {
      sheets.balances.getRange(`B2:D${lastRows.balances}`).setNumberFormat('#,##0.00');
    }
    logMessage('INFO', `Wrote ${balances.length} rows to Balances!A2:D, sample: ${JSON.stringify(balances.slice(0, 5))}`);
    batchCacheLastRows(sheetLastRows);
    logMessage('INFO', `updateBalances completed, took ${new Date().getTime() - startTime}ms`);
  } catch (e) {
    logMessage('ERROR', `updateBalances failed: ${e.message}`);
    throw e;
  }
}

function updateTransactionsAndBalances(suppressUi = false) {
  const startTime = new Date().getTime();
  logMessage('INFO', `updateTransactionsAndBalances started, suppressUi: ${suppressUi}`);
  try {
    updateTransactions();
    updateBalances();
    logMessage('INFO', `updateTransactionsAndBalances completed, took ${new Date().getTime() - startTime}ms`);
    if (!suppressUi && isUiAvailable()) {
      SpreadsheetApp.getUi().alert('Success', 'Transactions and Balances updated!', SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } catch (e) {
    logMessage('ERROR', `updateTransactionsAndBalances failed: ${e.message}`);
    if (!suppressUi && isUiAvailable()) {
      SpreadsheetApp.getUi().alert('Error', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    throw e;
  }
}

function createRequiredSheets(isTestMode = false) {
  const startTime = new Date().getTime();
  logMessage('INFO', `createRequiredSheets started, isTestMode: ${isTestMode}`);
  try {
    spreadsheet.toast('Creating sheets...', 'Progress', 5);
    removeAllSheets(false);

    logMessage('DEBUG', 'Attempting to clear TransactionIds cache');
    try {
      PropertiesService.getScriptProperties().deleteProperty('TransactionIds');
      logMessage('INFO', 'Cleared TransactionIds cache');
    } catch (e) {
      logMessage('ERROR', `Failed to clear TransactionIds cache: ${e.message}`);
    }

    const sheetsToCreate = SHEET_CONFIG.map(sheetInfo => ({ ...sheetInfo, sheet: null }));
    const sheets = {};
    sheetsToCreate.forEach(sheetInfo => {
      let sheet = spreadsheet.getSheetByName(sheetInfo.name);
      if (!sheet) {
        sheet = spreadsheet.insertSheet(sheetInfo.name);
        logMessage('INFO', `Inserted new sheet: ${sheetInfo.name}`);
      }
      sheets[sheetInfo.name] = sheet;
      sheetInfo.sheet = sheet;
    });

    const sheetLastRows = {};
    const paymentMethods = ['Bank Transfer', 'Cheque', 'Cash', 'UPI', 'Others'];
    sheetsToCreate.forEach(sheetInfo => {
      const sheet = sheetInfo.sheet;
      const headerRange = sheet.getRange(1, 1, 1, sheetInfo.headers.length);
      headerRange.setValues([sheetInfo.headers])
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      logMessage('INFO', `Set headers for ${sheetInfo.name}: ${sheetInfo.headers.join(', ')}`);
      sheet.setColumnWidths(1, sheetInfo.headers.length, 120);

      if (sheetInfo.name === 'Sales') {
        const range = sheet.getRange('A2:I' + CONFIG.MAX_ROWS);
        range.setNumberFormats(range.getValues().map(row => ['dd/MM/yyyy', '', '', '0.00', '0.00', '0.00', '0.00', '', '0.00']));
        const salesPaymentRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(paymentMethods)
          .setAllowInvalid(false)
          .build();
        sheet.getRange('H2:H' + CONFIG.MAX_ROWS).setDataValidation(salesPaymentRule);
        sheetLastRows['Sales'] = { lastRow: 1, columns: [3] };
        logMessage('INFO', `Formatted Sales!A2:I${CONFIG.MAX_ROWS}, set H2:H${CONFIG.MAX_ROWS} data validation`);
        Utilities.sleep(50);
      } else if (sheetInfo.name === 'Payments') {
        const range = sheet.getRange('A2:E' + CONFIG.MAX_ROWS);
        range.setNumberFormats(range.getValues().map(row => ['dd/MM/yyyy', '', '', '', '0.00']));
        const paymentsPaymentRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(paymentMethods)
          .setAllowInvalid(false)
          .build();
        sheet.getRange('D2:D' + CONFIG.MAX_ROWS).setDataValidation(paymentsPaymentRule);
        sheetLastRows['Payments'] = { lastRow: 1, columns: [3, 5] };
        logMessage('INFO', `Formatted Payments!A2:E${CONFIG.MAX_ROWS}, set D2:D${CONFIG.MAX_ROWS} data validation`);
        Utilities.sleep(50);
      } else if (sheetInfo.name === 'Balances') {
        sheet.getRange('B2:D' + CONFIG.MAX_ROWS).setNumberFormat('#,##0.00');
        sheetLastRows['Balances'] = { lastRow: 1, columns: [1] };
        logMessage('INFO', `Formatted Balances!B2:D${CONFIG.MAX_ROWS}`);
      } else if (sheetInfo.name === 'Transactions') {
        const range = sheet.getRange('B2:J' + CONFIG.MAX_ROWS);
        range.setNumberFormats(range.getValues().map(row => ['dd/MM/yyyy', '', '', '0.00', '0.00', '0.00', '0.00', '', '0.00']));
        sheetLastRows['Transactions'] = { lastRow: 1, columns: [3] };
        logMessage('INFO', `Formatted Transactions!B2:J${CONFIG.MAX_ROWS}`);
        Utilities.sleep(50);
      } else if (sheetInfo.name === 'Logs') {
        sheet.getRange(1, 1, 1, 2).setValues([['Timestamp', 'Message']])
          .setFontWeight('bold')
          .setHorizontalAlignment('center');
        sheet.setColumnWidths(1, 2, 200);
        sheetLastRows['Logs'] = { lastRow: 1, columns: [2] };
        logMessage('INFO', 'Sheet Logs setup completed');
      } else if (sheetInfo.name === 'Customers') {
        sheetLastRows['Customers'] = { lastRow: 1, columns: [2] };
      }
      logMessage('INFO', `Sheet ${sheetInfo.name} setup completed`);
    });

    const customersSheet = sheets['Customers'];
    const validationRange = customersSheet.getRange(`B2:B${CONFIG.MAX_ROWS}`);
    const customerRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(validationRange)
      .setAllowInvalid(false)
      .build();
    sheets['Sales'].getRange('C2:C' + CONFIG.MAX_ROWS).setDataValidation(customerRule);
    sheets['Payments'].getRange('C2:C' + CONFIG.MAX_ROWS).setDataValidation(customerRule);
    logMessage('INFO', `Set Sales!C2:C${CONFIG.MAX_ROWS} and Payments!C2:C${CONFIG.MAX_ROWS} data validation`);
    Utilities.sleep(50);

    const transactionsSheet = sheets['Transactions'];
    const transLastRow = transactionsSheet.getLastRow();
    if (transLastRow > 1) {
      Utils.clearRange(transactionsSheet, 2, 1, transLastRow - 1, 11);
    }
    logMessage('INFO', `Setting Transactions initial state`);

    const protection = transactionsSheet.protect().setDescription('Prevent manual edits');
    protection.removeEditors(protection.getEditors());
    if (protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }
    logMessage('INFO', `Protected Transactions sheet`);

    formatDateColumns();
    batchCacheLastRows(sheetLastRows);

    const defaultSheet = spreadsheet.getSheetByName('Sheet1');
    if (defaultSheet) {
      spreadsheet.deleteSheet(defaultSheet);
      logMessage('INFO', 'Deleted default Sheet1');
    }

    logMessage('INFO', `createRequiredSheets completed, took ${new Date().getTime() - startTime}ms`);
    if (!isTestMode && isUiAvailable()) {
      SpreadsheetApp.getUi().alert('Success', 'Required sheets created!', SpreadsheetApp.getUi().ButtonSet.OK);
    }
    spreadsheet.toast('Sheets created successfully', 'Success', 5);
  } catch (e) {
    logMessage('ERROR', `createRequiredSheets failed: ${e.message}`);
    spreadsheet.toast('Failed to create sheets', 'Error', 5);
    throw e;
  }
}

function testCreateRequiredSheets() {
  const startTime = new Date().getTime();
  logMessage('INFO', 'testCreateRequiredSheets started');
  try {
    spreadsheet.toast('Running tests...', 'Progress', 5); // Fixed syntax error
    createRequiredSheets(true);

    let errors = [];
    SHEET_CONFIG.forEach(sheetInfo => {
      const sheet = spreadsheet.getSheetByName(sheetInfo.name);
      if (!sheet) {
        errors.push(`Sheet ${sheetInfo.name} not created.`);
        return;
      }
      const headers = sheet.getRange(1, 1, 1, sheetInfo.headers.length).getValues()[0];
      if (JSON.stringify(headers) !== JSON.stringify(sheetInfo.headers)) {
        errors.push(`Headers for ${sheetInfo.name} incorrect. Expected: ${sheetInfo.headers.join(', ')}, Got: ${headers.join(', ')}`);
      }
      const headerRange = sheet.getRange(1, 1, 1, sheetInfo.headers.length);
      const isBold = headerRange.getFontWeights()[0].every(weight => weight === 'bold');
      const isCentered = headerRange.getHorizontalAlignments()[0].every(align => align === 'center');
      if (!isBold) errors.push(`Headers in ${sheetInfo.name} not bold.`);
      if (!isCentered) errors.push(`Headers in ${sheetInfo.name} not centered.`);
      for (let i = 1; i <= sheetInfo.headers.length; i++) {
        const columnWidth = sheet.getColumnWidth(i);
        if (columnWidth < 120) {
          errors.push(`Column ${i} in ${sheetInfo.name} width too small: ${columnWidth}.`);
        }
      }
      if (sheetInfo.name !== 'Logs') {
        const row2 = sheet.getRange(2, 1, 1, sheetInfo.headers.length).getValues()[0];
        if (row2.some(cell => cell !== '' && cell != null)) {
          errors.push(`Row 2 in ${sheetInfo.name} not blank: ${row2.join(', ')}`);
        }
      }
    });

    if (spreadsheet.getSheetByName('Sheet1')) {
      errors.push('Sheet1 not deleted.');
    }

    if (errors.length > 0) {
      const errorMessage = 'Blank sheet setup failed:\n' + errors.join('\n');
      logMessage('ERROR', errorMessage);
      throw new Error(errorMessage);
    }

    const customersSheet = spreadsheet.getSheetByName('Customers');
    customersSheet.getRange('A2:B3').setValues([
      ['CUST0001', 'John Doe'],
      ['CUST0002', 'Jane Smith']
    ]);
    invalidateCustomerCache();
    invalidateLastRowCache('Customers', [2]);
    logMessage('INFO', `Set Customers!A2:B3`);

    const salesSheet = spreadsheet.getSheetByName('Sales');
    salesSheet.getRange('A2:J3').setValues([
      ['01/06/2025', 'SALE0001', 'John Doe', 100, 50, 1000, 6000, 'Others', 2000, 'First sale'],
      ['02/06/2025', 'SALE0002', 'Jane Smith', 200, 45, 0, 9000, 'UPI', 5000, '']
    ]);
    salesSheet.getRange('D2:I3').setNumberFormat('0.00');
    invalidateLastRowCache('Sales', [3]);
    logMessage('INFO', `Set Sales!A2:J3`);

    const paymentsSheet = spreadsheet.getSheetByName('Payments');
    paymentsSheet.getRange('A2:F4').setValues([
      ['01/06/2025', 'PAY0001', 'John Doe', 'Bank Transfer', 3000, 'Payment for SALE0001'],
      ['03/06/2025', 'PAY0002', 'Jane Smith', 'Cheque', 4000, ''],
      ['04/06/2025', 'PAY0003', 'John Doe', 'Cash', 1000, 'Manual']
    ]);
    paymentsSheet.getRange('E2:E4').setNumberFormat('0.00');
    invalidateLastRowCache('Payments', [3, 5]);
    logMessage('INFO', `Set Payments!A2:F4`);

    updateTransactions();
    updateBalances();

    const transactionsSheet = spreadsheet.getSheetByName('Transactions');
    const transData = transactionsSheet.getRange('A2:K6').getValues().map(row => {
      row[1] = normalizeDate(row[1]);
      return row;
    });
    const expectedTransCount = 5;
    if (transData.length < expectedTransCount) {
      errors.push(`Transactions sheet has fewer rows than expected. Expected: ${expectedTransCount}, Got: ${transData.length}.`);
    } else {
      const expectedTransactions = [
        { id: 'SALE0001', type: 'Sale', date: '01/06/2025', customer: 'John Doe', quantity: 100, rate: 50, vehicleRent: 1000, amount: 6000, paymentMethod: 'Others', paymentReceived: 2000, remarks: 'First sale' },
        { id: 'SALE0002', type: 'Sale', date: '02/06/2025', customer: 'Jane Smith', quantity: 200, rate: 45, vehicleRent: 0, amount: 9000, paymentMethod: 'UPI', paymentReceived: 5000, remarks: '' },
        { id: 'PAY0001', type: 'Payment', date: '01/06/2025', customer: 'John Doe', quantity: 0, rate: 0, vehicleRent: 0, amount: 0, paymentMethod: 'Bank Transfer', paymentReceived: 3000, remarks: 'Payment for SALE0001' },
        { id: 'PAY0002', type: 'Payment', date: '03/06/2025', customer: 'Jane Smith', quantity: 0, rate: 0, vehicleRent: 0, amount: 0, paymentMethod: 'Cheque', paymentReceived: 4000, remarks: '' },
        { id: 'PAY0003', type: 'Payment', date: '04/06/2025', customer: 'John Doe', quantity: 0, rate: 0, vehicleRent: 0, amount: 0, paymentMethod: 'Cash', paymentReceived: 1000, remarks: 'Manual' }
      ];
      transData.forEach(row => {
        const trans = row;
        const expected = expectedTransactions.find(item => item.id === trans[2]);
        if (!expected) {
          errors.push(`Transaction ${trans[2]} not found.`);
          return;
        }
        const transDate = normalizeDate(trans[1]);
        if (trans[0] !== expected.type) errors.push(`Transaction ${expected.id} type incorrect. Expected: ${expected.type}, Got: ${trans[0]}`);
        if (transDate !== expected.date) errors.push(`Transaction ${expected.id} date incorrect. Expected: ${expected.date}, Got: ${transDate}`);
        if (trans[3] !== expected.customer) errors.push(`Transaction ${expected.id} customer incorrect. Expected: ${expected.customer}, Got: ${trans[3]}`);
        if (Number(trans[4]) !== expected.quantity) errors.push(`Transaction ${expected.id} quantity incorrect. Expected: ${expected.quantity}, Got: ${trans[4]}`);
        if (Number(trans[5]) !== expected.rate) errors.push(`Transaction ${expected.id} rate incorrect. Expected: ${expected.rate}, Got: ${trans[5]}`);
        if (Number(trans[6]) !== expected.vehicleRent) errors.push(`Transaction ${expected.id} vehicle rent incorrect. Expected: ${expected.vehicleRent}, Got: ${trans[6]}`);
        if (Number(trans[7]) !== expected.amount) errors.push(`Transaction ${expected.id} amount incorrect. Expected: ${expected.amount}, Got: ${trans[7]}`);
        if (trans[8] !== expected.paymentMethod) errors.push(`Transaction ${expected.id} payment method incorrect. Expected: ${expected.paymentMethod}, Got: ${trans[8]}`);
        if (Number(trans[9]) !== expected.paymentReceived) errors.push(`Transaction ${expected.id} payment received incorrect. Expected: ${expected.paymentReceived}, Got: ${trans[9]}`);
        if (trans[10] !== expected.remarks) errors.push(`Transaction ${expected.id} remarks incorrect. Expected: ${expected.remarks}, Got: ${trans[10]}`);
      });
    }
    logMessage('INFO', `Transactions data: ${JSON.stringify(transData)}`);

    const balancesSheet = spreadsheet.getSheetByName('Balances');
    const balancesData = balancesSheet.getRange('A2:D4').getValues();
    logMessage('INFO', `Balances data: ${JSON.stringify(balancesData)}`);

    const expectedBalances = [
      { customer: 'John Doe', sales: 6000, payments: 6000, balance: 0 },
      { customer: 'Jane Smith', sales: 9000, payments: 9000, balance: 0 }
    ];
    expectedBalances.forEach(expected => {
      const balance = balancesData.find(row => row[0] === expected.customer);
      if (!balance) {
        errors.push(`Balance for ${expected.customer} not found.`);
      } else {
        if (Number(balance[1]) !== expected.sales) errors.push(`Incorrect sales for ${expected.customer}. Expected: ${expected.sales}, Got: ${balance[1]}`);
        if (Number(balance[2]) !== expected.payments) errors.push(`Incorrect payments for ${expected.customer}. Expected: ${expected.payments}, Got: ${balance[2]}`);
        if (Number(balance[3]) !== expected.balance) errors.push(`Incorrect balance for ${expected.customer}. Expected: ${expected.balance}, Got: ${balance[3]}`);
      }
    });

    logMessage('INFO', `Sales!C2:C3: ${JSON.stringify(salesSheet.getRange('C2:C3').getValues())}`);

    if (errors.length === 0) {
      logMessage('INFO', 'Tests passed successfully');
      spreadsheet.toast('Tests passed successfully', 'Success', 5);
    } else {
      const errorMessage = 'Test failed:\n' + errors.join('\n');
      logMessage('ERROR', errorMessage);
      spreadsheet.toast('Test failed!', 'Error', 5);
      throw new Error(errorMessage);
    }
  } catch (e) {
    logMessage('ERROR', `testCreateRequiredSheets failed: ${e.message}`);
    spreadsheet.toast('Test failed', 'Error', 5);
    throw e;
  }
}

function processFormSubmission(sheetName, row, suppressUi = true) {
  const startTime = new Date().getTime();
  logMessage('INFO', `processFormSubmission started for ${sheetName}, row: ${row}, suppressUi: ${suppressUi}`);
  try {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) throw new Error(`Sheet ${sheetName} not found`);
    const range = sheet.getRange(row, 1, 1, sheetName === 'Sales' ? 10 : 6);
    updateTransactions(range);
    const customer = sheet.getRange(row, 3).getValue();
    if (customer) updateBalances([customer]);
    logMessage('INFO', `processFormSubmission completed, took ${new Date().getTime() - startTime}ms`);
  } catch (e) {
    logMessage('ERROR', `processFormSubmission failed: ${e.message}`);
    if (!suppressUi && isUiAvailable()) {
      SpreadsheetApp.getUi().alert('Error', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    throw e;
  }
}

function onOpen() {
  try {
    logMessage('INFO', 'Script version: 1.2');
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('BrickSync')
      .addItem('Create Required Sheets', 'createRequiredSheets')
      .addItem('Remove All Sheets', 'removeAllSheets')
      .addItem('Test Sheet Setup', 'testCreateRequiredSheets')
      .addItem('Update Transactions and Balances', 'updateTransactionsAndBalances')
      .addToUi();
  } catch (e) {
    logMessage('ERROR', `onOpen failed: ${e.message}`);
    throw e;
  }
}

function onEdit(e) {
  try {
    const sheet = e.source.getActiveSheet();
    const sheetName = sheet.getName();
    const range = e.range;
    const colStart = range.getColumn();
    const rowStart = range.getRow();
    if (rowStart < 2) return;

    if (sheetName === 'Customers' && colStart === 2) {
      invalidateCustomerCache();
      invalidateLastRowCache(sheetName, [2]);
    } else if (sheetName === 'Sales' && colStart <= 3) {
      invalidateLastRowCache(sheetName, [3]);
    } else if (sheetName === 'Payments' && (colStart <= 3 || colStart === 5)) {
      invalidateLastRowCache(sheetName, [3, 5]);
    } else if (sheetName === 'Transactions' && colStart === 3) {
      invalidateLastRowCache(sheetName, [3]);
      logMessage('DEBUG', 'Attempting to clear TransactionIds cache on Transactions edit');
      try {
        PropertiesService.getScriptProperties().deleteProperty('TransactionIds');
        logMessage('INFO', 'Cleared TransactionIds cache on Transactions edit');
      } catch (err) {
        logMessage('ERROR', `Failed to clear TransactionIds cache on Transactions edit: ${err.message}`);
      }
    } else {
      return;
    }

    for (let row = rowStart; row <= rowStart + range.getNumRows() - 1; row++) {
      const dateFormat = sheet.getRange(`A${row}`).getNumberFormat();
      if (dateFormat !== CONFIG.DATE_FORMAT) {
        if (sheetName === 'Sales') {
          sheet.getRange(`A${row}:I${row}`).setNumberFormat(['dd/MM/yyyy', '', '', '0.00', '0.00', '0.00', '0.00', '', '0.00'].join(';'));
        } else if (sheetName === 'Payments') {
          sheet.getRange(`A${row}:E${row}`).setNumberFormat(['dd/MM/yyyy', '', '', '', '0.00'].join(';'));
        }
      }
    }
  } catch (e) {
    logMessage('ERROR', `onEdit failed: ${e.message}`);
    if (isUiAvailable()) SpreadsheetApp.getUi().alert('Error', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    throw e;
  }
}

function removeAllSheets(showPopup = true) {
  const startTime = new Date().getTime();
  try {
    spreadsheet.toast('Removing all sheets...', 'Progress', 5);
    let sheets = spreadsheet.getSheets();
    let tempSheet = spreadsheet.getSheetByName('TempSheet');
    if (tempSheet) {
      try {
        spreadsheet.deleteSheet(tempSheet);
      } catch (e) {}
    }
    spreadsheet.insertSheet('TempSheet');
    tempSheet = spreadsheet.getSheetByName('TempSheet');
    if (!tempSheet) throw new Error('Failed to create TempSheet');
    for (let i = 0; i < sheets.length; i++) {
      const sheet = sheets[i];
      if (sheet.getName() !== 'TempSheet') {
        let attempts = 0;
        while (attempts < 3) {
          try {
            spreadsheet.deleteSheet(sheet);
            break;
          } catch (e) {
            attempts++;
            if (attempts === 3) throw e;
            Utilities.sleep(100);
          }
        }
      }
    }
    tempSheet = spreadsheet.getSheetByName('TempSheet');
    if (tempSheet) {
      tempSheet.setName('Sheet1');
    }
    logMessage('INFO', 'Deleted default sheet');
    if (showPopup && isUiAvailable()) {
      SpreadsheetApp.getUi().alert('Success', 'All sheets removed successfully!', SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } catch (e) {
    Logger.log(`removeAllSheets failed: ${e.message}`);
    tempSheet = spreadsheet.getSheetByName('TempSheet');
    if (tempSheet) {
      try {
        spreadsheet.deleteSheet(tempSheet);
      } catch (err) {}
    }
    spreadsheet.toast('Failed to remove sheets', 'Error', 5);
    throw e;
  }
}

function isUiAvailable(suppressUi = false) {
  const startTime = new Date().getTime();
  logMessage('INFO', `isUiAvailable started, suppressUi: ${suppressUi}`);
  if (suppressUi) {
    logMessage('INFO', `isUiAvailable completed, result: false, took ${new Date().getTime() - startTime}ms (suppressed)`);
    return false;
  }
  try {
    SpreadsheetApp.getUi();
    logMessage('INFO', `isUiAvailable completed, result: true, took ${new Date().getTime() - startTime}ms`);
    return true;
  } catch (e) {
    logMessage('INFO', `isUiAvailable completed, result: false, took ${new Date().getTime() - startTime}ms`);
    return false;
  }
}
