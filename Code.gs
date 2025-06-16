let isRemovingSheets = false;

const VALID_PAYMENT_METHODS = ['Cash', 'UPI', 'Bank Transfer', 'Cheque', 'Others'];

function formatDate(dateStr) {
  try {
    let date;
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(dateStr)) {
      const parts = dateStr.split('/');
      const day = parseInt(parts[0], 10);
      const month = parseInt(parts[1], 10) - 1;
      const year = parseInt(parts[2], 10);
      date = new Date(year, month, day);
    } else if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
      date = new Date(dateStr);
    } else {
      throw new Error('Invalid date format. Use dd/mm/yyyy or yyyy-mm-dd.');
    }
    if (isNaN(date.getTime()) || date.getDate() !== parseInt(dateStr.split(/\/|-/)[0], 10)) {
      throw new Error('Invalid date.');
    }
    const day = ('0' + date.getDate()).slice(-2);
    const month = ('0' + (date.getMonth() + 1)).slice(-2);
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
  } catch (e) {
    logMessage('ERROR', `formatDate failed: ${e.message} Input: ${dateStr}`);
    throw e;
  }
}

function doGet(e) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().setSpreadsheetLocale('en_IN');
    const template = HtmlService.createTemplateFromFile('index');
    return template.evaluate()
      .setTitle('BrickSync')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .addMetaTag('Content-Security-Policy', "default-src 'self'; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline'; connect-src 'self'");
  } catch (e) {
    logMessage('ERROR', `doGet failed: ${e.message}`);
    return HtmlService.createHtmlOutput(`<h3>Error</h3><p>${e.message}</p>`);
  }
}

function doPost(e) {
  try {
    if (!e.postData || !e.postData.contents) {
      throw new Error('No data provided in POST request');
    }
    const params = JSON.parse(e.postData.contents);
    const action = params.action || 'unknown';
    let response;
    switch (action) {
      case 'getData':
        response = getSheetData(params.sheet);
        break;
      case 'addData':
        response = addSheetData(params.sheet, params.data);
        break;
      default:
        throw new Error(`Unknown action: ${action}`);
    }
    return ContentService.createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    logMessage('ERROR', `doPost failed: ${e.message}`);
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: e.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function include(filename) {
  try {
    return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
  } catch (e) {
    logMessage('ERROR', `Include failed for ${filename}: ${e.message}`);
    return `<p>Error loading ${filename}</p>`;
  }
}

function getSheetSafely(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet ${sheetName} not found`);
  return sheet;
}

function getLastContentRow(sheet, columns) {
  if (!sheet) return 0;
  columns = columns || ['A'];
  let maxLastRow = 0;
  const maxRows = Math.min(sheet.getMaxRows(), CONFIG.MAX_ROWS);
  for (const column of columns.slice(0, 3)) {
    const values = sheet.getRange(column + '2:' + column + maxRows).getValues();
    let lastRow = 0;
    for (let i = 0; i < values.length; i++) {
      if (values[i][0] !== '' && values[i][0] != null) lastRow = i + 2;
    }
    maxLastRow = Math.max(maxLastRow, lastRow);
  }
  return maxLastRow;
}

function getSheetData(sheetName) {
  try {
    logMessage('INFO', `Fetching data from ${sheetName}`);
    const sheet = getSheetSafely(sheetName);
    const lastRow = getLastContentRow(sheet, ['A', 'B', 'C']);
    if (lastRow < 2) {
      logMessage('INFO', `No data found in ${sheetName} (last row < 2)`);
      return { success: true, data: [] };
    }
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data = lastRow > 2 ? sheet.getRange(2, 1, lastRow - 1, headers.length).getValues() : [];
    const records = data.map(row => {
      const record = {};
      headers.forEach((header, i) => {
        record[header] = row[i] !== null && row[i] !== undefined ? row[i] : '';
      });
      return record;
    });
    logMessage('INFO', `Fetched ${records.length} records from ${sheetName}`);
    return { success: true, data: records };
  } catch (e) {
    logMessage('ERROR', `getSheetData failed for ${sheetName}: ${e.message}`);
    if (e.message.includes('not found')) {
      createRequiredSheets();
      return { success: false, message: `Sheet ${sheetName} was missing and has been recreated. Please try again.` };
    }
    return { success: false, message: e.message };
  }
}

function addSheetData(sheetName, data) {
  try {
    if (!data) throw new Error('No data provided');
    const isBatch = Array.isArray(data);
    const records = isBatch ? data : [data];
    logMessage('INFO', `Adding ${records.length} record(s) to ${sheetName}`);
    switch (sheetName) {
      case 'Customers': return batchAddCustomers(records);
      case 'Sales': return batchAddSales(records);
      case 'Payments': return batchAddPayments(records);
      default: throw new Error(`Cannot write to ${sheetName}`);
    }
  } catch (e) {
    logMessage('ERROR', `addSheetData failed for ${sheetName}: ${e.message}`);
    return { success: false, message: e.message };
  }
}

function getCustomers() { return getSheetData('Customers'); }
function getSales() { return getSheetData('Sales'); }
function getPayments() { return getSheetData('Payments'); }
function getTransactions() { return getSheetData('Transactions'); }
function getBalances() {
  try {
    logMessage('INFO', 'Fetching data from Balances');
    const sheet = getSheetSafely('Balances');
    const lastRow = getLastContentRow(sheet, ['A']);
    if (lastRow < 2) {
      logMessage('INFO', 'No data found in Balances (last row < 2)');
      return { success: true, data: [] };
    }
    const headers = ['Customer', 'Total Sales', 'Total Payments', 'Pending Balance'];
    const data = lastRow > 2 ? sheet.getRange(2, 1, lastRow - 1, headers.length).getValues() : [];
    const records = data.map(row => ({
      customer: row[0] || '',
      sales: parseFloat(row[1]) || 0,
      payments: parseFloat(row[2]) || 0,
      balance: parseFloat(row[3]) || 0
    }));
    logMessage('INFO', `Fetched ${records.length} records from Balances`);
    return { success: true, data: records };
  } catch (e) {
    logMessage('ERROR', `getBalances failed: ${e.message}`);
    if (e.message.includes('not found')) {
      createRequiredSheets();
      return { success: false, message: 'Sheet Balances was missing and has been recreated. Please try again.' };
    }
    return { success: false, message: e.message };
  }
}

function addCustomer(customerData) {
  try {
    if (!customerData.name) throw new Error('Missing required field: name');
    const sheet = getSheetSafely('Customers');
    const lastRow = getLastContentRow(sheet, ['A']) + 1;
    const customerId = 'CUST' + lastRow.toString().padStart(4, '0');
    sheet.getRange('A' + lastRow + ':B' + lastRow).setValues([[customerId, customerData.name]]);
    SpreadsheetApp.flush();
    logMessage('INFO', `Added customer ${customerData.name} at A${lastRow}`);
    return { success: true, message: 'Customer added', customerId: customerId };
  } catch (e) {
    logMessage('ERROR', `addCustomer failed: ${e.message}`);
    return { success: false, message: e.message };
  }
}

function batchAddCustomers(customers) {
  try {
    if (!customers.length) throw new Error('No customers provided');
    const sheet = getSheetSafely('Customers');
    const lastRow = getLastContentRow(sheet, ['A']) + 1;
    const values = customers.map((customer, i) => {
      if (!customer.name) throw new Error(`Missing name at index ${i}`);
      return ['CUST' + (lastRow + i).toString().padStart(4, '0'), customer.name];
    });
    sheet.getRange('A' + lastRow + ':B' + (lastRow + values.length - 1)).setValues(values);
    SpreadsheetApp.flush();
    logMessage('INFO', `Added ${values.length} customers at A${lastRow}`);
    return { success: true, message: `${values.length} customers added` };
  } catch (e) {
    logMessage('ERROR', `batchAddCustomers failed: ${e.message}`);
    return { success: false, message: e.message };
  }
}

function deleteCustomer(customerData) {
  try {
    if (!customerData.customerId) throw new Error('Missing required field: customerId');
    const sheet = getSheetSafely('Customers');
    const data = sheet.getRange('A2:A' + sheet.getLastRow()).getValues();
    let rowIndex = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === customerData.customerId) {
        rowIndex = i + 2;
        break;
      }
    }
    if (rowIndex === -1) throw new Error(`Customer ID ${customerData.customerId} not found`);
    Utils.clearFullRow(sheet, rowIndex);
    sheet.deleteRow(rowIndex);
    SpreadsheetApp.flush();
    logMessage('INFO', `Deleted customer ${customerData.customerId} at row ${rowIndex}`);
    return { success: true, message: 'Customer deleted' };
  } catch (e) {
    logMessage('ERROR', `deleteCustomer failed: ${e.message}`);
    return { success: false, message: e.message };
  }
}

function updateCustomer(customerData) {
  try {
    if (!customerData.customerId || !customerData.name) {
      throw new Error('Missing required fields: customerId, name');
    }
    const sheet = getSheetSafely('Customers');
    const data = sheet.getRange('A2:B' + sheet.getLastRow()).getValues();
    let found = false;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === customerData.customerId) {
        sheet.getRange('B' + (i + 2)).setValue(customerData.name);
        found = true;
        break;
      }
    }
    if (!found) throw new Error(`Customer ID ${customerData.customerId} not found`);
    SpreadsheetApp.flush();
    logMessage('INFO', `Updated customer ${customerData.customerId}`);
    return { success: true, message: 'Customer updated' };
  } catch (e) {
    logMessage('ERROR', `updateCustomer failed: ${e.message}`);
    return { success: false, message: e.message };
  }
}

function addSale(saleData) {
  try {
    if (!saleData || !saleData.customer || !saleData.date) {
      throw new Error('Missing required fields: customer, date');
    }
    saleData.date = formatDate(saleData.date);
    if (saleData.quantity < 0 || saleData.rate < 0 || saleData.vehicleRent < 0 || saleData.paymentReceived < 0) {
      throw new Error('Negative values not allowed');
    }
    if (saleData.paymentMethod && !VALID_PAYMENT_METHODS.includes(saleData.paymentMethod)) {
      throw new Error(`Invalid payment method: ${saleData.paymentMethod}`);
    }
    const salesSheet = getSheetSafely('Sales');
    const customersSheet = getSheetSafely('Customers');
    let customerId = '';
    const customerRange = customersSheet.getRange('B2:B' + CONFIG.DROPDOWN_ROWS);
    const foundCell = customerRange.createTextFinder(saleData.customer)
      .matchEntireCell(true)
      .findNext();
    if (foundCell) {
      customerId = foundCell.offset(0, -1).getValue();
    } else {
      const lastRow = getLastContentRow(customersSheet, ['A']) + 1;
      customerId = 'CUST' + lastRow.toString().padStart(4, '0');
      customersSheet.getRange('A' + lastRow + ':B' + lastRow).setValues([[customerId, saleData.customer]]);
    }
    const lastRow = getLastContentRow(salesSheet, ['A']) + 1;
    salesSheet.getRange('A' + lastRow + ':J' + lastRow).setValues([[
      saleData.date,
      saleData.saleId || 'SALE' + lastRow.toString().padStart(4, '0'),
      saleData.customer,
      saleData.quantity || 0,
      saleData.rate || 0,
      saleData.vehicleRent || 0,
      saleData.amount || 0,
      saleData.paymentMethod || '',
      saleData.paymentReceived || 0,
      saleData.remarks || ''
    ]]);
    salesSheet.getRange('W' + lastRow).setValue(customerId);
    const formulas = [
      `=IF(A${lastRow}<>"","Sale","")`,
      `=K${lastRow}`, `=A${lastRow}`, `=B${lastRow}`, `=C${lastRow}`, `=D${lastRow}`, `=E${lastRow}`, `=F${lastRow}`, `=G${lastRow}`, `=H${lastRow}`, `=I${lastRow}`, `=J${lastRow}`
    ];
    salesSheet.getRange('K' + lastRow + ':V' + lastRow).setFormulas([formulas]);
    SpreadsheetApp.flush();
    logMessage('INFO', `Added sale at Sales!A${lastRow} Data: ${JSON.stringify(saleData)}`);
    return { success: true, message: 'Sale added' };
  } catch (e) {
    logMessage('ERROR', `addSale failed: ${e.message}`);
    return { success: false, message: e.message };
  }
}

function batchAddSales(sales) {
  try {
    if (!sales.length) throw new Error('No sales provided');
    const results = sales.map((sale, i) => {
      try {
        return addSale(sale);
      } catch (e) {
        logMessage('ERROR', `Sale at index ${i} failed: ${e.message}`);
        return { success: false, message: `Index ${i}: ${e.message}` };
      }
    });
    const successCount = results.filter(r => r.success).length;
    SpreadsheetApp.flush();
    logMessage('INFO', `Added ${successCount}/${sales.length} sales`);
    return { success: successCount === sales.length, message: `Added ${successCount}/${sales.length} sales`, results: results };
  } catch (e) {
    logMessage('ERROR', `batchAddSales failed: ${e.message}`);
    return { success: false, message: e.message };
  }
}

function deleteSale(saleId) {
  try {
    const sheet = getSheetSafely('Sales');
    const data = sheet.getRange('B2:B' + sheet.getLastRow()).getValues();
    let rowIndex = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === saleId) {
        rowIndex = i + 2;
        break;
      }
    }
    if (rowIndex === -1) throw new Error(`Sale ID ${saleId} not found`);
    Utils.clearFullRow(sheet, rowIndex);
    sheet.deleteRow(rowIndex);
    SpreadsheetApp.flush();
    logMessage('INFO', `Deleted sale ${saleId} at row ${rowIndex}`);
    return { success: true, message: 'Sale deleted' };
  } catch (e) {
    logMessage('ERROR', `deleteSale failed: ${e.message}`);
    return { success: false, message: e.message };
  }
}

function addPayment(paymentData) {
  try {
    if (!paymentData || !paymentData.customer || !paymentData.date) {
      throw new Error('Missing required fields: customer, date');
    }
    paymentData.date = formatDate(paymentData.date);
    if (paymentData.paymentReceived < 0) throw new Error('Negative payment not allowed');
    if (paymentData.paymentMethod && !VALID_PAYMENT_METHODS.includes(paymentData.paymentMethod)) {
      throw new Error(`Invalid payment method: ${paymentData.paymentMethod}`);
    }
    const paymentsSheet = getSheetSafely('Payments');
    const customersSheet = getSheetSafely('Customers');
    let customerId = '';
    const customerRange = customersSheet.getRange('B2:B' + CONFIG.DROPDOWN_ROWS);
    const foundCell = customerRange.createTextFinder(paymentData.customer)
      .matchEntireCell(true)
      .findNext();
    if (foundCell) {
      customerId = foundCell.offset(0, -1).getValue();
    } else {
      const lastRow = getLastContentRow(customersSheet, ['A']) + 1;
      customerId = 'CUST' + lastRow.toString().padStart(4, '0');
      customersSheet.getRange('A' + lastRow + ':B' + lastRow).setValues([[customerId, paymentData.customer]]);
    }
    const lastRow = getLastContentRow(paymentsSheet, ['A']) + 1;
    paymentsSheet.getRange('A' + lastRow + ':F' + lastRow).setValues([[
      paymentData.date,
      paymentData.paymentId || 'PAY' + lastRow.toString().padStart(4, '0'),
      paymentData.customer,
      paymentData.paymentMethod || '',
      paymentData.paymentReceived || 0,
      paymentData.remarks || ''
    ]]);
    paymentsSheet.getRange('H' + lastRow + ':K' + lastRow).setValues([[0, 0, 0, 0]]);
    const formulas = [
      `=IF(A${lastRow}<>"","Payment","")`, '', '', '', '',
      `=G${lastRow}`, `=A${lastRow}`, `=B${lastRow}`, `=C${lastRow}`,
      `=IF(A${lastRow}<>"",0,"")`, `=IF(A${lastRow}<>"",0,"")`, `=IF(A${lastRow}<>"",0,"")`, `=IF(A${lastRow}<>"",0,"")`,
      `=D${lastRow}`, `=E${lastRow}`, `=F${lastRow}`
    ];
    paymentsSheet.getRange('G' + lastRow + ':V' + lastRow).setFormulas([formulas]);
    paymentsSheet.getRange('W' + lastRow).setValue(customerId);
    SpreadsheetApp.flush();
    logMessage('INFO', `Added payment at Payments!A${lastRow} Data: ${JSON.stringify(paymentData)}`);
    return { success: true, message: 'Payment added' };
  } catch (e) {
    logMessage('ERROR', `addPayment failed: ${e.message}`);
    return { success: false, message: e.message };
  }
}

function batchAddPayments(payments) {
  try {
    if (!payments.length) throw new Error('No payments provided');
    const results = payments.map((payment, i) => {
      try {
        return addPayment(payment);
      } catch (e) {
        logMessage('ERROR', `Payment at index ${i} failed: ${e.message}`);
        return { success: false, message: `Index ${i}: ${e.message}` };
      }
    });
    const successCount = results.filter(r => r.success).length;
    SpreadsheetApp.flush();
    logMessage('INFO', `Added ${successCount}/${payments.length} payments`);
    return { success: successCount === payments.length, message: `Added ${successCount}/${payments.length} payments`, results: results };
  } catch (e) {
    logMessage('ERROR', `batchAddPayments failed: ${e.message}`);
    return { success: false, message: e.message };
  }
}

function deletePayment(paymentId) {
  try {
    const sheet = getSheetSafely('Payments');
    const data = sheet.getRange('B2:B' + sheet.getLastRow()).getValues();
    let rowIndex = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === paymentId) {
        rowIndex = i + 2;
        break;
      }
    }
    if (rowIndex === -1) throw new Error(`Payment ID ${paymentId} not found`);
    Utils.clearFullRow(sheet, rowIndex);
    sheet.deleteRow(rowIndex);
    SpreadsheetApp.flush();
    logMessage('INFO', `Deleted payment ${paymentId} at row ${rowIndex}`);
    return { success: true, message: 'Payment deleted' };
  } catch (e) {
    logMessage('ERROR', `deletePayment failed: ${e.message}`);
    return { success: false, message: e.message };
  }
}

function initializeApp() {
  try {
    createRequiredSheets();
    logMessage('INFO', 'App initialized');
  } catch (e) {
    logMessage('ERROR', `initializeApp failed: ${e.message}`);
  }
}

function logMessage(level, message) {
  try {
    let logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Logs');
    if (!logSheet && !isRemovingSheets) {
      logSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Logs');
      logSheet.getRange(1, 1, 1, 2).setValues([['Timestamp', 'Message']]);
      logSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
      SpreadsheetApp.flush();
    }
    if (logSheet) {
      const timestamp = new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' });
      logSheet.appendRow([timestamp, `${level}: ${message}`]);
    } else {
      Logger.log(`${level}: ${message}`);
    }
  } catch (e) {
    Logger.log(`logMessage failed: ${e.message} | Original: ${message}`);
  }
}
