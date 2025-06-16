var VALID_PAYMENT_METHODS = ['Cash', 'UPI', 'Bank Transfer', 'Cheque', 'Others'];

function doGet(e) {
  try {
    var template = HtmlService.createTemplateFromFile('index');
    var htmlOutput = template.evaluate()
      .setTitle('BrickSync')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0'); // Moved viewport here
    return htmlOutput;
  } catch (err) {
    logMessage('ERROR', 'doGet failed: ' + err.message);
    return HtmlService.createHtmlOutput('Error loading application: ' + err.message);
  }
}

function doPost(e) {
  try {
    if (!Session.getActiveUser().getEmail()) { // Added authentication
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Unauthorized access' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    var params = JSON.parse(e.postData.contents);
    var action = params.action;
    var data = params.data;

    if (action === 'getData') {
      return ContentService.createTextOutput(JSON.stringify(getSheetData(data.sheetName)))
        .setMimeType(ContentService.MimeType.JSON);
    } else if (action === 'addData') {
      var result;
      if (data.sheetName === 'Customers') {
        result = addCustomer(data.record);
      } else if (data.sheetName === 'Sales') {
        result = addSale(data.record);
      } else if (data.sheetName === 'Payments') {
        result = addPayment(data.record);
      }
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Invalid action' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    logMessage('ERROR', 'doPost failed: ' + err.message);
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSheetData(sheetName) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      logMessage('WARN', 'Sheet ' + sheetName + ' was missing and has been recreated');
      createRequiredSheets();
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    }
    var headers = getHeaders(sheetName);
    var data = sheet.getDataRange().getValues();
    var records = [];
    for (var i = 1; i < data.length && i <= CONFIG.MAX_ROWS; i++) {
      var record = {};
      headers.forEach(function(header, index) {
        record[header] = data[i][index];
      });
      records.push(record);
    }
    return { success: true, data: records };
  } catch (err) {
    logMessage('ERROR', 'getSheetData failed for ' + sheetName + ': ' + err.message);
    return { success: false, message: err.message };
  }
}

function getLastContentRow(sheet, columns) {
  try {
    var lastRow = Math.min(sheet.getLastRow(), CONFIG.MAX_ROWS);
    var values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    for (var i = values.length - 1; i >= 0; i--) {
      for (var j = 0; j < columns.length; j++) {
        if (values[i][columns[j] - 1]) {
          return i + 2;
        }
      }
    }
    return 1;
  } catch (err) {
    logMessage('ERROR', 'getLastContentRow failed: ' + err.message);
    return 1;
  }
}

function getCustomers() {
  return getSheetData('Customers');
}

function getSales() {
  return getSheetData('Sales');
}

function getPayments() {
  return getSheetData('Payments');
}

function getTransactions() {
  return getSheetData('Transactions');
}

function getBalances() {
  return getSheetData('Balances');
}

function generateId(prefix) { // Added for UUID generation
  return prefix + '-' + Utilities.getUuid().slice(0, 8);
}

function addCustomer(customer) {
  var lock = LockService.getScriptLock(); // Added LockService
  try {
    lock.waitLock(10000);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customers');
    if (!sheet) {
      createRequiredSheets();
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customers');
    }
    var name = customer.name ? customer.name.trim() : '';
    if (!name) {
      return { success: false, message: 'Customer name is required' };
    }
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1].toLowerCase() === name.toLowerCase()) {
        return { success: false, message: 'Customer already exists' };
      }
    }
    var customerId = generateId('CUST'); // Changed to UUID
    sheet.appendRow([customerId, name]);
    SpreadsheetApp.flush();
    logMessage('INFO', 'Added customer: ' + name);
    Utils.invalidateCustomerCache();
    return { success: true, message: 'Customer added successfully' };
  } catch (err) {
    logMessage('ERROR', 'addCustomer failed: ' + err.message);
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function batchAddCustomers(customers) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customers');
    if (!sheet) {
      createRequiredSheets();
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customers');
    }
    var existingNames = sheet.getDataRange().getValues().slice(1).map(function(row) { return row[1].toLowerCase(); });
    var newCustomers = [];
    customers.forEach(function(customer) {
      var name = customer.name ? customer.name.trim() : '';
      if (name && existingNames.indexOf(name.toLowerCase()) === -1) {
        newCustomers.push([generateId('CUST'), name]); // Changed to UUID
        existingNames.push(name.toLowerCase());
      }
    });
    if (newCustomers.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newCustomers.length, 2).setValues(newCustomers);
      SpreadsheetApp.flush();
      logMessage('INFO', 'Batch added ' + newCustomers.length + ' customers');
      Utils.invalidateCustomerCache();
      return { success: true, message: newCustomers.length + ' customers added successfully' };
    }
    return { success: false, message: 'No new customers to add' };
  } catch (err) {
    logMessage('ERROR', 'batchAddCustomers failed: ' + err.message);
    return { success: false, message: err.message };
  }
}

function deleteCustomer(customerId) {
  var lock = LockService.getScriptLock(); // Added LockService
  try {
    lock.waitLock(10000);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customers');
    if (!sheet) {
      return { success: false, message: 'Customers sheet not found' };
    }
    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === customerId) {
        rowIndex = i + 1;
        break;
      }
    }
    if (rowIndex === -1) {
      return { success: false, message: 'Customer not found' };
    }
    var salesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sales');
    var paymentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Payments');
    var salesData = salesSheet.getDataRange().getValues();
    var paymentsData = paymentsSheet.getDataRange().getValues();
    for (var i = 1; i < salesData.length; i++) {
      if (salesData[i][2] === data[rowIndex - 1][1]) {
        return { success: false, message: 'Cannot delete customer with associated sales' };
      }
    }
    for (var i = 1; i < paymentsData.length; i++) {
      if (paymentsData[i][2] === data[rowIndex - 1][1]) {
        return { success: false, message: 'Cannot delete customer with associated payments' };
      }
    }
    sheet.deleteRow(rowIndex);
    SpreadsheetApp.flush();
    logMessage('INFO', 'Deleted customer: ' + customerId);
    Utils.invalidateCustomerCache();
    return { success: true, message: 'Customer deleted successfully' };
  } catch (err) {
    logMessage('ERROR', 'deleteCustomer failed: ' + err.message);
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function updateCustomer(customer) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customers');
    if (!sheet) {
      return { success: false, message: 'Customers sheet not found' };
    }
    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === customer.customerId) {
        rowIndex = i + 1;
        break;
      }
    }
    if (rowIndex === -1) {
      return { success: false, message: 'Customer not found' };
    }
    var newName = customer.name ? customer.name.trim() : '';
    if (!newName) {
      return { success: false, message: 'Customer name is required' };
    }
    for (var i = 1; i < data.length; i++) {
      if (i !== rowIndex - 1 && data[i][1].toLowerCase() === newName.toLowerCase()) {
        return { success: false, message: 'Customer name already exists' };
      }
    }
    var oldName = data[rowIndex - 1][1];
    sheet.getRange(rowIndex, 2).setValue(newName);
    var salesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sales');
    var paymentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Payments');
    var salesData = salesSheet.getDataRange().getValues();
    var paymentsData = paymentsSheet.getDataRange().getValues();
    for (var i = 1; i < salesData.length; i++) {
      if (salesData[i][2] === oldName) {
        salesSheet.getRange(i + 1, 3).setValue(newName);
      }
    }
    for (var i = 1; i < paymentsData.length; i++) {
      if (paymentsData[i][2] === oldName) {
        paymentsSheet.getRange(i + 1, 3).setValue(newName);
      }
    }
    SpreadsheetApp.flush();
    logMessage('INFO', 'Updated customer: ' + customer.customerId + ' to ' + newName);
    Utils.invalidateCustomerCache();
    return { success: true, message: 'Customer updated successfully' };
  } catch (err) {
    logMessage('ERROR', 'updateCustomer failed: ' + err.message);
    return { success: false, message: err.message };
  }
}

function addSale(sale) {
  var lock = LockService.getScriptLock(); // Added LockService
  try {
    lock.waitLock(10000);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sales');
    if (!sheet) {
      createRequiredSheets();
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sales');
    }
    var date = formatDate(sale.date);
    if (!date) {
      return { success: false, message: 'Invalid date format' };
    }
    var customer = sale.customer ? sale.customer.trim() : '';
    if (!customer) {
      return { success: false, message: 'Customer is required' };
    }
    var customersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customers');
    var customerExists = false;
    var customerData = customersSheet.getRange(2, 2, Math.min(customersSheet.getLastRow() - 1, CONFIG.DROPDOWN_ROWS)).getValues(); // Used DROPDOWN_ROWS
    for (var i = 0; i < customerData.length; i++) {
      if (customerData[i][0] && customerData[i][0].toLowerCase() === customer.toLowerCase()) {
        customer = customerData[i][0];
        customerExists = true;
        break;
      }
    }
    if (!customerExists) {
      var addCustomerResult = addCustomer({ name: customer });
      if (!addCustomerResult.success) {
        return addCustomerResult;
      }
    }
    var quantity = parseFloat(sale.quantity);
    var rate = parseFloat(sale.rate);
    var vehicleRent = parseFloat(sale.vehicleRent) || 0;
    var paymentReceived = parseFloat(sale.paymentReceived) || 0;
    if (isNaN(quantity) || quantity <= 0) {
      return { success: false, message: 'Invalid quantity' };
    }
    if (isNaN(rate) || rate <= 0) {
      return { success: false, message: 'Invalid rate' };
    }
    if (isNaN(vehicleRent) || vehicleRent < 0) {
      return { success: false, message: 'Invalid vehicle rent' };
    }
    if (isNaN(paymentReceived) || paymentReceived < 0) {
      return { success: false, message: 'Invalid payment received' };
    }
    var paymentMethod = sale.paymentMethod ? sale.paymentMethod.trim() : '';
    if (paymentMethod && VALID_PAYMENT_METHODS.indexOf(paymentMethod) === -1) {
      return { success: false, message: 'Invalid payment method' };
    }
    if (paymentMethod && paymentReceived === 0) {
      return { success: false, message: 'Payment received must be greater than 0 when payment method is specified' };
    }
    var amount = (quantity * rate) + vehicleRent;
    var saleId = generateId('SALE'); // Changed to UUID
    var remarks = sale.remarks ? sale.remarks.trim() : '';
    var row = [date, saleId, customer, quantity, rate, vehicleRent, amount, paymentMethod, paymentReceived, remarks];
    sheet.appendRow(row);
    sheet.getRange(sheet.getLastRow(), 7).setFormula('=D' + sheet.getLastRow() + '*E' + sheet.getLastRow() + '+F' + sheet.getLastRow());
    SpreadsheetApp.flush();
    updateTransactionsAndBalances(null, [customer]);
    logMessage('INFO', 'Added sale: ' + saleId + ' for customer: ' + customer);
    return { success: true, message: 'Sale added successfully' };
  } catch (err) {
    logMessage('ERROR', 'addSale failed: ' + err.message);
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function batchAddSales(sales) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sales');
    if (!sheet) {
      createRequiredSheets();
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sales');
    }
    var customersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customers');
    var customerData = customersSheet.getDataRange().getValues().slice(1).map(function(row) { return row[1].toLowerCase(); });
    var newSales = [];
    var newCustomers = [];
    sales.forEach(function(sale) {
      var date = formatDate(sale.date);
      if (!date) return;
      var customer = sale.customer ? sale.customer.trim() : '';
      if (!customer) return;
      var quantity = parseFloat(sale.quantity);
      var rate = parseFloat(sale.rate);
      var vehicleRent = parseFloat(sale.vehicleRent) || 0;
      var paymentReceived = parseFloat(sale.paymentReceived) || 0;
      var paymentMethod = sale.paymentMethod ? sale.paymentMethod.trim() : '';
      var remarks = sale.remarks ? sale.remarks.trim() : '';
      if (isNaN(quantity) || quantity <= 0 || isNaN(rate) || rate <= 0 || isNaN(vehicleRent) || vehicleRent < 0 || isNaN(paymentReceived) || paymentReceived < 0) {
        return;
      }
      if (paymentMethod && VALID_PAYMENT_METHODS.indexOf(paymentMethod) === -1) return;
      if (paymentMethod && paymentReceived === 0) return;
      var amount = (quantity * rate) + vehicleRent;
      var saleId = generateId('SALE'); // Changed to UUID
      if (customerData.indexOf(customer.toLowerCase()) === -1) {
        newCustomers.push({ name: customer });
        customerData.push(customer.toLowerCase());
      }
      newSales.push([date, saleId, customer, quantity, rate, vehicleRent, amount, paymentMethod, paymentReceived, remarks]);
    });
    if (newCustomers.length > 0) {
      var addCustomersResult = batchAddCustomers(newCustomers);
      if (!addCustomersResult.success) {
        return addCustomersResult;
      }
    }
    if (newSales.length > 0) {
      var startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, newSales.length, newSales[0].length).setValues(newSales);
      for (var i = 0; i < newSales.length; i++) {
        sheet.getRange(startRow + i, 7).setFormula('=D' + (startRow + i) + '*E' + (startRow + i) + '+F' + (startRow + i));
      }
      SpreadsheetApp.flush();
      var customers = [...new Set(newSales.map(function(sale) { return sale[2]; }))];
      updateTransactionsAndBalances(null, customers);
      logMessage('INFO', 'Batch added ' + newSales.length + ' sales');
      return { success: true, message: newSales.length + ' sales added successfully' };
    }
    return { success: false, message: 'No valid sales to add' };
  } catch (err) {
    logMessage('ERROR', 'batchAddSales failed: ' + err.message);
    return { success: false, message: err.message };
  }
}

function deleteSale(saleId) {
  var lock = LockService.getScriptLock(); // Added LockService
  try {
    lock.waitLock(10000);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sales');
    if (!sheet) {
      return { success: false, message: 'Sales sheet not found' };
    }
    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    var customer = '';
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === saleId) {
        rowIndex = i + 1;
        customer = data[i][2];
        break;
      }
    }
    if (rowIndex === -1) {
      return { success: false, message: 'Sale not found' };
    }
    sheet.deleteRow(rowIndex);
    SpreadsheetApp.flush();
    updateTransactionsAndBalances(null, [customer]);
    logMessage('INFO', 'Deleted sale: ' + saleId);
    return { success: true, message: 'Sale deleted successfully' };
  } catch (err) {
    logMessage('ERROR', 'deleteSale failed: ' + err.message);
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function addPayment(payment) {
  var lock = LockService.getScriptLock(); // Added LockService
  try {
    lock.waitLock(10000);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Payments');
    if (!sheet) {
      createRequiredSheets();
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Payments');
    }
    var date = formatDate(payment.date);
    if (!date) {
      return { success: false, message: 'Invalid date format' };
    }
    var customer = payment.customer ? payment.customer.trim() : '';
    if (!customer) {
      return { success: false, message: 'Customer is required' };
    }
    var customersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customers');
    var customerExists = false;
    var customerData = customersSheet.getRange(2, 2, Math.min(customersSheet.getLastRow() - 1, CONFIG.DROPDOWN_ROWS)).getValues(); // Used DROPDOWN_ROWS
    for (var i = 0; i < customerData.length; i++) {
      if (customerData[i][0] && customerData[i][0].toLowerCase() === customer.toLowerCase()) {
        customer = customerData[i][0];
        customerExists = true;
        break;
      }
    }
    if (!customerExists) {
      var addCustomerResult = addCustomer({ name: customer });
      if (!addCustomerResult.success) {
        return addCustomerResult;
      }
    }
    var paymentMethod = payment.paymentMethod ? payment.paymentMethod.trim() : '';
    if (!paymentMethod || VALID_PAYMENT_METHODS.indexOf(paymentMethod) === -1) {
      return { success: false, message: 'Invalid payment method' };
    }
    var paymentReceived = parseFloat(payment.paymentReceived);
    if (isNaN(paymentReceived) || paymentReceived <= 0) {
      return { success: false, message: 'Payment received must be greater than 0' };
    }
    var paymentId = generateId('PAY'); // Changed to UUID
    var remarks = payment.remarks ? payment.remarks.trim() : '';
    sheet.appendRow([date, paymentId, customer, paymentMethod, paymentReceived, remarks]);
    SpreadsheetApp.flush();
    updateTransactionsAndBalances(null, [customer]);
    logMessage('INFO', 'Added payment: ' + paymentId + ' for customer: ' + customer);
    return { success: true, message: 'Payment added successfully' };
  } catch (err) {
    logMessage('ERROR', 'addPayment failed: ' + err.message);
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function batchAddPayments(payments) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Payments');
    if (!sheet) {
      createRequiredSheets();
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Payments');
    }
    var customersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customers');
    var customerData = customersSheet.getDataRange().getValues().slice(1).map(function(row) { return row[1].toLowerCase(); });
    var newPayments = [];
    var newCustomers = [];
    payments.forEach(function(payment) {
      var date = formatDate(payment.date);
      if (!date) return;
      var customer = payment.customer ? payment.customer.trim() : '';
      if (!customer) return;
      var paymentMethod = payment.paymentMethod ? payment.paymentMethod.trim() : '';
      var paymentReceived = parseFloat(payment.paymentReceived);
      var remarks = payment.remarks ? payment.remarks.trim() : '';
      if (!paymentMethod || VALID_PAYMENT_METHODS.indexOf(paymentMethod) === -1 || isNaN(paymentReceived) || paymentReceived <= 0) {
        return;
      }
      var paymentId = generateId('PAY'); // Changed to UUID
      if (customerData.indexOf(customer.toLowerCase()) === -1) {
        newCustomers.push({ name: customer });
        customerData.push(customer.toLowerCase());
      }
      newPayments.push([date, paymentId, customer, paymentMethod, paymentReceived, remarks]);
    });
    if (newCustomers.length > 0) {
      var addCustomersResult = batchAddCustomers(newCustomers);
      if (!addCustomersResult.success) {
        return addCustomersResult;
      }
    }
    if (newPayments.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newPayments.length, newPayments[0].length).setValues(newPayments);
      SpreadsheetApp.flush();
      var customers = [...new Set(newPayments.map(function(payment) { return payment[2]; }))];
      updateTransactionsAndBalances(null, customers);
      logMessage('INFO', 'Batch added ' + newPayments.length + ' payments');
      return { success: true, message: newPayments.length + ' payments added successfully' };
    }
    return { success: false, message: 'No valid payments to add' };
  } catch (err) {
    logMessage('ERROR', 'batchAddPayments failed: ' + err.message);
    return { success: false, message: err.message };
  }
}

function deletePayment(paymentId) {
  var lock = LockService.getScriptLock(); // Added LockService
  try {
    lock.waitLock(10000);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Payments');
    if (!sheet) {
      return { success: false, message: 'Payments sheet not found' };
    }
    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    var customer = '';
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === paymentId) {
        rowIndex = i + 1;
        customer = data[i][2];
        break;
      }
    }
    if (rowIndex === -1) {
      return { success: false, message: 'Payment not found' };
    }
    sheet.deleteRow(rowIndex);
    SpreadsheetApp.flush();
    updateTransactionsAndBalances(null, [customer]);
    logMessage('INFO', 'Deleted payment: ' + paymentId);
    return { success: true, message: 'Payment deleted successfully' };
  } catch (err) {
    logMessage('ERROR', 'deletePayment failed: ' + err.message);
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function formatDate(dateStr) {
  try {
    if (!dateStr) return null;
    var regex = /^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/;
    var match = dateStr.match(regex);
    if (match) {
      var day = parseInt(match[1], 10);
      var month = parseInt(match[2], 10) - 1;
      var year = parseInt(match[3], 10);
      var date = new Date(year, month, day);
      if (date.getFullYear() === year && date.getMonth() === month && date.getDate() === day) {
        return Utilities.formatDate(date, Session.getScriptTimeZone(), CONFIG.DATE_FORMAT);
      }
    }
    regex = /^(\d{4})-(\d{1,2})-(\d{1,2})$/;
    match = dateStr.match(regex);
    if (match) {
      var year = parseInt(match[1], 10);
      var month = parseInt(match[2], 10) - 1;
      var day = parseInt(match[3], 10);
      var date = new Date(year, month, day);
      if (date.getFullYear() === year && date.getMonth() === month && date.getDate() === day) {
        return Utilities.formatDate(date, Session.getScriptTimeZone(), CONFIG.DATE_FORMAT);
      }
    }
    return null;
  } catch (err) {
    logMessage('ERROR', 'formatDate failed: ' + err.message);
    return null;
  }
}

function logMessage(level, message) {
  try {
    var sanitizedMessage = message.replace(/[<>]/g, ''); // Added sanitization
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Logs');
    if (!sheet) {
      Logger.log('[' + level + '] ' + sanitizedMessage);
      return;
    }
    sheet.appendRow([new Date(), '[' + level + '] ' + sanitizedMessage]);
    SpreadsheetApp.flush();
  } catch (err) {
    Logger.log('[' + level + '] ' + sanitizedMessage);
  }
}

function initializeApp() {
  try {
    createRequiredSheets();
    logMessage('INFO', 'Application initialized');
  } catch (err) {
    logMessage('ERROR', 'initializeApp failed: ' + err.message);
  }
}
