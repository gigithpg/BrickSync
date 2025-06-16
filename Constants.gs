const CONFIG = {
  MAX_ROWS: 500,
  DATE_FORMAT: 'dd/MM/yyyy'
};
const SHEET_CONFIG = [
  { name: 'Customers', headers: ['Customer ID', 'Name'] },
  { name: 'Sales', headers: ['Date', 'Sale ID', 'Customer', 'Quantity', 'Rate', 'Vehicle Rent', 'Amount', 'Payment Method', 'Payment Received', 'Remarks'] },
  { name: 'Payments', headers: ['Date', 'Payment ID', 'Customer', 'Payment Method', 'Payment Received', 'Remarks'] },
  { name: 'Transactions', headers: ['Type', 'Date', 'Transaction ID', 'Customer', 'Quantity', 'Rate', 'Vehicle Rent', 'Amount', 'Payment Method', 'Payment Received', 'Remarks'] },
  { name: 'Balances', headers: ['Customer', 'Total Sales', 'Total Payments', 'Pending Balance'] },
  { name: 'Logs', headers: ['Timestamp', 'Message'] }
];
