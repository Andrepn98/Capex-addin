// ============================================
// ORNA MODULE: Financial Analysis
// ============================================

var csvData = null;
var parsedTransactions = [];

function initFinancialModule() {
  // File upload handlers
  var dropZone = document.getElementById('dropZone');
  var csvFile = document.getElementById('csvFile');
  
  dropZone.addEventListener('dragover', function(e) {
    e.preventDefault();
    dropZone.classList.add('dragover');
  });

  dropZone.addEventListener('dragleave', function(e) {
    e.preventDefault();
    dropZone.classList.remove('dragover');
  });

  dropZone.addEventListener('drop', function(e) {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    var files = e.dataTransfer.files;
    if (files.length > 0 && files[0].name.endsWith('.csv')) {
      processFile(files[0]);
    } else {
      toast("Please upload a CSV file");
    }
  });

  csvFile.addEventListener('change', function(e) {
    if (e.target.files[0]) {
      processFile(e.target.files[0]);
    }
  });

  document.getElementById('processBtn').onclick = processTransactions;
  setFinStatus("ok", "Upload a CSV file to begin");
}

function setFinStatus(type, text) {
  document.getElementById("finStatusDot").className = "dot" + (type ? " " + type : "");
  document.getElementById("finStatusText").textContent = text;
}

function updateProgress(step) {
  document.getElementById("step1").className = "progress-step" + (step >= 1 ? " done" : "");
  document.getElementById("step2").className = "progress-step" + (step >= 2 ? " done" : step === 1 ? " active" : "");
  document.getElementById("step3").className = "progress-step" + (step >= 3 ? " done" : step === 2 ? " active" : "");
}

// ============================================
// FILE HANDLING
// ============================================
function processFile(file) {
  var reader = new FileReader();
  reader.onload = function(e) {
    csvData = e.target.result;
    document.getElementById('fileName').textContent = "✓ " + file.name;
    document.getElementById('fileName').classList.add('show');
    document.getElementById('processBtn').disabled = false;
    updateProgress(1);
    setFinStatus("ok", "File loaded: " + file.name);
    toast("CSV file loaded successfully");
  };
  reader.onerror = function() {
    toast("Error reading file");
    setFinStatus("error", "Error reading file");
  };
  reader.readAsText(file);
}

// ============================================
// CSV PARSING - IMPROVED VERSION
// ============================================
function parseCSV(text) {
  var lines = text.trim().split(/\r?\n/);
  if (lines.length < 2) return [];

  // Parse header row
  var headers = parseCSVLine(lines[0]).map(function(h) {
    return h.toLowerCase().trim().replace(/['"]/g, '');
  });

  console.log("CSV Headers found:", headers);

  // Find column indices - more flexible matching
  var txnIdIdx = findColumnIndex(headers, ['transaction id', 'txn_id', 'id', 'trans id', 'reference', 'ref', 'transaction_id']);
  var dateIdx = findColumnIndex(headers, ['date', 'transaction date', 'trans date', 'posting date', 'value date', 'txn date']);
  var descIdx = findColumnIndex(headers, ['description', 'desc', 'narrative', 'details', 'memo', 'transaction description', 'particulars', 'name']);
  var amountIdx = findColumnIndex(headers, ['amount', 'value', 'sum', 'transaction amount', 'net']);
  var debitIdx = findColumnIndex(headers, ['debit', 'withdrawal', 'out', 'dr', 'debit amount', 'withdrawals']);
  var creditIdx = findColumnIndex(headers, ['credit', 'deposit', 'in', 'cr', 'credit amount', 'deposits']);
  var accountIdx = findColumnIndex(headers, ['account', 'account name', 'bank', 'source', 'account_name']);
  var currencyIdx = findColumnIndex(headers, ['currency', 'ccy', 'cur']);
  var typeIdx = findColumnIndex(headers, ['type', 'transaction type', 'txn type', 'trans type', 'category']);

  console.log("Column indices - txnId:", txnIdIdx, "date:", dateIdx, "desc:", descIdx, "amount:", amountIdx, "debit:", debitIdx, "credit:", creditIdx);

  var transactions = [];
  
  for (var i = 1; i < lines.length; i++) {
    if (!lines[i].trim()) continue;
    
    var cols = parseCSVLine(lines[i]);
    
    // Get original transaction ID (preserve it, don't generate)
    var txnId = txnIdIdx >= 0 && cols[txnIdIdx] ? cols[txnIdIdx].trim() : 'TXN_' + String(i).padStart(5, '0');
    
    // Get date - FIXED: properly parse the date column
    var dateVal = dateIdx >= 0 ? cols[dateIdx] : '';
    var parsedDate = parseDateValue(dateVal);
    
    // Get description
    var descVal = descIdx >= 0 ? cols[descIdx] : '';
    
    // Get debit and credit separately (keep original values)
    var debitVal = debitIdx >= 0 ? parseNumber(cols[debitIdx]) : 0;
    var creditVal = creditIdx >= 0 ? parseNumber(cols[creditIdx]) : 0;
    
    // Calculate net amount: credit - debit (or use amount column if no debit/credit)
    var netAmount = 0;
    if (debitIdx >= 0 || creditIdx >= 0) {
      netAmount = creditVal - debitVal;
    } else if (amountIdx >= 0) {
      netAmount = parseNumber(cols[amountIdx]);
    }

    // Get account name - FIXED: read actual column, not date serial
    var accountVal = accountIdx >= 0 && cols[accountIdx] ? cols[accountIdx].trim() : 'Unknown';

    // Get currency
    var currencyVal = currencyIdx >= 0 && cols[currencyIdx] ? cols[currencyIdx].trim() : 'USD';

    // Get transaction type
    var typeVal = typeIdx >= 0 && cols[typeIdx] ? cols[typeIdx].trim() : '';

    // Only add if there's meaningful data
    if (dateVal || descVal || netAmount !== 0 || debitVal !== 0 || creditVal !== 0) {
      transactions.push({
        txn_id: txnId,
        account_name: accountVal,
        date: parsedDate,
        description: descVal || '',
        debit: debitVal,
        credit: creditVal,
        amount: netAmount, // This will be a formula in Excel: =credit-debit
        currency: currencyVal,
        txn_type: typeVal,
        merchant_key: extractMerchant(descVal),
        category_group: 'Unmapped',
        category: 'Uncategorized',
        method: 'rule',
        confidence: 0,
        month: null
      });
    }
  }

  return transactions;
}

function parseCSVLine(line) {
  var result = [];
  var current = '';
  var inQuotes = false;
  
  for (var i = 0; i < line.length; i++) {
    var char = line[i];
    if (char === '"') {
      inQuotes = !inQuotes;
    } else if (char === ',' && !inQuotes) {
      result.push(current.trim().replace(/^["']|["']$/g, ''));
      current = '';
    } else {
      current += char;
    }
  }
  result.push(current.trim().replace(/^["']|["']$/g, ''));
  return result;
}

function findColumnIndex(headers, possibleNames) {
  // Exact match first
  for (var i = 0; i < possibleNames.length; i++) {
    var idx = headers.indexOf(possibleNames[i]);
    if (idx >= 0) return idx;
  }
  // Partial match
  for (var i = 0; i < headers.length; i++) {
    for (var j = 0; j < possibleNames.length; j++) {
      if (headers[i].includes(possibleNames[j]) || possibleNames[j].includes(headers[i])) {
        return i;
      }
    }
  }
  return -1;
}

function parseNumber(str) {
  if (!str || str === '') return 0;
  // Remove currency symbols, spaces, and handle parentheses for negatives
  var cleaned = String(str).replace(/[$€£¥,\s]/g, '');
  if (cleaned.match(/^\(.*\)$/)) {
    cleaned = '-' + cleaned.replace(/[()]/g, '');
  }
  var num = parseFloat(cleaned);
  return isNaN(num) ? 0 : num;
}

// FIXED: Properly parse date values
function parseDateValue(str) {
  if (!str || str === '') return null;
  
  str = String(str).trim();
  
  // Check if it's already an Excel serial number (number)
  var numVal = parseFloat(str);
  if (!isNaN(numVal) && numVal > 1000 && numVal < 100000) {
    // It's an Excel serial number, convert to date
    var epoch = new Date(1899, 11, 30);
    var date = new Date(epoch.getTime() + numVal * 24 * 60 * 60 * 1000);
    return date;
  }
  
  // Try standard date parsing
  var date = new Date(str);
  if (!isNaN(date.getTime())) return date;
  
  // Try DD/MM/YYYY or DD-MM-YYYY
  var parts = str.split(/[\/\-\.]/);
  if (parts.length === 3) {
    var day, month, year;
    
    // Determine format based on values
    if (parseInt(parts[0]) > 12) {
      // DD/MM/YYYY
      day = parseInt(parts[0]);
      month = parseInt(parts[1]) - 1;
      year = parseInt(parts[2]);
    } else if (parseInt(parts[1]) > 12) {
      // MM/DD/YYYY
      month = parseInt(parts[0]) - 1;
      day = parseInt(parts[1]);
      year = parseInt(parts[2]);
    } else {
      // Assume MM/DD/YYYY for US format
      month = parseInt(parts[0]) - 1;
      day = parseInt(parts[1]);
      year = parseInt(parts[2]);
    }
    
    // Handle 2-digit year
    if (year < 100) {
      year += year > 50 ? 1900 : 2000;
    }
    
    date = new Date(year, month, day);
    if (!isNaN(date.getTime())) return date;
  }
  
  return null;
}

function extractMerchant(description) {
  if (!description) return '';
  return description
    .replace(/[0-9#*]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .split(' ')
    .slice(0, 3)
    .join(' ')
    .toUpperCase();
}

// ============================================
// CATEGORIZATION RULES
// ============================================
var categoryRules = [
  // Revenue
  { keywords: ['sales', 'revenue', 'income', 'payment received', 'invoice paid', 'stripe', 'paypal received', 'deposit'], group: 'Revenue', category: 'Sales Revenue' },
  { keywords: ['interest income', 'interest earned'], group: 'Revenue', category: 'Interest Income' },
  { keywords: ['refund received', 'rebate'], group: 'Revenue', category: 'Other Income' },
  
  // COGS
  { keywords: ['inventory', 'stock purchase', 'raw material', 'supplies', 'manufacturing', 'cogs'], group: 'COGS', category: 'Inventory/Materials' },
  { keywords: ['freight', 'shipping cost', 'logistics'], group: 'COGS', category: 'Freight & Shipping' },
  
  // Operating Expenses
  { keywords: ['salary', 'payroll', 'wages', 'bonus', 'compensation'], group: 'Opex', category: 'Salaries & Wages' },
  { keywords: ['rent', 'lease', 'office space'], group: 'Opex', category: 'Rent' },
  { keywords: ['utility', 'electric', 'gas', 'water', 'internet', 'phone'], group: 'Opex', category: 'Utilities' },
  { keywords: ['software', 'subscription', 'saas', 'cloud', 'aws', 'azure', 'google cloud', 'microsoft', 'adobe', 'slack', 'zoom', 'shopify', 'loom', 'notion', 'figma', 'canva', 'hubspot', 'mailchimp', 'sendinblue', 'zoho', 'algolia', 'huggingface'], group: 'Opex', category: 'Software & Subscriptions' },
  { keywords: ['marketing', 'advertising', 'ads', 'facebook ads', 'google ads', 'promotion', 'dimabay', 'sortlist'], group: 'Opex', category: 'Marketing & Advertising' },
  { keywords: ['travel', 'flight', 'hotel', 'uber', 'lyft', 'taxi', 'airbnb'], group: 'Opex', category: 'Travel & Entertainment' },
  { keywords: ['meal', 'restaurant', 'food', 'lunch', 'dinner', 'coffee'], group: 'Opex', category: 'Meals & Entertainment' },
  { keywords: ['insurance', 'premium'], group: 'Opex', category: 'Insurance' },
  { keywords: ['legal', 'attorney', 'lawyer', 'law firm'], group: 'Opex', category: 'Legal & Professional' },
  { keywords: ['accounting', 'bookkeeping', 'audit', 'cpa'], group: 'Opex', category: 'Accounting' },
  { keywords: ['consulting', 'consultant', 'advisory'], group: 'Opex', category: 'Consulting' },
  { keywords: ['contractor', 'freelance', 'upwork', 'fiverr', 'toptal'], group: 'Opex', category: 'Contractors' },
  { keywords: ['office supplies', 'staples', 'office depot'], group: 'Opex', category: 'Office Supplies' },
  { keywords: ['maintenance', 'repair', 'fix'], group: 'Opex', category: 'Maintenance & Repairs' },
  { keywords: ['bank fee', 'service charge', 'wire fee', 'transaction fee', 'fx fee', 'currency'], group: 'Opex', category: 'Bank Fees' },
  
  // Taxes
  { keywords: ['tax', 'irs', 'vat', 'gst', 'sales tax', 'income tax', 'payroll tax'], group: 'Taxes', category: 'Taxes' },
  
  // Financing
  { keywords: ['loan', 'interest payment', 'mortgage', 'financing'], group: 'Financing', category: 'Loan Interest' },
  { keywords: ['dividend', 'distribution'], group: 'Financing', category: 'Dividends' },
  
  // Transfers
  { keywords: ['transfer', 'internal transfer', 'move funds', 'from savings', 'to savings'], group: 'Transfers', category: 'Internal Transfer' },
  { keywords: ['owner draw', 'owner contribution', 'capital'], group: 'Transfers', category: 'Owner Transactions' }
];

function categorizeTransaction(txn) {
  var descLower = (txn.description || '').toLowerCase();
  var merchantLower = (txn.merchant_key || '').toLowerCase();
  
  for (var i = 0; i < categoryRules.length; i++) {
    var rule = categoryRules[i];
    for (var j = 0; j < rule.keywords.length; j++) {
      if (descLower.includes(rule.keywords[j]) || merchantLower.includes(rule.keywords[j])) {
        return {
          category_group: rule.group,
          category: rule.category,
          method: 'rule',
          confidence: 0.8
        };
      }
    }
  }
  
  return {
    category_group: 'Unmapped',
    category: 'Uncategorized',
    method: 'rule',
    confidence: 0
  };
}

// ============================================
// PROCESS TRANSACTIONS
// ============================================
function processTransactions() {
  if (!csvData) {
    toast("Please upload a CSV file first");
    return;
  }

  setFinStatus("processing", "Processing transactions...");
  updateProgress(2);
  toast("Processing transactions...");

  // Parse CSV
  parsedTransactions = parseCSV(csvData);
  
  console.log("Parsed transactions:", parsedTransactions.length);
  if (parsedTransactions.length > 0) {
    console.log("First transaction:", parsedTransactions[0]);
  }
  
  if (parsedTransactions.length === 0) {
    setFinStatus("error", "No transactions found in CSV");
    toast("No transactions found");
    return;
  }

  // Categorize each transaction
  for (var i = 0; i < parsedTransactions.length; i++) {
    var cat = categorizeTransaction(parsedTransactions[i]);
    parsedTransactions[i].category_group = cat.category_group;
    parsedTransactions[i].category = cat.category;
    parsedTransactions[i].method = cat.method;
    parsedTransactions[i].confidence = cat.confidence;
    
    // Add month (end of month)
    if (parsedTransactions[i].date) {
      var d = parsedTransactions[i].date;
      parsedTransactions[i].month = new Date(d.getFullYear(), d.getMonth() + 1, 0);
    }
  }

  // Write to Excel
  writeToExcel();
}

// ============================================
// WRITE TO EXCEL
// ============================================
function writeToExcel() {
  Excel.run(function(context) {
    var sheets = context.workbook.worksheets;
    sheets.load("items/name");

    return context.sync().then(function() {
      // Delete existing sheets if they exist
      var sheetsToDelete = ['CLEANED_TXN', 'P&L'];
      for (var i = 0; i < sheets.items.length; i++) {
        if (sheetsToDelete.indexOf(sheets.items[i].name) >= 0) {
          sheets.items[i].delete();
        }
      }
      return context.sync();
    }).then(function() {
      // ========== CREATE CLEANED_TXN SHEET ==========
      var txnSheet = sheets.add("CLEANED_TXN");
      
      // NEW COLUMN STRUCTURE with debit, credit, and amount formula
      var headers = [
        "txn_id",           // A - Original transaction ID (preserved)
        "account_name",     // B - Account/bank name
        "date",             // C - Transaction date
        "description",      // D - Transaction description
        "txn_type",         // E - Transaction type
        "debit",            // F - Debit amount (outflow)
        "credit",           // G - Credit amount (inflow)
        "amount",           // H - Net amount (formula: =G-F)
        "currency",         // I - Currency
        "merchant_key",     // J - Normalized merchant
        "category_group",   // K - Category group
        "category",         // L - Category
        "method",           // M - Categorization method
        "confidence",       // N - Confidence score
        "month"             // O - Month (for P&L)
      ];
      
      // Write headers
      var headerRange = txnSheet.getRange("A1:O1");
      headerRange.values = [headers];
      headerRange.format.font.bold = true;
      headerRange.format.fill.color = "#217346";
      headerRange.format.font.color = "#FFFFFF";

      // Write transaction data
      var dataRows = [];
      for (var i = 0; i < parsedTransactions.length; i++) {
        var txn = parsedTransactions[i];
        var rowNum = i + 2; // Excel row number (1-indexed, after header)
        
        dataRows.push([
          txn.txn_id,                                    // A - Original ID
          txn.account_name,                              // B - Account name
          txn.date ? excelDate(txn.date) : "",          // C - Date
          txn.description,                               // D - Description
          txn.txn_type || "",                           // E - Transaction type
          txn.debit || 0,                               // F - Debit
          txn.credit || 0,                              // G - Credit
          null,                                          // H - Amount (will be formula)
          txn.currency,                                  // I - Currency
          txn.merchant_key,                              // J - Merchant
          txn.category_group,                            // K - Category group
          txn.category,                                  // L - Category
          txn.method,                                    // M - Method
          txn.confidence,                                // N - Confidence
          txn.month ? excelDate(txn.month) : ""         // O - Month
        ]);
      }

      if (dataRows.length > 0) {
        var dataRange = txnSheet.getRange("A2:O" + (dataRows.length + 1));
        dataRange.values = dataRows;
        
        // Add amount formula for each row: =G-F (credit - debit)
        for (var i = 0; i < dataRows.length; i++) {
          var rowNum = i + 2;
          txnSheet.getRange("H" + rowNum).formulas = [["=G" + rowNum + "-F" + rowNum]];
        }
        
        // Format date column
        txnSheet.getRange("C2:C" + (dataRows.length + 1)).numberFormat = [["yyyy-mm-dd"]];
        
        // Format month column
        txnSheet.getRange("O2:O" + (dataRows.length + 1)).numberFormat = [["yyyy-mm"]];
        
        // Format currency columns
        txnSheet.getRange("F2:H" + (dataRows.length + 1)).numberFormat = [["#,##0.00"]];
        
        // Create table
        var tableRange = txnSheet.getRange("A1:O" + (dataRows.length + 1));
        var table = txnSheet.tables.add(tableRange, true);
        table.name = "tbl_cleaned_txn";
        
        // Auto-fit columns
        txnSheet.getRange("A:O").format.autofitColumns();
      }

      // ========== CREATE P&L SHEET ==========
      var plSheet = sheets.add("P&L");
      
      // Get unique months
      var months = [];
      var monthSet = {};
      for (var i = 0; i < parsedTransactions.length; i++) {
        if (parsedTransactions[i].month) {
          var monthKey = parsedTransactions[i].month.getFullYear() + "-" + 
                        String(parsedTransactions[i].month.getMonth() + 1).padStart(2, '0');
          if (!monthSet[monthKey]) {
            monthSet[monthKey] = parsedTransactions[i].month;
            months.push({ key: monthKey, date: parsedTransactions[i].month });
          }
        }
      }
      months.sort(function(a, b) { return a.key.localeCompare(b.key); });

      // P&L Categories
      var plCategories = [
        { name: "Revenue", group: "Revenue", isTotal: false },
        { name: "COGS", group: "COGS", isTotal: false },
        { name: "Gross Profit", group: null, isTotal: true },
        { name: "", group: null, isTotal: false },
        { name: "Operating Expenses", group: "Opex", isTotal: false },
        { name: "Operating Income", group: null, isTotal: true },
        { name: "", group: null, isTotal: false },
        { name: "Taxes", group: "Taxes", isTotal: false },
        { name: "Financing", group: "Financing", isTotal: false },
        { name: "Net Income", group: null, isTotal: true },
        { name: "", group: null, isTotal: false },
        { name: "Transfers (Excluded)", group: "Transfers", isTotal: false },
        { name: "Unmapped", group: "Unmapped", isTotal: false }
      ];

      // Write P&L headers
      plSheet.getRange("A1").values = [["P&L Summary"]];
      plSheet.getRange("A1").format.font.bold = true;
      plSheet.getRange("A1").format.font.size = 16;

      plSheet.getRange("A3").values = [["Category"]];
      plSheet.getRange("A3").format.font.bold = true;

      // Month headers
      for (var m = 0; m < months.length; m++) {
        var colLetter = getColumnLetter(m + 1);
        plSheet.getRange(colLetter + "3").values = [[months[m].key]];
        plSheet.getRange(colLetter + "3").format.font.bold = true;
        plSheet.getRange(colLetter + "3").format.horizontalAlignment = "Center";
      }

      // Write P&L rows
      var rowNum = 4;
      var rowRefs = {};
      
      for (var c = 0; c < plCategories.length; c++) {
        var cat = plCategories[c];
        
        if (cat.name === "") {
          rowNum++;
          continue;
        }

        plSheet.getRange("A" + rowNum).values = [[cat.name]];
        
        if (cat.isTotal) {
          plSheet.getRange("A" + rowNum).format.font.bold = true;
          plSheet.getRange("A" + rowNum + ":" + getColumnLetter(months.length) + rowNum).format.fill.color = "#FFF2CC";
        }

        // Write SUMIFS formulas for each month
        for (var m = 0; m < months.length; m++) {
          var colLetter = getColumnLetter(m + 1);
          var monthVal = months[m].key;
          
          if (cat.group) {
            // SUMIFS formula referencing tbl_cleaned_txn[amount] column (H)
            var formula = '=SUMIFS(tbl_cleaned_txn[amount],tbl_cleaned_txn[category_group],"' + cat.group + '",tbl_cleaned_txn[month],"' + monthVal + '*")';
            plSheet.getRange(colLetter + rowNum).formulas = [[formula]];
          } else if (cat.isTotal) {
            if (cat.name === "Gross Profit") {
              var formula = "=" + colLetter + rowRefs["Revenue"] + "-" + colLetter + rowRefs["COGS"];
              plSheet.getRange(colLetter + rowNum).formulas = [[formula]];
            } else if (cat.name === "Operating Income") {
              var formula = "=" + colLetter + rowRefs["Gross Profit"] + "-" + colLetter + rowRefs["Operating Expenses"];
              plSheet.getRange(colLetter + rowNum).formulas = [[formula]];
            } else if (cat.name === "Net Income") {
              var formula = "=" + colLetter + rowRefs["Operating Income"] + "-" + colLetter + rowRefs["Taxes"] + "-" + colLetter + rowRefs["Financing"];
              plSheet.getRange(colLetter + rowNum).formulas = [[formula]];
            }
          }
          
          plSheet.getRange(colLetter + rowNum).numberFormat = [["#,##0.00"]];
        }

        rowRefs[cat.name] = rowNum;
        rowNum++;
      }

      // Format
      plSheet.getRange("A:A").format.columnWidth = 150;
      for (var m = 0; m < months.length; m++) {
        plSheet.getRange(getColumnLetter(m + 1) + ":" + getColumnLetter(m + 1)).format.columnWidth = 100;
      }

      // Activate CLEANED_TXN sheet
      txnSheet.activate();

      return context.sync();
    }).then(function() {
      updateProgress(3);
      setFinStatus("ok", "Created CLEANED_TXN (" + parsedTransactions.length + " rows) and P&L!");
      toast("Analysis complete!");
    });
  }).catch(function(error) {
    setFinStatus("error", "Error: " + error.message);
    toast("Error creating sheets");
    console.error(error);
  });
}
