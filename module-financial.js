// ============================================
// ORNA MODULE: Financial Analysis
// Version 2.0 - Fixed column mapping
// ============================================

var csvData = null;
var parsedTransactions = [];

function initFinancialModule() {
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
// CSV PARSING - VERSION 2 (FIXED)
// ============================================
function parseCSV(text) {
  var lines = text.trim().split(/\r?\n/);
  if (lines.length < 2) return [];

  // Parse header row
  var rawHeaders = parseCSVLine(lines[0]);
  var headers = rawHeaders.map(function(h) {
    return h.toLowerCase().trim().replace(/['"]/g, '');
  });

  console.log("=== CSV DEBUG ===");
  console.log("Raw headers:", rawHeaders);
  console.log("Normalized headers:", headers);

  // Find column indices - VERY flexible matching
  var colMap = {
    txnId: findColumn(headers, ['transaction id', 'txn_id', 'id', 'trans id', 'reference', 'ref', 'transaction_id', 'transactionid']),
    date: findColumn(headers, ['time', 'date', 'transaction date', 'trans date', 'posting date', 'value date', 'txn date', 'transactiondate', 'posted', 'created at']),
    desc: findColumn(headers, ['description', 'desc', 'narrative', 'details', 'memo', 'transaction description', 'particulars', 'name', 'payee']),
    amount: findColumn(headers, ['amount', 'value', 'sum', 'transaction amount', 'net', 'total']),
    debit: findColumn(headers, ['debit net amount', 'debit', 'withdrawal', 'out', 'dr', 'debit amount', 'withdrawals', 'outflow', 'expense']),
    credit: findColumn(headers, ['credit net amount', 'credit', 'deposit', 'in', 'cr', 'credit amount', 'deposits', 'inflow', 'income']),
    currency: findColumn(headers, ['wallet currency', 'currency', 'ccy', 'cur']),
    type: findColumn(headers, ['financial transaction type', 'type', 'transaction type', 'txn type', 'trans type', 'category', 'transactiontype'])
  };

  console.log("Column mapping:", colMap);
  
  // Show first data row for debugging
  if (lines.length > 1) {
    console.log("First data row:", parseCSVLine(lines[1]));
  }

  var transactions = [];
  
  for (var i = 1; i < lines.length; i++) {
    if (!lines[i].trim()) continue;
    
    var cols = parseCSVLine(lines[i]);
    
    // Get original transaction ID
    var txnId = colMap.txnId >= 0 && cols[colMap.txnId] ? String(cols[colMap.txnId]).trim() : 'TXN_' + String(i).padStart(5, '0');
    
    // Get date - CRITICAL FIX: parse properly
    var dateRaw = colMap.date >= 0 ? cols[colMap.date] : '';
    var parsedDate = parseDateValue(dateRaw);
    
    // Get description
    var descVal = colMap.desc >= 0 ? cols[colMap.desc] : '';
    
    // Get transaction type
    var typeVal = colMap.type >= 0 && cols[colMap.type] ? String(cols[colMap.type]).trim() : '';
    
    // Get debit and credit
    var debitVal = colMap.debit >= 0 ? parseNumber(cols[colMap.debit]) : 0;
    var creditVal = colMap.credit >= 0 ? parseNumber(cols[colMap.credit]) : 0;
    
    // If no separate debit/credit columns, use amount column
    // Negative amounts = debit (expense), Positive amounts = credit (income)
    if (colMap.debit < 0 && colMap.credit < 0 && colMap.amount >= 0) {
      var amt = parseNumber(cols[colMap.amount]);
      if (amt < 0) {
        debitVal = Math.abs(amt);
        creditVal = 0;
      } else {
        debitVal = 0;
        creditVal = amt;
      }
    }

    // Get currency
    var currencyVal = colMap.currency >= 0 && cols[colMap.currency] ? String(cols[colMap.currency]).trim() : 'USD';

    // Only add if there's meaningful data
    if (parsedDate || descVal || debitVal !== 0 || creditVal !== 0) {
      transactions.push({
        txn_id: txnId,
        date: parsedDate,
        description: descVal || '',
        txn_type: typeVal,
        debit: debitVal,
        credit: creditVal,
        currency: currencyVal,
        merchant_key: extractMerchant(descVal),
        category_group: 'Unmapped',
        category: 'Uncategorized',
        method: 'rule',
        confidence: 0,
        month: null
      });
    }
  }

  console.log("Parsed " + transactions.length + " transactions");
  if (transactions.length > 0) {
    console.log("Sample transaction:", transactions[0]);
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

function findColumn(headers, possibleNames) {
  // Exact match first
  for (var j = 0; j < possibleNames.length; j++) {
    for (var i = 0; i < headers.length; i++) {
      if (headers[i] === possibleNames[j]) {
        return i;
      }
    }
  }
  // Partial/contains match
  for (var j = 0; j < possibleNames.length; j++) {
    for (var i = 0; i < headers.length; i++) {
      if (headers[i].includes(possibleNames[j]) || possibleNames[j].includes(headers[i])) {
        return i;
      }
    }
  }
  return -1;
}

function parseNumber(str) {
  if (str === null || str === undefined || str === '') return 0;
  var cleaned = String(str).replace(/[$€£¥,\s]/g, '');
  if (cleaned.match(/^\(.*\)$/)) {
    cleaned = '-' + cleaned.replace(/[()]/g, '');
  }
  var num = parseFloat(cleaned);
  return isNaN(num) ? 0 : num;
}

// FIXED: Robust date parsing
function parseDateValue(str) {
  if (str === null || str === undefined || str === '') return null;
  
  str = String(str).trim();
  
  // Handle ISO 8601 format with timezone: 2025-10-02T06:06:23+0800
  if (str.match(/^\d{4}-\d{2}-\d{2}T/)) {
    var date = new Date(str);
    if (!isNaN(date.getTime())) {
      return date;
    }
    // If direct parsing fails, extract just the date part
    var datePart = str.split('T')[0];
    var parts = datePart.split('-');
    if (parts.length === 3) {
      var date = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
      if (!isNaN(date.getTime())) return date;
    }
  }
  
  // Check if it's an Excel serial number (pure number between reasonable date range)
  var numVal = parseFloat(str);
  if (!isNaN(numVal) && numVal > 1000 && numVal < 100000 && str.match(/^[\d.]+$/)) {
    // Excel serial number: days since 1899-12-30
    var epoch = new Date(1899, 11, 30);
    var date = new Date(epoch.getTime() + numVal * 24 * 60 * 60 * 1000);
    if (!isNaN(date.getTime())) {
      return date;
    }
  }
  
  // Try ISO format (YYYY-MM-DD)
  if (str.match(/^\d{4}-\d{2}-\d{2}/)) {
    var date = new Date(str);
    if (!isNaN(date.getTime())) return date;
  }
  
  // Try standard Date parsing
  var date = new Date(str);
  if (!isNaN(date.getTime()) && date.getFullYear() > 1900 && date.getFullYear() < 2100) {
    return date;
  }
  
  // Try various formats: DD/MM/YYYY, MM/DD/YYYY, DD-MM-YYYY, etc.
  var parts = str.split(/[\/\-\.]/);
  if (parts.length === 3) {
    var p0 = parseInt(parts[0]);
    var p1 = parseInt(parts[1]);
    var p2 = parseInt(parts[2]);
    
    var day, month, year;
    
    // Determine year position
    if (p2 > 100) {
      year = p2;
      if (p0 > 12) {
        day = p0; month = p1 - 1;
      } else if (p1 > 12) {
        month = p0 - 1; day = p1;
      } else {
        day = p0; month = p1 - 1;
      }
    } else if (p0 > 100) {
      year = p0; month = p1 - 1; day = p2;
    } else {
      year = p2 < 50 ? 2000 + p2 : 1900 + p2;
      day = p0; month = p1 - 1;
    }
    
    var date = new Date(year, month, day);
    if (!isNaN(date.getTime())) return date;
  }
  
  console.log("Could not parse date: " + str);
  return null;
}

function extractMerchant(description) {
  if (!description) return '';
  return String(description)
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
  // Revenue - items that bring money IN
  { keywords: ['sales', 'revenue', 'payment received', 'invoice paid', 'stripe payout', 'paypal transfer', 'client payment'], group: 'Revenue', category: 'Sales Revenue' },
  { keywords: ['interest income', 'interest earned', 'interest credit'], group: 'Revenue', category: 'Interest Income' },
  { keywords: ['refund received', 'rebate', 'cashback'], group: 'Revenue', category: 'Other Income' },
  
  // COGS
  { keywords: ['inventory', 'stock purchase', 'raw material', 'supplies', 'manufacturing', 'cogs', 'cost of goods'], group: 'COGS', category: 'Inventory/Materials' },
  { keywords: ['freight', 'shipping cost', 'logistics', 'fedex', 'ups', 'dhl'], group: 'COGS', category: 'Freight & Shipping' },
  
  // Operating Expenses
  { keywords: ['salary', 'payroll', 'wages', 'bonus', 'compensation', 'gusto', 'adp'], group: 'Opex', category: 'Salaries & Wages' },
  { keywords: ['rent', 'lease', 'office space', 'wework', 'regus'], group: 'Opex', category: 'Rent' },
  { keywords: ['utility', 'electric', 'gas', 'water', 'internet', 'phone', 'comcast', 'verizon', 'at&t'], group: 'Opex', category: 'Utilities' },
  { keywords: ['software', 'subscription', 'saas', 'cloud', 'aws', 'azure', 'google cloud', 'microsoft', 'adobe', 'slack', 'zoom', 'shopify', 'loom', 'notion', 'figma', 'canva', 'hubspot', 'mailchimp', 'sendinblue', 'zoho', 'algolia', 'huggingface', 'openai', 'anthropic', 'github', 'atlassian', 'jira', 'asana', 'monday', 'dropbox', 'box', 'salesforce', 'quickbooks', 'xero', 'stripe', 'brex', 'ramp', 'paddle', 'gumroad', 'substack', 'convertkit', 'activecampaign', 'intercom', 'zendesk', 'freshdesk', 'twilio', 'sendgrid', 'postmark', 'cloudflare', 'vercel', 'netlify', 'heroku', 'digitalocean', 'linode', 'airtable', 'zapier', 'make', 'n8n', 'semrush', 'ahrefs', 'moz'], group: 'Opex', category: 'Software & Subscriptions' },
  { keywords: ['marketing', 'advertising', 'ads', 'facebook ads', 'google ads', 'promotion', 'dimabay', 'sortlist', 'linkedin ads', 'twitter ads', 'tiktok ads', 'meta ads', 'adwords', 'ppc', 'seo', 'content marketing', 'influencer', 'pr ', 'public relations', 'branding'], group: 'Opex', category: 'Marketing & Advertising' },
  { keywords: ['travel', 'flight', 'hotel', 'uber', 'lyft', 'taxi', 'airbnb', 'booking.com', 'expedia', 'airlines', 'train', 'rental car', 'hertz', 'enterprise'], group: 'Opex', category: 'Travel & Entertainment' },
  { keywords: ['meal', 'restaurant', 'food', 'lunch', 'dinner', 'coffee', 'doordash', 'grubhub', 'ubereats', 'postmates', 'starbucks', 'catering'], group: 'Opex', category: 'Meals & Entertainment' },
  { keywords: ['insurance', 'premium', 'liability', 'health insurance', 'dental', 'vision'], group: 'Opex', category: 'Insurance' },
  { keywords: ['legal', 'attorney', 'lawyer', 'law firm', 'paralegal', 'court', 'filing fee'], group: 'Opex', category: 'Legal & Professional' },
  { keywords: ['accounting', 'bookkeeping', 'audit', 'cpa', 'tax prep', 'tax preparation'], group: 'Opex', category: 'Accounting' },
  { keywords: ['consulting', 'consultant', 'advisory', 'coach', 'mentor'], group: 'Opex', category: 'Consulting' },
  { keywords: ['contractor', 'freelance', 'upwork', 'fiverr', 'toptal', '1099', 'gig'], group: 'Opex', category: 'Contractors' },
  { keywords: ['office supplies', 'staples', 'office depot', 'amazon', 'supplies'], group: 'Opex', category: 'Office Supplies' },
  { keywords: ['maintenance', 'repair', 'fix', 'service', 'cleaning'], group: 'Opex', category: 'Maintenance & Repairs' },
  { keywords: ['bank fee', 'service charge', 'wire fee', 'transaction fee', 'fx fee', 'currency conversion', 'atm fee', 'overdraft', 'monthly fee', 'account fee'], group: 'Opex', category: 'Bank Fees' },
  
  // Taxes
  { keywords: ['tax payment', 'irs', 'vat', 'gst', 'sales tax', 'income tax', 'payroll tax', 'quarterly tax', 'estimated tax', 'state tax', 'federal tax'], group: 'Taxes', category: 'Taxes' },
  
  // Financing
  { keywords: ['loan payment', 'interest payment', 'mortgage', 'financing', 'principal', 'line of credit', 'loc payment'], group: 'Financing', category: 'Loan Interest' },
  { keywords: ['dividend', 'distribution', 'shareholder'], group: 'Financing', category: 'Dividends' },
  
  // Transfers (excluded from P&L)
  { keywords: ['transfer', 'internal transfer', 'move funds', 'from savings', 'to savings', 'between accounts', 'ach transfer', 'wire transfer'], group: 'Transfers', category: 'Internal Transfer' },
  { keywords: ['owner draw', 'owner contribution', 'capital contribution', 'equity', 'investment'], group: 'Transfers', category: 'Owner Transactions' }
];

function categorizeTransaction(txn) {
  var descLower = (txn.description || '').toLowerCase();
  var merchantLower = (txn.merchant_key || '').toLowerCase();
  var typeLower = (txn.txn_type || '').toLowerCase();
  
  // Check transaction type first
  if (typeLower === 'deposit' || typeLower === 'credit' || typeLower === 'income') {
    // Likely revenue unless description says otherwise
    if (!descLower.includes('refund') && !descLower.includes('transfer')) {
      return { category_group: 'Revenue', category: 'Sales Revenue', method: 'rule', confidence: 0.7 };
    }
  }
  
  // Check against rules
  for (var i = 0; i < categoryRules.length; i++) {
    var rule = categoryRules[i];
    for (var j = 0; j < rule.keywords.length; j++) {
      var kw = rule.keywords[j];
      if (descLower.includes(kw) || merchantLower.includes(kw)) {
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

  parsedTransactions = parseCSV(csvData);
  
  if (parsedTransactions.length === 0) {
    setFinStatus("error", "No transactions found in CSV");
    toast("No transactions found");
    return;
  }

  // Categorize and add month
  for (var i = 0; i < parsedTransactions.length; i++) {
    var cat = categorizeTransaction(parsedTransactions[i]);
    parsedTransactions[i].category_group = cat.category_group;
    parsedTransactions[i].category = cat.category;
    parsedTransactions[i].method = cat.method;
    parsedTransactions[i].confidence = cat.confidence;
    
    // Add month (end of month for grouping)
    if (parsedTransactions[i].date) {
      var d = parsedTransactions[i].date;
      parsedTransactions[i].month = new Date(d.getFullYear(), d.getMonth() + 1, 0);
    }
  }

  writeToExcel();
}

// ============================================
// WRITE TO EXCEL - FIXED VERSION
// ============================================
function writeToExcel() {
  Excel.run(function(context) {
    var sheets = context.workbook.worksheets;
    sheets.load("items/name");

    return context.sync().then(function() {
      // Delete existing sheets
      var sheetsToDelete = ['CLEANED_TXN', 'P&L'];
      for (var i = 0; i < sheets.items.length; i++) {
        if (sheetsToDelete.indexOf(sheets.items[i].name) >= 0) {
          sheets.items[i].delete();
        }
      }
      return context.sync();
    }).then(function() {
      
      // ========== SHEET 1: CLEANED_TXN ==========
      var txnSheet = sheets.add("CLEANED_TXN");
      
      // UPDATED COLUMNS (removed account_name)
      var headers = [
        "txn_id",           // A - Original transaction ID
        "date",             // B - Transaction date
        "description",      // C - Description
        "txn_type",         // D - Transaction type
        "debit",            // E - Debit (outflow)
        "credit",           // F - Credit (inflow)  
        "amount",           // G - Net amount (FORMULA: =F-E)
        "currency",         // H - Currency
        "merchant_key",     // I - Merchant
        "category_group",   // J - Category group
        "category",         // K - Category
        "method",           // L - Method
        "confidence",       // M - Confidence
        "month"             // N - Month (for P&L)
      ];
      
      // Write headers
      var headerRange = txnSheet.getRange("A1:N1");
      headerRange.values = [headers];
      headerRange.format.font.bold = true;
      headerRange.format.fill.color = "#217346";
      headerRange.format.font.color = "#FFFFFF";

      // Prepare data rows
      var dataRows = [];
      for (var i = 0; i < parsedTransactions.length; i++) {
        var txn = parsedTransactions[i];
        
        dataRows.push([
          txn.txn_id,                                    // A
          txn.date ? excelDate(txn.date) : "",          // B - Date
          txn.description,                               // C
          txn.txn_type || "",                           // D
          txn.debit || 0,                               // E
          txn.credit || 0,                              // F
          null,                                          // G - Will be formula
          txn.currency,                                  // H
          txn.merchant_key,                              // I
          txn.category_group,                            // J
          txn.category,                                  // K
          txn.method,                                    // L
          txn.confidence,                                // M
          txn.month ? excelDate(txn.month) : ""         // N
        ]);
      }

      if (dataRows.length > 0) {
        var lastRow = dataRows.length + 1;
        var dataRange = txnSheet.getRange("A2:N" + lastRow);
        dataRange.values = dataRows;
        
        // Add AMOUNT formula for each row: =F-E (credit - debit)
        for (var i = 0; i < dataRows.length; i++) {
          var rowNum = i + 2;
          txnSheet.getRange("G" + rowNum).formulas = [["=F" + rowNum + "-E" + rowNum]];
        }
        
        // Format date column B
        txnSheet.getRange("B2:B" + lastRow).numberFormat = [["yyyy-mm-dd"]];
        
        // Format month column N
        txnSheet.getRange("N2:N" + lastRow).numberFormat = [["yyyy-mm"]];
        
        // Format currency columns E, F, G
        txnSheet.getRange("E2:G" + lastRow).numberFormat = [["#,##0.00"]];
        
        // Create Excel Table
        var tableRange = txnSheet.getRange("A1:N" + lastRow);
        var txnTable = txnSheet.tables.add(tableRange, true);
        txnTable.name = "tbl_cleaned_txn";
        
        // Auto-fit columns
        txnSheet.getRange("A:N").format.autofitColumns();
      }

      // ========== SHEET 2: P&L ==========
      var plSheet = sheets.add("P&L");
      
      // Collect unique months from parsed data
      var months = [];
      var monthSet = {};
      for (var i = 0; i < parsedTransactions.length; i++) {
        if (parsedTransactions[i].month) {
          var d = parsedTransactions[i].month;
          var monthKey = d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, '0');
          if (!monthSet[monthKey]) {
            monthSet[monthKey] = true;
            months.push(monthKey);
          }
        }
      }
      months.sort();
      
      console.log("P&L Months found:", months);

      // P&L Structure
      var plRows = [
        { name: "Revenue", group: "Revenue", isTotal: false },
        { name: "COGS", group: "COGS", isTotal: false },
        { name: "Gross Profit", group: null, isTotal: true, formula: "Revenue-COGS" },
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

      // Title
      plSheet.getRange("A1").values = [["P&L Summary"]];
      plSheet.getRange("A1").format.font.bold = true;
      plSheet.getRange("A1").format.font.size = 16;

      // Column headers
      plSheet.getRange("A3").values = [["Category"]];
      plSheet.getRange("A3").format.font.bold = true;
      
      // Month headers
      for (var m = 0; m < months.length; m++) {
        var col = getColumnLetter(m + 1);
        plSheet.getRange(col + "3").values = [[months[m]]];
        plSheet.getRange(col + "3").format.font.bold = true;
        plSheet.getRange(col + "3").format.horizontalAlignment = "Center";
      }
      
      // Total column
      var totalCol = getColumnLetter(months.length + 1);
      plSheet.getRange(totalCol + "3").values = [["TOTAL"]];
      plSheet.getRange(totalCol + "3").format.font.bold = true;

      // Write P&L rows with SUMIFS formulas
      var rowNum = 4;
      var rowRefs = {};
      
      for (var r = 0; r < plRows.length; r++) {
        var plRow = plRows[r];
        
        if (plRow.name === "") {
          rowNum++;
          continue;
        }

        plSheet.getRange("A" + rowNum).values = [[plRow.name]];
        
        if (plRow.isTotal) {
          plSheet.getRange("A" + rowNum).format.font.bold = true;
          var lastCol = getColumnLetter(months.length + 1);
          plSheet.getRange("A" + rowNum + ":" + lastCol + rowNum).format.fill.color = "#FFF2CC";
        }

        // Write formulas for each month
        for (var m = 0; m < months.length; m++) {
          var col = getColumnLetter(m + 1);
          var monthVal = months[m];
          
          if (plRow.group) {
            // SUMIFS formula using TEXT to match month format
            var formula = '=SUMIFS(tbl_cleaned_txn[amount],tbl_cleaned_txn[category_group],"' + plRow.group + '",TEXT(tbl_cleaned_txn[month],"YYYY-MM"),"' + monthVal + '")';
            plSheet.getRange(col + rowNum).formulas = [[formula]];
          } else if (plRow.isTotal) {
            // Calculate totals from other rows
            if (plRow.name === "Gross Profit") {
              plSheet.getRange(col + rowNum).formulas = [["=" + col + rowRefs["Revenue"] + "+" + col + rowRefs["COGS"]]];
            } else if (plRow.name === "Operating Income") {
              plSheet.getRange(col + rowNum).formulas = [["=" + col + rowRefs["Gross Profit"] + "+" + col + rowRefs["Operating Expenses"]]];
            } else if (plRow.name === "Net Income") {
              plSheet.getRange(col + rowNum).formulas = [["=" + col + rowRefs["Operating Income"] + "+" + col + rowRefs["Taxes"] + "+" + col + rowRefs["Financing"]]];
            }
          }
          
          plSheet.getRange(col + rowNum).numberFormat = [["#,##0.00"]];
        }
        
        // Total column - sum all months
        if ((plRow.group || plRow.isTotal) && months.length > 0) {
          var firstCol = getColumnLetter(1);
          var lastMonthCol = getColumnLetter(months.length);
          plSheet.getRange(totalCol + rowNum).formulas = [["=SUM(" + firstCol + rowNum + ":" + lastMonthCol + rowNum + ")"]];
          plSheet.getRange(totalCol + rowNum).numberFormat = [["#,##0.00"]];
          plSheet.getRange(totalCol + rowNum).format.font.bold = true;
        }

        rowRefs[plRow.name] = rowNum;
        rowNum++;
      }

      // Format P&L sheet
      plSheet.getRange("A:A").format.columnWidth = 180;
      for (var m = 0; m <= months.length; m++) {
        plSheet.getRange(getColumnLetter(m + 1) + ":" + getColumnLetter(m + 1)).format.columnWidth = 100;
      }

      // Activate CLEANED_TXN
      txnSheet.activate();

      return context.sync();
    }).then(function() {
      updateProgress(3);
      setFinStatus("ok", "Done! " + parsedTransactions.length + " transactions processed");
      toast("Analysis complete!");
    });
  }).catch(function(error) {
    setFinStatus("error", "Error: " + error.message);
    toast("Error: " + error.message);
    console.error(error);
  });
}
