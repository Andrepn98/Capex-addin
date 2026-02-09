// ============================================
// ORNA CORE - Shared utilities
// ============================================

// Page Navigation
function showPage(pageId) {
  document.querySelectorAll('.page').forEach(function(page) {
    page.classList.remove('active');
  });
  document.getElementById(pageId).classList.add('active');
}

// Toast notifications
var toastTimer = null;
function toast(msg) {
  var el = document.getElementById("toast");
  el.textContent = msg;
  el.classList.add("show");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(function() { el.classList.remove("show"); }, 2500);
}

// Helper: Get Excel column letter from index (0-based)
function getColumnLetter(index) {
  var letter = "";
  index = index + 1;
  while (index > 0) {
    var remainder = (index - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    index = Math.floor((index - 1) / 26);
  }
  return letter;
}

// Helper: Convert JS Date to Excel serial number
function excelDate(date) {
  if (!date || !(date instanceof Date) || isNaN(date.getTime())) {
    return "";
  }
  var epoch = new Date(1899, 11, 30);
  var msPerDay = 24 * 60 * 60 * 1000;
  return Math.floor((date - epoch) / msPerDay);
}

// Office.js Initialization
Office.onReady(function(info) {
  if (info.host === Office.HostType.Excel) {
    console.log("ORNA: Connected to Excel");
    
    // Initialize Capex Module
    if (typeof initCapexModule === 'function') {
      initCapexModule();
    }
    
    // Initialize Financial Module
    if (typeof initFinancialModule === 'function') {
      initFinancialModule();
    }
  }
});
