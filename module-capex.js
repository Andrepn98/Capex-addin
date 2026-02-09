// ============================================
// ORNA MODULE: Capex & D&A Calculator
// ============================================

function initCapexModule() {
  document.getElementById("createTemplate").onclick = createCapexTemplate;
  document.getElementById("addAsset").onclick = addAsset;
  setCapexStatus("ok", "Ready");
}

function setCapexStatus(type, text) {
  document.getElementById("capexStatusDot").className = "dot" + (type ? " " + type : "");
  document.getElementById("capexStatusText").textContent = text;
}

// ============================================
// CREATE CAPEX TEMPLATE
// ============================================
function createCapexTemplate() {
  setCapexStatus("processing", "Creating template…");
  toast("Creating template…");

  Excel.run(function(context) {
    var startDateInput = document.getElementById("startDate").value;
    var numPeriods = parseInt(document.getElementById("numPeriods").value) || 36;
    
    var sheets = context.workbook.worksheets;
    sheets.load("items/name");
    
    return context.sync().then(function() {
      // Delete existing Capex sheet if exists
      for (var i = 0; i < sheets.items.length; i++) {
        if (sheets.items[i].name === "Capex") {
          sheets.items[i].delete();
          break;
        }
      }
      
      var sheet = sheets.add("Capex");
      var lastCol = getColumnLetter(7 + numPeriods - 1);
      var lastAssetRow = 50;
      
      // ROW 1: Title
      sheet.getRange("A1").values = [["Capex and D&A Calculator"]];
      sheet.getRange("A1").format.font.bold = true;
      sheet.getRange("A1").format.font.size = 16;
      sheet.getRange("A1").format.font.color = "#217346";
      
      // ROW 3: Start date row
      sheet.getRange("G3").values = [["Start date"]];
      sheet.getRange("G3").format.font.bold = true;
      sheet.getRange("H3").values = [[startDateInput]];
      sheet.getRange("H3").numberFormat = [["yyyy-mm-dd"]];
      
      for (var i = 1; i < numPeriods; i++) {
        var col = getColumnLetter(7 + i);
        var prevCol = getColumnLetter(6 + i);
        sheet.getRange(col + "3").formulas = [["=EDATE(" + prevCol + "3,1)"]];
        sheet.getRange(col + "3").numberFormat = [["yyyy-mm-dd"]];
      }
      
      // ROW 4: End date row
      sheet.getRange("G4").values = [["End date"]];
      sheet.getRange("G4").format.font.bold = true;
      for (var i = 0; i < numPeriods; i++) {
        var col = getColumnLetter(7 + i);
        sheet.getRange(col + "4").formulas = [["=EOMONTH(" + col + "3,0)"]];
        sheet.getRange(col + "4").numberFormat = [["yyyy-mm-dd"]];
      }
      
      // ROW 5: Period numbers
      sheet.getRange("G5").values = [["Period"]];
      sheet.getRange("G5").format.font.bold = true;
      for (var i = 0; i < numPeriods; i++) {
        var col = getColumnLetter(7 + i);
        sheet.getRange(col + "5").values = [[i + 1]];
      }
      
      // ROW 6: Total D&A
      sheet.getRange("A6").values = [["Total D&A"]];
      sheet.getRange("A6").format.font.bold = true;
      sheet.getRange("A6:G6").format.fill.color = "#FFF2CC";
      
      for (var p = 0; p < numPeriods; p++) {
        var col = getColumnLetter(7 + p);
        sheet.getRange(col + "6").formulas = [["=SUM(" + col + "9:" + col + lastAssetRow + ")"]];
        sheet.getRange(col + "6").numberFormat = [["#,##0.00"]];
        sheet.getRange(col + "6").format.font.bold = true;
        sheet.getRange(col + "6").format.fill.color = "#FFF2CC";
      }
      
      // ROW 7: Capex Total
      sheet.getRange("A7").values = [["Capex Total"]];
      sheet.getRange("A7").format.font.bold = true;
      sheet.getRange("A7:G7").format.fill.color = "#DDEBF7";
      
      for (var p = 0; p < numPeriods; p++) {
        var col = getColumnLetter(7 + p);
        sheet.getRange(col + "7").formulas = [["=SUMIF($C$9:$C$" + lastAssetRow + "," + col + "4,$E$9:$E$" + lastAssetRow + ")"]];
        sheet.getRange(col + "7").numberFormat = [["#,##0.00"]];
        sheet.getRange(col + "7").format.font.bold = true;
        sheet.getRange(col + "7").format.fill.color = "#DDEBF7";
      }
      
      // ROW 8: Column Headers
      var headers = sheet.getRange("A8:G8");
      headers.values = [["Assets", "Date", "End period", "Period", "Cost", "Useful Life", "Salvage"]];
      headers.format.font.bold = true;
      headers.format.fill.color = "#217346";
      headers.format.font.color = "#FFFFFF";
      
      for (var p = 0; p < numPeriods; p++) {
        var col = getColumnLetter(7 + p);
        sheet.getRange(col + "8").format.fill.color = "#217346";
      }
      
      // ROWS 9-50: Asset rows with formulas
      for (var row = 9; row <= lastAssetRow; row++) {
        sheet.getRange("C" + row).formulas = [['=IF(B' + row + '="","-",EOMONTH(B' + row + ',0))']];
        sheet.getRange("D" + row).formulas = [['=IF(C' + row + '="-",0,MATCH(C' + row + ',$H$4:$' + lastCol + '$4,0))']];
        
        for (var p = 0; p < numPeriods; p++) {
          var col = getColumnLetter(7 + p);
          var formula = '=IF(OR($A' + row + '="",$D' + row + '=0),0,IF(AND(' + (p+1) + '>=$D' + row + ',' + (p+1) + '<$D' + row + '+$F' + row + '),($E' + row + '-$G' + row + ')/$F' + row + ',0))';
          sheet.getRange(col + row).formulas = [[formula]];
          sheet.getRange(col + row).numberFormat = [["#,##0.00"]];
        }
      }
      
      // Formatting
      sheet.getRange("A:G").format.autofitColumns();
      sheet.activate();
      return context.sync();
    }).then(function() {
      setCapexStatus("ok", "Template created!");
      toast("Template created!");
    });
  }).catch(function(error) {
    setCapexStatus("error", "Error: " + error.message);
    toast("Error creating template");
    console.error(error);
  });
}

// ============================================
// ADD ASSET
// ============================================
function addAsset() {
  var assetName = document.getElementById("assetName").value;
  var assetCost = document.getElementById("assetCost").value;

  if (!assetName || !assetCost) {
    setCapexStatus("warn", "Please fill in asset name and cost");
    toast("Missing required fields");
    return;
  }

  setCapexStatus("processing", "Adding asset…");
  toast("Adding asset…");

  Excel.run(function(context) {
    var assetDate = document.getElementById("assetDate").value;
    var usefulLife = document.getElementById("usefulLife").value;
    var salvageValue = document.getElementById("salvageValue").value || "0";
    
    var sheet = context.workbook.worksheets.getItem("Capex");
    var searchRange = sheet.getRange("A9:A50");
    searchRange.load("values");
    
    return context.sync().then(function() {
      var targetRow = -1;
      for (var i = 0; i < searchRange.values.length; i++) {
        if (!searchRange.values[i][0]) {
          targetRow = 9 + i;
          break;
        }
      }
      
      if (targetRow === -1) {
        setCapexStatus("warn", "No empty rows available");
        return context.sync();
      }
      
      sheet.getRange("A" + targetRow).values = [[assetName]];
      sheet.getRange("B" + targetRow).values = [[assetDate]];
      sheet.getRange("E" + targetRow).values = [[parseFloat(assetCost)]];
      sheet.getRange("F" + targetRow).values = [[parseInt(usefulLife)]];
      sheet.getRange("G" + targetRow).values = [[parseFloat(salvageValue)]];
      
      return context.sync();
    }).then(function() {
      document.getElementById("assetName").value = "";
      document.getElementById("assetCost").value = "";
      setCapexStatus("ok", "Asset added!");
      toast("Asset added!");
    });
  }).catch(function(error) {
    setCapexStatus("error", "Error: " + error.message);
    toast("Error adding asset");
    console.error(error);
  });
}
