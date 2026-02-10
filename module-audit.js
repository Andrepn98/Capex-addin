// ============================================
// ORNA MODULE: Financial Model Audit Tool v2.0
// ============================================
// Based on VBA modModelAudit v2.0
//
// FEATURES:
//   - Auto-detects period row (1,2,3... or sequential dates)
//   - Detects total/summary columns (not flagged as breaks)
//   - Uses row-dominant formula pattern (most common R1C1)
//   - Skips blank rows
//   - Pulls source row labels into the map
//   - Color-coded visual audit map
//
// MAP LEGEND:
//   .  = Consistent formula (matches row dominant pattern)
//   H  = Hardcoded constant in formula/period zone
//   X  = Pattern break (formula differs from dominant)
//   E  = Cell contains Excel error (#REF! #DIV/0! etc.)
//   S  = Summary/total column (expected break, not an issue)
//   ~  = Label zone cell (before period start column)
// ============================================

// ============================================
// CONFIGURATION
// ============================================
var auditConfig = {
    audPrefix: "AUD_",
    dashboardName: "AUDIT_DASHBOARD",
    periodScanRows: 15,      // Rows scanned for period detection
    labelCols: 4,            // Columns A-D scanned for row labels
    minFormulasForDominant: 3, // Need >= this many formulas to define "dominant"
    maxSheetNameLength: 31
};

// Colors (hex)
var auditColors = {
    clean: "#DAF2DA",        // Light green
    hardcode: "#FFC000",     // Orange
    break: "#FF8080",        // Soft red
    error: "#FF0000",        // Red
    summary: "#FFF2CC",      // Pale gold
    label: "#F2F2F2",        // Light gray
    headerBg: "#44546A",     // Dark blue
    white: "#FFFFFF"
};

// ============================================
// GLOBAL STATE
// ============================================
var auditState = {
    results: [],             // SheetResult objects
    totalIssues: 0,
    totalHardcodes: 0,
    totalBreaks: 0,
    totalErrors: 0
};

// ============================================
// INITIALIZE MODULE
// ============================================
function initAuditModule() {
    document.getElementById('runFullAudit').onclick = runModelAudit;
    document.getElementById('runQuickAudit').onclick = runQuickAudit;
    document.getElementById('auditSelection').onclick = auditSelectedRange;
    setAuditStatus("ok", "Ready to audit");
}

function setAuditStatus(type, text) {
    var dot = document.getElementById("auditStatusDot");
    var txt = document.getElementById("auditStatusText");
    if (dot) dot.className = "dot" + (type ? " " + type : "");
    if (txt) txt.textContent = text;
}

function updateAuditProgress(percent, message) {
    var bar = document.getElementById("auditProgress");
    var txt = document.getElementById("auditProgressText");
    if (bar) bar.style.width = percent + "%";
    if (txt) txt.textContent = message;
}

// ============================================
// MAIN ENTRY POINT
// ============================================
function runModelAudit() {
    setAuditStatus("processing", "Running model audit...");
    updateAuditProgress(0, "Initializing...");
    toast("Starting model audit...");

    // Reset state
    auditState = {
        results: [],
        totalIssues: 0,
        totalHardcodes: 0,
        totalBreaks: 0,
        totalErrors: 0
    };

    Excel.run(function(context) {
        var workbook = context.workbook;
        var sheets = workbook.worksheets;
        sheets.load("items/name");

        return context.sync().then(function() {
            updateAuditProgress(5, "Analyzing workbook structure...");

            // Filter sheets to audit
            var sheetsToAudit = sheets.items.filter(function(sheet) {
                return !isAuditSheet(sheet.name);
            });

            if (sheetsToAudit.length === 0) {
                throw new Error("No sheets to audit.");
            }

            // Purge old audit sheets
            return purgeOldAuditSheets(context, sheets.items);
        }).then(function() {
            updateAuditProgress(10, "Auditing sheets...");

            // Re-load sheets after purge
            var sheets = context.workbook.worksheets;
            sheets.load("items/name");
            return context.sync().then(function() {
                return sheets.items;
            });
        }).then(function(allSheets) {
            var sheetsToAudit = allSheets.filter(function(sheet) {
                return !isAuditSheet(sheet.name);
            });

            // Audit each sheet sequentially
            var progressPerSheet = 70 / Math.max(sheetsToAudit.length, 1);
            var auditPromise = Promise.resolve();

            sheetsToAudit.forEach(function(sheet, index) {
                auditPromise = auditPromise.then(function() {
                    updateAuditProgress(10 + (index * progressPerSheet), "Auditing: " + sheet.name);
                    return auditOneSheet(context, sheet);
                }).then(function(result) {
                    if (result) {
                        auditState.results.push(result);
                        auditState.totalHardcodes += result.nHardcode;
                        auditState.totalBreaks += result.nBreak;
                        auditState.totalErrors += result.nError;
                        auditState.totalIssues += result.nIssues;
                    }
                });
            });

            return auditPromise;
        }).then(function() {
            updateAuditProgress(85, "Building dashboard...");
            return buildDashboard(context);
        }).then(function() {
            updateAuditProgress(100, "Complete!");

            // Activate dashboard
            var dashboard = context.workbook.worksheets.getItem(auditConfig.dashboardName);
            dashboard.activate();

            return context.sync();
        });
    }).then(function() {
        setAuditStatus("ok", "Audit complete! " + auditState.totalIssues + " issues found");
        toast("Audit complete! See " + auditConfig.dashboardName);
        displayAuditSummary();
    }).catch(function(error) {
        setAuditStatus("error", "Error: " + error.message);
        toast("Audit failed: " + error.message);
        console.error(error);
    });
}

// ============================================
// QUICK AUDIT (Errors only)
// ============================================
function runQuickAudit() {
    setAuditStatus("processing", "Running quick audit...");
    updateAuditProgress(0, "Scanning for errors...");
    toast("Quick audit - checking for errors only...");

    auditState = {
        results: [],
        totalIssues: 0,
        totalHardcodes: 0,
        totalBreaks: 0,
        totalErrors: 0
    };

    Excel.run(function(context) {
        var sheets = context.workbook.worksheets;
        sheets.load("items/name");

        return context.sync().then(function() {
            var sheetsToAudit = sheets.items.filter(function(s) {
                return !isAuditSheet(s.name);
            });

            var promises = sheetsToAudit.map(function(sheet, index) {
                return quickAuditSheet(context, sheet).then(function(result) {
                    if (result) {
                        auditState.results.push(result);
                        auditState.totalErrors += result.nError;
                        auditState.totalIssues += result.nError;
                    }
                    updateAuditProgress((index + 1) / sheetsToAudit.length * 90, "Checked: " + sheet.name);
                });
            });

            return Promise.all(promises);
        }).then(function() {
            updateAuditProgress(100, "Done!");
            return context.sync();
        });
    }).then(function() {
        setAuditStatus("ok", "Quick audit done! " + auditState.totalErrors + " errors found");
        toast("Quick audit complete!");
        displayAuditSummary();
    }).catch(function(error) {
        setAuditStatus("error", "Error: " + error.message);
        console.error(error);
    });
}

// ============================================
// AUDIT SELECTED RANGE
// ============================================
function auditSelectedRange() {
    setAuditStatus("processing", "Auditing selection...");

    Excel.run(function(context) {
        var range = context.workbook.getSelectedRange();
        range.load("address, values, formulas, formulasR1C1, worksheet/name, rowCount, columnCount");

        return context.sync().then(function() {
            var issues = [];
            var sheetName = range.worksheet.name;
            var values = range.values;
            var formulas = range.formulas;
            var formulasR1C1 = range.formulasR1C1;

            // Build dominant pattern per row
            for (var r = 0; r < values.length; r++) {
                var dominant = getDominantPattern(formulasR1C1, r, 0, values[r].length - 1, {});

                for (var c = 0; c < values[r].length; c++) {
                    var value = values[r][c];
                    var formula = formulas[r][c];
                    var formulaR1C1 = formulasR1C1[r][c];
                    var cellAddr = getCellAddressFromSelection(range.address, r, c);
                    var mark = classifyCell(value, formula, formulaR1C1, dominant, false);

                    if (mark === "H" || mark === "X" || mark === "E") {
                        issues.push({
                            type: mark,
                            sheet: sheetName,
                            cell: cellAddr,
                            detail: mark === "H" ? String(value) : formula
                        });
                    }
                }
            }

            return issues;
        });
    }).then(function(issues) {
        setAuditStatus("ok", "Found " + issues.length + " issues in selection");
        
        if (issues.length === 0) {
            toast("No issues found in selection!");
        } else {
            toast(issues.length + " issues found");
            
            // Update summary display
            auditState.results = [{
                sheetName: "Selection",
                nHardcode: issues.filter(function(i) { return i.type === "H"; }).length,
                nBreak: issues.filter(function(i) { return i.type === "X"; }).length,
                nError: issues.filter(function(i) { return i.type === "E"; }).length,
                nIssues: issues.length,
                issues: issues
            }];
            auditState.totalIssues = issues.length;
            displayAuditSummary();
        }
    }).catch(function(error) {
        setAuditStatus("error", "Error: " + error.message);
        console.error(error);
    });
}

// ============================================
// AUDIT ONE SHEET (Full Analysis)
// ============================================
function auditOneSheet(context, sheet) {
    return new Promise(function(resolve, reject) {
        var result = {
            sheetName: sheet.name,
            auditName: safeAuditName(sheet.name),
            periodRow: 0,
            periodCol: 0,
            lastDataRow: 0,
            lastDataCol: 0,
            nHardcode: 0,
            nBreak: 0,
            nError: 0,
            nIssues: 0,
            issues: [],
            totalCols: {}
        };

        var usedRange = sheet.getUsedRange();
        usedRange.load("values, formulas, formulasR1C1, rowCount, columnCount, address");

        context.sync().then(function() {
            var values = usedRange.values;
            var formulas = usedRange.formulas;
            var formulasR1C1 = usedRange.formulasR1C1;
            var nRows = usedRange.rowCount;
            var nCols = usedRange.columnCount;

            result.lastDataRow = nRows;
            result.lastDataCol = nCols;

            // Detect period row and column
            var periodInfo = detectPeriod(values, nRows, nCols);
            result.periodRow = periodInfo.row;
            result.periodCol = periodInfo.col;

            var inspectFrom = result.periodCol > 0 ? result.periodCol - 1 : 0; // 0-based

            // Detect total/summary columns
            if (result.periodRow > 0 && result.periodCol > 0) {
                result.totalCols = detectTotalColumns(values, result.periodRow - 1, inspectFrom, nCols);
            }

            // Prepare audit map data
            var mapData = [];
            var mapHeaders = ["Row", "Label", "Chk"];
            
            // Add column headers for the grid
            for (var gc = inspectFrom; gc < nCols; gc++) {
                mapHeaders.push(numberToColumnLetter(gc + 1));
            }
            mapData.push(mapHeaders);

            // Main audit loop - row by row
            for (var r = 0; r < nRows; r++) {
                // Skip blank rows
                if (isRowBlank(values, r, inspectFrom, nCols)) continue;

                // Get row label
                var rowLabel = getRowLabel(values, r);

                // Get dominant pattern for this row
                var dominant = getDominantPattern(formulasR1C1, r, inspectFrom, nCols - 1, result.totalCols);

                var rowIssues = 0;
                var rowMarks = [r + 1, rowLabel, ""]; // Row number, label, check mark (filled later)

                for (var c = inspectFrom; c < nCols; c++) {
                    var value = values[r][c];
                    var formula = formulas[r][c];
                    var formulaR1C1 = formulasR1C1[r][c];
                    var isTotalCol = result.totalCols[c] === true;
                    
                    var mark = classifyCell(value, formula, formulaR1C1, dominant, isTotalCol);
                    rowMarks.push(mark);

                    // Collect issues
                    if (mark === "H") {
                        result.nHardcode++;
                        rowIssues++;
                        result.issues.push({
                            type: "H",
                            cell: numberToColumnLetter(c + 1) + (r + 1),
                            detail: truncateValue(value)
                        });
                    } else if (mark === "X") {
                        result.nBreak++;
                        rowIssues++;
                        result.issues.push({
                            type: "X",
                            cell: numberToColumnLetter(c + 1) + (r + 1),
                            detail: truncateFormula(formula)
                        });
                    } else if (mark === "E") {
                        result.nError++;
                        rowIssues++;
                        result.issues.push({
                            type: "E",
                            cell: numberToColumnLetter(c + 1) + (r + 1),
                            detail: String(value)
                        });
                    }
                }

                // Row health indicator
                rowMarks[2] = rowIssues === 0 ? "✓" : String(rowIssues);
                mapData.push(rowMarks);
            }

            result.nIssues = result.nHardcode + result.nBreak + result.nError;

            // Create audit map sheet
            return createAuditMapSheet(context, result, mapData, inspectFrom);
        }).then(function() {
            resolve(result);
        }).catch(function(error) {
            console.error("Error auditing " + sheet.name + ": " + error.message);
            resolve(result); // Return partial result
        });
    });
}

// ============================================
// QUICK AUDIT SHEET (Errors Only)
// ============================================
function quickAuditSheet(context, sheet) {
    return new Promise(function(resolve) {
        var result = {
            sheetName: sheet.name,
            nError: 0
        };

        var usedRange = sheet.getUsedRange();
        usedRange.load("values, rowCount, columnCount");

        context.sync().then(function() {
            var values = usedRange.values;
            
            for (var r = 0; r < values.length; r++) {
                for (var c = 0; c < values[r].length; c++) {
                    if (isErrorValue(values[r][c])) {
                        result.nError++;
                    }
                }
            }
            
            resolve(result);
        }).catch(function() {
            resolve(result);
        });
    });
}

// ============================================
// PERIOD DETECTION
// ============================================
function detectPeriod(values, nRows, nCols) {
    var result = { row: 0, col: 0 };
    var scanRows = Math.min(auditConfig.periodScanRows, nRows);
    var bestRun = 0;

    for (var r = 0; r < scanRows; r++) {
        for (var c = 0; c < nCols - 2; c++) {
            var v1 = values[r][c];
            var v2 = values[r][c + 1];
            var v3 = values[r][c + 2];

            // Look for 1, 2, 3 sequence
            if (isNumeric(v1) && isNumeric(v2) && isNumeric(v3)) {
                if (Math.floor(v1) === 1 && Math.floor(v2) === 2 && Math.floor(v3) === 3) {
                    // Count how long the run goes
                    var runLen = 3;
                    for (var k = c + 3; k < nCols; k++) {
                        var vk = values[r][k];
                        if (isNumeric(vk)) {
                            if (Math.floor(vk) === runLen + 1) {
                                runLen++;
                            } else if (Math.floor(vk) > runLen + 1 && Math.floor(vk) <= runLen + 2) {
                                // Allow small gap (total column)
                                runLen = Math.floor(vk);
                            } else {
                                break;
                            }
                        }
                        // Non-numeric in between is OK (total column)
                    }

                    if (runLen > bestRun) {
                        bestRun = runLen;
                        result.row = r + 1;  // 1-based
                        result.col = c + 1;  // 1-based
                    }
                }
            }
        }
    }

    return result;
}

// ============================================
// TOTAL/SUMMARY COLUMN DETECTION
// ============================================
function detectTotalColumns(values, periodRowIndex, startCol, endCol) {
    var totalCols = {};
    var periodRow = values[periodRowIndex];
    
    if (!periodRow) return totalCols;

    // Find extent of period numbering
    var endPeriodCol = startCol;
    for (var c = startCol; c < endCol; c++) {
        if (isNumeric(periodRow[c]) && periodRow[c] !== null && periodRow[c] !== "") {
            endPeriodCol = c;
        }
    }

    // Any non-numeric column between start and end is a total column
    var scanTo = Math.min(endPeriodCol + 1, endCol - 1);
    
    for (var c = startCol; c <= scanTo; c++) {
        var v = periodRow[c];
        if (v === null || v === "" || v === undefined) {
            totalCols[c] = true;
        } else if (!isNumeric(v)) {
            totalCols[c] = true;
        }
    }

    return totalCols;
}

// ============================================
// DOMINANT PATTERN DETECTION
// ============================================
function getDominantPattern(formulasR1C1, rowIndex, startCol, endCol, totalCols) {
    var patternCounts = {};
    var nFormulas = 0;

    for (var c = startCol; c <= endCol; c++) {
        if (totalCols[c]) continue;

        var raw = formulasR1C1[rowIndex][c];
        if (raw === null || raw === undefined) continue;
        if (typeof raw !== "string") continue;
        
        var key = String(raw);
        if (key.length === 0 || key.charAt(0) !== "=") continue;

        nFormulas++;
        patternCounts[key] = (patternCounts[key] || 0) + 1;
    }

    if (nFormulas < auditConfig.minFormulasForDominant) {
        return "";
    }

    // Find most common pattern
    var maxCount = 0;
    var maxKey = "";
    for (var key in patternCounts) {
        if (patternCounts[key] > maxCount) {
            maxCount = patternCounts[key];
            maxKey = key;
        }
    }

    // Only accept as dominant if it covers a decent share
    if (maxCount >= nFormulas * 0.4) {
        return maxKey;
    }

    return ""; // Too fragmented
}

// ============================================
// CELL CLASSIFICATION
// ============================================
function classifyCell(value, formula, formulaR1C1, dominant, isTotalCol) {
    var isFormulaCell = (typeof formula === "string" && formula.length > 0 && formula.charAt(0) === "=");

    // Empty cell
    if ((value === null || value === "" || value === undefined) && !isFormulaCell) {
        return "";
    }

    // Error value
    if (isErrorValue(value)) {
        return "E";
    }

    // Total/summary column
    if (isTotalCol) {
        return "S";
    }

    // Hardcode (not a formula, has a value)
    if (!isFormulaCell) {
        return "H";
    }

    // Formula - check against dominant pattern
    if (dominant.length > 0 && formulaR1C1 !== dominant) {
        return "X";
    }

    return ".";
}

// ============================================
// CREATE AUDIT MAP SHEET
// ============================================
function createAuditMapSheet(context, result, mapData, inspectFrom) {
    return new Promise(function(resolve, reject) {
        var sheets = context.workbook.worksheets;
        var audSheet = sheets.add(result.auditName);

        // Title
        audSheet.getRange("A1").values = [["AUDIT MAP: " + result.sheetName]];
        audSheet.getRange("A1").format.font.bold = true;
        audSheet.getRange("A1").format.font.size = 13;

        // Info row
        var infoText = "";
        if (result.periodCol > 0) {
            infoText = "Period detected: row " + result.periodRow + ", starts col " + numberToColumnLetter(result.periodCol);
        } else {
            infoText = "No period row detected - auditing from col " + numberToColumnLetter(inspectFrom + 1);
        }
        
        var totalColsList = Object.keys(result.totalCols);
        if (totalColsList.length > 0) {
            infoText += "  |  Total cols: " + totalColsList.map(function(c) {
                return numberToColumnLetter(parseInt(c) + 1);
            }).join(" ");
        }
        
        audSheet.getRange("A2").values = [[infoText]];
        audSheet.getRange("A2").format.font.color = "#646464";

        // Legend
        var legends = [". = OK", "H = Hardcode", "X = Break", "E = Error", "S = Total col"];
        for (var i = 0; i < legends.length; i++) {
            var cell = audSheet.getRange(numberToColumnLetter(i + 1) + "3");
            cell.values = [[legends[i]]];
            cell.format.font.size = 8;
            cell.format.font.color = "#787878";
        }

        // Write map data starting at row 5
        if (mapData.length > 0) {
            var dataRange = audSheet.getRange("A5").getResizedRange(mapData.length - 1, mapData[0].length - 1);
            dataRange.values = mapData;

            // Format header row
            var headerRange = audSheet.getRange("A5").getResizedRange(0, mapData[0].length - 1);
            headerRange.format.font.bold = true;
            headerRange.format.fill.color = auditColors.headerBg;
            headerRange.format.font.color = auditColors.white;

            // Format data rows (apply colors based on marks)
            for (var r = 1; r < mapData.length; r++) {
                var excelRow = 5 + r;
                
                // Row check column color
                var checkCell = audSheet.getRange("C" + excelRow);
                if (mapData[r][2] === "✓") {
                    checkCell.format.font.color = "#008C00";
                } else {
                    checkCell.format.font.color = "#C80000";
                    checkCell.format.font.bold = true;
                }

                // Data cells (starting from column D)
                for (var c = 3; c < mapData[r].length; c++) {
                    var mark = mapData[r][c];
                    var colLetter = numberToColumnLetter(c + 1);
                    var cell = audSheet.getRange(colLetter + excelRow);

                    cell.format.horizontalAlignment = "Center";
                    cell.format.font.size = 9;
                    cell.format.font.bold = true;

                    switch (mark) {
                        case ".":
                            cell.format.fill.color = auditColors.clean;
                            cell.format.font.color = "#64AA64";
                            break;
                        case "H":
                            cell.format.fill.color = auditColors.hardcode;
                            cell.format.font.color = "#824600";
                            break;
                        case "X":
                            cell.format.fill.color = auditColors.break;
                            cell.format.font.color = auditColors.white;
                            break;
                        case "E":
                            cell.format.fill.color = auditColors.error;
                            cell.format.font.color = auditColors.white;
                            break;
                        case "S":
                            cell.format.fill.color = auditColors.summary;
                            cell.format.font.color = "#8C8250";
                            break;
                    }
                }
            }
        }

        // Set column widths
        audSheet.getRange("A:A").format.columnWidth = 40;
        audSheet.getRange("B:B").format.columnWidth = 150;
        audSheet.getRange("C:C").format.columnWidth = 30;

        // Freeze panes
        audSheet.freezePanes.freezeRows(5);

        context.sync().then(function() {
            resolve();
        }).catch(reject);
    });
}

// ============================================
// BUILD DASHBOARD
// ============================================
function buildDashboard(context) {
    return new Promise(function(resolve, reject) {
        var sheets = context.workbook.worksheets;
        var dash = sheets.add(auditConfig.dashboardName);
        dash.position = 0;

        // Title
        dash.getRange("A1").values = [["MODEL AUDIT DASHBOARD"]];
        dash.getRange("A1").format.font.bold = true;
        dash.getRange("A1").format.font.size = 16;
        dash.getRange("A1").format.font.color = "#1E3C64";

        dash.getRange("A2").values = [["Generated: " + new Date().toLocaleString()]];
        dash.getRange("A2").format.font.color = "#787878";

        // Scoreboard
        writeScoreBox(dash, "I1", "Total Issues", auditState.totalIssues, 
            auditState.totalIssues === 0 ? auditColors.clean : "#FFC7CE");
        writeScoreBox(dash, "K1", "Hardcodes", auditState.totalHardcodes, "#FFE6B4");
        writeScoreBox(dash, "M1", "Breaks", auditState.totalBreaks, "#FFC7CE");
        writeScoreBox(dash, "O1", "Errors", auditState.totalErrors, 
            auditState.totalErrors === 0 ? auditColors.clean : auditColors.error);

        // Sheet summary table
        var summaryHeaders = [["Sheet", "Period Col", "Hardcodes (H)", "Breaks (X)", "Errors (E)", "Total Issues", "Audit Map"]];
        dash.getRange("A4:G4").values = summaryHeaders;
        dash.getRange("A4:G4").format.font.bold = true;
        dash.getRange("A4:G4").format.fill.color = auditColors.headerBg;
        dash.getRange("A4:G4").format.font.color = auditColors.white;

        // Write sheet results
        for (var i = 0; i < auditState.results.length; i++) {
            var res = auditState.results[i];
            var rowNum = 5 + i;
            var rowRange = dash.getRange("A" + rowNum + ":G" + rowNum);

            var periodText = res.periodCol > 0 
                ? numberToColumnLetter(res.periodCol) + " (row " + res.periodRow + ")"
                : "N/A";

            rowRange.values = [[
                res.sheetName,
                periodText,
                res.nHardcode || 0,
                res.nBreak || 0,
                res.nError || 0,
                res.nIssues || 0,
                "→ " + (res.auditName || "N/A")
            ]];

            // Color total issues cell
            var totalCell = dash.getRange("F" + rowNum);
            if ((res.nIssues || 0) === 0) {
                totalCell.format.fill.color = auditColors.clean;
            } else if ((res.nIssues || 0) <= 10) {
                totalCell.format.fill.color = "#FFEB9C";
            } else {
                totalCell.format.fill.color = "#FFC7CE";
            }
            totalCell.format.font.bold = true;

            // Add hyperlink styling to audit map
            if (res.auditName) {
                var linkCell = dash.getRange("G" + rowNum);
                linkCell.format.font.color = "#0563C1";
                linkCell.format.font.underline = "Single";
            }

            // Alternating row colors
            if (i % 2 === 1) {
                dash.getRange("A" + rowNum + ":E" + rowNum).format.fill.color = "#F2F2F2";
            }
        }

        // Totals row
        var totRow = 5 + auditState.results.length;
        dash.getRange("A" + totRow + ":G" + totRow).values = [[
            "TOTAL", "", 
            auditState.totalHardcodes, 
            auditState.totalBreaks, 
            auditState.totalErrors, 
            auditState.totalIssues, 
            ""
        ]];
        dash.getRange("A" + totRow + ":G" + totRow).format.font.bold = true;

        // Issue detail section
        var issueHeaderRow = totRow + 2;
        dash.getRange("A" + issueHeaderRow).values = [["ISSUE DETAIL"]];
        dash.getRange("A" + issueHeaderRow).format.font.bold = true;
        dash.getRange("A" + issueHeaderRow).format.font.size = 13;

        var issueTableHeader = issueHeaderRow + 1;
        dash.getRange("A" + issueTableHeader + ":E" + issueTableHeader).values = [[
            "Type", "Sheet", "Cell", "Detail / Formula", "Jump"
        ]];
        dash.getRange("A" + issueTableHeader + ":E" + issueTableHeader).format.font.bold = true;
        dash.getRange("A" + issueTableHeader + ":E" + issueTableHeader).format.fill.color = auditColors.headerBg;
        dash.getRange("A" + issueTableHeader + ":E" + issueTableHeader).format.font.color = auditColors.white;

        // Write all issues
        var issueRow = issueTableHeader + 1;
        for (var i = 0; i < auditState.results.length; i++) {
            var res = auditState.results[i];
            if (!res.issues) continue;

            for (var j = 0; j < res.issues.length; j++) {
                var issue = res.issues[j];
                var detail = issue.detail || "";
                if (detail.length > 80) detail = detail.substring(0, 80) + "...";

                dash.getRange("A" + issueRow + ":E" + issueRow).values = [[
                    issue.type,
                    res.sheetName,
                    issue.cell,
                    detail,
                    "Go"
                ]];

                // Type color
                var typeCell = dash.getRange("A" + issueRow);
                switch (issue.type) {
                    case "H":
                        typeCell.format.fill.color = auditColors.hardcode;
                        break;
                    case "X":
                        typeCell.format.fill.color = auditColors.break;
                        typeCell.format.font.color = auditColors.white;
                        break;
                    case "E":
                        typeCell.format.fill.color = auditColors.error;
                        typeCell.format.font.color = auditColors.white;
                        break;
                }

                // Jump link styling
                dash.getRange("E" + issueRow).format.font.color = "#0563C1";

                issueRow++;
            }
        }

        // Column widths
        dash.getRange("A:A").format.columnWidth = 60;
        dash.getRange("B:B").format.columnWidth = 120;
        dash.getRange("C:C").format.columnWidth = 60;
        dash.getRange("D:D").format.columnWidth = 250;
        dash.getRange("E:E").format.columnWidth = 40;
        dash.getRange("F:F").format.columnWidth = 80;
        dash.getRange("G:G").format.columnWidth = 120;

        context.sync().then(resolve).catch(reject);
    });
}

// ============================================
// HELPER: Write Score Box
// ============================================
function writeScoreBox(sheet, startCell, title, value, bgColor) {
    var titleCell = sheet.getRange(startCell);
    titleCell.values = [[title]];
    titleCell.format.font.size = 9;
    titleCell.format.font.bold = true;
    titleCell.format.font.color = "#505050";

    // Value is one row below
    var col = startCell.replace(/[0-9]/g, '');
    var row = parseInt(startCell.replace(/[A-Z]/gi, ''));
    var valueCell = sheet.getRange(col + (row + 1));
    valueCell.values = [[value]];
    valueCell.format.font.size = 22;
    valueCell.format.font.bold = true;
    valueCell.format.fill.color = bgColor;
    valueCell.format.horizontalAlignment = "Center";
}

// ============================================
// UTILITY FUNCTIONS
// ============================================
function isAuditSheet(name) {
    var upper = name.toUpperCase();
    return upper.indexOf(auditConfig.audPrefix.toUpperCase()) === 0 ||
           upper === auditConfig.dashboardName.toUpperCase();
}

function safeAuditName(srcName) {
    var name = auditConfig.audPrefix + srcName;
    if (name.length > auditConfig.maxSheetNameLength) {
        name = name.substring(0, auditConfig.maxSheetNameLength);
    }
    // Remove invalid characters
    name = name.replace(/[:\\\/?*\[\]]/g, "_");
    return name;
}

function purgeOldAuditSheets(context, sheets) {
    var toDelete = sheets.filter(function(s) {
        return isAuditSheet(s.name);
    });

    toDelete.forEach(function(s) {
        s.delete();
    });

    return context.sync();
}

function isNumeric(value) {
    return typeof value === "number" && !isNaN(value);
}

function isErrorValue(value) {
    if (typeof value === "string") {
        return value.charAt(0) === "#" && (
            value === "#DIV/0!" || value === "#N/A" || value === "#NAME?" ||
            value === "#NULL!" || value === "#NUM!" || value === "#REF!" ||
            value === "#VALUE!" || value === "#CALC!" || value === "#SPILL!"
        );
    }
    return false;
}

function isRowBlank(values, rowIndex, startCol, endCol) {
    for (var c = startCol; c < endCol; c++) {
        var v = values[rowIndex][c];
        if (v !== null && v !== "" && v !== undefined) {
            return false;
        }
    }
    return true;
}

function getRowLabel(values, rowIndex) {
    var best = "";
    for (var c = 0; c < Math.min(auditConfig.labelCols, values[rowIndex].length); c++) {
        var v = values[rowIndex][c];
        if (v !== null && v !== "" && v !== undefined && !isErrorValue(v)) {
            var s = String(v).trim();
            if (s.length > best.length) best = s;
        }
    }
    if (best.length > 40) best = best.substring(0, 40) + "...";
    return best;
}

function numberToColumnLetter(num) {
    var result = "";
    while (num > 0) {
        var remainder = (num - 1) % 26;
        result = String.fromCharCode(65 + remainder) + result;
        num = Math.floor((num - 1) / 26);
    }
    return result;
}

function getCellAddressFromSelection(rangeAddress, row, col) {
    var startCell = rangeAddress.split("!").pop().split(":")[0];
    var match = startCell.match(/([A-Z]+)(\d+)/);
    if (match) {
        var startCol = columnLetterToNumber(match[1]);
        var startRow = parseInt(match[2]);
        return numberToColumnLetter(startCol + col) + (startRow + row);
    }
    return numberToColumnLetter(col + 1) + (row + 1);
}

function columnLetterToNumber(letters) {
    var num = 0;
    for (var i = 0; i < letters.length; i++) {
        num = num * 26 + (letters.charCodeAt(i) - 64);
    }
    return num;
}

function truncateValue(value) {
    var s = String(value);
    return s.length > 50 ? s.substring(0, 50) + "..." : s;
}

function truncateFormula(formula) {
    return formula.length > 120 ? formula.substring(0, 120) + "..." : formula;
}

// ============================================
// UI DISPLAY
// ============================================
function displayAuditSummary() {
    var summaryDiv = document.getElementById('auditSummary');
    if (!summaryDiv) return;

    var html = '<div class="audit-summary-cards">';
    html += '<div class="audit-card critical"><span class="count">' + auditState.totalErrors + '</span><span class="label">Errors</span></div>';
    html += '<div class="audit-card warning"><span class="count">' + auditState.totalHardcodes + '</span><span class="label">Hardcodes</span></div>';
    html += '<div class="audit-card warning"><span class="count">' + auditState.totalBreaks + '</span><span class="label">Breaks</span></div>';
    html += '</div>';

    // Top issues
    if (auditState.results.length > 0) {
        html += '<div class="audit-issues-list">';
        html += '<p class="kicker">Sheets Audited</p>';

        auditState.results.slice(0, 5).forEach(function(res) {
            var statusClass = (res.nIssues || 0) === 0 ? "info" : ((res.nIssues || 0) > 10 ? "critical" : "warning");
            html += '<div class="audit-issue-item ' + statusClass + '">';
            html += '<span class="issue-type">' + res.sheetName + '</span>';
            html += '<span class="issue-location">' + (res.nIssues || 0) + ' issues</span>';
            html += '</div>';
        });

        if (auditState.results.length > 5) {
            html += '<p class="more-issues">... and ' + (auditState.results.length - 5) + ' more sheets</p>';
        }
        html += '</div>';
    }

    summaryDiv.innerHTML = html;
    summaryDiv.style.display = 'block';
}
