// ============================================
// ORNA MODULE: Financial Model Audit Tool v5
// ============================================
// Based on VBA modModelAudit v5
//
// IMPROVEMENTS:
//   - Only audits rows that have labels in columns A-D
//   - X breaks show: expected pattern vs actual formula
//   - Uses conditional formatting (fast, no Union)
//   - Selective sheet audit option
//   - Issues capped at 2000
//   - Progress feedback
//
// MAP LEGEND:
//   .  = OK (matches row dominant pattern)
//   H  = Hardcoded constant in formula zone
//   X  = Pattern break (formula differs from dominant)
//   E  = Excel error (#REF! #DIV/0! etc.)
//   S  = Summary/total column
// ============================================

// ============================================
// CONFIGURATION
// ============================================
var auditConfig = {
    audPrefix: "AUD_",
    dashboardName: "AUDIT_DASHBOARD",
    periodScanRows: 15,
    labelCols: 4,
    minFormulasForDominant: 3,
    maxSheetNameLength: 31,
    maxIssues: 2000
};

// Colors (hex)
var auditColors = {
    clean: "#DAF2DA",
    hardcode: "#FFC000",
    break: "#FF8080",
    error: "#FF0000",
    summary: "#FFF2CC",
    headerBg: "#44546A",
    white: "#FFFFFF",
    lightGray: "#F2F2F2"
};

// ============================================
// GLOBAL STATE
// ============================================
var auditState = {
    results: [],
    issues: [],
    totalH: 0,
    totalX: 0,
    totalE: 0,
    sheetsToAudit: [],
    allSheetNames: []
};

// ============================================
// INITIALIZE MODULE
// ============================================
function initAuditModule() {
    document.getElementById('runFullAudit').onclick = runFullAudit;
    document.getElementById('runQuickAudit').onclick = runSelectiveAudit;
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
// FULL AUDIT - All Sheets
// ============================================
function runFullAudit() {
    setAuditStatus("processing", "Running full audit...");
    updateAuditProgress(0, "Initializing...");
    toast("Starting full audit - all sheets...");

    resetAuditState();

    Excel.run(function(context) {
        var sheets = context.workbook.worksheets;
        sheets.load("items/name");

        return context.sync().then(function() {
            // Get all non-audit sheets
            var sheetsToAudit = sheets.items.filter(function(s) {
                return !isAuditSheet(s.name);
            });

            if (sheetsToAudit.length === 0) {
                throw new Error("No sheets to audit.");
            }

            auditState.sheetsToAudit = sheetsToAudit.map(function(s) { return s.name; });
            return runAuditCore(context);
        });
    }).then(function() {
        showAuditComplete();
    }).catch(function(error) {
        setAuditStatus("error", "Error: " + error.message);
        toast("Audit failed: " + error.message);
        console.error(error);
    });
}

// ============================================
// SELECTIVE AUDIT - Pick Sheets
// ============================================
function runSelectiveAudit() {
    setAuditStatus("processing", "Loading sheet list...");

    Excel.run(function(context) {
        var sheets = context.workbook.worksheets;
        sheets.load("items/name");

        return context.sync().then(function() {
            var eligibleSheets = sheets.items.filter(function(s) {
                return !isAuditSheet(s.name);
            });

            if (eligibleSheets.length === 0) {
                throw new Error("No sheets to audit.");
            }

            auditState.allSheetNames = eligibleSheets.map(function(s) { return s.name; });
            
            // Show sheet picker UI
            showSheetPicker(auditState.allSheetNames);
        });
    }).catch(function(error) {
        setAuditStatus("error", "Error: " + error.message);
        console.error(error);
    });
}

function showSheetPicker(sheetNames) {
    var summaryDiv = document.getElementById('auditSummary');
    if (!summaryDiv) return;

    var html = '<div class="sheet-picker">';
    html += '<p class="kicker">Select Sheets to Audit</p>';
    html += '<div class="sheet-picker-list">';
    
    sheetNames.forEach(function(name, index) {
        html += '<label class="sheet-checkbox">';
        html += '<input type="checkbox" id="sheet_' + index + '" checked>';
        html += '<span>' + escapeHtml(name) + '</span>';
        html += '</label>';
    });
    
    html += '</div>';
    html += '<div class="sheet-picker-actions">';
    html += '<button class="btn btn-sm" onclick="selectAllSheets(true)">Select All</button>';
    html += '<button class="btn btn-sm" onclick="selectAllSheets(false)">Deselect All</button>';
    html += '</div>';
    html += '<button class="btn btn-primary" onclick="runSelectedSheetsAudit()" style="width:100%;margin-top:12px;">Run Audit</button>';
    html += '</div>';

    summaryDiv.innerHTML = html;
    summaryDiv.style.display = 'block';
    setAuditStatus("ok", "Select sheets and click Run Audit");
}

function selectAllSheets(checked) {
    auditState.allSheetNames.forEach(function(name, index) {
        var cb = document.getElementById('sheet_' + index);
        if (cb) cb.checked = checked;
    });
}

function runSelectedSheetsAudit() {
    var selectedSheets = [];
    
    auditState.allSheetNames.forEach(function(name, index) {
        var cb = document.getElementById('sheet_' + index);
        if (cb && cb.checked) {
            selectedSheets.push(name);
        }
    });

    if (selectedSheets.length === 0) {
        toast("Please select at least one sheet");
        return;
    }

    resetAuditState();
    auditState.sheetsToAudit = selectedSheets;

    setAuditStatus("processing", "Running selective audit...");
    updateAuditProgress(0, "Initializing...");
    toast("Auditing " + selectedSheets.length + " sheet(s)...");

    Excel.run(function(context) {
        return runAuditCore(context);
    }).then(function() {
        showAuditComplete();
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
                            actual: mark === "H" ? String(value) : formula,
                            expected: mark === "X" ? ("Dominant: " + dominant) : (mark === "H" ? "(row expects formula)" : formula)
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
            auditState.issues = issues;
            auditState.totalH = issues.filter(function(i) { return i.type === "H"; }).length;
            auditState.totalX = issues.filter(function(i) { return i.type === "X"; }).length;
            auditState.totalE = issues.filter(function(i) { return i.type === "E"; }).length;
            displayAuditSummary();
        }
    }).catch(function(error) {
        setAuditStatus("error", "Error: " + error.message);
        console.error(error);
    });
}

// ============================================
// CORE AUDIT ENGINE
// ============================================
function runAuditCore(context) {
    return new Promise(function(resolve, reject) {
        var sheets = context.workbook.worksheets;
        sheets.load("items/name");

        context.sync().then(function() {
            // Purge old audit sheets
            return purgeOldAuditSheets(context, sheets.items);
        }).then(function() {
            // Re-load sheets
            sheets.load("items/name");
            return context.sync();
        }).then(function() {
            // Audit each selected sheet sequentially
            var sheetIndex = 0;
            var totalSheets = auditState.sheetsToAudit.length;

            function auditNextSheet() {
                if (sheetIndex >= totalSheets) {
                    // All sheets audited, build dashboard
                    updateAuditProgress(85, "Building dashboard...");
                    return buildDashboard(context).then(function() {
                        updateAuditProgress(100, "Complete!");
                        
                        // Activate dashboard
                        var dashboard = context.workbook.worksheets.getItem(auditConfig.dashboardName);
                        dashboard.activate();
                        
                        return context.sync();
                    });
                }

                var sheetName = auditState.sheetsToAudit[sheetIndex];
                var progress = 10 + (sheetIndex / totalSheets * 70);
                updateAuditProgress(progress, "Auditing: " + sheetName + " (" + (sheetIndex + 1) + "/" + totalSheets + ")");

                var sheet = null;
                try {
                    sheet = context.workbook.worksheets.getItem(sheetName);
                } catch (e) {
                    sheetIndex++;
                    return auditNextSheet();
                }

                return auditOneSheet(context, sheet).then(function(result) {
                    if (result) {
                        auditState.results.push(result);
                        auditState.totalH += result.nH;
                        auditState.totalX += result.nX;
                        auditState.totalE += result.nE;
                    }
                    sheetIndex++;
                    return auditNextSheet();
                }).catch(function(err) {
                    console.error("Error auditing " + sheetName + ": " + err.message);
                    sheetIndex++;
                    return auditNextSheet();
                });
            }

            return auditNextSheet();
        }).then(resolve).catch(reject);
    });
}

// ============================================
// AUDIT ONE SHEET (Only Labeled Rows)
// ============================================
function auditOneSheet(context, sheet) {
    return new Promise(function(resolve, reject) {
        var result = {
            sheetName: sheet.name,
            auditName: safeAuditName(sheet.name),
            periodRow: 0,
            periodCol: 0,
            nH: 0,
            nX: 0,
            nE: 0,
            labeledRows: [],
            totalCols: {}
        };

        var usedRange = sheet.getUsedRange();
        usedRange.load("values, formulas, formulasR1C1, rowCount, columnCount");

        context.sync().then(function() {
            var values = usedRange.values;
            var formulas = usedRange.formulas;
            var formulasR1C1 = usedRange.formulasR1C1;
            var nRows = usedRange.rowCount;
            var nCols = usedRange.columnCount;

            // Detect period
            var periodInfo = detectPeriod(values, nRows, nCols);
            result.periodRow = periodInfo.row;
            result.periodCol = periodInfo.col;

            var pCol = result.periodCol > 0 ? result.periodCol - 1 : auditConfig.labelCols; // 0-based

            // Detect total columns
            if (result.periodRow > 0) {
                result.totalCols = detectTotalColumns(values, result.periodRow - 1, pCol, nCols);
            }

            // Find labeled rows with data
            var labeledRows = [];
            for (var r = 0; r < nRows; r++) {
                var label = getRowLabel(values, r);
                if (label.length === 0) continue;

                // Check if row has any data in the audit zone
                var hasData = false;
                for (var c = pCol; c < nCols; c++) {
                    if (values[r][c] !== null && values[r][c] !== "" && values[r][c] !== undefined) {
                        hasData = true;
                        break;
                    }
                    var f = formulas[r][c];
                    if (typeof f === "string" && f.length > 0 && f.charAt(0) === "=") {
                        hasData = true;
                        break;
                    }
                }

                if (hasData) {
                    labeledRows.push({
                        rowIndex: r,
                        rowNum: r + 1,
                        label: label
                    });
                }
            }

            result.labeledRows = labeledRows;

            if (labeledRows.length === 0) {
                resolve(result);
                return;
            }

            // Audit only labeled rows
            var mapData = [];
            var gridCols = nCols - pCol;
            
            // Header row
            var headerRow = ["Row", "Label", "#"];
            for (var gc = 0; gc < gridCols; gc++) {
                headerRow.push(numberToColumnLetter(pCol + gc + 1));
            }
            mapData.push(headerRow);

            // Audit each labeled row
            for (var li = 0; li < labeledRows.length; li++) {
                var lr = labeledRows[li];
                var r = lr.rowIndex;
                var dominant = getDominantPattern(formulasR1C1, r, pCol, nCols - 1, result.totalCols);
                
                var rowIssues = 0;
                var rowMarks = [lr.rowNum, lr.label, ""];

                for (var c = pCol; c < nCols; c++) {
                    var value = values[r][c];
                    var formula = formulas[r][c];
                    var formulaR1C1 = formulasR1C1[r][c];
                    var isTotalCol = result.totalCols[c] === true;

                    var mark = classifyCell(value, formula, formulaR1C1, dominant, isTotalCol);
                    rowMarks.push(mark);

                    // Collect issues (with expected vs actual)
                    if (mark === "H" || mark === "X" || mark === "E") {
                        rowIssues++;
                        
                        if (auditState.issues.length < auditConfig.maxIssues) {
                            var issue = {
                                type: mark,
                                sheet: sheet.name,
                                cell: numberToColumnLetter(c + 1) + (r + 1),
                                actual: "",
                                expected: ""
                            };

                            if (mark === "H") {
                                issue.actual = truncateValue(value);
                                issue.expected = "(row expects formula)";
                                result.nH++;
                            } else if (mark === "X") {
                                issue.actual = truncateFormula(formula);
                                issue.expected = "Dominant: " + dominant;
                                result.nX++;
                            } else if (mark === "E") {
                                issue.actual = String(value);
                                issue.expected = truncateFormula(formula);
                                result.nE++;
                            }

                            auditState.issues.push(issue);
                        }
                    }
                }

                // Row health indicator
                rowMarks[2] = rowIssues === 0 ? "✓" : String(rowIssues);
                mapData.push(rowMarks);
            }

            // Create audit map sheet
            return createAuditMapSheet(context, result, mapData, pCol, gridCols);
        }).then(function() {
            resolve(result);
        }).catch(function(error) {
            console.error("Error auditing " + sheet.name + ": " + error.message);
            resolve(result);
        });
    });
}

// ============================================
// CREATE AUDIT MAP SHEET (with Conditional Formatting)
// ============================================
function createAuditMapSheet(context, result, mapData, pCol, gridCols) {
    return new Promise(function(resolve, reject) {
        var sheets = context.workbook.worksheets;
        var audSheet = sheets.add(result.auditName);

        // Title and info
        audSheet.getRange("A1").values = [["AUDIT: " + result.sheetName]];
        audSheet.getRange("A1").format.font.bold = true;
        audSheet.getRange("A1").format.font.size = 13;

        var infoText = result.periodCol > 0 
            ? "Period: row " + result.periodRow + " col " + numberToColumnLetter(result.periodCol)
            : "No period detected";
        audSheet.getRange("A2").values = [[infoText]];
        audSheet.getRange("A2").format.font.color = "#787878";

        audSheet.getRange("A3").values = [[". OK | H Hardcode | X Break | E Error | S Total"]];
        audSheet.getRange("A3").format.font.size = 8;
        audSheet.getRange("A3").format.font.color = "#8C8C8C";

        if (mapData.length > 1) {
            // Write data starting at row 5
            var dataRange = audSheet.getRange("A5").getResizedRange(mapData.length - 1, mapData[0].length - 1);
            dataRange.values = mapData;

            // Format header row (row 5)
            var headerRange = audSheet.getRange("A5").getResizedRange(0, mapData[0].length - 1);
            headerRange.format.font.bold = true;

            // Format grid header (columns D onwards)
            if (mapData[0].length > 3) {
                var gridHeaderRange = audSheet.getRange("D5").getResizedRange(0, gridCols - 1);
                gridHeaderRange.format.fill.color = auditColors.headerBg;
                gridHeaderRange.format.font.color = auditColors.white;
                gridHeaderRange.format.font.size = 7;
                gridHeaderRange.format.horizontalAlignment = "Center";
            }

            // Apply conditional formatting to the grid (FAST - no Union needed)
            if (mapData.length > 1 && gridCols > 0) {
                var gridDataRange = audSheet.getRange("D6").getResizedRange(mapData.length - 2, gridCols - 1);
                gridDataRange.format.horizontalAlignment = "Center";
                gridDataRange.format.font.size = 9;
                gridDataRange.format.font.bold = true;

                // Clear existing and add conditional formatting rules
                gridDataRange.conditionalFormats.clearAll();

                // Rule 1: . = OK (green)
                var cfOk = gridDataRange.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
                cfOk.cellValue.format.fill.color = auditColors.clean;
                cfOk.cellValue.format.font.color = "#64AA64";
                cfOk.cellValue.rule = { formula1: '"."', operator: Excel.ConditionalCellValueOperator.equalTo };

                // Rule 2: H = Hardcode (orange)
                var cfH = gridDataRange.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
                cfH.cellValue.format.fill.color = auditColors.hardcode;
                cfH.cellValue.format.font.color = "#824600";
                cfH.cellValue.rule = { formula1: '"H"', operator: Excel.ConditionalCellValueOperator.equalTo };

                // Rule 3: X = Break (red)
                var cfX = gridDataRange.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
                cfX.cellValue.format.fill.color = auditColors.break;
                cfX.cellValue.format.font.color = auditColors.white;
                cfX.cellValue.rule = { formula1: '"X"', operator: Excel.ConditionalCellValueOperator.equalTo };

                // Rule 4: E = Error (bright red)
                var cfE = gridDataRange.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
                cfE.cellValue.format.fill.color = auditColors.error;
                cfE.cellValue.format.font.color = auditColors.white;
                cfE.cellValue.rule = { formula1: '"E"', operator: Excel.ConditionalCellValueOperator.equalTo };

                // Rule 5: S = Summary (gold)
                var cfS = gridDataRange.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
                cfS.cellValue.format.fill.color = auditColors.summary;
                cfS.cellValue.format.font.color = "#8C8250";
                cfS.cellValue.rule = { formula1: '"S"', operator: Excel.ConditionalCellValueOperator.equalTo };
            }

            // Format check column (column C)
            if (mapData.length > 1) {
                var checkRange = audSheet.getRange("C6").getResizedRange(mapData.length - 2, 0);
                checkRange.format.horizontalAlignment = "Center";
                checkRange.format.font.size = 8;

                checkRange.conditionalFormats.clearAll();
                
                var cfCheck = checkRange.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
                cfCheck.cellValue.format.font.color = "#008C00";
                cfCheck.cellValue.rule = { formula1: '"✓"', operator: Excel.ConditionalCellValueOperator.equalTo };

                var cfCheckBad = checkRange.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
                cfCheckBad.cellValue.format.font.color = "#C80000";
                cfCheckBad.cellValue.format.font.bold = true;
                cfCheckBad.cellValue.rule = { formula1: '"✓"', operator: Excel.ConditionalCellValueOperator.notEqualTo };
            }

            // Format row number column (column A)
            if (mapData.length > 1) {
                var rowNumRange = audSheet.getRange("A6").getResizedRange(mapData.length - 2, 0);
                rowNumRange.format.font.size = 8;
                rowNumRange.format.font.color = "#969696";
            }

            // Format label column (column B)
            if (mapData.length > 1) {
                var labelRange = audSheet.getRange("B6").getResizedRange(mapData.length - 2, 0);
                labelRange.format.font.size = 8;
            }
        }

        // Column widths
        audSheet.getRange("A:A").format.columnWidth = 40;
        audSheet.getRange("B:B").format.columnWidth = 160;
        audSheet.getRange("C:C").format.columnWidth = 30;

        // Freeze panes
        audSheet.freezePanes.freezeRows(5);

        context.sync().then(resolve).catch(reject);
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

        var totalIssues = auditState.totalH + auditState.totalX + auditState.totalE;

        // Title
        dash.getRange("A1").values = [["MODEL AUDIT DASHBOARD"]];
        dash.getRange("A1").format.font.bold = true;
        dash.getRange("A1").format.font.size = 16;
        dash.getRange("A1").format.font.color = "#1E3C64";

        dash.getRange("A2").values = [[new Date().toLocaleString()]];
        dash.getRange("A2").format.font.color = "#969696";

        // Scoreboard
        writeScoreBox(dash, "H1", "Total", totalIssues, totalIssues === 0 ? auditColors.clean : "#FFC7CE");
        writeScoreBox(dash, "J1", "H", auditState.totalH, "#FFE6B4");
        writeScoreBox(dash, "L1", "X", auditState.totalX, "#FFC7CE");
        writeScoreBox(dash, "N1", "E", auditState.totalE, auditState.totalE === 0 ? auditColors.clean : "#FF3232");

        // Sheet summary table
        var summaryHeaders = [["Sheet", "Period", "H", "X", "E", "Total", "Map"]];
        dash.getRange("A4:G4").values = summaryHeaders;
        dash.getRange("A4:G4").format.font.bold = true;
        dash.getRange("A4:G4").format.fill.color = auditColors.headerBg;
        dash.getRange("A4:G4").format.font.color = auditColors.white;
        dash.getRange("A4:G4").format.horizontalAlignment = "Center";

        // Write sheet results
        if (auditState.results.length > 0) {
            var sheetData = auditState.results.map(function(res) {
                var periodText = res.periodCol > 0 
                    ? numberToColumnLetter(res.periodCol) + " r" + res.periodRow 
                    : "N/A";
                var total = (res.nH || 0) + (res.nX || 0) + (res.nE || 0);
                return [res.sheetName, periodText, res.nH || 0, res.nX || 0, res.nE || 0, total, "-> Map"];
            });

            var dataRange = dash.getRange("A5").getResizedRange(sheetData.length - 1, 6);
            dataRange.values = sheetData;

            // Conditional formatting for total column
            var totalColRange = dash.getRange("F5").getResizedRange(sheetData.length - 1, 0);
            totalColRange.format.font.bold = true;
            totalColRange.conditionalFormats.clearAll();

            var cfZero = totalColRange.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
            cfZero.cellValue.format.fill.color = auditColors.clean;
            cfZero.cellValue.rule = { formula1: "0", operator: Excel.ConditionalCellValueOperator.equalTo };

            var cfLow = totalColRange.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
            cfLow.cellValue.format.fill.color = "#FFEB9C";
            cfLow.cellValue.rule = { formula1: "1", formula2: "10", operator: Excel.ConditionalCellValueOperator.between };

            var cfHigh = totalColRange.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
            cfHigh.cellValue.format.fill.color = "#FFC7CE";
            cfHigh.cellValue.rule = { formula1: "10", operator: Excel.ConditionalCellValueOperator.greaterThan };

            // Map column styling
            var mapColRange = dash.getRange("G5").getResizedRange(sheetData.length - 1, 0);
            mapColRange.format.font.color = "#0563C1";
            mapColRange.format.font.underline = "Single";
        }

        // Totals row
        var totRow = 5 + auditState.results.length;
        dash.getRange("A" + totRow + ":G" + totRow).values = [[
            "TOTAL", "", auditState.totalH, auditState.totalX, auditState.totalE, totalIssues, ""
        ]];
        dash.getRange("A" + totRow + ":G" + totRow).format.font.bold = true;

        // Issue detail section
        var issueHeaderRow = totRow + 2;
        dash.getRange("A" + issueHeaderRow).values = [["ISSUE DETAIL"]];
        dash.getRange("A" + issueHeaderRow).format.font.bold = true;
        dash.getRange("A" + issueHeaderRow).format.font.size = 13;

        if (auditState.issues.length > 0) {
            var cappedNote = auditState.issues.length >= auditConfig.maxIssues ? " (capped at " + auditConfig.maxIssues + ")" : "";
            dash.getRange("A" + (issueHeaderRow + 1)).values = [[auditState.issues.length + " issues" + cappedNote]];
            dash.getRange("A" + (issueHeaderRow + 1)).format.font.color = "#969696";
        }

        var issueTableRow = issueHeaderRow + 2;
        var issueHeaders = [["Type", "Sheet", "Cell", "Formula / Value", "Expected Pattern", "Go"]];
        dash.getRange("A" + issueTableRow + ":F" + issueTableRow).values = issueHeaders;
        dash.getRange("A" + issueTableRow + ":F" + issueTableRow).format.font.bold = true;
        dash.getRange("A" + issueTableRow + ":F" + issueTableRow).format.fill.color = auditColors.headerBg;
        dash.getRange("A" + issueTableRow + ":F" + issueTableRow).format.font.color = auditColors.white;

        // Write issues
        if (auditState.issues.length > 0) {
            var issueData = auditState.issues.map(function(issue) {
                return [issue.type, issue.sheet, issue.cell, issue.actual, issue.expected, "Go"];
            });

            var issueDataRange = dash.getRange("A" + (issueTableRow + 1)).getResizedRange(issueData.length - 1, 5);
            issueDataRange.values = issueData;

            // Conditional formatting for type column
            var typeColRange = dash.getRange("A" + (issueTableRow + 1)).getResizedRange(issueData.length - 1, 0);
            typeColRange.conditionalFormats.clearAll();

            var cfTypeH = typeColRange.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
            cfTypeH.cellValue.format.fill.color = auditColors.hardcode;
            cfTypeH.cellValue.rule = { formula1: '"H"', operator: Excel.ConditionalCellValueOperator.equalTo };

            var cfTypeX = typeColRange.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
            cfTypeX.cellValue.format.fill.color = auditColors.break;
            cfTypeX.cellValue.format.font.color = auditColors.white;
            cfTypeX.cellValue.rule = { formula1: '"X"', operator: Excel.ConditionalCellValueOperator.equalTo };

            var cfTypeE = typeColRange.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
            cfTypeE.cellValue.format.fill.color = auditColors.error;
            cfTypeE.cellValue.format.font.color = auditColors.white;
            cfTypeE.cellValue.rule = { formula1: '"E"', operator: Excel.ConditionalCellValueOperator.equalTo };

            // Go column styling
            var goColRange = dash.getRange("F" + (issueTableRow + 1)).getResizedRange(issueData.length - 1, 0);
            goColRange.format.font.color = "#0563C1";
        }

        // Column widths
        dash.getRange("A:A").format.columnWidth = 50;
        dash.getRange("B:B").format.columnWidth = 130;
        dash.getRange("C:C").format.columnWidth = 50;
        dash.getRange("D:D").format.columnWidth = 280;
        dash.getRange("E:E").format.columnWidth = 200;
        dash.getRange("F:F").format.columnWidth = 40;
        dash.getRange("G:G").format.columnWidth = 80;

        context.sync().then(resolve).catch(reject);
    });
}

// ============================================
// HELPER FUNCTIONS
// ============================================
function resetAuditState() {
    auditState = {
        results: [],
        issues: [],
        totalH: 0,
        totalX: 0,
        totalE: 0,
        sheetsToAudit: [],
        allSheetNames: auditState.allSheetNames || []
    };
}

function showAuditComplete() {
    var totalIssues = auditState.totalH + auditState.totalX + auditState.totalE;
    setAuditStatus("ok", "Audit complete! " + totalIssues + " issues found");
    toast("Done! See " + auditConfig.dashboardName);
    displayAuditSummary();
}

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
    name = name.replace(/[:\\\/?*\[\]]/g, "_");
    return name;
}

function purgeOldAuditSheets(context, sheets) {
    var toDelete = sheets.filter(function(s) { return isAuditSheet(s.name); });
    toDelete.forEach(function(s) { s.delete(); });
    return context.sync();
}

function detectPeriod(values, nRows, nCols) {
    var result = { row: 0, col: 0 };
    var scanRows = Math.min(auditConfig.periodScanRows, nRows);
    var bestRun = 0;

    for (var r = 0; r < scanRows; r++) {
        for (var c = 0; c < nCols - 2; c++) {
            var v1 = values[r][c], v2 = values[r][c + 1], v3 = values[r][c + 2];
            
            if (isNumeric(v1) && isNumeric(v2) && isNumeric(v3)) {
                if (Math.floor(v1) === 1 && Math.floor(v2) === 2 && Math.floor(v3) === 3) {
                    var runLen = 3;
                    for (var k = c + 3; k < nCols; k++) {
                        var vk = values[r][k];
                        if (isNumeric(vk)) {
                            if (Math.floor(vk) >= runLen + 1 && Math.floor(vk) <= runLen + 2) {
                                runLen = Math.floor(vk);
                            } else {
                                break;
                            }
                        }
                    }
                    if (runLen > bestRun) {
                        bestRun = runLen;
                        result.row = r + 1;
                        result.col = c + 1;
                    }
                }
            }
        }
    }
    return result;
}

function detectTotalColumns(values, periodRowIndex, startCol, endCol) {
    var totalCols = {};
    var periodRow = values[periodRowIndex];
    if (!periodRow) return totalCols;

    var scanTo = Math.min(endCol, periodRow.length);
    for (var c = startCol; c < scanTo; c++) {
        var v = periodRow[c];
        if (v === null || v === "" || v === undefined || !isNumeric(v)) {
            totalCols[c] = true;
        }
    }
    return totalCols;
}

function getDominantPattern(formulasR1C1, rowIndex, startCol, endCol, totalCols) {
    var patternCounts = {};
    var nFormulas = 0;

    for (var c = startCol; c <= endCol; c++) {
        if (totalCols[c]) continue;
        var raw = formulasR1C1[rowIndex][c];
        if (raw === null || raw === undefined || typeof raw !== "string") continue;
        var key = String(raw);
        if (key.length === 0 || key.charAt(0) !== "=") continue;
        nFormulas++;
        patternCounts[key] = (patternCounts[key] || 0) + 1;
    }

    if (nFormulas < auditConfig.minFormulasForDominant) return "";

    var maxCount = 0, maxKey = "";
    for (var key in patternCounts) {
        if (patternCounts[key] > maxCount) {
            maxCount = patternCounts[key];
            maxKey = key;
        }
    }

    return maxCount >= nFormulas * 0.4 ? maxKey : "";
}

function classifyCell(value, formula, formulaR1C1, dominant, isTotalCol) {
    var isFormula = typeof formula === "string" && formula.length > 0 && formula.charAt(0) === "=";

    if ((value === null || value === "" || value === undefined) && !isFormula) return "";
    if (isErrorValue(value)) return "E";
    if (isTotalCol) return "S";
    if (!isFormula) return "H";
    if (dominant.length > 0 && formulaR1C1 !== dominant) return "X";
    return ".";
}

function getRowLabel(values, rowIndex) {
    var best = "";
    var labelCols = Math.min(auditConfig.labelCols, values[rowIndex].length);
    for (var c = 0; c < labelCols; c++) {
        var v = values[rowIndex][c];
        if (v !== null && v !== "" && v !== undefined && !isErrorValue(v)) {
            var s = String(v).trim();
            if (s.length > best.length) best = s;
        }
    }
    if (best.length > 42) best = best.substring(0, 42) + "...";
    return best;
}

function isNumeric(value) { return typeof value === "number" && !isNaN(value); }

function isErrorValue(value) {
    if (typeof value !== "string") return false;
    return value.charAt(0) === "#" && ["#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!", "#CALC!", "#SPILL!"].indexOf(value) >= 0;
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
    return s.length > 80 ? s.substring(0, 80) + "..." : s;
}

function truncateFormula(formula) {
    return formula && formula.length > 80 ? formula.substring(0, 80) + "..." : (formula || "");
}

function escapeHtml(text) {
    var div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function writeScoreBox(sheet, startCell, title, value, bgColor) {
    var titleCell = sheet.getRange(startCell);
    titleCell.values = [[title]];
    titleCell.format.font.size = 9;
    titleCell.format.font.bold = true;

    var col = startCell.replace(/[0-9]/g, '');
    var row = parseInt(startCell.replace(/[A-Z]/gi, ''));
    var valueCell = sheet.getRange(col + (row + 1));
    valueCell.values = [[value]];
    valueCell.format.font.size = 20;
    valueCell.format.font.bold = true;
    valueCell.format.fill.color = bgColor;
    valueCell.format.horizontalAlignment = "Center";
}

// ============================================
// UI DISPLAY
// ============================================
function displayAuditSummary() {
    var summaryDiv = document.getElementById('auditSummary');
    if (!summaryDiv) return;

    var totalIssues = auditState.totalH + auditState.totalX + auditState.totalE;

    var html = '<div class="audit-summary-cards">';
    html += '<div class="audit-card critical"><span class="count">' + auditState.totalE + '</span><span class="label">Errors</span></div>';
    html += '<div class="audit-card warning"><span class="count">' + auditState.totalH + '</span><span class="label">Hardcodes</span></div>';
    html += '<div class="audit-card warning"><span class="count">' + auditState.totalX + '</span><span class="label">Breaks</span></div>';
    html += '</div>';

    if (auditState.results.length > 0) {
        html += '<div class="audit-issues-list">';
        html += '<p class="kicker">Sheets Audited (' + auditState.results.length + ')</p>';

        auditState.results.slice(0, 5).forEach(function(res) {
            var total = (res.nH || 0) + (res.nX || 0) + (res.nE || 0);
            var statusClass = total === 0 ? "info" : (total > 10 ? "critical" : "warning");
            html += '<div class="audit-issue-item ' + statusClass + '">';
            html += '<span class="issue-type">' + escapeHtml(res.sheetName) + '</span>';
            html += '<span class="issue-location">' + total + ' issues</span>';
            html += '</div>';
        });

        if (auditState.results.length > 5) {
            html += '<p class="more-issues">... and ' + (auditState.results.length - 5) + ' more sheets</p>';
        }
        html += '</div>';
    }

    if (auditState.issues.length > 0) {
        html += '<div class="audit-issues-list" style="margin-top:12px;">';
        html += '<p class="kicker">Top Issues</p>';
        
        auditState.issues.slice(0, 5).forEach(function(issue) {
            var typeClass = issue.type === "E" ? "critical" : "warning";
            html += '<div class="audit-issue-item ' + typeClass + '">';
            html += '<span class="issue-type">' + issue.type + ': ' + issue.cell + '</span>';
            html += '<span class="issue-location">' + escapeHtml(issue.sheet) + '</span>';
            html += '</div>';
        });

        if (auditState.issues.length > 5) {
            html += '<p class="more-issues">... and ' + (auditState.issues.length - 5) + ' more issues</p>';
        }
        html += '</div>';
    }

    summaryDiv.innerHTML = html;
    summaryDiv.style.display = 'block';
}
