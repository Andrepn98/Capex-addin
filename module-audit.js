// ============================================
// ORNA MODULE: Financial Model Audit Tool v8
// ============================================
// - Colors on audit map: H=Yellow, X=Red, E=Orange
// - Dashboard summary uses COUNTIF formulas (auto-update when you delete issues)
// - Issue detail as Excel Table with AutoFilter
// ============================================

var auditConfig = {
    audPrefix: "AUD_",
    dashboardName: "AUDIT_DASHBOARD",
    periodScanRows: 15,
    labelCols: 4,
    minFormulasForDominant: 3,
    maxIssues: 2000
};

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
// INITIALIZE
// ============================================
function initAuditModule() {
    var btn1 = document.getElementById('runFullAudit');
    var btn2 = document.getElementById('runQuickAudit');
    var btn3 = document.getElementById('auditSelection');
    
    if (btn1) btn1.onclick = runFullAudit;
    if (btn2) btn2.onclick = runSelectiveAudit;
    if (btn3) btn3.onclick = auditSelectedRange;
    
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

function resetState() {
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

// ============================================
// FULL AUDIT
// ============================================
function runFullAudit() {
    setAuditStatus("processing", "Starting full audit...");
    updateAuditProgress(0, "Initializing...");
    toast("Starting full audit...");
    resetState();

    Excel.run(function(context) {
        var sheets = context.workbook.worksheets;
        sheets.load("items/name");

        return context.sync().then(function() {
            var eligible = [];
            for (var i = 0; i < sheets.items.length; i++) {
                var name = sheets.items[i].name;
                if (!isAuditSheet(name)) {
                    eligible.push(name);
                }
            }

            if (eligible.length === 0) {
                throw new Error("No sheets to audit.");
            }

            auditState.sheetsToAudit = eligible;
            return doAudit(context);
        });
    }).then(function() {
        var total = auditState.totalH + auditState.totalX + auditState.totalE;
        setAuditStatus("ok", "Done! " + total + " issues found");
        toast("Audit complete!");
        showSummaryUI();
    }).catch(function(error) {
        setAuditStatus("error", "Error: " + error.message);
        toast("Audit failed: " + error.message);
        console.error("Audit error:", error);
    });
}

// ============================================
// SELECTIVE AUDIT
// ============================================
function runSelectiveAudit() {
    setAuditStatus("processing", "Loading sheets...");

    Excel.run(function(context) {
        var sheets = context.workbook.worksheets;
        sheets.load("items/name");

        return context.sync().then(function() {
            var eligible = [];
            for (var i = 0; i < sheets.items.length; i++) {
                var name = sheets.items[i].name;
                if (!isAuditSheet(name)) {
                    eligible.push(name);
                }
            }

            if (eligible.length === 0) {
                throw new Error("No sheets available.");
            }

            auditState.allSheetNames = eligible;
            showSheetPicker(eligible);
            setAuditStatus("ok", "Select sheets to audit");
        });
    }).catch(function(error) {
        setAuditStatus("error", "Error: " + error.message);
        console.error(error);
    });
}

function showSheetPicker(names) {
    var div = document.getElementById('auditSummary');
    if (!div) return;

    var html = '<div class="sheet-picker">';
    html += '<p class="kicker">Select Sheets to Audit</p>';
    html += '<div class="sheet-picker-list">';
    
    for (var i = 0; i < names.length; i++) {
        html += '<label class="sheet-checkbox">';
        html += '<input type="checkbox" id="sheetcb_' + i + '" checked>';
        html += '<span>' + escapeHtml(names[i]) + '</span>';
        html += '</label>';
    }
    
    html += '</div>';
    html += '<div class="sheet-picker-actions">';
    html += '<button class="btn btn-sm" onclick="toggleAllSheets(true)">Select All</button>';
    html += '<button class="btn btn-sm" onclick="toggleAllSheets(false)">Deselect All</button>';
    html += '</div>';
    html += '<button class="btn btn-primary" onclick="runPickedSheets()" style="width:100%;margin-top:12px;">Run Audit</button>';
    html += '</div>';

    div.innerHTML = html;
    div.style.display = 'block';
}

function toggleAllSheets(checked) {
    for (var i = 0; i < auditState.allSheetNames.length; i++) {
        var cb = document.getElementById('sheetcb_' + i);
        if (cb) cb.checked = checked;
    }
}

function runPickedSheets() {
    var picked = [];
    for (var i = 0; i < auditState.allSheetNames.length; i++) {
        var cb = document.getElementById('sheetcb_' + i);
        if (cb && cb.checked) {
            picked.push(auditState.allSheetNames[i]);
        }
    }

    if (picked.length === 0) {
        toast("Select at least one sheet");
        return;
    }

    resetState();
    auditState.sheetsToAudit = picked;

    setAuditStatus("processing", "Auditing " + picked.length + " sheets...");
    updateAuditProgress(0, "Starting...");
    toast("Auditing " + picked.length + " sheet(s)...");

    Excel.run(function(context) {
        return doAudit(context);
    }).then(function() {
        var total = auditState.totalH + auditState.totalX + auditState.totalE;
        setAuditStatus("ok", "Done! " + total + " issues");
        toast("Audit complete!");
        showSummaryUI();
    }).catch(function(error) {
        setAuditStatus("error", "Error: " + error.message);
        console.error(error);
    });
}

// ============================================
// AUDIT SELECTION ONLY
// ============================================
function auditSelectedRange() {
    setAuditStatus("processing", "Auditing selection...");
    resetState();

    Excel.run(function(context) {
        var range = context.workbook.getSelectedRange();
        range.load("address, values, formulas, formulasR1C1, worksheet/name");

        return context.sync().then(function() {
            var sheetName = range.worksheet.name;
            var values = range.values;
            var formulas = range.formulas;
            var formulasR1C1 = range.formulasR1C1;

            for (var r = 0; r < values.length; r++) {
                var dom = getDominant(formulasR1C1, r, 0, values[r].length - 1, {});

                for (var c = 0; c < values[r].length; c++) {
                    var mark = classify(values[r][c], formulas[r][c], formulasR1C1[r][c], dom, false);

                    if (mark === "H" || mark === "X" || mark === "E") {
                        var addr = getCellAddr(range.address, r, c);
                        auditState.issues.push({
                            type: mark,
                            sheet: sheetName,
                            cell: addr,
                            actual: mark === "H" ? String(values[r][c]) : formulas[r][c],
                            expected: mark === "X" ? dom : (mark === "H" ? "(expects formula)" : "")
                        });

                        if (mark === "H") auditState.totalH++;
                        else if (mark === "X") auditState.totalX++;
                        else if (mark === "E") auditState.totalE++;
                    }
                }
            }
        });
    }).then(function() {
        var total = auditState.issues.length;
        setAuditStatus("ok", "Found " + total + " issues");
        toast(total + " issues in selection");
        showSummaryUI();
    }).catch(function(error) {
        setAuditStatus("error", "Error: " + error.message);
        console.error(error);
    });
}

// ============================================
// MAIN AUDIT LOGIC
// ============================================
function doAudit(context) {
    var sheets = context.workbook.worksheets;
    sheets.load("items/name");

    return context.sync().then(function() {
        // Delete old audit sheets
        var toDelete = [];
        for (var i = 0; i < sheets.items.length; i++) {
            if (isAuditSheet(sheets.items[i].name)) {
                toDelete.push(sheets.items[i]);
            }
        }
        for (var j = 0; j < toDelete.length; j++) {
            toDelete[j].delete();
        }
        return context.sync();
    }).then(function() {
        // Process sheets one by one
        var idx = 0;
        var total = auditState.sheetsToAudit.length;

        function next() {
            if (idx >= total) {
                updateAuditProgress(90, "Building dashboard...");
                return makeDashboard(context);
            }

            var name = auditState.sheetsToAudit[idx];
            var pct = 10 + (idx / total) * 75;
            updateAuditProgress(pct, "Auditing: " + name);

            return auditSheet(context, name).then(function(res) {
                if (res) {
                    auditState.results.push(res);
                    auditState.totalH += res.nH;
                    auditState.totalX += res.nX;
                    auditState.totalE += res.nE;
                }
                idx++;
                return next();
            }).catch(function(err) {
                console.error("Sheet error:", name, err);
                idx++;
                return next();
            });
        }

        return next();
    });
}

// ============================================
// AUDIT SINGLE SHEET
// ============================================
function auditSheet(context, sheetName) {
    var result = {
        sheetName: sheetName,
        auditName: makeAuditName(sheetName),
        periodRow: 0,
        periodCol: 0,
        nH: 0,
        nX: 0,
        nE: 0
    };

    var sheet;
    try {
        sheet = context.workbook.worksheets.getItem(sheetName);
    } catch (e) {
        return Promise.resolve(null);
    }

    var usedRange = sheet.getUsedRangeOrNullObject();
    usedRange.load("isNullObject, values, formulas, formulasR1C1, rowCount, columnCount");

    return context.sync().then(function() {
        if (usedRange.isNullObject) {
            return result;
        }

        var values = usedRange.values;
        var formulas = usedRange.formulas;
        var formulasR1C1 = usedRange.formulasR1C1;
        var nRows = usedRange.rowCount;
        var nCols = usedRange.columnCount;

        var period = detectPeriod(values, nRows, nCols);
        result.periodRow = period.row;
        result.periodCol = period.col;

        var startCol = period.col > 0 ? period.col - 1 : auditConfig.labelCols;

        var totals = {};
        if (period.row > 0) {
            totals = detectTotals(values, period.row - 1, startCol, nCols);
        }

        // Find labeled rows
        var labeledRows = [];
        for (var r = 0; r < nRows; r++) {
            var lbl = getLabel(values, r);
            if (lbl.length === 0) continue;

            var hasData = false;
            for (var c = startCol; c < nCols; c++) {
                if (values[r][c] !== null && values[r][c] !== "" && values[r][c] !== undefined) {
                    hasData = true;
                    break;
                }
                if (typeof formulas[r][c] === "string" && formulas[r][c].charAt(0) === "=") {
                    hasData = true;
                    break;
                }
            }

            if (hasData) {
                labeledRows.push({ idx: r, num: r + 1, label: lbl });
            }
        }

        if (labeledRows.length === 0) {
            return result;
        }

        // Build map data and track cell positions for coloring
        var gridCols = nCols - startCol;
        var mapData = [];
        var colorCells = { H: [], X: [], E: [], ok: [], S: [] };

        // Header
        var hdr = ["Row", "Label", "#"];
        for (var gc = 0; gc < gridCols; gc++) {
            hdr.push(colLetter(startCol + gc + 1));
        }
        mapData.push(hdr);

        // Audit each labeled row
        for (var li = 0; li < labeledRows.length; li++) {
            var lr = labeledRows[li];
            var r = lr.idx;
            var dom = getDominant(formulasR1C1, r, startCol, nCols - 1, totals);

            var rowIssues = 0;
            var rowData = [lr.num, lr.label, ""];

            for (var c = startCol; c < nCols; c++) {
                var val = values[r][c];
                var frm = formulas[r][c];
                var frc = formulasR1C1[r][c];
                var isTotal = totals[c] === true;

                var mark = classify(val, frm, frc, dom, isTotal);
                rowData.push(mark);

                // Track position for coloring (row in map = li+1 because of header, col = c-startCol+3 for D onwards)
                var mapRowIdx = li + 1; // 0 is header
                var mapColIdx = c - startCol + 3; // columns D, E, F...
                
                if (mark === "H") colorCells.H.push({ r: mapRowIdx, c: mapColIdx });
                else if (mark === "X") colorCells.X.push({ r: mapRowIdx, c: mapColIdx });
                else if (mark === "E") colorCells.E.push({ r: mapRowIdx, c: mapColIdx });
                else if (mark === "S") colorCells.S.push({ r: mapRowIdx, c: mapColIdx });
                else if (mark === ".") colorCells.ok.push({ r: mapRowIdx, c: mapColIdx });

                if ((mark === "H" || mark === "X" || mark === "E") && auditState.issues.length < auditConfig.maxIssues) {
                    rowIssues++;
                    var issue = {
                        type: mark,
                        sheet: sheetName,
                        cell: colLetter(c + 1) + (r + 1),
                        actual: mark === "H" ? trunc(val) : truncF(frm),
                        expected: mark === "X" ? ("Dominant: " + dom) : (mark === "H" ? "(expects formula)" : truncF(frm))
                    };
                    auditState.issues.push(issue);

                    if (mark === "H") result.nH++;
                    else if (mark === "X") result.nX++;
                    else if (mark === "E") result.nE++;
                }
            }

            rowData[2] = rowIssues === 0 ? "✓" : String(rowIssues);
            mapData.push(rowData);
        }

        // Create audit map sheet with colors
        return createMapSheet(context, result, mapData, gridCols, colorCells);
    }).then(function() {
        return result;
    });
}

// ============================================
// CREATE AUDIT MAP SHEET WITH COLORS
// ============================================
function createMapSheet(context, result, mapData, gridCols, colorCells) {
    var ws = context.workbook.worksheets.add(result.auditName);

    // Title
    ws.getRange("A1").values = [["AUDIT: " + result.sheetName]];
    ws.getRange("A1").format.font.bold = true;
    ws.getRange("A1").format.font.size = 13;

    // Period info
    var info = result.periodCol > 0 
        ? "Period: row " + result.periodRow + " col " + colLetter(result.periodCol)
        : "No period detected";
    ws.getRange("A2").values = [[info]];
    ws.getRange("A2").format.font.color = "#888888";

    // Legend with colors
    ws.getRange("A3").values = [[". OK"]];
    ws.getRange("A3").format.fill.color = "#DAF2DA";
    ws.getRange("B3").values = [["H Hardcode"]];
    ws.getRange("B3").format.fill.color = "#FFFF00"; // Yellow
    ws.getRange("C3").values = [["X Break"]];
    ws.getRange("C3").format.fill.color = "#FF6B6B"; // Red
    ws.getRange("D3").values = [["E Error"]];
    ws.getRange("D3").format.fill.color = "#FFA500"; // Orange
    ws.getRange("E3").values = [["S Total"]];
    ws.getRange("E3").format.fill.color = "#FFF2CC";
    ws.getRange("A3:E3").format.font.size = 8;
    ws.getRange("A3:E3").format.font.bold = true;

    // Write map data at row 5
    if (mapData.length > 0 && mapData[0].length > 0) {
        var dataRng = ws.getRange("A5").getResizedRange(mapData.length - 1, mapData[0].length - 1);
        dataRng.values = mapData;

        // Header row formatting
        var hdrRng = ws.getRange("A5").getResizedRange(0, mapData[0].length - 1);
        hdrRng.format.font.bold = true;

        // Grid header (columns D+)
        if (gridCols > 0) {
            var gridHdr = ws.getRange("D5").getResizedRange(0, gridCols - 1);
            gridHdr.format.fill.color = "#44546A";
            gridHdr.format.font.color = "#FFFFFF";
            gridHdr.format.font.size = 7;
            gridHdr.format.horizontalAlignment = "Center";
        }

        // Data area base formatting
        if (mapData.length > 1 && gridCols > 0) {
            var gridData = ws.getRange("D6").getResizedRange(mapData.length - 2, gridCols - 1);
            gridData.format.horizontalAlignment = "Center";
            gridData.format.font.size = 9;
            gridData.format.font.bold = true;

            // Check column
            var chkCol = ws.getRange("C6").getResizedRange(mapData.length - 2, 0);
            chkCol.format.horizontalAlignment = "Center";
            chkCol.format.font.size = 8;

            // Row number column
            var rowNumCol = ws.getRange("A6").getResizedRange(mapData.length - 2, 0);
            rowNumCol.format.font.size = 8;
            rowNumCol.format.font.color = "#999999";

            // Label column
            var lblCol = ws.getRange("B6").getResizedRange(mapData.length - 2, 0);
            lblCol.format.font.size = 8;
        }

        // Apply colors to individual cells
        // H = Yellow
        for (var i = 0; i < colorCells.H.length; i++) {
            var pos = colorCells.H[i];
            var cellAddr = colLetter(pos.c + 1) + (pos.r + 5); // +5 because data starts at row 6
            ws.getRange(cellAddr).format.fill.color = "#FFFF00";
            ws.getRange(cellAddr).format.font.color = "#000000";
        }

        // X = Red
        for (var i = 0; i < colorCells.X.length; i++) {
            var pos = colorCells.X[i];
            var cellAddr = colLetter(pos.c + 1) + (pos.r + 5);
            ws.getRange(cellAddr).format.fill.color = "#FF6B6B";
            ws.getRange(cellAddr).format.font.color = "#FFFFFF";
        }

        // E = Orange
        for (var i = 0; i < colorCells.E.length; i++) {
            var pos = colorCells.E[i];
            var cellAddr = colLetter(pos.c + 1) + (pos.r + 5);
            ws.getRange(cellAddr).format.fill.color = "#FFA500";
            ws.getRange(cellAddr).format.font.color = "#FFFFFF";
        }

        // S = Pale gold
        for (var i = 0; i < colorCells.S.length; i++) {
            var pos = colorCells.S[i];
            var cellAddr = colLetter(pos.c + 1) + (pos.r + 5);
            ws.getRange(cellAddr).format.fill.color = "#FFF2CC";
            ws.getRange(cellAddr).format.font.color = "#8C8250";
        }

        // . = Green
        for (var i = 0; i < colorCells.ok.length; i++) {
            var pos = colorCells.ok[i];
            var cellAddr = colLetter(pos.c + 1) + (pos.r + 5);
            ws.getRange(cellAddr).format.fill.color = "#DAF2DA";
            ws.getRange(cellAddr).format.font.color = "#228B22";
        }

        // Color check marks
        for (var row = 0; row < mapData.length - 1; row++) {
            var checkAddr = "C" + (row + 6);
            var checkVal = mapData[row + 1][2];
            if (checkVal === "✓") {
                ws.getRange(checkAddr).format.font.color = "#008000";
            } else {
                ws.getRange(checkAddr).format.font.color = "#CC0000";
                ws.getRange(checkAddr).format.font.bold = true;
            }
        }
    }

    // Column widths
    ws.getRange("A:A").format.columnWidth = 35;
    ws.getRange("B:B").format.columnWidth = 150;
    ws.getRange("C:C").format.columnWidth = 25;

    return context.sync();
}

// ============================================
// CREATE DASHBOARD WITH COUNTIF FORMULAS AND TABLE
// ============================================
function makeDashboard(context) {
    var ws = context.workbook.worksheets.add(auditConfig.dashboardName);
    ws.position = 0;

    var totalIssues = auditState.totalH + auditState.totalX + auditState.totalE;
    var issueCount = auditState.issues.length;

    // Title
    ws.getRange("A1").values = [["MODEL AUDIT DASHBOARD"]];
    ws.getRange("A1").format.font.bold = true;
    ws.getRange("A1").format.font.size = 16;
    ws.getRange("A1").format.font.color = "#1E3C64";

    ws.getRange("A2").values = [[new Date().toLocaleString()]];
    ws.getRange("A2").format.font.color = "#888888";

    // Calculate where issue table starts
    var sheetTableEndRow = 5 + auditState.results.length;
    var issueHeaderRow = sheetTableEndRow + 3;
    var issueDataStartRow = issueHeaderRow + 1;
    var issueDataEndRow = issueDataStartRow + issueCount - 1;
    var issueTypeColAddr = "$A$" + issueDataStartRow + ":$A$" + (issueDataStartRow + Math.max(issueCount - 1, 0));

    // Scoreboard with COUNTIF formulas (will auto-update when you delete rows)
    ws.getRange("H1").values = [["Total"]];
    ws.getRange("H1").format.font.bold = true;
    ws.getRange("H1").format.font.size = 9;
    // Total formula = count of H + X + E in issue table
    if (issueCount > 0) {
        ws.getRange("H2").formulas = [["=COUNTA(" + issueTypeColAddr + ")"]];
    } else {
        ws.getRange("H2").values = [[0]];
    }
    ws.getRange("H2").format.font.size = 20;
    ws.getRange("H2").format.font.bold = true;
    ws.getRange("H2").format.fill.color = totalIssues === 0 ? "#DAF2DA" : "#FFC7CE";

    ws.getRange("J1").values = [["H (Hardcode)"]];
    ws.getRange("J1").format.font.bold = true;
    ws.getRange("J1").format.font.size = 9;
    if (issueCount > 0) {
        ws.getRange("J2").formulas = [['=COUNTIF(' + issueTypeColAddr + ',"H")']];
    } else {
        ws.getRange("J2").values = [[0]];
    }
    ws.getRange("J2").format.font.size = 20;
    ws.getRange("J2").format.font.bold = true;
    ws.getRange("J2").format.fill.color = "#FFFF00";

    ws.getRange("L1").values = [["X (Break)"]];
    ws.getRange("L1").format.font.bold = true;
    ws.getRange("L1").format.font.size = 9;
    if (issueCount > 0) {
        ws.getRange("L2").formulas = [['=COUNTIF(' + issueTypeColAddr + ',"X")']];
    } else {
        ws.getRange("L2").values = [[0]];
    }
    ws.getRange("L2").format.font.size = 20;
    ws.getRange("L2").format.font.bold = true;
    ws.getRange("L2").format.fill.color = "#FF6B6B";
    ws.getRange("L2").format.font.color = "#FFFFFF";

    ws.getRange("N1").values = [["E (Error)"]];
    ws.getRange("N1").format.font.bold = true;
    ws.getRange("N1").format.font.size = 9;
    if (issueCount > 0) {
        ws.getRange("N2").formulas = [['=COUNTIF(' + issueTypeColAddr + ',"E")']];
    } else {
        ws.getRange("N2").values = [[0]];
    }
    ws.getRange("N2").format.font.size = 20;
    ws.getRange("N2").format.font.bold = true;
    ws.getRange("N2").format.fill.color = "#FFA500";
    ws.getRange("N2").format.font.color = "#FFFFFF";

    // Sheet summary table
    ws.getRange("A4:G4").values = [["Sheet", "Period", "H", "X", "E", "Total", "Map"]];
    ws.getRange("A4:G4").format.font.bold = true;
    ws.getRange("A4:G4").format.fill.color = "#44546A";
    ws.getRange("A4:G4").format.font.color = "#FFFFFF";

    // Write sheet results
    if (auditState.results.length > 0) {
        var data = [];
        for (var i = 0; i < auditState.results.length; i++) {
            var res = auditState.results[i];
            var pTxt = res.periodCol > 0 ? colLetter(res.periodCol) + " r" + res.periodRow : "N/A";
            var tot = res.nH + res.nX + res.nE;
            data.push([res.sheetName, pTxt, res.nH, res.nX, res.nE, tot, "-> " + res.auditName]);
        }
        ws.getRange("A5").getResizedRange(data.length - 1, 6).values = data;
    }

    // Totals row for sheets
    var totRow = 5 + auditState.results.length;
    ws.getRange("A" + totRow + ":G" + totRow).values = [["TOTAL", "", auditState.totalH, auditState.totalX, auditState.totalE, totalIssues, ""]];
    ws.getRange("A" + totRow + ":G" + totRow).format.font.bold = true;

    // Issue detail section
    ws.getRange("A" + issueHeaderRow).values = [["ISSUE DETAIL"]];
    ws.getRange("A" + issueHeaderRow).format.font.bold = true;
    ws.getRange("A" + issueHeaderRow).format.font.size = 13;

    // Issue table header
    var tblHdrRow = issueHeaderRow + 1;
    ws.getRange("A" + tblHdrRow + ":F" + tblHdrRow).values = [["Type", "Sheet", "Cell", "Actual", "Expected", "Go"]];
    ws.getRange("A" + tblHdrRow + ":F" + tblHdrRow).format.font.bold = true;
    ws.getRange("A" + tblHdrRow + ":F" + tblHdrRow).format.fill.color = "#44546A";
    ws.getRange("A" + tblHdrRow + ":F" + tblHdrRow).format.font.color = "#FFFFFF";

    // Write issues
    if (issueCount > 0) {
        var issData = [];
        for (var j = 0; j < issueCount; j++) {
            var iss = auditState.issues[j];
            issData.push([iss.type, iss.sheet, iss.cell, iss.actual || "", iss.expected || "", "Go"]);
        }
        var issDataRange = ws.getRange("A" + (tblHdrRow + 1)).getResizedRange(issData.length - 1, 5);
        issDataRange.values = issData;

        // Color the Type column
        for (var k = 0; k < issData.length; k++) {
            var typeCell = ws.getRange("A" + (tblHdrRow + 1 + k));
            var t = issData[k][0];
            if (t === "H") {
                typeCell.format.fill.color = "#FFFF00";
            } else if (t === "X") {
                typeCell.format.fill.color = "#FF6B6B";
                typeCell.format.font.color = "#FFFFFF";
            } else if (t === "E") {
                typeCell.format.fill.color = "#FFA500";
                typeCell.format.font.color = "#FFFFFF";
            }
        }

        // Style Go column
        var goCol = ws.getRange("F" + (tblHdrRow + 1)).getResizedRange(issData.length - 1, 0);
        goCol.format.font.color = "#0563C1";

        // Create Table with AutoFilter
        var tableRange = ws.getRange("A" + tblHdrRow + ":F" + (tblHdrRow + issData.length));
        var issueTable = ws.tables.add(tableRange, true);
        issueTable.name = "IssueTable";
        issueTable.style = "TableStyleMedium2";
    }

    // Column widths
    ws.getRange("A:A").format.columnWidth = 45;
    ws.getRange("B:B").format.columnWidth = 100;
    ws.getRange("C:C").format.columnWidth = 50;
    ws.getRange("D:D").format.columnWidth = 200;
    ws.getRange("E:E").format.columnWidth = 150;
    ws.getRange("F:F").format.columnWidth = 35;
    ws.getRange("G:G").format.columnWidth = 100;

    // Activate dashboard
    ws.activate();

    updateAuditProgress(100, "Complete!");
    return context.sync();
}

// ============================================
// HELPER FUNCTIONS
// ============================================
function isAuditSheet(name) {
    var u = name.toUpperCase();
    return u.indexOf(auditConfig.audPrefix.toUpperCase()) === 0 || u === auditConfig.dashboardName.toUpperCase();
}

function makeAuditName(name) {
    var n = auditConfig.audPrefix + name;
    if (n.length > 31) n = n.substring(0, 31);
    return n.replace(/[:\\\/?*\[\]]/g, "_");
}

function detectPeriod(values, nRows, nCols) {
    var res = { row: 0, col: 0 };
    var scanRows = Math.min(auditConfig.periodScanRows, nRows);
    var best = 0;

    for (var r = 0; r < scanRows; r++) {
        for (var c = 0; c < nCols - 2; c++) {
            var v1 = values[r][c], v2 = values[r][c + 1], v3 = values[r][c + 2];
            if (isNum(v1) && isNum(v2) && isNum(v3)) {
                if (Math.floor(v1) === 1 && Math.floor(v2) === 2 && Math.floor(v3) === 3) {
                    var run = 3;
                    for (var k = c + 3; k < nCols; k++) {
                        var vk = values[r][k];
                        if (isNum(vk) && Math.floor(vk) >= run + 1 && Math.floor(vk) <= run + 2) {
                            run = Math.floor(vk);
                        } else if (!isNum(vk)) {
                            // skip non-numeric (total col)
                        } else {
                            break;
                        }
                    }
                    if (run > best) {
                        best = run;
                        res.row = r + 1;
                        res.col = c + 1;
                    }
                }
            }
        }
    }
    return res;
}

function detectTotals(values, pRowIdx, startCol, nCols) {
    var tot = {};
    var row = values[pRowIdx];
    if (!row) return tot;
    for (var c = startCol; c < nCols; c++) {
        var v = row[c];
        if (v === null || v === "" || v === undefined || !isNum(v)) {
            tot[c] = true;
        }
    }
    return tot;
}

function getDominant(formulasR1C1, r, startCol, endCol, totals) {
    var counts = {};
    var nF = 0;

    for (var c = startCol; c <= endCol; c++) {
        if (totals[c]) continue;
        var f = formulasR1C1[r][c];
        if (typeof f !== "string" || f.length === 0 || f.charAt(0) !== "=") continue;
        nF++;
        counts[f] = (counts[f] || 0) + 1;
    }

    if (nF < auditConfig.minFormulasForDominant) return "";

    var maxCnt = 0, maxKey = "";
    for (var k in counts) {
        if (counts[k] > maxCnt) {
            maxCnt = counts[k];
            maxKey = k;
        }
    }

    return maxCnt >= nF * 0.4 ? maxKey : "";
}

function classify(val, frm, frc, dom, isTot) {
    var isF = typeof frm === "string" && frm.length > 0 && frm.charAt(0) === "=";
    if ((val === null || val === "" || val === undefined) && !isF) return "";
    if (isErr(val)) return "E";
    if (isTot) return "S";
    if (!isF) return "H";
    if (dom.length > 0 && frc !== dom) return "X";
    return ".";
}

function getLabel(values, r) {
    var best = "";
    var cols = Math.min(auditConfig.labelCols, values[r].length);
    for (var c = 0; c < cols; c++) {
        var v = values[r][c];
        if (v !== null && v !== "" && v !== undefined && !isErr(v)) {
            var s = String(v).trim();
            if (s.length > best.length) best = s;
        }
    }
    return best.length > 42 ? best.substring(0, 42) + "..." : best;
}

function isNum(v) { return typeof v === "number" && !isNaN(v); }

function isErr(v) {
    if (typeof v !== "string") return false;
    return v.charAt(0) === "#";
}

function colLetter(n) {
    var s = "";
    while (n > 0) {
        var m = (n - 1) % 26;
        s = String.fromCharCode(65 + m) + s;
        n = Math.floor((n - 1) / 26);
    }
    return s;
}

function getCellAddr(addr, r, c) {
    var start = addr.split("!").pop().split(":")[0];
    var m = start.match(/([A-Z]+)(\d+)/);
    if (m) {
        var sc = 0;
        for (var i = 0; i < m[1].length; i++) {
            sc = sc * 26 + (m[1].charCodeAt(i) - 64);
        }
        return colLetter(sc + c) + (parseInt(m[2]) + r);
    }
    return colLetter(c + 1) + (r + 1);
}

function trunc(v) {
    var s = String(v);
    return s.length > 60 ? s.substring(0, 60) + "..." : s;
}

function truncF(f) {
    return f && f.length > 60 ? f.substring(0, 60) + "..." : (f || "");
}

function escapeHtml(t) {
    var d = document.createElement('div');
    d.textContent = t;
    return d.innerHTML;
}

// ============================================
// UI DISPLAY
// ============================================
function showSummaryUI() {
    var div = document.getElementById('auditSummary');
    if (!div) return;

    var total = auditState.totalH + auditState.totalX + auditState.totalE;

    var html = '<div class="audit-summary-cards">';
    html += '<div class="audit-card" style="background:#FFA500;color:#fff;"><span class="count">' + auditState.totalE + '</span><span class="label">Errors (E)</span></div>';
    html += '<div class="audit-card" style="background:#FFFF00;color:#000;"><span class="count">' + auditState.totalH + '</span><span class="label">Hardcodes (H)</span></div>';
    html += '<div class="audit-card" style="background:#FF6B6B;color:#fff;"><span class="count">' + auditState.totalX + '</span><span class="label">Breaks (X)</span></div>';
    html += '</div>';

    if (auditState.results.length > 0) {
        html += '<div class="audit-issues-list">';
        html += '<p class="kicker">Sheets (' + auditState.results.length + ')</p>';
        for (var i = 0; i < Math.min(5, auditState.results.length); i++) {
            var res = auditState.results[i];
            var t = res.nH + res.nX + res.nE;
            var cls = t === 0 ? "info" : (t > 10 ? "critical" : "warning");
            html += '<div class="audit-issue-item ' + cls + '">';
            html += '<span class="issue-type">' + escapeHtml(res.sheetName) + '</span>';
            html += '<span class="issue-location">' + t + '</span>';
            html += '</div>';
        }
        if (auditState.results.length > 5) {
            html += '<p class="more-issues">+' + (auditState.results.length - 5) + ' more</p>';
        }
        html += '</div>';
    }

    div.innerHTML = html;
    div.style.display = 'block';
}
