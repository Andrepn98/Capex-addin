// ============================================
// ORNA MODULE: Financial Audit Tool
// Version 1.0 - Comprehensive Excel Workbook Auditing
// ============================================

// ============================================
// GLOBAL VARIABLES
// ============================================
var auditResults = {
    issues: [],
    sheetSummaries: [],
    externalLinks: [],
    namedRanges: [],
    circularRefs: [],
    metadata: {}
};

var auditSettings = {
    complexityThreshold: 7,
    maxFormulasDisplay: 250,
    checkVolatileFunctions: true,
    checkPatternBreaks: true,
    checkHardcodes: true,
    checkErrors: true
};

// ============================================
// INITIALIZE MODULE
// ============================================
function initAuditModule() {
    document.getElementById('runFullAudit').onclick = runFullAudit;
    document.getElementById('runQuickAudit').onclick = runQuickAudit;
    document.getElementById('auditSelection').onclick = auditSelectedRange;
    setAuditStatus("ok", "Ready to audit");
}

function setAuditStatus(type, text) {
    document.getElementById("auditStatusDot").className = "dot" + (type ? " " + type : "");
    document.getElementById("auditStatusText").textContent = text;
}

function updateAuditProgress(percent, message) {
    document.getElementById("auditProgress").style.width = percent + "%";
    document.getElementById("auditProgressText").textContent = message;
}

// ============================================
// MAIN AUDIT FUNCTIONS
// ============================================
function runFullAudit() {
    setAuditStatus("processing", "Running full audit...");
    updateAuditProgress(0, "Initializing...");
    toast("Starting full audit...");

    // Reset results
    auditResults = {
        issues: [],
        sheetSummaries: [],
        externalLinks: [],
        namedRanges: [],
        circularRefs: [],
        metadata: {}
    };

    Excel.run(function(context) {
        var workbook = context.workbook;
        var sheets = workbook.worksheets;
        sheets.load("items/name");

        // Load named ranges
        var names = workbook.names;
        names.load("items");

        return context.sync().then(function() {
            updateAuditProgress(10, "Analyzing workbook structure...");

            // Store metadata
            auditResults.metadata = {
                sheetCount: sheets.items.length,
                auditDate: new Date().toISOString()
            };

            // Audit named ranges
            return auditNamedRanges(context, names.items);
        }).then(function() {
            updateAuditProgress(20, "Auditing sheets...");

            // Filter sheets to audit
            var sheetsToAudit = sheets.items.filter(function(sheet) {
                return shouldAuditSheet(sheet.name);
            });

            // Audit each sheet sequentially
            var sheetPromises = [];
            var progressPerSheet = 60 / Math.max(sheetsToAudit.length, 1);

            sheetsToAudit.forEach(function(sheet, index) {
                sheetPromises.push(
                    auditWorksheet(context, sheet).then(function(summary) {
                        auditResults.sheetSummaries.push(summary);
                        updateAuditProgress(20 + (index + 1) * progressPerSheet, 
                            "Audited: " + sheet.name);
                    })
                );
            });

            return Promise.all(sheetPromises);
        }).then(function() {
            updateAuditProgress(85, "Generating reports...");

            // Create audit report sheets
            return createAuditReports(context);
        }).then(function() {
            updateAuditProgress(100, "Complete!");
            return context.sync();
        });
    }).then(function() {
        setAuditStatus("ok", "Audit complete! " + auditResults.issues.length + " issues found");
        toast("Audit complete!");
        displayAuditSummary();
    }).catch(function(error) {
        setAuditStatus("error", "Error: " + error.message);
        toast("Audit failed: " + error.message);
        console.error(error);
    });
}

function runQuickAudit() {
    setAuditStatus("processing", "Running quick audit...");
    updateAuditProgress(0, "Initializing...");
    toast("Starting quick audit...");

    auditResults = {
        issues: [],
        sheetSummaries: [],
        externalLinks: [],
        namedRanges: [],
        circularRefs: [],
        metadata: {}
    };

    Excel.run(function(context) {
        var sheets = context.workbook.worksheets;
        sheets.load("items/name");

        return context.sync().then(function() {
            var sheetsToAudit = sheets.items.filter(function(sheet) {
                return shouldAuditSheet(sheet.name);
            });

            var promises = sheetsToAudit.map(function(sheet, index) {
                return auditWorksheetQuick(context, sheet).then(function(summary) {
                    auditResults.sheetSummaries.push(summary);
                    updateAuditProgress((index + 1) / sheetsToAudit.length * 80, 
                        "Audited: " + sheet.name);
                });
            });

            return Promise.all(promises);
        }).then(function() {
            updateAuditProgress(90, "Generating report...");
            return createQuickReport(context);
        }).then(function() {
            updateAuditProgress(100, "Complete!");
            return context.sync();
        });
    }).then(function() {
        setAuditStatus("ok", "Quick audit done! " + auditResults.issues.length + " issues");
        toast("Quick audit complete!");
        displayAuditSummary();
    }).catch(function(error) {
        setAuditStatus("error", "Error: " + error.message);
        console.error(error);
    });
}

function auditSelectedRange() {
    setAuditStatus("processing", "Auditing selection...");

    auditResults.issues = [];

    Excel.run(function(context) {
        var range = context.workbook.getSelectedRange();
        range.load("address, values, formulas, formulasR1C1, worksheet/name");

        return context.sync().then(function() {
            var sheetName = range.worksheet.name;
            var values = range.values;
            var formulas = range.formulas;
            var formulasR1C1 = range.formulasR1C1;

            for (var r = 0; r < values.length; r++) {
                for (var c = 0; c < values[r].length; c++) {
                    var cellAddress = getCellAddress(range.address, r, c);
                    auditSingleCell(
                        values[r][c],
                        formulas[r][c],
                        formulasR1C1[r][c],
                        sheetName,
                        cellAddress,
                        r, c,
                        formulas
                    );
                }
            }

            return context.sync();
        });
    }).then(function() {
        setAuditStatus("ok", "Found " + auditResults.issues.length + " issues in selection");
        toast("Selection audit complete!");
        displayAuditSummary();
    }).catch(function(error) {
        setAuditStatus("error", "Error: " + error.message);
        console.error(error);
    });
}

// ============================================
// WORKSHEET AUDIT
// ============================================
function auditWorksheet(context, sheet) {
    return new Promise(function(resolve, reject) {
        var summary = {
            sheetName: sheet.name,
            totalCells: 0,
            formulaCells: 0,
            hardcodes: 0,
            patternBreaks: 0,
            shiftedRefs: 0,
            complexFormulas: 0,
            errorCells: 0,
            volatileFunctions: 0
        };

        var usedRange = sheet.getUsedRange();
        usedRange.load("address, values, formulas, formulasR1C1, rowCount, columnCount");

        context.sync().then(function() {
            var values = usedRange.values;
            var formulas = usedRange.formulas;
            var formulasR1C1 = usedRange.formulasR1C1;

            summary.totalCells = usedRange.rowCount * usedRange.columnCount;

            // Analyze each cell
            for (var r = 0; r < values.length; r++) {
                for (var c = 0; c < values[r].length; c++) {
                    var value = values[r][c];
                    var formula = formulas[r][c];
                    var formulaR1C1 = formulasR1C1[r][c];
                    var cellAddress = getCellAddress(usedRange.address, r, c);

                    // Check if formula
                    if (typeof formula === 'string' && formula.startsWith('=')) {
                        summary.formulaCells++;

                        // Check for errors
                        if (isErrorValue(value)) {
                            summary.errorCells++;
                            addIssue("Error", "Critical", sheet.name, cellAddress,
                                getErrorDescription(value), String(value), formula,
                                getErrorRecommendation(value));
                        }

                        // Check for volatile functions
                        if (containsVolatileFunction(formula)) {
                            summary.volatileFunctions++;
                            addIssue("Volatile Function", "Warning", sheet.name, cellAddress,
                                "Contains volatile function that recalculates constantly",
                                truncateValue(value), truncateFormula(formula),
                                "Consider replacing with static value");
                        }

                        // Check complexity
                        var complexity = calculateFormulaComplexity(formula);
                        if (complexity >= auditSettings.complexityThreshold) {
                            summary.complexFormulas++;
                            addIssue("Complex Formula", "Info", sheet.name, cellAddress,
                                "Complexity score: " + complexity + "/10",
                                truncateValue(value), truncateFormula(formula),
                                "Consider breaking into helper columns");
                        }

                        // Check pattern break (compare with left neighbor)
                        if (c > 0) {
                            var leftFormulaR1C1 = formulasR1C1[r][c - 1];
                            if (typeof leftFormulaR1C1 === 'string' && leftFormulaR1C1.startsWith('=')) {
                                if (formulaR1C1 !== leftFormulaR1C1) {
                                    summary.patternBreaks++;
                                    addIssue("Pattern Break", "Warning", sheet.name, cellAddress,
                                        "Formula pattern differs from left neighbor",
                                        truncateValue(value), truncateFormula(formula),
                                        "Verify this is intentional");
                                }
                            }
                        }

                        // Check shifted references
                        if (hasShiftedColumnReference(formula, c + 1)) {
                            summary.shiftedRefs++;
                            addIssue("Shifted Reference", "Warning", sheet.name, cellAddress,
                                "Inconsistent column references",
                                truncateValue(value), truncateFormula(formula),
                                "Check if references should be absolute ($)");
                        }

                    } else if (value !== null && value !== "") {
                        // Non-formula cell - check for hardcode
                        if (isHardcodeInFormulaArea(formulas, r, c)) {
                            summary.hardcodes++;
                            addIssue("Hardcode", "Warning", sheet.name, cellAddress,
                                "Constant value in formula area",
                                truncateValue(value), "",
                                "Consider using formula or input reference");
                        }
                    }
                }
            }

            resolve(summary);
        }).catch(function(error) {
            console.error("Error auditing sheet " + sheet.name + ": " + error.message);
            resolve(summary); // Return partial summary on error
        });
    });
}

function auditWorksheetQuick(context, sheet) {
    return new Promise(function(resolve, reject) {
        var summary = {
            sheetName: sheet.name,
            totalCells: 0,
            formulaCells: 0,
            hardcodes: 0,
            errorCells: 0
        };

        var usedRange = sheet.getUsedRange();
        usedRange.load("values, formulas, rowCount, columnCount");

        context.sync().then(function() {
            var values = usedRange.values;
            var formulas = usedRange.formulas;

            summary.totalCells = usedRange.rowCount * usedRange.columnCount;

            for (var r = 0; r < values.length; r++) {
                for (var c = 0; c < values[r].length; c++) {
                    var value = values[r][c];
                    var formula = formulas[r][c];
                    var cellAddress = getCellAddressFromRC(r, c);

                    if (typeof formula === 'string' && formula.startsWith('=')) {
                        summary.formulaCells++;

                        if (isErrorValue(value)) {
                            summary.errorCells++;
                            addIssue("Error", "Critical", sheet.name, cellAddress,
                                getErrorDescription(value), String(value), formula,
                                getErrorRecommendation(value));
                        }
                    } else if (value !== null && value !== "") {
                        if (isHardcodeInFormulaArea(formulas, r, c)) {
                            summary.hardcodes++;
                            addIssue("Hardcode", "Warning", sheet.name, cellAddress,
                                "Constant in formula area", truncateValue(value), "",
                                "Review this hardcoded value");
                        }
                    }
                }
            }

            resolve(summary);
        }).catch(function(error) {
            resolve(summary);
        });
    });
}

// ============================================
// NAMED RANGES AUDIT
// ============================================
function auditNamedRanges(context, names) {
    return new Promise(function(resolve) {
        names.forEach(function(name) {
            var nameInfo = {
                name: name.name,
                refersTo: "",
                isValid: true,
                scope: "Workbook"
            };

            try {
                nameInfo.refersTo = name.formula;
                
                // Check for #REF! errors
                if (nameInfo.refersTo.indexOf("#REF!") >= 0) {
                    nameInfo.isValid = false;
                    addIssue("Invalid Named Range", "Critical", "(Workbook)", name.name,
                        "Named range refers to #REF!", nameInfo.refersTo, "",
                        "Delete or fix the named range");
                }
            } catch (e) {
                nameInfo.isValid = false;
            }

            auditResults.namedRanges.push(nameInfo);
        });

        resolve();
    });
}

// ============================================
// CELL ANALYSIS FUNCTIONS
// ============================================
function auditSingleCell(value, formula, formulaR1C1, sheetName, cellAddress, row, col, allFormulas) {
    // Error check
    if (isErrorValue(value)) {
        addIssue("Error", "Critical", sheetName, cellAddress,
            getErrorDescription(value), String(value), formula,
            getErrorRecommendation(value));
        return;
    }

    if (typeof formula === 'string' && formula.startsWith('=')) {
        // Volatile function check
        if (containsVolatileFunction(formula)) {
            addIssue("Volatile Function", "Warning", sheetName, cellAddress,
                "Contains volatile function", truncateValue(value), truncateFormula(formula),
                "Consider static alternative");
        }

        // Complexity check
        var complexity = calculateFormulaComplexity(formula);
        if (complexity >= auditSettings.complexityThreshold) {
            addIssue("Complex Formula", "Info", sheetName, cellAddress,
                "Complexity: " + complexity + "/10", truncateValue(value), truncateFormula(formula),
                "Break into helper columns");
        }

        // Shifted reference check
        if (hasShiftedColumnReference(formula, col + 1)) {
            addIssue("Shifted Reference", "Warning", sheetName, cellAddress,
                "Inconsistent column refs", truncateValue(value), truncateFormula(formula),
                "Use absolute references ($)");
        }
    } else if (value !== null && value !== "") {
        // Hardcode check
        if (allFormulas && isHardcodeInFormulaArea(allFormulas, row, col)) {
            addIssue("Hardcode", "Warning", sheetName, cellAddress,
                "Constant in formula area", truncateValue(value), "",
                "Use formula or input cell");
        }
    }
}

// ============================================
// FORMULA ANALYSIS FUNCTIONS
// ============================================
function containsVolatileFunction(formula) {
    var volatileFuncs = ['NOW(', 'TODAY(', 'RAND(', 'RANDBETWEEN(', 'INDIRECT(', 'OFFSET(', 'INFO(', 'CELL('];
    var upperFormula = formula.toUpperCase();
    
    for (var i = 0; i < volatileFuncs.length; i++) {
        if (upperFormula.indexOf(volatileFuncs[i]) >= 0) {
            return true;
        }
    }
    return false;
}

function calculateFormulaComplexity(formula) {
    var score = 0;

    // Length factor (0-2)
    if (formula.length > 200) score += 2;
    else if (formula.length > 100) score += 1;
    else if (formula.length > 50) score += 0.5;

    // Nesting depth (0-3)
    var maxDepth = 0, currentDepth = 0;
    for (var i = 0; i < formula.length; i++) {
        if (formula[i] === '(') {
            currentDepth++;
            if (currentDepth > maxDepth) maxDepth = currentDepth;
        } else if (formula[i] === ')') {
            currentDepth--;
        }
    }
    if (maxDepth > 5) score += 3;
    else if (maxDepth > 3) score += 2;
    else if (maxDepth > 1) score += 1;

    // Function count (0-2)
    var funcMatches = formula.match(/[A-Z][A-Z0-9_]*\(/gi);
    var funcCount = funcMatches ? funcMatches.length : 0;
    if (funcCount > 10) score += 2;
    else if (funcCount > 5) score += 1;
    else if (funcCount > 2) score += 0.5;

    // Reference count (0-2)
    var refMatches = formula.match(/\$?[A-Z]{1,3}\$?\d+/gi);
    var refCount = refMatches ? refMatches.length : 0;
    if (refCount > 20) score += 2;
    else if (refCount > 10) score += 1;
    else if (refCount > 5) score += 0.5;

    return Math.min(10, Math.round(score));
}

function hasShiftedColumnReference(formula, currentCol) {
    if (!formula || !formula.startsWith('=')) return false;

    var refPattern = /\$?([A-Z]{1,3})\$?\d+/gi;
    var matches = formula.match(refPattern);
    
    if (!matches || matches.length === 0) return false;

    var foundSameCol = false;
    var foundDiffCol = false;

    matches.forEach(function(match) {
        var colLetters = match.replace(/[\$\d]/g, '');
        var refCol = columnLetterToNumber(colLetters);
        
        if (refCol === currentCol) {
            foundSameCol = true;
        } else if (Math.abs(refCol - currentCol) <= 3) {
            foundDiffCol = true;
        }
    });

    return foundSameCol && foundDiffCol;
}

function isHardcodeInFormulaArea(formulas, row, col) {
    var formulaNeighbors = 0;
    var directions = [[-1, 0], [1, 0], [0, -1], [0, 1]]; // up, down, left, right

    directions.forEach(function(dir) {
        var r = row + dir[0];
        var c = col + dir[1];
        
        if (r >= 0 && r < formulas.length && c >= 0 && c < formulas[0].length) {
            var neighborFormula = formulas[r][c];
            if (typeof neighborFormula === 'string' && neighborFormula.startsWith('=')) {
                formulaNeighbors++;
            }
        }
    });

    return formulaNeighbors >= 2;
}

// ============================================
// ERROR HANDLING
// ============================================
function isErrorValue(value) {
    if (typeof value === 'string') {
        return value.startsWith('#') && (
            value === '#DIV/0!' ||
            value === '#N/A' ||
            value === '#NAME?' ||
            value === '#NULL!' ||
            value === '#NUM!' ||
            value === '#REF!' ||
            value === '#VALUE!' ||
            value === '#CALC!' ||
            value === '#SPILL!'
        );
    }
    return false;
}

function getErrorDescription(value) {
    var errors = {
        '#DIV/0!': 'Division by zero',
        '#N/A': 'Value not available',
        '#NAME?': 'Unrecognized name',
        '#NULL!': 'Incorrect range reference',
        '#NUM!': 'Invalid numeric value',
        '#REF!': 'Invalid cell reference',
        '#VALUE!': 'Wrong value type',
        '#CALC!': 'Calculation error',
        '#SPILL!': 'Spill range blocked'
    };
    return errors[value] || 'Unknown error';
}

function getErrorRecommendation(value) {
    var recommendations = {
        '#DIV/0!': 'Check denominator; use IFERROR wrapper',
        '#N/A': 'Verify lookup value exists; use IFNA',
        '#NAME?': 'Check for typos or undefined names',
        '#NULL!': 'Check range intersection syntax',
        '#NUM!': 'Check for invalid numeric arguments',
        '#REF!': 'Update deleted cell references',
        '#VALUE!': 'Check data types in formula',
        '#CALC!': 'Review calculation logic',
        '#SPILL!': 'Clear blocking cells in spill range'
    };
    return recommendations[value] || 'Review formula logic';
}

// ============================================
// HELPER FUNCTIONS
// ============================================
function shouldAuditSheet(sheetName) {
    var upperName = sheetName.toUpperCase();
    return upperName !== 'AUDIT_MASTER' &&
           upperName !== 'AUDIT_ISSUES' &&
           !upperName.startsWith('AUDIT_') &&
           upperName !== 'ISSUES' &&
           upperName !== 'EXTERNAL_LINKS' &&
           upperName !== 'NAMED_RANGES';
}

function addIssue(type, severity, sheet, cell, description, value, formula, recommendation) {
    auditResults.issues.push({
        type: type,
        severity: severity,
        sheet: sheet,
        cell: cell,
        description: description,
        value: value,
        formula: formula,
        recommendation: recommendation
    });
}

function truncateValue(value) {
    var str = String(value);
    return str.length > 50 ? str.substring(0, 50) + '...' : str;
}

function truncateFormula(formula) {
    return formula.length > auditSettings.maxFormulasDisplay 
        ? formula.substring(0, auditSettings.maxFormulasDisplay) + '...' 
        : formula;
}

function columnLetterToNumber(letters) {
    var num = 0;
    letters = letters.toUpperCase();
    for (var i = 0; i < letters.length; i++) {
        num = num * 26 + (letters.charCodeAt(i) - 64);
    }
    return num;
}

function numberToColumnLetter(num) {
    var result = '';
    while (num > 0) {
        var remainder = (num - 1) % 26;
        result = String.fromCharCode(65 + remainder) + result;
        num = Math.floor((num - 1) / 26);
    }
    return result;
}

function getCellAddress(rangeAddress, row, col) {
    // Extract start cell from range address
    var startCell = rangeAddress.split('!').pop().split(':')[0];
    var match = startCell.match(/([A-Z]+)(\d+)/);
    if (match) {
        var startCol = columnLetterToNumber(match[1]);
        var startRow = parseInt(match[2]);
        return numberToColumnLetter(startCol + col) + (startRow + row);
    }
    return getCellAddressFromRC(row, col);
}

function getCellAddressFromRC(row, col) {
    return numberToColumnLetter(col + 1) + (row + 1);
}

// ============================================
// REPORT GENERATION
// ============================================
function createAuditReports(context) {
    return new Promise(function(resolve, reject) {
        var sheets = context.workbook.worksheets;

        // Delete existing audit sheets
        deleteSheetIfExists(context, 'AUDIT_MASTER');
        deleteSheetIfExists(context, 'AUDIT_ISSUES');

        context.sync().then(function() {
            // Create master sheet
            var masterSheet = sheets.add('AUDIT_MASTER');
            masterSheet.position = 0;

            writeMasterDashboard(masterSheet);

            // Create issues sheet
            var issuesSheet = sheets.add('AUDIT_ISSUES');
            writeIssuesSheet(issuesSheet);

            return context.sync();
        }).then(function() {
            resolve();
        }).catch(function(error) {
            reject(error);
        });
    });
}

function createQuickReport(context) {
    return new Promise(function(resolve, reject) {
        var sheets = context.workbook.worksheets;

        deleteSheetIfExists(context, 'AUDIT_QUICK');

        context.sync().then(function() {
            var reportSheet = sheets.add('AUDIT_QUICK');
            reportSheet.position = 0;

            writeQuickReport(reportSheet);

            return context.sync();
        }).then(function() {
            resolve();
        }).catch(reject);
    });
}

function deleteSheetIfExists(context, sheetName) {
    try {
        var sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
        sheet.load('isNullObject');
        
        return context.sync().then(function() {
            if (!sheet.isNullObject) {
                sheet.delete();
            }
        });
    } catch (e) {
        // Sheet doesn't exist, that's fine
    }
}

function writeMasterDashboard(sheet) {
    // Title
    sheet.getRange("A1").values = [["ORNA FINANCIAL AUDIT REPORT"]];
    sheet.getRange("A1").format.font.bold = true;
    sheet.getRange("A1").format.font.size = 20;
    sheet.getRange("A1").format.font.color = "#2F6DB3";

    sheet.getRange("A2").values = [["Generated: " + new Date().toLocaleString()]];
    sheet.getRange("A2").format.font.color = "#808080";

    // Summary cards
    var criticalCount = auditResults.issues.filter(function(i) { return i.severity === 'Critical'; }).length;
    var warningCount = auditResults.issues.filter(function(i) { return i.severity === 'Warning'; }).length;
    var infoCount = auditResults.issues.filter(function(i) { return i.severity === 'Info'; }).length;

    sheet.getRange("A4:B4").values = [["TOTAL ISSUES", auditResults.issues.length]];
    sheet.getRange("A4").format.font.bold = true;
    sheet.getRange("B4").format.font.bold = true;
    sheet.getRange("B4").format.font.size = 18;

    sheet.getRange("C4:D4").values = [["CRITICAL", criticalCount]];
    sheet.getRange("C4").format.font.bold = true;
    sheet.getRange("D4").format.font.bold = true;
    sheet.getRange("D4").format.font.size = 18;
    sheet.getRange("D4").format.font.color = "#C00000";
    if (criticalCount > 0) {
        sheet.getRange("C4:D4").format.fill.color = "#FFC7CE";
    }

    sheet.getRange("E4:F4").values = [["WARNING", warningCount]];
    sheet.getRange("E4").format.font.bold = true;
    sheet.getRange("F4").format.font.bold = true;
    sheet.getRange("F4").format.font.size = 18;
    sheet.getRange("F4").format.font.color = "#9C6500";
    if (warningCount > 0) {
        sheet.getRange("E4:F4").format.fill.color = "#FFEB9C";
    }

    sheet.getRange("G4:H4").values = [["INFO", infoCount]];
    sheet.getRange("G4").format.font.bold = true;
    sheet.getRange("H4").format.font.bold = true;
    sheet.getRange("H4").format.font.size = 18;
    sheet.getRange("H4").format.font.color = "#006400";

    // Sheet summary table
    sheet.getRange("A7").values = [["SHEET SUMMARY"]];
    sheet.getRange("A7").format.font.bold = true;
    sheet.getRange("A7").format.font.size = 14;

    var headers = [["Sheet", "Cells", "Formulas", "Hardcodes", "Errors", "Pattern Breaks", "Complex"]];
    sheet.getRange("A8:G8").values = headers;
    sheet.getRange("A8:G8").format.font.bold = true;
    sheet.getRange("A8:G8").format.fill.color = "#2F6DB3";
    sheet.getRange("A8:G8").format.font.color = "#FFFFFF";

    // Write sheet data
    var dataStartRow = 9;
    auditResults.sheetSummaries.forEach(function(summary, index) {
        var row = dataStartRow + index;
        var rowData = [[
            summary.sheetName,
            summary.totalCells,
            summary.formulaCells || 0,
            summary.hardcodes || 0,
            summary.errorCells || 0,
            summary.patternBreaks || 0,
            summary.complexFormulas || 0
        ]];
        sheet.getRange("A" + row + ":G" + row).values = rowData;

        // Highlight rows with issues
        if ((summary.hardcodes || 0) + (summary.errorCells || 0) + (summary.patternBreaks || 0) > 0) {
            sheet.getRange("A" + row + ":G" + row).format.fill.color = "#FFF2CC";
        }
    });

    // Auto-fit columns
    sheet.getRange("A:H").format.autofitColumns();
}

function writeIssuesSheet(sheet) {
    // Title
    sheet.getRange("A1").values = [["ALL ISSUES"]];
    sheet.getRange("A1").format.font.bold = true;
    sheet.getRange("A1").format.font.size = 16;

    sheet.getRange("A2").values = [["Total: " + auditResults.issues.length + " issues"]];

    // Headers
    var headers = [["Severity", "Type", "Sheet", "Cell", "Description", "Value", "Formula", "Recommendation"]];
    sheet.getRange("A4:H4").values = headers;
    sheet.getRange("A4:H4").format.font.bold = true;
    sheet.getRange("A4:H4").format.fill.color = "#2F6DB3";
    sheet.getRange("A4:H4").format.font.color = "#FFFFFF";

    // Write issues
    auditResults.issues.forEach(function(issue, index) {
        var row = 5 + index;
        var rowData = [[
            issue.severity,
            issue.type,
            issue.sheet,
            issue.cell,
            issue.description,
            issue.value,
            issue.formula,
            issue.recommendation
        ]];
        sheet.getRange("A" + row + ":H" + row).values = rowData;

        // Color by severity
        var color = "#FFFFFF";
        if (issue.severity === 'Critical') color = "#FFC7CE";
        else if (issue.severity === 'Warning') color = "#FFEB9C";
        else if (issue.severity === 'Info') color = "#C6EFCE";
        
        sheet.getRange("A" + row + ":H" + row).format.fill.color = color;
    });

    // Auto-fit and set column widths
    sheet.getRange("A:A").format.columnWidth = 80;
    sheet.getRange("B:B").format.columnWidth = 120;
    sheet.getRange("C:C").format.columnWidth = 100;
    sheet.getRange("D:D").format.columnWidth = 60;
    sheet.getRange("E:E").format.columnWidth = 200;
    sheet.getRange("F:F").format.columnWidth = 100;
    sheet.getRange("G:G").format.columnWidth = 200;
    sheet.getRange("H:H").format.columnWidth = 180;
}

function writeQuickReport(sheet) {
    sheet.getRange("A1").values = [["QUICK AUDIT REPORT"]];
    sheet.getRange("A1").format.font.bold = true;
    sheet.getRange("A1").format.font.size = 16;

    sheet.getRange("A2").values = [["Issues found: " + auditResults.issues.length]];

    // Simple issues list
    var headers = [["Severity", "Type", "Sheet", "Cell", "Description"]];
    sheet.getRange("A4:E4").values = headers;
    sheet.getRange("A4:E4").format.font.bold = true;
    sheet.getRange("A4:E4").format.fill.color = "#217346";
    sheet.getRange("A4:E4").format.font.color = "#FFFFFF";

    auditResults.issues.forEach(function(issue, index) {
        var row = 5 + index;
        sheet.getRange("A" + row + ":E" + row).values = [[
            issue.severity, issue.type, issue.sheet, issue.cell, issue.description
        ]];
    });

    sheet.getRange("A:E").format.autofitColumns();
}

// ============================================
// UI DISPLAY
// ============================================
function displayAuditSummary() {
    var summaryDiv = document.getElementById('auditSummary');
    if (!summaryDiv) return;

    var critical = auditResults.issues.filter(function(i) { return i.severity === 'Critical'; }).length;
    var warning = auditResults.issues.filter(function(i) { return i.severity === 'Warning'; }).length;
    var info = auditResults.issues.filter(function(i) { return i.severity === 'Info'; }).length;

    var html = '<div class="audit-summary-cards">';
    html += '<div class="audit-card critical"><span class="count">' + critical + '</span><span class="label">Critical</span></div>';
    html += '<div class="audit-card warning"><span class="count">' + warning + '</span><span class="label">Warning</span></div>';
    html += '<div class="audit-card info"><span class="count">' + info + '</span><span class="label">Info</span></div>';
    html += '</div>';

    // Top issues list
    if (auditResults.issues.length > 0) {
        html += '<div class="audit-issues-list">';
        html += '<p class="kicker">Top Issues</p>';
        
        var topIssues = auditResults.issues.slice(0, 5);
        topIssues.forEach(function(issue) {
            var severityClass = issue.severity.toLowerCase();
            html += '<div class="audit-issue-item ' + severityClass + '">';
            html += '<span class="issue-type">' + issue.type + '</span>';
            html += '<span class="issue-location">' + issue.sheet + '!' + issue.cell + '</span>';
            html += '</div>';
        });

        if (auditResults.issues.length > 5) {
            html += '<p class="more-issues">... and ' + (auditResults.issues.length - 5) + ' more</p>';
        }
        html += '</div>';
    }

    summaryDiv.innerHTML = html;
    summaryDiv.style.display = 'block';
}
