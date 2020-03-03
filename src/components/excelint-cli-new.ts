// Process Excel files (input from .xls or .xlsx) with ExceLint.
// by Emery Berger, Microsoft Research / University of Massachusetts Amherst
// www.emeryberger.com

'use strict';
let fs = require('fs');
let path = require('path');
import { ExcelJSON } from './exceljson';
import { ExcelUtils } from './excelutils';
import { Colorize } from './colorize';
import { Timer } from './timer';
import { string } from 'prop-types';

type excelintVector = [number, number, number];

// Convert a rectangle into a list of indices.
function expand(first: excelintVector, second: excelintVector): Array<excelintVector> {
    const [fcol, frow] = first;
    const [scol, srow] = second;
    let expanded: Array<excelintVector> = [];
    for (let i = fcol; i <= scol; i++) {
        for (let j = frow; j <= srow; j++) {
            expanded.push([i, j, 0]);
        }
    }
    return expanded;
}

// Set to true to use the hard-coded example below.
const useExample = false;

const usageString = 'Usage: $0 <command> [options]';
const defaultFormattingDiscount = Colorize.getFormattingDiscount();
const defaultReportingThreshold = Colorize.getReportingThreshold();

// Process command-line arguments.
const args = require('yargs')
    .usage(usageString)
    .command('input', 'Input from FILENAME (.xls / .xlsx file).')
    .alias('i', 'input')
    .nargs('input', 1)
    .command('directory', 'Read from a directory of files (all ending in .xls / .xlsx).')
    .alias('d', 'directory')
    .command('formattingDiscount', 'Set discount for formatting differences (default = ' + defaultFormattingDiscount + ').')
    .command('reportingThreshold', 'Set the threshold % for reporting suspicious formulas (default = ' + defaultReportingThreshold + ').')
    .command('suppressOutput', 'Don\'t output the processed JSON to stdout.')
    .command('noElapsedTime', 'Suppress elapsed time output (for regression testing).')
    .command('sweep', 'Perform a parameter sweep and report the best settings overall.')
    .help('h')
    .alias('h', 'help')
    .argv;

if (args.help) {
    process.exit(0);
}

let allFiles = [];

if (args.directory) {
    // Load up all files to process.
    allFiles = fs.readdirSync(args.directory).filter((x: string) => x.endsWith('.xls') || x.endsWith('.xlsx'));
}
//console.log(JSON.stringify(allFiles));

// argument:
// input = filename. Default file is standard input.
let fname = '/dev/stdin';
if (args.input) {
    fname = args.input;
    allFiles = [fname];
}

// argument:
// formattingDiscount = amount of impact of formatting on fix reporting (0-100%).
let formattingDiscount = defaultFormattingDiscount;
if ('formattingDiscount' in args) {
    formattingDiscount = args.formattingDiscount;
}
// Ensure formatting discount is within range (0-100, inclusive).
if (formattingDiscount < 0) {
    formattingDiscount = 0;
}
if (formattingDiscount > 100) {
    formattingDiscount = 100;
}
Colorize.setFormattingDiscount(formattingDiscount);


// As above, but for reporting threshold.
let reportingThreshold = defaultReportingThreshold;
if ('reportingThreshold' in args) {
    reportingThreshold = args.reportingThreshold;
}
// Ensure formatting discount is within range (0-100, inclusive).
if (reportingThreshold < 0) {
    reportingThreshold = 0;
}
if (reportingThreshold > 100) {
    reportingThreshold = 100;
}
Colorize.setReportingThreshold(reportingThreshold);

//
// Ready to start processing.
//

let inp = null;

if (useExample) {
    // A simple example.
    inp = {
        workbookName: 'example',
        worksheets: [{
            sheetname: 'Sheet1',
            usedRangeAddress: 'Sheet1!E12:E21',
            formulas: [
                ['=D12'], ['=D13'],
                ['=D14'], ['=D15'],
                ['=D16'], ['=D17'],
                ['=D18'], ['=D19'],
                ['=D20'], ['=C21']
            ],
            values: [
                ['0'], ['0'],
                ['0'], ['0'],
                ['0'], ['0'],
                ['0'], ['0'],
                ['0'], ['0']
            ],
            styles: [
                [''], [''],
                [''], [''],
                [''], [''],
                [''], [''],
                [''], ['']
            ]
        }]
    };
}

let annotated_bugs = '{}';
try {
    annotated_bugs = fs.readFileSync('annotations-processed.json');
} catch (e) {
}

let bugs = JSON.parse(annotated_bugs);

let base = '';
if (args.directory) {
    base = args.directory + '/';
}

let parameters = [];
if (args.sweep) {
    const step = 10;
    for (let i = 0; i <= 100; i += step) {
        for (let j = 0; j <= 100; j += step) {
            parameters.push([i, j]);
        }
    }
} else {
    parameters = [[formattingDiscount, reportingThreshold]];
}

let f1scores = [];
let outputs = [];

for (let parms of parameters) {
    formattingDiscount = parms[0];
    Colorize.setFormattingDiscount(formattingDiscount);
    reportingThreshold = parms[1];
    Colorize.setReportingThreshold(reportingThreshold);

    let scores = [];

    for (let fname of allFiles) {
        // Read from file.
	console.warn('processing ' + fname);
        inp = ExcelJSON.processWorkbook(base, fname);

        let output = {
            'workbookName': path.basename(inp['workbookName']),
            'worksheets': {}
        };

        for (let i = 0; i < inp.worksheets.length; i++) {
            const sheet = inp.worksheets[i];

            // Skip empty sheets.
            if ((sheet.formulas.length === 0) && (sheet.values.length === 0)) {
                continue;
            }

            // Get rid of multiple exclamation points in the used range address,
            // as these interfere with later regexp parsing.
            let usedRangeAddress = sheet.usedRangeAddress;
            usedRangeAddress = usedRangeAddress.replace(/!(!+)/, '!');

            const myTimer = new Timer('excelint');

            // Get suspicious cells and proposed fixes, among others.
            let [suspicious_cells, grouped_formulas, grouped_data, proposed_fixes]
                = Colorize.process_suspicious(usedRangeAddress, sheet.formulas, sheet.values);

            // Adjust the fixes based on font stuff. We should allow parameterization here for weighting (as for thresholding).
            // NB: origin_col and origin_row currently hard-coded at 0,0.

            proposed_fixes = Colorize.adjust_proposed_fixes(proposed_fixes, sheet.styles, 0, 0);

            // Adjust the proposed fixes for real (just adjusting the scores downwards by the formatting discount).
            let adjusted_fixes = [];
            // tslint:disable-next-line: forin
            for (let ind = 0; ind < proposed_fixes.length; ind++) {
                const f = proposed_fixes[ind];
                const [score, first, second, sameFormat] = f;
                let adjusted_score = -score;
                if (!sameFormat) {
                    adjusted_score *= (100 - formattingDiscount) / 100;
                }
                if (adjusted_score * 100 >= reportingThreshold) {
                    adjusted_fixes.push([adjusted_score, first, second]);
                }
            }

	    let example_fixes_r1c1 = [];
	    {
		let totalNumericDiff = 0.0;
		if (adjusted_fixes.length > 0) {
		    for (let ind = 0; ind < adjusted_fixes.length; ind++) {
			let direction = "";
			if (adjusted_fixes[ind][1][0][0] === adjusted_fixes[ind][2][0][0]) {
			    direction = "vertical";
			} else {
			    direction = "horizontal";
			}
			let formulas = [];              // actual formulas
			let print_formulas = [];        // formulas with a preface (the cell name containing each)
			let r1c1_formulas = [];         // formulas in R1C1 format
			let r1c1_print_formulas = [];   // as above, but for R1C1 formulas
			let all_numbers = [];           // all the numeric constants in each formula
			let numbers = [];               // the sum of all the numeric constants in each formula
			let dependence_count = [];      // the number of dependent cells
			let absolute_refs = [];         // the number of absolute references in each formula
			for (let i = 0; i < 2; i++) {
			     // the coordinates of the cell containing the first formula in the proposed fix range
			    const formulaCoord = adjusted_fixes[ind][i+1][0];
			    const formulaX = formulaCoord[1]-1;                   // row
			    const formulaY = formulaCoord[0]-1;                   // column
			    const formula = sheet.formulas[formulaX][formulaY];   // the formula itself
			    const numeric_constants = ExcelUtils.numeric_constants(formula); // all numeric constants in the formula
			    all_numbers.push(numeric_constants);
			    numbers.push(numbers.reduce((a,b) => a + b, 0));      // the sum of all numeric constants
			    const dependences_wo_constants = ExcelUtils.all_cell_dependencies(formula, formulaY+1, formulaX+1, false);
			    dependence_count.push(dependences_wo_constants.length);
			    const r1c1 = ExcelUtils.formulaToR1C1(formula, formulaY+1, formulaX+1);
			    const preface = ExcelUtils.column_index_to_name(formulaY+1) + (formulaX+1) + ":";
			    const cellPlusFormula = preface + r1c1;
			    // Add the formulas plus their prefaces (the latter for printing).
			    r1c1_formulas.push(r1c1);
			    r1c1_print_formulas.push(cellPlusFormula);
			    formulas.push(formula);
			    print_formulas.push(preface + formula);
			    absolute_refs.push((formula.match(/\$/g) || []).length);
			    // console.log(preface + JSON.stringify(dependences_wo_constants));
			}
			totalNumericDiff = Math.abs(numbers[0] - numbers[1]);
			// Binning.
			let bin = [];
			if (dependence_count[0] !== dependence_count[1]) {
			    bin.push("different-dependent-count");
			}
			if (all_numbers[0].length !== all_numbers[1].length) {
			    bin.push("number-of-constants-mismatch");
			}
			if (r1c1_formulas[0].localeCompare(r1c1_formulas[1])) {
			    // reference mismatch, but can have false positives with constants.
			    bin.push("r1c1-mismatch");
			}
			if (absolute_refs[0] !== absolute_refs[1]) {
			    bin.push("absolute-ref-mismatch");
			}
			if (bin === []) {
			    bin.push("unclassified");
			}
			example_fixes_r1c1.push({ "bin" : bin,
						  "direction" : direction,
						  "numbers" : numbers,
						  "numeric_difference": totalNumericDiff,
						  "magnitude_numeric_difference": (totalNumericDiff === 0) ? 0 : Math.log10(totalNumericDiff),
						  "formulas": print_formulas,
						  "r1c1formulas" : r1c1_print_formulas });
			// example_fixes_r1c1.push([direction, formulas]);
		    }
		}
	    }

            let elapsed = myTimer.elapsedTime();
            if (args.noElapsedTime) {
                elapsed = 0; // Dummy value, used for regression testing.
            }
            // Compute number of cells containing formulas.
            const numFormulaCells = (sheet.formulas.flat().filter(x => x.length > 0)).length;

            // Count the number of non-empty cells.
            const numValueCells = (sheet.values.flat().filter(x => x.length > 0)).length;

            // Compute total number of cells in the sheet (rows * columns).
            const columns = sheet.values[0].length;
            const rows = sheet.values.length;
            const totalCells = rows * columns;

            const out = {
                'suspiciousnessThreshold': reportingThreshold,
                'formattingDiscount': formattingDiscount,
                'proposedFixes': adjusted_fixes,
		'exampleFixes' : example_fixes_r1c1,
//		'exampleFixesR1C1' : example_fixes_r1c1,
                'suspiciousRanges': adjusted_fixes.length,
		'weightedSuspiciousRanges' : 0, // actually calculated below.
                'suspiciousCells': 0, // actually calculated below.
                'elapsedTimeSeconds': elapsed / 1e6,
                'columns': columns,
                'rows': rows,
                'totalCells': totalCells,
                'numFormulaCells': numFormulaCells,
                'numValueCells': numValueCells
            };

            // Compute precision and recall of proposed fixes, if we have annotated ground truth.
            const workbookBasename = path.basename(inp['workbookName']);
            // Build list of bugs.
            let foundBugs: any = out['proposedFixes'].map(x => {
                if (x[0] >= (reportingThreshold / 100)) {
                    return expand(x[1][0], x[1][1]).concat(expand(x[2][0], x[2][1]));
                } else {
                    return [];
                }
            });
            const foundBugsArray: any = Array.from(new Set(foundBugs.flat(1).map(JSON.stringify)));
            foundBugs = foundBugsArray.map(JSON.parse);
            out['suspiciousCells'] = foundBugs.length;
	    let weightedSuspiciousRanges = out['proposedFixes'].map(x => x[0]).reduce((x, y) => x + y, 0);
	    out['weightedSuspiciousRanges'] = weightedSuspiciousRanges;
            if (workbookBasename in bugs) {
                if (sheet.sheetName in bugs[workbookBasename]) {
                    const trueBugs = bugs[workbookBasename][sheet.sheetName]['bugs'];
                    const totalTrueBugs = trueBugs.length;
                    const trueBugsJSON = trueBugs.map(x => JSON.stringify(x));
                    const foundBugsJSON = foundBugs.map(x => JSON.stringify(x));
                    const truePositives = trueBugsJSON.filter(value => foundBugsJSON.includes(value)).map(x => JSON.parse(x));
                    const falsePositives = foundBugsJSON.filter(value => !trueBugsJSON.includes(value)).map(x => JSON.parse(x));
                    const falseNegatives = trueBugsJSON.filter(value => !foundBugsJSON.includes(value)).map(x => JSON.parse(x));
                    let precision = 0;
                    let recall = 0;
                    out['truePositives'] = truePositives.length;
                    out['falsePositives'] = falsePositives.length;
                    out['falseNegatives'] = falseNegatives.length;

                    // We adopt the methodology used by the ExceLint paper (OOPSLA 18):
                    //   "When a tool flags nothing, we define precision to
                    //    be 1, since the tool makes no mistakes. When a benchmark contains no errors but the tool flags
                    //    anything, we define precision to be 0 since nothing that it flags can be a real error."

                    if (foundBugs.length === 0) {
                        out['precision'] = 1;
                    }
                    if ((truePositives.length === 0) && (foundBugs.length > 0)) {
                        out['precision'] = 0;
                    }
                    if ((truePositives.length > 0) && (foundBugs.length > 0)) {
                        precision = truePositives.length / (truePositives.length + falsePositives.length);
                        out['precision'] = precision;
                    }
                    if (falseNegatives.length + trueBugs.length > 0) {
                        recall = truePositives.length / (falseNegatives.length + truePositives.length);
                        out['recall'] = recall;
                    } else {
                        // No bugs to find means perfect recall. NOTE: this is not described in the paper.
                        out['recall'] = 1;
                    }
                    scores.push(truePositives.length - falsePositives.length);
                    if (false) {
                        if (precision + recall > 0) {
                            // F1 score: https://en.wikipedia.org/wiki/F1_score
                            const f1score = (2 * precision * recall) / (precision + recall);
                            /// const f1score = precision; //// FIXME for testing (2 * precision * recall) / (precision + recall);
                            scores.push(f1score);
                        }
                    }
                }
            }
            output.worksheets[sheet.sheetName] = out;
        }
        outputs.push(output);
    }
    let averageScores = 0;
    let sumScores = 0;
    if (scores.length > 0) {
        averageScores = scores.reduce((a, b) => a + b, 0) / scores.length;
        sumScores = scores.reduce((a, b) => a + b, 0);
    }
    f1scores.push([formattingDiscount, reportingThreshold, sumScores]);
}
f1scores.sort((a, b) => { if (a[2] < b[2]) { return -1; } if (a[2] > b[2]) { return 1; } return 0; });
// Now find the lowest threshold with the highest F1 score.
const maxScore = f1scores.reduce((a, b) => { if (a[2] > b[2]) { return a[2]; } else { return b[2]; } });
//console.log('maxScore = ' + maxScore);
// Find the first one with the max.
const firstMax = f1scores.find(item => { return item[2] === maxScore; });
//console.log('first max = ' + firstMax);
if (!args.suppressOutput) {
    console.log(JSON.stringify(outputs, null, '\t'));
}
// console.log(JSON.stringify(f1scores));

