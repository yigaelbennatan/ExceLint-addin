/* global Office, Excel */

import * as React from 'react';
import { Header } from './Header';
import { Content } from './Content';
import Progress from './Progress';
import { Colorize } from './colorize';

import * as OfficeHelpers from '@microsoft/office-js-helpers';

export interface AppProps {
    title: string;
    isOfficeInitialized: boolean;
}

export interface AppState {
}

export default class App extends React.Component<AppProps, AppState> {

    private savedProperties : any = [];
    private savedColors : any = [];
    
    constructor(props, context) {
        super(props, context);
    }

    
    private async processRange(context : any, currentWorksheet : any, startCol : number, startRow : number, endCol : number, endRow : number) {
 	if ((endCol - startCol >= 0) && (endRow - startRow >= 0)) {
	    let startCell = Colorize.column_index_to_name(startCol) + startRow;
	    let endCell = Colorize.column_index_to_name(endCol) + endRow;
	    let range = [] as any;
	    if (startCell === endCell) {
		range = currentWorksheet.getCell(startRow, startCol);
	    } else {
		/* // iterative version
		for (let c = startCol; c <= endCol; c++) {
		    for (let r = startRow; r <= endRow; r++) {
			range = currentWorksheet.getCell(c, r);
			range.format.fill.load(['color']);
			console.log("synching for "+c+ " " +r);
			await context.sync();
			let color = range.format.fill.color;
			let startCell = Colorize.column_index_to_name(c) + r;
			this.savedColors.push([startCell, startCell, color]);
		    }
		}
		*/
		range = currentWorksheet.getRange(startCell + ":" + endCell);
	    }
//	    return;
	    // 	    await context.sync();
	    range.format.fill.load(['color']);
	    console.log("synching for "+startCell+" through " + endCell);
	    await context.sync();
	    let color = range.format.fill.color;
	    if (color !== null) {
		this.savedColors.push([startCell, endCell, color]);
	    } else {
		/* // column by column
		if (startCol !== endCol) {
		    let midCol = startCol + Math.floor((endCol - startCol) / 2);		    
		    await this.processRange(context, currentWorksheet, startCol, startRow, midCol, endRow);
		    await this.processRange(context, currentWorksheet, midCol + 1, startRow, endCol, endRow);
		} else {
		    let midRow  = startRow + Math.floor((endRow - startRow) / 2);
		    await this.processRange(context, currentWorksheet, startCol, startRow, endCol, midRow);
		    await this.processRange(context, currentWorksheet, startCol, midRow + 1, endCol, endRow);
		}
		*/
		// binary search
		let divisor = 2; // (Math.random() * (8 - 2)) + 2.0001;
 		let midCol = startCol + Math.floor((endCol - startCol) / divisor);
		let midRow  = startRow + Math.floor((endRow - startRow) / divisor);
		await this.processRange(context, currentWorksheet, startCol, startRow, midCol, midRow);
		await this.processRange(context, currentWorksheet, midCol + 1, startRow, endCol, midRow);
		await this.processRange(context, currentWorksheet, startCol, midRow + 1, midCol, endRow);
		await this.processRange(context, currentWorksheet, midCol + 1, midRow + 1, endCol, endRow);
	    }
	}
    }
    
    attemptToSaveColor = async () => {
	try {
	    await Excel.run(async context => {
		let startTime = performance.now();
		this.savedColors = [];
		// Load up the used range.
	    	let currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
		let usedRange = currentWorksheet.getUsedRange();
		await context.sync();
		usedRange.load(['address']);
 		usedRange.format.fill.load(['color']);
		//		let items = usedRange.format.fill;
		await context.sync();
		let color = usedRange.format.fill.color;
		let address = usedRange.address;
		let [sheetName, startCell, endCell] = Colorize.extract_sheet_range(address);
		let [startCol, startRow] = Colorize.cell_dependency(startCell, 0, 0);
		let [endCol, endRow] = Colorize.cell_dependency(endCell, 0, 0);
		// Are we done? (We got a color)
		if (color !== null) {
		    this.savedColors.push([startCell, endCell, color]);
		} else {
		    await this.processRange(context, currentWorksheet, startCol, startRow, endCol, endRow);
		}
		console.log(this.savedColors);
		let endTime = performance.now();
		let timeElapsedMS = endTime - startTime;
 		console.log("Time elapsed (ms) = " + timeElapsedMS);
	    });
	} catch (error) {
            OfficeHelpers.UI.notify(error);
            OfficeHelpers.Utilities.log(error);
        }
	
    }
    
    private process(f, currentWorksheet) {
	// Sort by COLUMNS (first dimension).
	let identified_ranges = Colorize.identify_ranges(f, (a, b) => { if (a[0] == b[0]) { return a[1] - b[1]; } else { return a[0] - b[0]; }});

	// Now group them (by COLUMNS).
	let grouped_ranges = Colorize.group_ranges(identified_ranges);
	// console.log(grouped_ranges);
	// FINALLY, process the ranges.
	Object.keys(grouped_ranges).forEach(color => {
	    if (!(color === undefined)) {
 		let v = grouped_ranges[color];
		for (let theRange of v) {
		    let r = theRange;
		    let col0 = Colorize.column_index_to_name(r[0][0]);
		    let row0 = r[0][1];
		    let col1 = Colorize.column_index_to_name(r[1][0]);
		    let row1 = r[1][1];
		    
		    let range = currentWorksheet.getRange(col0 + row0 + ":" + col1 + row1);
		    range.format.fill.color = color;
		}
	    }
	})
    }
    
    clearColor = async () => {
        try {
            await Excel.run(async context => {

		// Clear all formats and borders.
		
	    	let currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
		let everythingRange = currentWorksheet.getRange();
		await context.sync();
		everythingRange.clear(Excel.ClearApplyTo.formats);
 		everythingRange.format.borders.load(['items']);
		await context.sync();
		let items = everythingRange.format.borders.items;
		
		for (let border of items) {
		    border.set ({ "style" : "None",
				  "tintAndShade" : 0 });
		}
		
	    });
	} catch (error) {
            OfficeHelpers.UI.notify(error);
            OfficeHelpers.Utilities.log(error);
        }
    }
    
    setColor = async () => {
        try {
            await Excel.run(async context => {
		let app = context.workbook.application;
		console.log("ExceLint: starting processing.");
		let startTime = performance.now();
		
	    	let currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
		let usedRange = currentWorksheet.getUsedRange();
		let everythingRange = currentWorksheet.getRange();
		// Now get the addresses, the formulas, and the values.
                usedRange.load(['address', 'formulas', 'values']);
		currentWorksheet.charts.load(['items']);
		
		await context.sync();
		console.log("ExceLint: done with sync 1.");

		let address = usedRange.address;
		
		// Now we can get the formula ranges (all cells with formulas),
		// and the numeric ranges (all cells with numbers). These come in as 2-D arrays.
		let formulaRanges = usedRange.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);
		let numericRanges = usedRange.getSpecialCellsOrNullObject(Excel.SpecialCellType.constants,
									  Excel.SpecialCellValueType.numbers);
		let formulas = usedRange.formulas;
		let values = usedRange.values;
 		numericRanges.format.borders.load(['items']);
		formulaRanges.format.borders.load(['items']);
		

		await context.sync();
		console.log("ExceLint: done with sync 2.");

		
		// FIX ME - need a button to restore all formatting.
		// First, clear all formatting. Really we want to just clear colors but fine for now (FIXME later)
		everythingRange.clear(Excel.ClearApplyTo.formats);
		
		// Make all numbers yellow; this will be the default value for unreferenced data.
		numericRanges.format.fill.color = "yellow";

		// Give every numeric data item a dashed border.
		let items = numericRanges.format.borders.items;
		for (let border of items) {
		    border.set ({ "weight" : "Thin",
				  "style" : "Dash",
				  "tintAndShade" : -1 });
		}

		// Give every formula a solid border.
		items = formulaRanges.format.borders.items;
		for (let border of items) {
		    border.set ({ "weight" : "Thin",
				  "style" : "Continuous",
				  "tintAndShade" : -1 });
		}

		let [sheetName, startCell] = Colorize.extract_sheet_cell(address);
		let vec = Colorize.cell_dependency(startCell, 0, 0);
		
 		let processed_formulas = Colorize.process_formulas(formulas, vec[0]-1, vec[1]-1);
		let processed_data = Colorize.color_all_data(formulas, processed_formulas, vec[0], vec[1]);
		
		this.process(processed_data, currentWorksheet);
		this.process(processed_formulas, currentWorksheet);
		
		
		await context.sync();
		console.log("ExceLint: done with sync 3.");
		
		let endTime = performance.now();
		let timeElapsedMS = endTime - startTime;
 		console.log("Time elapsed (ms) = " + timeElapsedMS);
            });
        } catch (error) {
            OfficeHelpers.UI.notify(error);
            OfficeHelpers.Utilities.log(error);
        }
    }

    render() {
        const {
            title,
            isOfficeInitialized,
        } = this.props;

        if (!isOfficeInitialized) {
            return (
                    <Progress
                title={title}
                logo='assets/logo-filled.png'
                message='Please sideload your addin to see app body.'
                    />
            );
        }

        return (
		<div className='ms-welcome'>
                <Header title='ExceLint' />
                <Content message1='Click the button below to reveal the deep structure of this spreadsheet.' buttonLabel1='Reveal structure' click1={this.setColor} message2='Click the button below to clear colors and borders.' buttonLabel2='Clear' click2={this.clearColor} />
		</div>
        );
    }
}