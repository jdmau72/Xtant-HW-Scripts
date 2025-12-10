// values used by the whole script
const MAX_EXCEL_SHEET_SIZE = 1048576;
let appliedRange: ExcelScript.Range;
let selectedSheet: ExcelScript.Worksheet;

let lotAdjCol: string;
let notesCol: string;
let expDateCol: string;
let itemNumCol: string;


/**
@param {boolean} ExpirationDate Do you want to check for expiring items?
*/
function main(workbook: ExcelScript.Workbook, 
						hasExpirationDate: true | false = false) {


	//selectedSheet = workbook.getActiveWorksheet();
	
	// first copy this worksheet to a new sheet, so it can be rolled back if need be
    let selectedSheet = workbook.getActiveWorksheet().copy();
    selectedSheet.activate();
	
	// set the indices of each column
	let headerRange = selectedSheet.getRange("A1:Z1");
	let adjColIndex = findColumnIndex("Status", headerRange);
	let lotColIndex = findColumnIndex("Lot", headerRange);
	let itemColIndex = findColumnIndex("Item Number", headerRange);
	let qtyColIndex = findColumnIndex("QTY", headerRange);
	let binColIndex = findColumnIndex("Bin Number", headerRange);
	let locColIndex = findColumnIndex("Location", headerRange);
	let expDateColIndex = findColumnIndex("EXP Date", headerRange);
				
	// set column values that are used later
	let lotAdjCol = indexToLetter(adjColIndex);
	// lotAdjCol = "E";

	let expDateCol = indexToLetter(expDateColIndex);
	// expDateCol = "G";

	let notesCol = indexToLetter(expDateColIndex + 1); // want to put the notes column to the right of EXP date
	// notesCol = "H";

	let itemNumCol = indexToLetter(itemColIndex);
	// itemNumCol = "B";

	// remove the Location column first
	//selectedSheet.getRange("A:A").delete(ExcelScript.DeleteShiftDirection.left);

	// clear Status column and replace with LOT ADJ  
	selectedSheet.getRange(`${lotAdjCol}:${lotAdjCol}`).clear(ExcelScript.ClearApplyTo.contents);
	selectedSheet.getRange(`${lotAdjCol}1`).setValue("LOT ADJ");

	// create NOTES column
	selectedSheet.getRange(`${notesCol}1`).copyFrom(selectedSheet.getRange(`${expDateCol}1`), ExcelScript.RangeCopyType.formats, false, false);
	selectedSheet.getRange(`${notesCol}1`).setValue("NOTES");

	// if EXP date is false, hide the EXP date column
	if (!hasExpirationDate) { 
		selectedSheet.getRange(`${expDateCol}:${expDateCol}`).setColumnHidden(true); 
	}

	
	
	// set text size, spacing, make it neat 
	selectedSheet.getRange().getFormat().getFont().setSize(14);
	selectedSheet.getRange().getFormat().autofitColumns();
 
	// code from previous AutoFormat code -------------------------------------------- \/ \/ \/
	// gets the range of columns actually used for nicer formatting
	let usedRange = selectedSheet.getUsedRange();
	let usedRangeAddress = usedRange.getAddress().split("!")[1].split(":");
	let leftColumn = usedRangeAddress[0][0];
	let rightColumn = usedRangeAddress[1][0];
	// console.log(usedRangeAddress);
	appliedRange = selectedSheet.getRange(
		leftColumn + "1:" + rightColumn + MAX_EXCEL_SHEET_SIZE.toString()
	);

	// now sets the formats on this range
	defineConditionalFormat_Adj("#FFFF00", false);	// adj first rule, overwritten by others
	if (hasExpirationDate) { defineConditionalFormat_ExpDate("#FF6565"); }
	defineConditionalFormat("fail", "#FF0000");
	defineConditionalFormat("EXTRA", "#92D050");
	defineConditionalFormat("QAHOLD", "FF6565");
	defineConditionalFormat("MIA", "#BFBFBF");
	defineConditionalFormat_Adj("#FFFF00", true); // final adj lets it apply to just the LOTADJ column


	defineConditionalFormat_ItemBorders_Bottom();
	// overwrite applied range to ignore first row, otherwise border gets drawn below where it should
	appliedRange = selectedSheet.getRange(
		leftColumn + "2:" + rightColumn + MAX_EXCEL_SHEET_SIZE.toString()
	);
	defineConditionalFormat_ItemBorders_Top();
	
}




function defineConditionalFormat(keyword: string, color: string){
	// Create custom from all cells on selectedSheet
	let conditionalFormatting: ExcelScript.ConditionalFormat;
	let condition: string;

	condition = '=ISNUMBER(SEARCH(\"' + keyword + '\",$' + notesCol + '1))';

	conditionalFormatting = appliedRange.addConditionalFormat(ExcelScript.ConditionalFormatType.custom);
	conditionalFormatting.getCustom().getRule().setFormula(condition);
	conditionalFormatting.getCustom().getFormat().getFill().setColor(color);
	conditionalFormatting.setStopIfTrue(false);
	conditionalFormatting.setPriority(0);
}




function defineConditionalFormat_Adj(color: string, colOnly: boolean = false) {
	// Create custom from all cells on selectedSheet
	let conditionalFormatting: ExcelScript.ConditionalFormat;
	let condition: string;

	condition = '=AND(ISNUMBER(SEARCH("LOT ADJ",$' + lotAdjCol + '1))=FALSE, OR(ISTEXT($' + lotAdjCol + '1),ISNUMBER($' + lotAdjCol + '1)))';

	// if its just for the adjustment col, uses only the adjustment column
	if (colOnly == true) {
		let adjColRange = selectedSheet.getRange(lotAdjCol + "1:" + lotAdjCol + MAX_EXCEL_SHEET_SIZE.toString())
		conditionalFormatting = adjColRange.addConditionalFormat(ExcelScript.ConditionalFormatType.custom);
	
	// otherwise uses the whole range of the table
	} else {
		conditionalFormatting = appliedRange.addConditionalFormat(ExcelScript.ConditionalFormatType.custom);
	}

	conditionalFormatting.getCustom().getRule().setFormula(condition);
	conditionalFormatting.getCustom().getFormat().getFill().setColor(color);
	conditionalFormatting.setStopIfTrue(false);
	conditionalFormatting.setPriority(0);
}




function defineConditionalFormat_ExpDate(color: string){
	let conditionalFormatting: ExcelScript.ConditionalFormat;
	let condition: string;

	// checks if expired or if today is within 1 month of the expiration date
	condition = '=AND( ISNUMBER(DATEVALUE($' + expDateCol + '1)), (TODAY() - ($' + expDateCol + '1 - EDATE(MONTH($' + 									expDateCol + '1), 1))) > DATE(1900, 1, 0))'; 

	conditionalFormatting = appliedRange.addConditionalFormat(ExcelScript.ConditionalFormatType.custom);
	conditionalFormatting.getCustom().getRule().setFormula(condition);
	conditionalFormatting.getCustom().getFormat().getFill().setColor(color);
	conditionalFormatting.setStopIfTrue(false);
	conditionalFormatting.setPriority(0);

}



function defineConditionalFormat_ItemBorders_Bottom() {
	let conditionalFormatting: ExcelScript.ConditionalFormat;
	let condition: string;

	// checks if item number BELOW it is same, if not, creates border
	condition = `=\$${itemNumCol}1 <> \$${itemNumCol}2`;

	conditionalFormatting = appliedRange.addConditionalFormat(ExcelScript.ConditionalFormatType.custom);
	conditionalFormatting.getCustom().getRule().setFormula(condition);
	conditionalFormatting.getCustom().getFormat().getConditionalRangeBorderBottom().setStyle(ExcelScript.ConditionalRangeBorderLineStyle.continuous);
	conditionalFormatting.setStopIfTrue(false);
	conditionalFormatting.setPriority(0);
}


function defineConditionalFormat_ItemBorders_Top(){
	let conditionalFormatting: ExcelScript.ConditionalFormat;
	let condition: string;

	// checks if item number ABOVE it is same, if not, creates border
	condition = `=\$${itemNumCol}2 <> \$${itemNumCol}1`;

	conditionalFormatting = appliedRange.addConditionalFormat(ExcelScript.ConditionalFormatType.custom);
	conditionalFormatting.getCustom().getRule().setFormula(condition);
	conditionalFormatting.getCustom().getFormat().getConditionalRangeBorderTop().setStyle(ExcelScript.ConditionalRangeBorderLineStyle.continuous);
	conditionalFormatting.setStopIfTrue(false);
	conditionalFormatting.setPriority(0);
}


// HELPER FUNCTIONS ---------------------------------------------------------------------------------------- \\
function findColumnIndex(searchTerm: string, headerRange: ExcelScript.Range = selectedSheet.getRange("A1:Z1")){
    // gets the values of the headers
    let headers = headerRange.getValues()[0];
    let targetCol = 0;

    for (let col = 0; col < headers.length; col++){
        if (headers[col].toString().toLowerCase() == searchTerm.toLowerCase()) {
          targetCol = col;
        }
    }
    return targetCol;
}

function indexToLetter(i: number) {
	let letters = ' ABCDEFGHIJKLMNOPQRSTUVWXYZ';
	return letters[i];
}

function findColumnLetter(searchTerm: string, headerRange: ExcelScript.Range = selectedSheet.getRange("A1:Z1")){
	let colIndex = findColumnIndex(searchTerm, headerRange);
	return (indexToLetter(colIndex));
}