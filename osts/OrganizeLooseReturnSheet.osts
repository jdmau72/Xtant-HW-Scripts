
const ADJ_COL = "E";
const F_COL = "F";
const T_COL = "H";

// formula for getting associated value in other table:
// =IF(ISBLANK(Main!I21)=FALSE, Main!I21, "")


function main(workbook: ExcelScript.Workbook){
  // Get the active cell and worksheet.
  let selectedCell = workbook.getActiveCell();
  let sheet = workbook.getFirstWorksheet();

  
  let adjSheet = sheet.copy(ExcelScript.WorksheetPositionType.after, sheet);
  adjSheet.setName("ADJ IN");
  adjSheet.getUsedRange().clearAllConditionalFormats(); // need to clear, otherwise things break
  organize_ADJ(adjSheet);



  // trying to get every combination of values for transfer to and from
  let fromColumn = sheet.getRange(`${F_COL}3:${F_COL}99`).getValues();
  let toColumn = sheet.getRange(`${T_COL}3:${T_COL}99`).getValues();
  
  // iterates through the to and from column, concats, then checks if unique
  let uniqueTransferNames: string[] = []
  for (let i = 0; i < fromColumn.length; i++){
      if (fromColumn[i][0].toString() !== "" && toColumn[i][0].toString() !== ""){
        let transferName: string = fromColumn[i][0].toString() + "_" + toColumn[i][0].toString();
      
        if ((uniqueTransferNames.includes(transferName)) == false){
          uniqueTransferNames.push(transferName);
        }
    }
  }
  console.log(uniqueTransferNames);

  // for each unique combination of transfers (to & from),
  // creates a new sheet
  // then organizes that sheet
  for (let i = 0; i < uniqueTransferNames.length; i++){
    let tn = uniqueTransferNames[i].split("_");
    let tFrom = tn[0];
    let tTo = tn[1];
    // console.log(tFrom + "   " + tTo);

  let transferSheet = sheet.copy(ExcelScript.WorksheetPositionType.after, sheet);
  transferSheet.setName(tFrom + " -> " + tTo);
  //transferSheet.getUsedRange().clearAllConditionalFormats(); // need to clear, otherwise things break
  organize_Transfer(transferSheet, tFrom, tTo);

  }
}

function organize_Transfer(sheet: ExcelScript.Worksheet, tFrom: string, tTo: string) {

  let toColumn = sheet.getRange(`${T_COL}3:${T_COL}99`).getValues();
  let fromColumn = sheet.getRange(`${F_COL}3:${F_COL}99`).getValues();
  for (let row = 0; row < toColumn.length; row++) {
    let toValue = toColumn[row][0].toString();
    let fromValue = fromColumn[row][0].toString();


    // gets entire row at that index
    let currentItemRangeQuery: string = `A${row + 3}:L${row + 3} `;
    let currentItemRange = sheet.getRange(currentItemRangeQuery)

    // checks if the transfer From and transfer to fields are matching the parameters
    if (!(toValue === tTo && fromValue === tFrom)) {
      currentItemRange.clear();
      currentItemRange.setRowHidden(true);
    }
  }
}


function organize_ADJ(sheet: ExcelScript.Worksheet){
  
  let checkedColumn = sheet.getRange(`${ADJ_COL}3:${ADJ_COL}99`).getValues();
  for(let row = 0; row < checkedColumn.length; row++){
    let cellValue = checkedColumn[row][0].toString();
    //console.log(cellValue);
    
    // gets entire row at that index
    let currentItemRangeQuery: string = `A${row+3}:L${row+3} `;
    let currentItemRange = sheet.getRange(currentItemRangeQuery)

    // if ADJ IN column empty, then will clear and hide the row
    if (!(cellValue.toLowerCase() === "y" || cellValue.toLowerCase() === "yes")){
      currentItemRange.clear();
      currentItemRange.setRowHidden(true);
    }
  }  
}




  
    




