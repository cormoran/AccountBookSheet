module Lib {
  export const LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  export function spreadSheetAutoResizeColumns(sheet: GoogleAppsScript.Spreadsheet.Sheet, minWidth: number = 0, maxWidth: number = 0) {
    const range = sheet.getDataRange();
    const numColumns = range.getNumColumns();
    for (let i = 0; i < numColumns; i++) {
      sheet.autoResizeColumn(i + 1);
      const width = sheet.getColumnWidth(i + 1);
      if (minWidth > 0 && width < minWidth) {
        sheet.setColumnWidth(i + 1, minWidth);
      } else if (maxWidth > 0 && width > maxWidth) {
        sheet.setColumnWidth(i + 1, maxWidth);
      }
    }
  }
  export function spreadSheetSortSheets(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet) {
    const sheets = spreadsheet.getSheets();
    sheets.sort((a, b) => a.getName().localeCompare(b.getName()));
    const isSorted = sheets.map((sheet, i) => sheet.getIndex() == i + 1).reduce((a, b) => a && b, true);
    if (!isSorted) {
      console.log("Sorting sheets");
      sheets.forEach((sheet, i) => {
        sheet.activate();
        spreadsheet.moveActiveSheet(i + 1);
      });
    } else {
      console.log("Skip sorting sheets. Already sorted.");
    }
  }

  export function spreadSheetGetOrCreateSheet(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet, sheetName: string) {
    return spreadsheet.getSheetByName(sheetName) ?? spreadsheet.insertSheet(sheetName);
  }
}
