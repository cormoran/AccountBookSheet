function GET_SHEET_NAME(pos: number) {
  return SpreadsheetApp.getActive().getSheets()[pos - 1].getName();
}

function LIST_ALL_IMPORTED_ITEMS() {
  const importedRanges = SpreadsheetApp.getActive()
    .getSheets()
    .filter((sheet) => sheet.getName().startsWith("Import_"))
    .map((sheet) => `${sheet.getName()}!A2:${Lib.LETTERS[ExtendedColumns.__LENGTH - 1]}`);
  return `=QUERY({${importedRanges.join(";")}},
            "select * where Col${MoneyForwardExportedCSVColumns.ID + 1} <> ''
              and Col${MoneyForwardExportedCSVColumns.CalcTargetFlag + 1} > 0
              and Col${MoneyForwardExportedCSVColumns.Amount + 1} < 0
              and Col${MoneyForwardExportedCSVColumns.TransferFlag} = 0
            ")`;
}
