const GoogleDriveFolderId = Secret.GoogleDriveFolderId;

// ---

type FileImportingState = "finished" | "processing" | "error";
type ImportedFileInfo = { file: string; date: GoogleAppsScript.Base.Date | Date; row: number; state: FileImportingState };

enum MoneyForwardExportedCSVColumns {
  CalcTargetFlag,
  Date,
  Content,
  Amount,
  Source,
  Category,
  SubCategory,
  Memo,
  TransferFlag,
  ID,
  __LENGTH,
}
enum ExtendedColumns {
  NumPseudoSplit = MoneyForwardExportedCSVColumns.__LENGTH,
  IsSharedPay,
  MyPayRatio,
  __LENGTH,
}
const LETTERS = Lib.LETTERS;

/**
 * ImportState
 * | file name | last updated (ISOString) | Status             |
 * | ...       | ...                      | FileImportingState |
 */
const ImportStateSheetName = "0_ImportState";
/**
 * Category
 * | Category | SubCategory |
 * | ...      | ...         |
 */
const CategorySheetName = "0_Category";
enum CategorySheetColumns {
  Category,
  SubCategory,
  CategoryPlusSubCategory,
  __LENGTH,
}

function mainImportFromGoogleDrive() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetImportState = Lib.spreadSheetGetOrCreateSheet(spreadsheet, ImportStateSheetName);

  /**
   * Import specified CSV file on google drive to spreadsheet.
   * The sheet name is like "Import_2021_01". The sheet is created if not exists.
   * @param csvFile CSV file exported from MoneyForward.
   *                The file name should be like "収入・支出詳細_2021-11-01_2021-11-30.csv".
   *                The file should contain columns defined in MoneyForwardExportedCSVColumns with header.
   */
  function importFile(csvFile: GoogleAppsScript.Drive.File) {
    const match = /^[^_]+_(?<start>[0-9-]+)_(?<end>[0-9-]+)\.csv$/.exec(csvFile.getName());
    const [startYear, startMonth, startDate] = match?.groups?.start?.split("-")?.map((i) => parseInt(i)) ?? [];
    const [endYear, endMonth, endDate] = match?.groups?.end?.split("-")?.map((i) => parseInt(i)) ?? [];
    if (startYear == undefined || startMonth == undefined || startYear != endYear || startMonth != endMonth) {
      throw new Error(`Invalid file name format: ${csvFile.getName()}`);
    }
    const sheetName = `Import_${startYear}_${("0" + startMonth).slice(-2)}`;
    const sheet = spreadsheet.getSheetByName(sheetName) ?? spreadsheet.insertSheet(sheetName);

    //
    // parse existing rows
    //
    const existingData = sheet.getDataRange().getValues();
    const idToRow: { [id: string]: number } = {};
    existingData.forEach((row, i) => {
      // header is also treated as row
      idToRow[row[MoneyForwardExportedCSVColumns.ID]] = i + 1;
    });
    //
    // parse CSV file and upsert rows to sheet
    //
    const records = Utilities.parseCsv(csvFile.getBlob().getDataAsString("shift_jis"));
    records.forEach((row, i) => {
      if (row.length != MoneyForwardExportedCSVColumns.__LENGTH) {
        throw new Error(`Invalid CSV format: ${csvFile.getName()} row ${i} has ${row.length} columns\n${row}`);
      }
      // header is also treated as row
      const id = row[MoneyForwardExportedCSVColumns.ID];
      if (idToRow[id] == undefined) {
        sheet.appendRow(row);
        idToRow[id] = sheet.getLastRow();
      }
    });
    updateImportedSheetSetting(sheet);
    upsertCategoryListSheet(sheet);
  }

  /**
   * Configure sheet after importing CSV file.
   * - Set protection to all imported cells
   * - Add filter
   * @param sheet target sheet
   */
  function updateImportedSheetSetting(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
    const additionalColumns = ["擬似分割払い回数", "共同出費", "自己負担率"];
    sheet.getRange(1, MoneyForwardExportedCSVColumns.__LENGTH + 1, 1, additionalColumns.length).setValues([additionalColumns]);
    const importedRange = sheet.getRange(`A:${LETTERS[MoneyForwardExportedCSVColumns.__LENGTH - 1]}`);
    importedRange.protect().setWarningOnly(true);
    sheet.getFilter()?.remove();
    const dataRange = sheet.getRange(`A:${LETTERS[sheet.getDataRange().getNumColumns() - 1]}`);
    dataRange.createFilter();
    Lib.spreadSheetAutoResizeColumns(sheet, 50, 300);
  }

  function upsertCategoryListSheet(importedSheet: GoogleAppsScript.Spreadsheet.Sheet) {
    // assumption
    if (MoneyForwardExportedCSVColumns.Category + 1 != MoneyForwardExportedCSVColumns.SubCategory) throw new Error("Unexpected column order");
    if (CategorySheetColumns.Category + 1 != CategorySheetColumns.SubCategory) throw new Error("Unexpected column order");
    const sheetCategory = spreadsheet.getSheetByName(CategorySheetName) ?? spreadsheet.insertSheet(CategorySheetName);
    sheetCategory.getRange(1, 1, 1, CategorySheetColumns.__LENGTH).setValues([["Category", "SubCategory", "Category_SubCategory"]]);
    const concat = (row) => row.join("_");
    const existingSet = new Set<string>( // "Category_SubCategory")
      sheetCategory.getLastRow() > 1
        ? sheetCategory
            .getRange(2, CategorySheetColumns.Category + 1, sheetCategory.getLastRow() - 1, 2)
            .getValues()
            .map(concat)
        : []
    );
    const newPairs = importedSheet.getLastRow() > 1 ? importedSheet.getRange(2, MoneyForwardExportedCSVColumns.Category + 1, importedSheet.getLastRow() - 1, 2).getValues() : [];
    newPairs.forEach((categoryPair) => {
      const concatenated = concat(categoryPair);
      if (!existingSet.has(concatenated)) {
        sheetCategory.appendRow(categoryPair.concat([concatenated]));
        existingSet.add(concatenated);
      }
    });
    sheetCategory
      .getDataRange()
      .offset(1, 0, sheetCategory.getLastRow() - 1)
      .sort(CategorySheetColumns.CategoryPlusSubCategory + 1);
  }

  function loadImportStates() {
    const fileNameToLastUpdated: { [name: string]: ImportedFileInfo } = {};
    const data = sheetImportState.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      fileNameToLastUpdated[data[i][0]] = {
        file: data[i][0],
        date: new Date(data[i][1]),
        row: i + 1,
        state: data[i][2],
      };
    }
    return fileNameToLastUpdated;
  }

  function saveImportState(info: ImportedFileInfo) {
    sheetImportState.getRange(info.row, 1, 1, 3).setValues([[info.file, info.date, info.state]]);
  }

  function importNewFile(csvFile: GoogleAppsScript.Drive.File) {
    console.log("Importing new file", csvFile.getName());
    const info: ImportedFileInfo = { file: csvFile.getName(), date: csvFile.getLastUpdated(), row: sheetImportState.getLastRow() + 1, state: "processing" };
    saveImportState(info);
    importFile(csvFile);
    saveImportState({ ...info, state: "finished" });
  }

  function importUpdatedFile(csvFile: GoogleAppsScript.Drive.File, lastImported: ImportedFileInfo) {
    console.log("Importing updated file", csvFile.getName());
    const info: ImportedFileInfo = { ...lastImported, date: csvFile.getLastUpdated(), state: "processing" };
    saveImportState(info);
    importFile(csvFile);
    saveImportState({ ...info, state: "finished" });
  }

  function importMoneyForwardCsvFromGoogleDrive() {
    const folder = DriveApp.getFolderById(GoogleDriveFolderId);
    console.log("folder:", folder.getName());
    const alreadyImported = loadImportStates();

    const csvFiles = folder.getFilesByType("text/csv");
    while (csvFiles.hasNext()) {
      const csvFile = csvFiles.next();
      const lastUpdated = csvFile.getLastUpdated();
      const lastImportedOrNull = alreadyImported[csvFile.getName()];
      if (lastImportedOrNull == null) {
        importNewFile(csvFile);
      } else if (lastImportedOrNull.state != "finished" || lastImportedOrNull.date < lastUpdated) {
        importUpdatedFile(csvFile, lastImportedOrNull);
      } else {
        console.log("Skip since not updated", csvFile.getName());
      }
    }
  }

  importMoneyForwardCsvFromGoogleDrive();
  Lib.spreadSheetSortSheets(spreadsheet);
}

function mainRebuildSummarySheets() {
  function buildSummarySheet(rebuild: boolean = false) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const existingSheet = spreadsheet.getSheetByName("Z_Summary");
    if (existingSheet != null && !rebuild) {
      return;
    }
    const summarySheet = existingSheet ?? spreadsheet.insertSheet("Z_Summary");

    const data: any[][] = [["Year", "Month", "Year/Month", "Income", "Expense", "Balance"]].concat(
      spreadsheet
        .getSheets()
        .filter((sheet) => sheet.getName().startsWith("Import_"))
        .sort((a, b) => a.getName().localeCompare(b.getName()))
        .map((sheet) => {
          const name = sheet.getName();
          const [_, year, month] = name.split("_");
          const buildColumnRange = (letter) => `${name}!$${letter}$2:$${letter}`; // Import_2023_03!$D$2:$D
          const amount = buildColumnRange(LETTERS[MoneyForwardExportedCSVColumns.Amount]);
          const calcTargetFlag = buildColumnRange(LETTERS[MoneyForwardExportedCSVColumns.CalcTargetFlag]);
          const transferFlag = buildColumnRange(LETTERS[MoneyForwardExportedCSVColumns.TransferFlag]);
          return [
            year,
            month,
            `${year}/${month}`,
            `=SUMIFS(${amount}, ${amount}, ">0", ${calcTargetFlag}, 1, ${transferFlag}, 0)`,
            `=SUMIFS(${amount}, ${amount}, "<0", ${calcTargetFlag}, 1, ${transferFlag}, 0) * -1`,
            `=SUMIFS(${amount}, ${calcTargetFlag}, 1, ${transferFlag}, 0)`,
          ];
        })
    );
    summarySheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  }

  function buildCategorySummarySheet(rebuild: boolean = false) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const existingSheet = spreadsheet.getSheetByName("Z_Category");
    if (existingSheet != null && !rebuild) {
      return;
    }
    const summarySheet = existingSheet ?? spreadsheet.insertSheet("Z_Category");

    const data: any[][] = [["Year", "Month", "Year/Month", "Category", "SubCategory", "Expense"]].concat(
      spreadsheet
        .getSheets()
        .filter((sheet) => sheet.getName().startsWith("Import_"))
        .sort((a, b) => a.getName().localeCompare(b.getName()))
        .map((sheet) => {
          const name = sheet.getName();
          const [_, year, month] = name.split("_");
          const buildColumnRange = (letter) => `${name}!$${letter}$2:$${letter}`; // Import_2023_03!$D$2:$D
          const amount = buildColumnRange(LETTERS[MoneyForwardExportedCSVColumns.Amount]);
          const calcTargetFlag = buildColumnRange(LETTERS[MoneyForwardExportedCSVColumns.CalcTargetFlag]);
          const transferFlag = buildColumnRange(LETTERS[MoneyForwardExportedCSVColumns.TransferFlag]);
          return [
            year,
            month,
            `${year}/${month}`,
            `=SUMIFS(${amount}, ${amount}, ">0", ${calcTargetFlag}, 1, ${transferFlag}, 0)`,
            `=SUMIFS(${amount}, ${amount}, "<0", ${calcTargetFlag}, 1, ${transferFlag}, 0) * -1`,
            `=SUMIFS(${amount}, ${calcTargetFlag}, 1, ${transferFlag}, 0)`,
          ];
        })
    );
    summarySheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    Lib.spreadSheetSortSheets(spreadsheet);
  }

  function buildAllItemsSheet(rebuild: boolean = false) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const existingSheet = spreadsheet.getSheetByName("Z_All");
    if (existingSheet != null && !rebuild) {
      return;
    }
    const sheet = existingSheet ?? spreadsheet.insertSheet("Z_All");

    const importedSheets = SpreadsheetApp.getActive()
      .getSheets()
      .filter((sheet) => sheet.getName().startsWith("Import_"));

    sheet.getRange(1, 1).setValue(`=ARRAYFORMULA(${importedSheets[0].getName()}!A1:${LETTERS[ExtendedColumns.__LENGTH - 1]}1)`);
    const importedRanges = importedSheets.map((sheet) => `${sheet.getName()}!A2:${LETTERS[ExtendedColumns.__LENGTH - 1]}`);
    sheet.getRange(2, 1).setValue(`=QUERY({${importedRanges.join(";")}},
              "select * where Col${MoneyForwardExportedCSVColumns.ID + 1} <> ''
                and Col${MoneyForwardExportedCSVColumns.CalcTargetFlag + 1} > 0
                and Col${MoneyForwardExportedCSVColumns.Amount + 1} < 0
                and Col${MoneyForwardExportedCSVColumns.TransferFlag + 1} = 0
              ")`);

    sheet.getRange(1, ExtendedColumns.__LENGTH + 1, 1, 2).setValues([["実質出費", "繰り返し回数"]]);
    const r = (idx: number) => `${LETTERS[idx]}2:${LETTERS[idx]}`;
    const id = r(MoneyForwardExportedCSVColumns.ID);
    const amount = r(MoneyForwardExportedCSVColumns.Amount);
    const myPayRatio = r(ExtendedColumns.MyPayRatio);
    const split = r(ExtendedColumns.NumPseudoSplit);
    sheet
      .getRange(2, ExtendedColumns.__LENGTH + 1, 1, 2)
      .setValues([
        [
          `=ARRAY_CONSTRAIN(ARRAYFORMULA(${amount}*IF(ISBLANK(${myPayRatio}),1,${myPayRatio})*-1*IF(ISBLANK(${split}), 1, 1/${split})), COUNTA(${id}), 1)`,
          `=ARRAY_CONSTRAIN(ARRAYFORMULA(IF(ISBLANK(${split}), 1, ${split})), COUNTA(${id}), 1)`,
        ],
      ]);
    sheet.getFilter()?.remove();
    const dataRange = sheet.getRange(`A:${LETTERS[sheet.getDataRange().getNumColumns() - 1]}`);
    dataRange.createFilter();
    Lib.spreadSheetSortSheets(spreadsheet);
  }

  function buildAllItemsRepeatedSheet(rebuild: boolean = false) {
    buildAllItemsSheet(rebuild);
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const existingSheet = spreadsheet.getSheetByName("Z_AllRepeated");
    if (existingSheet != null && !rebuild) {
      return;
    }
    const sheet = existingSheet ?? spreadsheet.insertSheet("Z_AllRepeated");
    const titles = ["ID", "RepeatIndex", "PayMonth", "Category", "SubCategory", "Expense"];
    sheet.getRange(1, 1, 1, titles.length).setValues([titles]);
    const r = (idx: number) => `Z_All!${LETTERS[idx]}2:${LETTERS[idx]}`;
    const id = r(MoneyForwardExportedCSVColumns.ID);
    const date = r(MoneyForwardExportedCSVColumns.Date);
    const cat = r(MoneyForwardExportedCSVColumns.Category);
    const subcat = r(MoneyForwardExportedCSVColumns.SubCategory);
    const amount = r(ExtendedColumns.__LENGTH - 1 + 1);
    const repeat = r(ExtendedColumns.__LENGTH - 1 + 2);
    const constrain = (folmula) => `=ARRAY_CONSTRAIN(ARRAYFORMULA(${folmula}), SUM(${repeat}), 1)`;
    sheet
      .getRange(2, 1, 1, titles.length)
      .setValues([
        [
          constrain(`TOCOL(FLATTEN(SPLIT(REPT(${id}&"@", ${repeat}), "@")),1)`),
          constrain(`TOCOL(FLATTEN(SPLIT(MAP(${repeat}, LAMBDA(rep, JOIN(",", SEQUENCE(rep)))), ",")),1)`),
          constrain(`TOCOL(FLATTEN(SPLIT(MAP(${repeat}, ${date}, LAMBDA(rep, date, JOIN(",", MAP(SEQUENCE(rep), LAMBDA(i, EOMONTH(EDATE(date, i), -1) + 1))))), ",")),1)`),
          constrain(`TOCOL(FLATTEN(SPLIT(REPT((${cat})&"@", ${repeat}), "@")),1)`),
          constrain(`TOCOL(FLATTEN(SPLIT(REPT((${subcat})&"@", ${repeat}), "@")),1)`),
          constrain(`TOCOL(FLATTEN(SPLIT(REPT((${amount})&"@", ${repeat}), "@")),1)`),
        ],
      ]);
  }

  buildSummarySheet(true);
  buildCategorySummarySheet(true);
  buildAllItemsRepeatedSheet(true);
}
