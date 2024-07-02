function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Uncle Bob')
    .addItem('Format Documents', 'formatDocuments')
    .addItem('Chart Progress', 'chartProgress')
    .addToUi();
}

function formatDocuments() {
  const start = Date.now();
  const mainSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const fileId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const file = DriveApp.getFileById(fileId);
  const folder = file.getParents().next();
  const subFolders = folder.getFoldersByName('Q-A Sets');

  if (!subFolders.hasNext()) {
    throw new Error('Subdirectory Q-A Sets not found in the current folder.');
  }

  removeAllDataValidations();

  const { concatenatedData, fileNames } = fetchFilesAndConcatenateData(subFolders);

  const concatenatedSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Concatenated Q-A Data');
  concatenatedSheet.getRange(1, 1, concatenatedData.length, concatenatedData[0].length).setValues(concatenatedData);
  setupAndColorSheet(concatenatedSheet);
  const newSheetNames = splitAndSaveSheets(concatenatedSheet, fileNames);

  createListOfSheetNames(mainSheet, newSheetNames);

  const elapsedTime = Date.now() - start;
  const totalSeconds = Math.floor(elapsedTime / 1000);
  const minutes = Math.floor(totalSeconds / 60);
  const seconds = totalSeconds % 60;

  console.log(`Time elapsed: ${minutes} min ${seconds} sec`);
}

function removeAllDataValidations() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  sheets.forEach(sheet => {
    sheet.getDataRange().clearDataValidations();
  });
}

function fetchFilesAndConcatenateData(subFolders) {
  const files = subFolders.next().getFiles();
  const concatenatedData = [];
  const fileNames = [];

  while (files.hasNext()) {
    const file = files.next();
    fileNames.push(file.getName());
    const linkedSheet = SpreadsheetApp.openByUrl(file.getUrl()).getActiveSheet();
    const data = linkedSheet.getDataRange().getValues();
    concatenatedData.push(...data);
  }

  return { concatenatedData, fileNames };
}

function splitAndSaveSheets(concatenatedSheet, fileNames) {
  const totalRows = concatenatedSheet.getLastRow();
  const rowsPerSheet = Math.ceil(totalRows / fileNames.length);

  const newSheetNames = [];

  for (let i = 0; i < fileNames.length; i++) {
    const startRow = i * rowsPerSheet + 1;
    const endRow = Math.min(startRow + rowsPerSheet - 1, totalRows);
    const sheetData = concatenatedSheet.getRange(startRow, 1, endRow - startRow + 1, concatenatedSheet.getLastColumn()).getValues();

    const newSheetName = fileNames[i];
    const newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(newSheetName);
    newSheet.getRange(1, 1, sheetData.length, sheetData[0].length).setValues(sheetData);
    copyAndPasteWithFormatting(concatenatedSheet, newSheet, startRow, sheetData.length, concatenatedSheet.getLastColumn());

    newSheetNames.push(newSheetName);
  }

  SpreadsheetApp.getActiveSpreadsheet().deleteSheet(concatenatedSheet);
  return newSheetNames;
}

function copyAndPasteWithFormatting(sourceSheet, targetSheet, startRow, numRows, numCols) {
  const sourceRange = sourceSheet.getRange(startRow, 1, numRows, numCols);
  const targetRange = targetSheet.getRange(1, 1, numRows, numCols);
  sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_COLUMN_WIDTHS, false);
}

function setupAndColorSheet(sheet) {
  const lastRow = sheet.getLastRow();
  const checkboxColumns = ['B', 'C', 'D', 'E'];
  const contentColumn = 'C';
  const destinationColumn = 'F';

  sheet.getRange(contentColumn + "1:" + contentColumn + lastRow).copyTo(sheet.getRange(destinationColumn + "1:" + destinationColumn + lastRow), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

  checkboxColumns.forEach(column => {
    sheet.getRange(column + "1:" + column + lastRow).insertCheckboxes();
    sheet.setColumnWidth(column.charCodeAt(0) - 64, 50);
  });

  const rules = sheet.getConditionalFormatRules();
  const colors = ['#8FC08F', '#FFF89A', '#dd7e6b'];

  checkboxColumns.forEach((column, index) => {
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$' + column + '1=TRUE')
      .setBackground(colors[index])
      .setRanges([sheet.getRange("A1:A" + lastRow)])
      .build();
    rules.push(rule);
  });

  const fontColorRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$E1=TRUE')
    .setFontColor("#FFFFFF")
    .setRanges([sheet.getRange(destinationColumn + "1:" + destinationColumn + lastRow)])
    .build();
  rules.push(fontColorRule);
  sheet.setConditionalFormatRules(rules);

  colorCheckboxes(sheet, lastRow);
  applyBoldAndRemoveCheckboxes(sheet);
}

function colorCheckboxes(sheet, lastRow) {
  const range = sheet.getRange("B1:E" + lastRow);
  const values = range.getValues();
  const colors = range.getFontColors();

  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      if (typeof values[i][j] === 'boolean') {
        switch (j) {
          case 0:
            colors[i][j] = '#8fc08f'; // Applies green to column B
            break;
          case 1:
            colors[i][j] = '#E1C041'; // Applies yellow to column C
            break;
          case 2:
            colors[i][j] = '#dd7e6b'; // Applies red to column D
            break;
          case 3:
            colors[i][j] = '#000000'; // Explicitly set black (or any default) to column E
            break;
        }
      }
    }
  }

  range.setFontColors(colors);
}

function applyBoldAndRemoveCheckboxes(sheet) {
  const range = sheet.getDataRange();
  const values = range.getValues();

  const rowsToUpdate = [];
  const rowsToRemoveCheckboxes = [];

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (row[0] && !row[5]) {
      rowsToUpdate.push(i + 1);
      rowsToRemoveCheckboxes.push(i + 1);
    }
    if (!row[0] && !row[1]) {
      rowsToRemoveCheckboxes.push(i + 1);
    }
  }

  if (rowsToUpdate.length > 0) {
    sheet.getRangeList(rowsToUpdate.map(row => `A${row}`)).setFontWeight('bold');
  }

  if (rowsToRemoveCheckboxes.length > 0) {
    const clearRanges = ['B', 'C', 'D', 'E'].flatMap(col => rowsToRemoveCheckboxes.map(row => `${col}${row}`));
    const rangeList = sheet.getRangeList(clearRanges);
    rangeList.clearContent();
    rangeList.clearDataValidations();
  }
}

function createListOfSheetNames(mainSheet, newSheetNames) {
  const lastRow = mainSheet.getLastRow() + 1;
  const fileData = newSheetNames.map(name => [name]);

  mainSheet.getRange(lastRow, 1, fileData.length, 1).setValues(fileData).setFontSize(10).setFontWeight('normal').setWrap(true);

  mainSheet.getRange(lastRow, 2, fileData.length, 1).insertCheckboxes();
}

function chartProgress() {
  const ui = SpreadsheetApp.getUi();
  const mainSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const rows = mainSheet.getDataRange().getValues();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  let anyProcessed = false;

  rows.forEach((row, index) => {
    if (row[1] === true) { // Check if checkbox is ticked
      const sheetName = row[0];
      const linkedSheet = spreadsheet.getSheetByName(sheetName);

      if (linkedSheet) {
        anyProcessed = true;
        processQASheet(linkedSheet, mainSheet, index + 1);
      }
    }
  });

  if (!anyProcessed) {
    ui.alert('No Q-A sets selected to count questions. Please check at least one and ensure they contain valid sheet names.');
  }
}

function processQASheet(qaSheet, mainSheet, rowIndex) {
  const data = qaSheet.getDataRange().getValues();
  let totalQuestions = 0;
  let greenQuestions = 0;

  data.forEach(row => {
    // Assuming column B is used for marking green
    if (row[1] !== "" && row[1] !== undefined && row[1] !== null) {
      totalQuestions++;
      if (row[1] === true) {
        greenQuestions++;
      }
    }
  });

  const percentGreen = totalQuestions > 0 ? (greenQuestions / totalQuestions * 100) : 0;
  const formattedPercentGreen = percentGreen.toFixed(0) + '%';

  const color = getColorBasedOnPercentage(percentGreen);

  // Find the next available column in the main sheet for updating progress
  const rowRange = mainSheet.getRange(rowIndex, 3, 1, mainSheet.getLastColumn());
  const rowValues = rowRange.getValues()[0];
  let targetColumn = rowValues.findIndex(value => !value) + 3; // +3 because range starts at column C
  if (targetColumn < 3) {
    targetColumn = mainSheet.getLastColumn() + 1;
  }

  const currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yy");
  const outputText = `${currentDate}\n${totalQuestions} questions\n${formattedPercentGreen} green`;

  const targetCell = mainSheet.getRange(rowIndex, targetColumn);
  targetCell.setValue(outputText);
  targetCell.setBackground(color);
}

function getColorBasedOnPercentage(percentGreen) {
  if (percentGreen >= 90) return '#93c47d'; // Green
  if (percentGreen >= 80) return '#b6d7a8'; // Light Green
  if (percentGreen >= 70) return '#ffd966'; // Yellow
  if (percentGreen >= 60) return '#f6b26b'; // Orange
  return '#dd7e6b'; // Red
}
