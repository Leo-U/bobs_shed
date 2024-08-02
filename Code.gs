function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  ui.createAddonMenu('Learning Insights')
    .addItem('Format All Sheets', 'formatDocuments')
    .addItem('Format Additional Sheet', 'formatIndividualSheet')
    .addItem('Chart Progress', 'chartProgress')
    .addToUi();

  // Check and set the main chart sheet ID if not already set
  const properties = PropertiesService.getScriptProperties();
  if (!properties.getProperty('MAIN_CHART_SHEET_ID')) {
    // Assuming the main chart is the first sheet created, we designate it
    const firstSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    properties.setProperty('MAIN_CHART_SHEET_ID', firstSheet.getSheetId().toString());
  }
}

function formatIndividualSheet() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Enter the filename (without extension) to format:');

  if (response.getSelectedButton() == ui.Button.OK) {
    const filename = response.getResponseText();
    const file = findFileInQASetsFolder(filename);

    if (file) {
      // Start processing and display a toast message
      SpreadsheetApp.getActiveSpreadsheet().toast('Formatting documents. Please wait...', 'Status', -1);  // -1 indicates that it will stay until explicitly cleared
      
      const sourceSheet = SpreadsheetApp.open(file).getActiveSheet();
      const newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(filename);
      copyDataToNewSheet(sourceSheet, newSheet);
      setupAndColorSheet(newSheet);
      wrapText(newSheet);  // Wrap text for the entire sheet
      updateMainChart(filename);

      // Finish processing and clear the toast message
      SpreadsheetApp.flush();  // Apply all pending Spreadsheet changes
      SpreadsheetApp.getActiveSpreadsheet().toast('Formatting completed successfully.', 'Status', 3);  // 3 seconds before disappearing
    } else {
      ui.alert('File not found in the Q-A Sets folder.');
    }
  } else {
    ui.alert('Action canceled.');
  }
}

function wrapText(sheet) {
  const range = sheet.getDataRange();
  range.setWrap(true);  // Set text wrapping for all cells in the range
}

function findFileInQASetsFolder(filename) {
  const fileId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const file = DriveApp.getFileById(fileId);
  const folder = file.getParents().next();
  const subFolders = folder.getFoldersByName('Q-A Sets');

  if (!subFolders.hasNext()) {
    throw new Error('Subdirectory Q-A Sets not found in the current folder.');
  }

  const files = subFolders.next().getFilesByName(filename);
  return files.hasNext() ? files.next() : null;
}

function copyDataToNewSheet(sourceSheet, targetSheet) {
  const data = sourceSheet.getDataRange().getValues();
  const targetRange = targetSheet.getRange(1, 1, data.length, data[0].length);
  targetRange.setValues(data);
}

function updateMainChart(filename) {
  let mainSheet = findMainChartSheet();

  if (mainSheet) {
    const lastRow = mainSheet.getLastRow() + 1;
    mainSheet.getRange(lastRow, 1).setValue(filename);
    mainSheet.getRange(lastRow, 2).insertCheckboxes();
  } else {
    throw new Error('Main chart sheet not found.');
  }
}

function findMainChartSheet() {
  const properties = PropertiesService.getScriptProperties();
  const mainChartSheetId = properties.getProperty('MAIN_CHART_SHEET_ID');
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();

  for (const sheet of sheets) {
    if (sheet.getSheetId().toString() === mainChartSheetId) {
      return sheet;
    }
  }
  return null;
}

function formatDocuments() {
  const mainSheet = findMainChartSheet();
  if (!mainSheet) {
    throw new Error('Main chart sheet not found.');
  }
  const checkRange = mainSheet.getRange('B1:B5').getValues(); // Get values from the first five rows of column B

  // Check if there are any checkboxes (either TRUE or FALSE)
  const hasCheckboxes = checkRange.some(row => row[0] === true || row[0] === false);

  if (hasCheckboxes) {
    SpreadsheetApp.getUi().alert("One-time operation. To add more sheets, use 'Format Individual Sheet'");
    return; // Exit the function if any cell has a checkbox
  }

  // Continue with the rest of the function
  const fileId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const file = DriveApp.getFileById(fileId);
  const folder = file.getParents().next();
  const subFolders = folder.getFoldersByName('Q-A Sets');

  if (!subFolders.hasNext()) {
    throw new Error('Subdirectory Q-A Sets not found in the current folder.');
  }

  removeAllDataValidations();

  // Start processing and display a toast message
  SpreadsheetApp.getActiveSpreadsheet().toast('Formatting documents. Please wait...', 'Status', -1);  // -1 indicates that it will stay until explicitly cleared

  const { concatenatedData, fileNames, rowCounts } = fetchFilesAndConcatenateData(subFolders);

  const concatenatedSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Concatenated Q-A Data');
  concatenatedSheet.getRange(1, 1, concatenatedData.length, concatenatedData[0].length).setValues(concatenatedData);
  setupAndColorSheet(concatenatedSheet);
  wrapText(concatenatedSheet);  // Wrap text for the entire sheet
  const newSheetNames = splitAndSaveSheets(concatenatedSheet, fileNames, rowCounts);

  createListOfSheetNames(mainSheet, newSheetNames);

  // Finish processing and clear the toast message
  SpreadsheetApp.flush();  // Apply all pending Spreadsheet changes
  SpreadsheetApp.getActiveSpreadsheet().toast('Formatting completed successfully.', 'Status', 3);  // 3 seconds before disappearing
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
  const rowCounts = [];  // Array to store row counts for each file

  while (files.hasNext()) {
    const file = files.next();
    fileNames.push(file.getName());
    const linkedSheet = SpreadsheetApp.openByUrl(file.getUrl()).getActiveSheet();
    const data = linkedSheet.getDataRange().getValues();
    concatenatedData.push(...data);
    rowCounts.push(data.length);  // Store the number of rows added for this file
  }

  return { concatenatedData, fileNames, rowCounts };  // Include rowCounts in the return value
}

function splitAndSaveSheets(concatenatedSheet, fileNames, rowCounts) {
  let startRow = 1;
  const newSheetNames = [];

  for (let i = 0; i < fileNames.length; i++) {
    const numRows = rowCounts[i];  // Get the number of rows for the current file
    const endRow = startRow + numRows - 1;
    const newSheetName = fileNames[i];
    const newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(newSheetName);

    // Using copyTo for copying data along with formatting and validations
    const sourceRange = concatenatedSheet.getRange(startRow, 1, numRows, concatenatedSheet.getLastColumn());
    const targetRange = newSheet.getRange(1, 1, numRows, concatenatedSheet.getLastColumn());
    sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_COLUMN_WIDTHS, false);
    sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_CONDITIONAL_FORMATTING, false);
    sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    
    wrapText(newSheet);  // Wrap text for the entire new sheet
    newSheetNames.push(newSheetName);
    startRow = endRow + 1;  // Update startRow for the next file
  }

  SpreadsheetApp.getActiveSpreadsheet().deleteSheet(concatenatedSheet);
  return newSheetNames;
}

function setupAndColorSheet(sheet) {
  const lastRow = sheet.getLastRow();
  const checkboxColumns = ['B', 'C', 'D', 'E'];
  const contentColumn = 'B';
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
  const mainSheet = findMainChartSheet();
  if (!mainSheet) {
    ui.alert('Main chart sheet not found.');
    return;
  }
  
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

        // Provide feedback on the sheet being processed
        spreadsheet.toast(`Charting progress in ${sheetName}...`, 'Status', -1);
      }
    }
  });

  if (!anyProcessed) {
    ui.alert('No Q-A sets selected to count questions. Please check at least one and ensure they contain valid sheet names.');
  }

  // Finish processing and clear the toast message
  SpreadsheetApp.flush();  // Apply all pending Spreadsheet changes
  spreadsheet.toast('Progress charted successfully.', 'Status', 3);  // 3 seconds before disappearing
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
