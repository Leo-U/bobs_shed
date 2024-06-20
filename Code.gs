function formatRoadmapAndApplyFormatting() {
  const mainSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const fileId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const file = DriveApp.getFileById(fileId);
  const folders = file.getParents();

  if (!folders.hasNext()) {
    throw new Error('No parent folders found.');
  }

  const folder = folders.next();
  const subFolders = folder.getFoldersByName('Q-A Sets');

  if (!subFolders.hasNext()) {
    throw new Error('Subdirectory Q-A Sets not found in the current folder.');
  }

  const files = subFolders.next().getFiles();
  const fileData = [];
  const checkBoxes = [];

  while (files.hasNext()) {
    let file = files.next();
    const url = file.getUrl();
    const name = file.getName();
    const hyperlinkFormula = `=HYPERLINK("${url}", "${name}")`;
    fileData.push([hyperlinkFormula]);
    checkBoxes.push([true]);  // Initially set all checkboxes to unchecked
  }

  mainSheet.getRange('A1:B1').setValues([['Q-A sets', 'Run program']]).setFontWeight('bold');

  if (fileData.length > 0) {
    const range = mainSheet.getRange(2, 1, fileData.length, 1);
    range.setValues(fileData);
    range.setFontSize(10);
    range.setFontWeight('normal');
    range.setWrap(true);

    fileData.forEach((formula, index) => {
      const cell = mainSheet.getRange(index + 2, 1);
      cell.setFormula(formula[0]);
      const linkedSheetId = formula[0].match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/)[1];
      const linkedSheet = SpreadsheetApp.openById(linkedSheetId).getActiveSheet();
      setupAndColorSheet(linkedSheet);
    });

    const checkBoxRange = mainSheet.getRange(2, 2, checkBoxes.length, 1);
    checkBoxRange.insertCheckboxes();
  } else {
    mainSheet.getRange(2, 1, 1, 1).setValue('No files found').setFontSize(10).setFontWeight('normal').setWrap(true);
  }
}

function setupAndColorSheet(sheet) {
  const lastRow = sheet.getLastRow();
  const checkboxColumns = ['B', 'C', 'D', 'E'];
  const contentColumn = 'C';
  const destinationColumn = 'F';

  const contentRange = sheet.getRange(contentColumn + "1:" + contentColumn + lastRow);
  const destinationRange = sheet.getRange(destinationColumn + "1:" + destinationColumn + lastRow);
  contentRange.copyTo(destinationRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  
  checkboxColumns.forEach(column => {
    const checkboxRange = sheet.getRange(column + "1:" + column + lastRow);
    checkboxRange.insertCheckboxes();
    sheet.setColumnWidth(column.charCodeAt(0) - 64, 50);
  });

  let rules = sheet.getConditionalFormatRules();
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
  applyBoldAndRemoveCheckboxesEfficiently(sheet);
}

function colorCheckboxes(sheet, lastRow) {
  var range = sheet.getRange("B1:D" + lastRow);
  var values = range.getValues();
  var colors = range.getFontColors();

  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (typeof values[i][j] === 'boolean') {
        switch(j) {
          case 0:
            colors[i][j] = '#8fc08f';
            break;
          case 1:
            colors[i][j] = '#E1C041';
            break;
          case 2:
            colors[i][j] = '#dd7e6b';
            break;
        }
      }
    }
  }

  range.setFontColors(colors);
}

function applyBoldAndRemoveCheckboxesEfficiently(sheet) {
  const range = sheet.getDataRange();
  const values = range.getValues();

  let rowsToUpdate = [];
  let rowsToRemoveCheckboxes = [];

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
    const boldRanges = rowsToUpdate.map(row => `A${row}`);
    sheet.getRangeList(boldRanges).setFontWeight('bold');
  }

  rowsToRemoveCheckboxes = [...new Set(rowsToRemoveCheckboxes)];

  if (rowsToRemoveCheckboxes.length > 0) {
    const clearRanges = ['B', 'C', 'D', 'E'].flatMap(col => rowsToRemoveCheckboxes.map(row => `${col}${row}`));
    const rangeList = sheet.getRangeList(clearRanges);
    rangeList.clearContent();
    rangeList.clearDataValidations();
  }
}
