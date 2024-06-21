function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Uncle Bob')
    .addItem('Prepare Documents', 'formatRoadmapAndApplyFormatting')
    .addItem('Map Progress', 'countQuestions')
    .addToUi();
}

function formatRoadmapAndApplyFormatting() {
  const ui = SpreadsheetApp.getUi();
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
    checkBoxes.push([true]);
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
      const linkedSheet = SpreadsheetApp.openByUrl(cell.getFormula().match(/"(.*?)"/)[1]).getActiveSheet();
      if (linkedSheet.getRange('Z1').getValue() !== 'Formatted') {
        setupAndColorSheet(linkedSheet);
        linkedSheet.getRange('Z1').setValue('Formatted');
      }
    });

    const checkBoxRange = mainSheet.getRange(2, 2, checkBoxes.length, 1);
    checkBoxRange.insertCheckboxes();
  } else {
    mainSheet.getRange(2, 1, 1, 1).setValue('No files found').setFontSize(10).setFontWeight('normal').setWrap(true);
  }
}

function setupAndColorSheet(sheet) {
  if (sheet.getRange('Z1').getValue() === 'Formatted') return;  // Check if already formatted

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
  if (sheet.getRange('Z1').getValue() === 'Formatted') return;  // Check if already formatted

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
  if (sheet.getRange('Z1').getValue() === 'Formatted') return;  // Check if already formatted

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

function countQuestions() {
  const ui = SpreadsheetApp.getUi();
  const mainSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const rows = mainSheet.getDataRange().getValues();

  let anyProcessed = false;

  rows.forEach((row, index) => {
    // Check if the checkbox in column B is checked
    if (row[1] === true) {
      const cell = mainSheet.getRange('A' + (index + 1));
      const richText = cell.getRichTextValue();
      const linkUrl = richText.getLinkUrl();

      if (linkUrl) { // Ensure both checkbox is checked and hyperlink is present
        anyProcessed = true;
        const linkedSheet = SpreadsheetApp.openByUrl(linkUrl).getActiveSheet();
        processQASheet(linkedSheet, mainSheet, index + 1);
      }
    }
  });

  if (!anyProcessed) {
    ui.alert('No Q-A sets selected or valid to count questions. Please check at least one and ensure they contain valid links.');
  }
}

function processQASheet(qaSheet, mainSheet, rowIndex) {
  const data = qaSheet.getDataRange().getValues();
  let totalQuestions = 0;
  let greenQuestions = 0;

  // Counting total and green questions based on some conditions assumed in column B
  data.forEach(row => {
    if (row[1]) {  // Assuming the relevant data is in column B
      totalQuestions++;
      if (row[1] === true) {  // Assuming true represents a 'green' question
        greenQuestions++;
      }
    }
  });

  // Calculate percentage of green questions
  const percentGreen = totalQuestions > 0 ? (greenQuestions / totalQuestions * 100) : 0;
  const formattedPercentGreen = percentGreen.toFixed(0) + '%';

  // Find the first empty cell in the specified row to place the new data
  const rowRange = mainSheet.getRange(rowIndex, 3, 1, mainSheet.getLastColumn());
  const rowValues = rowRange.getValues()[0];
  let targetColumn = rowValues.findIndex(value => !value) + 3; // +3 because range starts at column C
  if (targetColumn === 2) {  // +3 made it out of bounds, meaning row is full
    targetColumn = mainSheet.getLastColumn() + 1;  // Move to the next available column outside current range
  }

  // Prepare data to be written
  const currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yy");
  const outputText = `${currentDate}\n${totalQuestions} t ${greenQuestions} g\n${formattedPercentGreen}`;

  // Write data to the next available column in the same row
  mainSheet.getRange(rowIndex, targetColumn).setValue(outputText);
}
