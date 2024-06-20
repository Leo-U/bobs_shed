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

  // Set headers and make them bold
  mainSheet.getRange('A1:B1').setValues([['Q-A sets', 'Run program']]).setFontWeight('bold');

  if (fileData.length > 0) {
    const range = mainSheet.getRange(2, 1, fileData.length, 1); // Start from row 2
    range.setValues(fileData);
    range.setFontSize(10);
    range.setFontWeight('normal');
    range.setWrap(true);

    fileData.forEach((formula, index) => {
      const cell = mainSheet.getRange(index + 2, 1); // Adjust index to start from row 2
      cell.setFormula(formula[0]);
      // Open the sheet linked in the cell and apply formatting
      const linkedSheetId = formula[0].match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/)[1];
      const linkedSheet = SpreadsheetApp.openById(linkedSheetId).getActiveSheet();
      setupAndColorSheet(linkedSheet);
    });

    // Add checkboxes in column B, starting from row 2
    const checkBoxRange = mainSheet.getRange(2, 2, checkBoxes.length, 1);
    checkBoxRange.insertCheckboxes();
  } else {
    // If no files found, inform in row 2 to keep the headers clean
    mainSheet.getRange(2, 1, 1, 1).setValue('No files found').setFontSize(10).setFontWeight('normal').setWrap(true);
  }
}

function setupAndColorSheet(sheet) {
  const lastRow = sheet.getLastRow();
  const checkboxColumns = ['B', 'C', 'D', 'E'];
  const contentColumn = 'C';
  const destinationColumn = 'F';

  // Copy content and insert checkboxes in bulk
  const contentRange = sheet.getRange(contentColumn + "1:" + contentColumn + lastRow);
  const destinationRange = sheet.getRange(destinationColumn + "1:" + destinationColumn + lastRow);
  contentRange.copyTo(destinationRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  
  checkboxColumns.forEach(column => {
    const checkboxRange = sheet.getRange(column + "1:" + column + lastRow);
    checkboxRange.insertCheckboxes();
    sheet.setColumnWidth(column.charCodeAt(0) - 64, 50);
  });

  // Apply conditional formatting rules
  let rules = sheet.getConditionalFormatRules();
  const colors = ['#8FC08F', '#FFF89A', '#dd7e6b']; // Define color codes

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

  // Additional setup and clean-up functions are omitted for brevity
}
