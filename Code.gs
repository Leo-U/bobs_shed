function formatRoadmap() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
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
  const headers = sheet.getRange('A1:B1');
  headers.setValues([['Q-A sets', 'Run program']]);
  headers.setFontWeight('bold');

  if (fileData.length > 0) {
    const range = sheet.getRange(2, 1, fileData.length, 1); // Start from row 2
    range.setValues(fileData);
    range.setFontSize(10);
    range.setFontWeight('normal');
    range.setWrap(true);

    fileData.forEach((formula, index) => {
      const cell = sheet.getRange(index + 2, 1); // Adjust index to start from row 2
      cell.setFormula(formula[0]);
    });

    // Add checkboxes in column B, starting from row 2
    const checkBoxRange = sheet.getRange(2, 2, checkBoxes.length, 1);
    checkBoxRange.insertCheckboxes();
  } else {
    // If no files found, inform in row 2 to keep the headers clean
    const range = sheet.getRange(2, 1, 1, 1);
    range.setValue('No files found');
    range.setFontSize(10);
    range.setFontWeight('normal');
    range.setWrap(true);
  }
}
