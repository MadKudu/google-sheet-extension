function onInstall(e) {
  onOpen(e);
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createAddonMenu()
    .addItem('Open', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('MadKudu')
    .setWidth(500);
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function alert(message) {
  SpreadsheetApp.getUi().alert(message);
}

function error(status, statusText, message, helpTip) {
  const template = HtmlService.createTemplateFromFile('Error');
  template.status     = status;
  template.statusText = statusText;
  template.message    = message;
  template.helpTip    = helpTip;
  const html = template.evaluate().setHeight(220);
  SpreadsheetApp.getUi()
          .showModalDialog(html, 'MadKudu Error');
}

function getActiveSheet() {
  return SpreadsheetApp.getActiveSheet().getSheetId();
}

function getSheetById(id) {
  const sheets = SpreadsheetApp.getActive().getSheets();
  for (var n in sheets) {
    if (sheets[n].getSheetId() === id) {
      return sheets[n];
    }
  }
  return null;
}

function getActiveSelection(sheetId) {
  const range = SpreadsheetApp.getActiveRange();
  return {
    row: range.getRowIndex(),
    column: range.getColumn(),
    values: range.getValues()
  };
}

function setCellValue(sheetId, row, col, val) {
  const sheet = getSheetById(sheetId);
  const cell = sheet.getRange(row, col);
  cell.setValue(val);
}

// cells should be an array of { row, col, val }
function setCells(sheetId, cells) {
  cells.forEach(function(cell) {
    setCellValue(sheetId, cell.row, cell.col, cell.val);
  });
  return true;
}

function createRowHeaders(sheetId, row, col, headers) {
  const sheet = getSheetById(sheetId);
  
  shouldInsertRow = row != 2;
  if (shouldInsertRow) {
    sheet.insertRowBefore(row);
  }
  const adjustedRow = (shouldInsertRow ? row : row - 1);
  // insert the headers
  headers.forEach(function (header, i) {
    if (i === 0) {
      // leave out the first column
      return;
    }
    const cell = sheet.getRange(adjustedRow, col + i);
    cell.setValue(header);
    cell.setFontWeight('bold');
  });
  
  return shouldInsertRow;
}

function saveUserProperty(key, value) {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty(key, value);
}

function deleteUserProperty(key) {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.deleteProperty(key);
}

function getUserProperty(key) {
  const userProperties = PropertiesService.getUserProperties();
  return userProperties.getProperty(key);
}
