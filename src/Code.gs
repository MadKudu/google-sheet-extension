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
  SpreadsheetApp
    .getUi()
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

function setCellValue(sheetId, row, col, val, isError) {
  const sheet = getSheetById(sheetId);
  const cell = sheet.getRange(row, col);
  
  if (isError) {
    cell.setBackground('red').setFontColor('white')
  } else {
    cell.setBackground('white').setFontColor('black')
  }
  
  cell.setFontWeight('normal')
  cell.setValue(val);
}

// cells should be an array of { row, col, val }
function setCells(sheetId, cells, isError) {
  cells.forEach(function(cell) {
    setCellValue(sheetId, cell.row, cell.col, cell.val, isError);
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

function getError500() {
  return { 
    message: 'An unexpected error occured, please retry this record', 
    status: 500, 
    error: true 
  }
}

function parseApiResults (res) {
  const status = res.getResponseCode();
  const body = res.getContentText();
  console.log(status + ' ' + body);
  var parsedRes = {};
  try {
    parsedRes = JSON.parse(body);
    parsedRes.error = res.getResponseCode() !== 200;
    parsedRes.status = res.getResponseCode();
  } catch (err) {
    // this will catch timeouts and others
    parsedRes = { message: 'An unexpected error occured, please retry this record', status: 500, error: true }
  }
  if (parsedRes.error) {
    console.error('error', parsedRes);
  }
  return parsedRes;
}

function getAuthorization() {
  const apiKey = getUserProperty('madkudu_api_key');
  return 'Basic ' + Utilities.base64Encode(apiKey);
}

function getTopicalPrediction(domain, modelId) {
  const baseUrl = 'https://api.madkudu.com/v1/';
  domain = domain.replace(/https?:\/\/(www\.)?/, '') // trim http and www
  if (domain[domain.length - 1] === '/') domain = domain.slice(0, domain.length - 1); // trim / at end if exists
  const payload = JSON.stringify({ domain: domain, show_scores: true });
  const url = baseUrl + 'models/' + modelId + '/predictions';
  const params = {
    method: 'post',
    payload: payload,
    headers: {
      'Authorization': getAuthorization(),
      'content-type': 'application/json'
    },
    muteHttpExceptions: true // use this to be able to catch the status of the error
  };
  try {
    const res = UrlFetchApp.fetch(url, params);
    return parseApiResults(res);
  } catch (err) {
    return getError500()
  }
}

function getCustomerFitPrediction(domainOrEmail, model) {
  const baseUrl = 'https://api.madkudu.com/v1/';
  const paramKey = (model === 'companies' ? 'domain' : 'email');
  if (paramKey === 'domain') {
    domainOrEmail = domainOrEmail.replace(/https?:\/\/(www\.)?/, ''); // trim http and www
    if (domainOrEmail[domainOrEmail.length - 1] === '/') {
      domainOrEmail = domainOrEmail.slice(0, domainOrEmail.length - 1); // trim / at end if exists
    }
  }
  const url = baseUrl + model + '?' + paramKey + '=' + encodeURIComponent(domainOrEmail);
  const params = {
    method: 'get',
    headers: {
      'Authorization': getAuthorization()
    },
    muteHttpExceptions: true, // use this to be able to catch the status of the error
  };
  try {
    const res = UrlFetchApp.fetch(url, params);
    return parseApiResults(res);
  } catch (err) {
    console.error('timeout', err);
    return getError500()
  }
}
