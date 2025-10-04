function myFRGUrl()
{
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  var spreadsheetId = spreadsheet.getId();

  var spreadsheetUrl = spreadsheet.getUrl();
  Logger.log(spreadsheetUrl);

  var frgSheet = spreadsheet.getSheetByName("FRG");
  var frgSheetId = frgSheet.getSheetId();

  var frgUrl = spreadsheetUrl + "?gid=" + frgSheetId + "#gid=" + frgSheetId;
  Logger.log(frgUrl);

  return [frgUrl, spreadsheetId];
}

//=======================================================================
// Create Menu
//=======================================================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('GraphGenerator')
      .addItem('BIF Generator', 'showCsvBIF')
      .addItem('HCD Generator', 'showCsvHCD')
      .addItem('FRG Generator', 'showCsvFRG')
      .addToUi();
}

//=======================================================================
// Connections Check (BIF)
//=======================================================================
function showCsvBIF() {
  var html = HtmlService.createTemplateFromFile("dialogBIF").evaluate();
  SpreadsheetApp.getUi().showModalDialog(html, "以下をコピーしてください。");
}

function getCsvBIF() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var values = spreadsheet.getSheetByName("[WS] GraphBIF").getRange("V:V").getValues();

  var dataArray = [];
  dataArray.push("##");
  dataArray.push("## BIF Generator");
  dataArray.push("##");
  for (var i = 0; i < values.length; i++) {
    if (values[i] == "")
      break;

    dataArray.push(values[i]);
  }
  dataArray.push("\r\n");
  return dataArray.join("\r\n");
}

//=======================================================================
// Projections Check (HCD)
//=======================================================================
function showCsvHCD() {
  var html = HtmlService.createTemplateFromFile("dialogHCD").evaluate();
  SpreadsheetApp.getUi().showModalDialog(html, "以下をコピーしてください。");
}

function getCsvHCD() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var values = spreadsheet.getSheetByName("[WS] GraphHCD").getRange("AB:AB").getValues();

  var dataArray = [];
  dataArray.push("##");
  dataArray.push("## HCD Generator");
  dataArray.push("##");
  for (var i = 0; i < values.length; i++) {
    if (values[i] == "")
      break;

    dataArray.push(values[i]);
  }
  dataArray.push("\r\n");
  return dataArray.join("\r\n");
}

//=======================================================================
// FRG Check
//=======================================================================
function showCsvFRG() {
  var html = HtmlService.createTemplateFromFile("dialogFRG").evaluate();
  SpreadsheetApp.getUi().showModalDialog(html, "以下をコピーしてください。");
}

function getCsvFRG() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var values = spreadsheet.getSheetByName("[WS] GraphFRG").getRange("M:M").getValues();

  var dataArray = [];
  dataArray.push("##");
  dataArray.push("## FRG Generator");
  dataArray.push("##");
  for (var i = 0; i < values.length; i++) {
    if (values[i] == "")
      break;

    dataArray.push(values[i]);
  }
  dataArray.push("\r\n");
  return dataArray.join("\r\n");
}
