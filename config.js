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
