/**
 * FRGシートの選択行をチェック
 */
function runFRGChecksOnSelectedRowWithUI() {
  const ui = SpreadsheetApp.getUi();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const activeCell = sheet.getActiveCell();
  if (!activeCell) {
    ui.alert('エラー', 'セルが選択されていません。処理対象の行のいずれかのセルを選択してください。', ui.ButtonSet.OK);
    return;
  }
  const currentRow = activeCell.getRow();

  const confirm = ui.alert(
      '確認',
      `選択中の ${currentRow} 行目に対してFRGチェックを実行しますか？\n処理には数分かかることがあります。`,
      ui.ButtonSet.YES_NO);

  if (confirm === ui.Button.YES) {
    SpreadsheetApp.getActiveSpreadsheet().toast('FRGチェック処理を開始します...', '処理中', -1);
    try {
      LLMChecksFRG(sheet, currentRow);

      SpreadsheetApp.getActiveSpreadsheet().toast('FRGチェック処理が完了しました。', '完了', 5);
      ui.alert('完了', `${currentRow} 行目のFRGチェック処理が完了しました。`, ui.ButtonSet.OK);
    } catch (e) {
      Logger.log(`Error in runFRGChecksOnSelectedRowWithUI: ${e.toString()}\nStack: ${e.stack}`);
      SpreadsheetApp.getActiveSpreadsheet().toast('エラーが発生しました。詳細はログを確認してください。', 'エラー', 10);
      ui.alert('エラー', `処理中にエラーが発生しました: ${e.message}\n詳細は[表示] > [ログ]で確認してください。`, ui.ButtonSet.OK);
    }
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast('処理はキャンセルされました。', 'キャンセル', 5);
  }
}

/**
 *  FRGシートの全データをチェック
 */
function runFRGChecksAllWithUI() {
  const ui = SpreadsheetApp.getUi();

  const confirm = ui.alert(
      '確認',
      '"FRG Review End Line"までのFRGチェックを実行します。\n処理には時間がかかることがあります。',
      ui.ButtonSet.YES_NO);

  if (confirm === ui.Button.YES) {
    SpreadsheetApp.getActiveSpreadsheet().toast('FRGチェック処理を開始します...', '処理中', -1);
    try {
      checkAllFRG();

      ui.alert('完了', `FRGチェック処理が完了しました。`, ui.ButtonSet.OK);
    } catch (e) {
      Logger.log(`Error in runFRGChecksOnSelectedRowWithUI: ${e.toString()}\nStack: ${e.stack}`);
      SpreadsheetApp.getActiveSpreadsheet().toast('エラーが発生しました。詳細はログを確認してください。', 'エラー', 10);
      ui.alert('エラー', `処理中にエラーが発生しました: ${e.message}\n詳細は[表示] > [ログ]で確認してください。`, ui.ButtonSet.OK);
    }
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast('処理はキャンセルされました。', 'キャンセル', 5);
  }
}

function checkAllFRG()
{
  var sheetProject = getSheet("Project");
  var sheetFRG = getSheet("FRG");

  var frgReviewEL = sheetProject.getRange("B8").getValue();
  Logger.log("FRG Review End Line: " + frgReviewEL);
  if (!checkReviewEL(frgReviewEL)) {
    throw new Error("Please enter a number greater than or equal to 1 in \"Review End Line\" of the Project sheet. (FRG)");
  }

  for (var i = 2; i <= frgReviewEL; i++) {
    Logger.log("######################### Row: " + i);

    LLMChecksFRG(sheetFRG, i);
  }
}

function getSheet(sheetName)
{
  var braSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return braSpreadsheet.getSheetByName(sheetName);
}

function checkReviewEL(value)
{
  if (Number.isInteger(value)) {
    if (value > 0)
      return true;
  }
  return false;
}
