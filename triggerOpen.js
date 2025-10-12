function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('GraphGenerator')
      .addItem('BIF Generator', 'showCsvBIF')
      .addItem('HCD Generator', 'showCsvHCD')
      .addItem('FRG Generator', 'showCsvFRG')
      .addToUi();

  ui.createMenu('FRG Checker (LLM)')
      .addItem('FRGシートの選択行をチェック', 'runFRGChecksOnSelectedRowWithUI')
      .addItem('FRGシートの全データをチェック', 'runFRGChecksAllWithUI')
      .addToUi();
}
