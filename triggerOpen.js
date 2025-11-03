function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('GraphGenerator')
      .addItem('BIF Generator', 'showCsvBIF')
      .addItem('HCD Generator', 'showCsvHCD')
      .addItem('FRG Generator', 'showCsvFRG')
      .addToUi();

  ui.createMenu('FRG Checker (LLM)')
      .addItem('Selected Line', 'runFRGChecksOnSelectedRowWithUI')
      .addItem('Above REL', 'runFRGChecksAllWithUI')
      .addToUi();
}
