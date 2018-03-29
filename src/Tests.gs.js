// noinspection JSUnusedGlobalSymbols
/**
 * Add menu to launch the tests
 */
function onOpen() {
  
  SpreadsheetApp.getUi()
    .createMenu('Test FormulaConverter')
    .addItem('Run tests', 'tests')
    .addToUi();
  
}

/**
 * run tests
 * 
 * @OnlyCurrentDoc
 */
function tests() {
  
  var sps = SpreadsheetApp.getActiveSpreadsheet();
  var sheetTests = sps.getSheetByName('tests');
  
  
  // Hyperlinks
  var rangeHyperLink = sheetTests.getRange('B2:C8');
  var rangeHyperLinkOutput = sheetTests.getRange('E2:F8');
  
  
  var processedHyperlinkValues = FormulaConverter.convertFormulasToHTML(rangeHyperLink.getFormulas(), rangeHyperLink.getValues());
  
  rangeHyperLinkOutput.setValues(processedHyperlinkValues);
  
}