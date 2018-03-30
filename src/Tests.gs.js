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
  
  
  var range = sheetTests.getRange('A1:C12');
  var rangeOutput = sheetTests.getRange('E2:E12');
  
  
  var processedHyperlinkValues = FormulaConverter.convertFormulasToHTML(range.getFormulas(), range.getValues());
  
  // select only output link tests
  var output = [];
  for (var i = 1; i < processedHyperlinkValues.length; i++){
    output.push([processedHyperlinkValues[i][2]]);
  }
  
  rangeOutput.setValues(output);
}
