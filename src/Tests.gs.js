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
  
  
  var range = sheetTests.getRange('B2:C');
  var rangeOutput = sheetTests.getRange('E2:E');
  
  
  var processedHyperlinkValues = FormulaConverter.convertFormulasToHTML({
    range: range.getA1Notation(),
    values: range.getValues(),
    formulas: range.getFormulas()
  });
  
  // select only output link tests
  var output = [];
  for (var i = 0; i < processedHyperlinkValues.length; i++){
    output.push([processedHyperlinkValues[i][1]]);
  }
  
  rangeOutput.setValues(output);
}


function testIndividualFn() {
  
  var param = {
    range: 'C2:C8',
    totalRows: 16
  };
  
  var output = FormulaConverter_._getBoundRange(param.range, param.totalRows);
  
  Logger.log(JSON.stringify({
    input: param,
    output: output
  }, null, '\t'));
  
}
