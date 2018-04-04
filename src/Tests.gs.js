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







// LOCAL TEST
function test_local() {
  console.log('START');
  
  // noinspection JSUnusedLocalSymbols
  var _param = {
    values: [
      ["Simple link (no formula)", "http://www.ikea.com/us/en/images/products/tjena-box-with-lid-green__0321624_PE515923_S4.JPG"],
      ["Simple HYPERLINK", "http://www.ikea.com/us/en/images/products/micke-desk-white__0324519_PE517088_S4.JPG"],
      ["HYPERLINK with link label", "Bouh"],
      ["HYPERLINK with cell ref for url", "Bouh"],
      ["HYPERLINK with 2 cells ref", "Simple link (no formula)"],
      ["ARRAYFORMULA + HYPERLINK", "Bouh"],
      ["idem", "Simple link (no formula)"],
      ["Simple IMAGE", ""],
      ["IMAGE with cell ref for url", ""],
      ["ARRAYFORMULA + IMAGE", ""],
      ["idem", ""],
      ["Simple HYPERLINK + IMAGE", ""],
      ["HYPERLINK + IMAGE with cell ref", ""],
      ["HYPERLINK + IMAGE with ARRAYFORMULA", ""],
      ["idem", ""]
    ],
    formulas: [
      ["", ""],
      ["", "=HYPERLINK(\"http://www.ikea.com/us/en/images/products/micke-desk-white__0324519_PE517088_S4.JPG\")"],
      ["", "=HYPERLINK(\"http://www.ikea.com/us/en/images/products/micke-desk-white__0324519_PE517088_S4.JPG\", \"Bouh\")"],
      ["", "=HYPERLINK(C2, \"Bouh\")"],
      ["", "=HYPERLINK(C2, B2)"],
      ["", "=ARRAYFORMULA(HYPERLINK(C2:C3, C5:C6))"],
      ["", ""],
      ["", "=IMAGE(\"http://www.ikea.com/us/en/images/products/micke-desk-white__0324519_PE517088_S4.JPG\")"],
      ["", "=IMAGE(C2)"],
      ["", "=ARRAYFORMULA(IMAGE(C2:C3))"],
      ["", ""],
      ["", "=HYPERLINK(\"http://www.ikea.com/us/en/images/products/tjena-box-with-lid-green__0321624_PE515923_S4.JPG\", IMAGE(\"http://www.ikea.com/us/en/images/products/tjena-box-with-lid-green__0321624_PE515923_S4.JPG\"))"],
      ["", "=HYPERLINK(C2, IMAGE(C2))"],
      ["", "=ARRAYFORMULA(HYPERLINK(C2:C3, IMAGE(C2:C3)))"],
      ["", ""]
    ],
    range: 'B2:C',
  };
  var param = {
    values: [
      ["Simple link (no formula)", "http://www.ikea.com/us/en/images/products/tjena-box-with-lid-green__0321624_PE515923_S4.JPG"],
      ["Simple HYPERLINK", "http://www.ikea.com/us/en/images/products/micke-desk-white__0324519_PE517088_S4.JPG"],
      ["HYPERLINK with link label", "Bouh"],
      ["HYPERLINK with cell ref for url", "Bouh"],
      ["HYPERLINK with 2 cells ref", "Simple link (no formula)"],
      ["ARRAYFORMULA + HYPERLINK", "Bouh"],
      ["idem", "Simple link (no formula)"],
      ["Simple IMAGE", ""],
      ["IMAGE with cell ref for url", ""],
      ["ARRAYFORMULA + IMAGE", ""],
      ["idem", ""],
      ["Simple HYPERLINK + IMAGE", ""],
      ["HYPERLINK + IMAGE with cell ref", ""],
      ["HYPERLINK + IMAGE with ARRAYFORMULA", ""],
      ["idem", ""],
    ],
    formulas: [
      ["", ""],
      ["", ""],
      ["", ""],
      ["", ""],
      ["", ""],
      ["", ""],
      ["", ""],
      ["", ""],
      ["", ""],
      ["", ""],
      ["", ""],
      ["", ""],
      ["", ""],
      ["", "=ARRAYFORMULA(HYPERLINK(C2:C3, IMAGE(C2:C3)))"],
      ["", ""],
    ],
    range: 'B2:C',
  };
  
  // test get data bound
  var converter = new FormulaConverter_(param.range, param.values, param.formulas);
  
  var res = converter.process();
  
  console.log('\n########### RESULTS ##########\n');
  console.log(res);
  
}

// test();



// extractParam(`"azert", "zaert", "qdsfgh"`);
// extractParam(`"aze,rt", "zaert", "qdsfgh"`);
// extractParam(`IMAGE("aze,rt", "zaert"), "qdsfgh"`);
// FormulaConverter_.extractParam(`IMAGE("az""e,rt"; "zaert"); "qdsfgh", 14`);

