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
  
  var currentSheetName = sps.getActiveSheet().getName();
  
  var testList = {
    'validation': {
      sheet: 'validation',
      rangeInput: 'B2:C',
      cellOutput: 'E2',
      res: {
        offset: 0,
        columns: [1],
      } 
    },
    'test YAMM': {
      sheet: 'test YAMM',
      rangeInput: 'A1:F',
      cellOutput: 'H2',
      res: {
        offset: 1,
        columns: [0, 1, 2, 3, 4, 5],
      }
    },
    'isolatedTest': {
      sheet: 'isolatedTest',
      rangeInput: 'B4:B5',
      cellOutput: 'D4',
      res: {
        offset: 0,
        columns: [0],
      }
    },
  };
  
  var currentTest = testList[currentSheetName] || testList.validation;
  
  var sheetTests = sps.getSheetByName(currentTest.sheet);
  var range = sheetTests.getRange(currentTest.rangeInput);
  var rangeOutput = sheetTests.getRange(currentTest.cellOutput);
  
  
  var processedHyperlinkValues = FormulaConverter.convertFormulasToHTML({
    range: range.getA1Notation(),
    values: range.getValues(),
    formulas: range.getFormulas()
  });
  
  // select only output link tests
  var output = [];
  for (var i = currentTest.res.offset; i < processedHyperlinkValues.length; i++){
    var row = [];
    
    for (var j = 0; j < currentTest.res.columns.length; j++){
      row.push(processedHyperlinkValues[i][ currentTest.res.columns[j] ]);
    }
    
    output.push(row);
  }
  
  // write results
  sheetTests
    .getRange(rangeOutput.getRow(), rangeOutput.getColumn(), output.length, output[0].length)
    .setValues(output);
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
  // noinspection JSUnusedLocalSymbols
  var __param = {
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
  var param = {
    values: [
      [""],
      [""],
    ],
    formulas: [
      ["=HYPERLINK(\"https://sites.google.com/a/revevol.eu/images-base/hurdal-cadre-de-lit-boites-de-rangement-brun__0255068_PE399082_S4.JPG\", \"TJENA - Box with lid green\")"],
      ["=HYPERLINK(\"https://sites.google.com/a/revevol.eu/images-base/hurdal-cadre-de-lit-boites-de-rangement-brun__0255068_PE399082_S4.JPG\", \"Bed frame with storage, white\")"],
    ],
    range: 'B4:B5',
  };
  
  // test get data bound
  var converter = new FormulaConverter_(param.range, param.values, param.formulas);
  
  var res = converter.process();
  
  console.log('\n########### RESULTS ##########\n');
  console.log(res);
  
}

// test_local();



/* Test param extraction
FormulaConverter_._extractParam(`"https://sites.google.com/a/revevol.eu/images-base/hurdal-cadre-de-lit-boites-de-rangement-brun__0255068_PE399082_S4.JPG", "Bed frame with storage, white"`);
FormulaConverter_._extractParam(`"aze,rt", "zaert", "qd,sfgh"`);
FormulaConverter_._extractParam(`"azert", "zaert", "qdsfgh"`);
FormulaConverter_._extractParam(`"aze,rt", "zaert", "qdsfgh"`);
FormulaConverter_._extractParam(`IMAGE("aze,rt", "zae,rt"), "qdsf,gh"`);
FormulaConverter_._extractParam(`IMAGE("az""e,rt"; "zaert"); ,"qdsfgh", 14`);
*/
