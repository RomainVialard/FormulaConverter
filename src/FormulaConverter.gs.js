/****************************************************************
 * FormulaConverter library
 * https://github.com/RomainVialard/FormulaConverter
 *
 * Generate HTML <a> and <img> tags
 * out of =HYPERLINK() & =IMAGE() formulas
 * Also supports arrayFormula() & cell / ranges references
 * But no support for formulas referencing other sheets
 *
 * convertFormulasToHTML()
 *
 * _insertHTML_fromFormula()
 * _cellA1ToIndex()
 * _colA1ToIndex()
 * _rowA1ToIndex()
 * _toImgHtml()
 * _toLinkHtml()
 *****************************************************************/


/**
 * Update a two-dimensional array of sheet values with =HYPERLINK() and =IMAGE() formulas
 * converted to HTML <a> and <img> tags
 * Nothing is returned as the a two-dimensional array of values given as parameter is directly updated
 *
 * @param {{}} param
 * @param {string || SpreadsheetApp.Range} param.range  - the Range passed, if not a string, formulas & values are not mandatory and will be fetched if not provided
 * @param {string[][]} [param.formulas]                 - a two-dimensional array of formulas in string format
 * @param {Array<Array>} [param.values]                 - a two-dimensional array of values
 * @param {number[]} [param.columnsIgnored]             - an array of relative indexes of all columns to skip (no conversion)
 * 
 * @return {object[][]}
 */
function convertFormulasToHTML(param) {
  var formulas = param.formulas;
  var values = param.values;
  var columnsIgnored = param.columnsIgnored;
  var range = param.range;
  
  if (!range) throw new Error(FormulaConverter_.ERROR.INVALID_RANGE);
  
  // Get data / formula / range directly from the spreadsheet range
  if (typeof range !== 'string') {
    !formulas && (formulas = range.getFormulas());
    !values && (values = range.getValues());
    
    range = range.getA1Notation();
  }
  
  var converter = new FormulaConverter_(range, values, formulas, columnsIgnored);
  
  return converter.process();
  
/*
  // Double loop for 2 dimensions array
  for (var i = 0; i < numberOfRows; i++) {
    for (var j = 0; j < numberOfCols; j++) {
      if (columnsIgnored && columnsIgnored_set[j]) continue;
      
      var formula = formulas[i][j];
      
      // If no formula, check if cell begins with http (a valid URL value)
      if (!formula && /^http/.test(values[i][j].toString())) {
        output[i][j] = FormulaConverter_._toLinkHtml(values[i][j]);
        
        continue;
      }
      
      // TODO: check regex (we only want the first parameter of the Image formula)
      var imageFormula = formula.match(/^=(?:arrayformula\(image\((.*?)[,;]?\)\)|image\(['"]?(.*?)[,;]?['"]?\))/i);
      
      if (imageFormula) {
        FormulaConverter_._insertHTML_fromFormula({row: i, col: j, rangeFormula: imageFormula[1] || imageFormula[2]}, values, output);
        
        continue;
      }
      
      
      // TODO: check regex
      // TODO: add HYPERLINK(url, IMAGE(url)) support (build clickable images)
      var hyperLinkFormula = formula.match(/=(?:arrayformula\(HYPERLINK\((.*?)(?:[,;]\s?(.*?))?\)\)|HYPERLINK\(["']*(.*?)["']*(?:[,;]\s?["']*(.*?))?["']*\))/i);
      if (!hyperLinkFormula) continue;
      
      // check if it's a simple hyperlink formula (just 2 strings, no cell reference)
      // in that case, process is much more simple, no need to call FormulaConverter_._insertHTML_fromFormula()
      // TODO: check if we can't use first regex for this as well
      var simpleHyperLink = /=(?:HYPERLINK\(["'](.*?)["'](?:[,;]\s?["'](.*?))?["']\))/i;
      var simple = formula.match(simpleHyperLink);
      
      if (simple) {
        output[i][j] = FormulaConverter_._toLinkHtml(simple[1], simple[2]);
      }
      else {
        FormulaConverter_._insertHTML_fromFormula({
          row: i,
          col: j,
          rangeFormula: hyperLinkFormula[1] || hyperLinkFormula[3],
          label: hyperLinkFormula[2] || hyperLinkFormula[4] || hyperLinkFormula[1] || hyperLinkFormula[3]
        }, values, output);
      }
      
    }
  }
  
  return output;
*/
}


// noinspection JSUnusedGlobalSymbols, ThisExpressionReferencesGlobalObjectJS
this['FormulaConverter'] = {
  // Add local alias to run the library as normal code
  convertFormulasToHTML: convertFormulasToHTML
};


//<editor-fold desc="# Private methods">

/**
 * @namespace FormulaConverter_
 */
/**
 * @typedef {{
 *   firstRow: number,
 *   firstCol: number,
 *   lastRow: number,
 *   lastCol: number,
 *   
 *   nbRows: number,
 *   nbColumns: number,
 *   
 *   [values]: Array<Array>,
 *   [formulas]: Array<Array>,
 *   [rangeA1]: string
 * }} FormulaConverter_.CellRange
 */
/**
 * @typedef {{
 *   row: number,
 *   col: number
 * }} FormulaConverter_.cellA1
 */
/**
 * @typedef {{
 *   rows: number,
 *   cols: number
 * }} FormulaConverter_.DataRange
 */

/**
 * Build helper formula converter on the givenDataRange (accessible data)
 * 
 * @param {string} range
 * @param {Array<Array<string || Object>>} values
 * @param {Array<Array<string>>} formulas
 * @param {Array<number | string>} [ignoredCols] - Exclude columns of the process, number are index relative to the range, string are absolute column labels ('A')
 * 
 * @constructor
 * @private
 */
var FormulaConverter_ = function (range, values, formulas, ignoredCols) {
  
  // Simple sanity check
  if (formulas.length !== values.length || formulas[0].length !== values[0].length) throw new Error(FormulaConverter_.ERROR.RANGES_DONT_MATCH);
  
  // Get range
  // noinspection JSCheckFunctionSignatures
  /**
   * @type {FormulaConverter_.CellRange}
   */
  this.dataRange = this.rangeA1ToIndex(range, {
    values: values,
    formulas: formulas
  });
  
  // init output
  this.output = [];
  for (var i = 0; i < this.dataRange.nbRows; i++) {
    this.output[i] = [];
    
    for (var j = 0; j < this.dataRange.nbColumns; j++) {
      this.output[i][j] = '';
    }
  }
  
  
  // Prepare quick ignored columns check
  this.columnsIgnored_set = {};
  
  ignoredCols && ignoredCols.forEach(function(col){
    this.columnsIgnored_set[
      typeof col === 'number'
        ? col
        : FormulaConverter_._colA1ToIndex(col) - this.dataRange.firstCol
      ] = true;
  }.bind(this));
  
};

/**
 * Convert all IMAGE / HYPERLINK formulas to HTML value
 */
FormulaConverter_.prototype.process = function() {
  
  // Double loop for 2 dimensions array
  for (var j = 0; j < this.dataRange.nbColumns; j++) {
    // Skip ignored columns
    if (this.columnsIgnored_set[j]) continue;
    
    // process all rows
    for (var i = 0; i < this.dataRange.nbRows; i++) {
      // Skip already processed cells
      if (this.output[i][j]) continue;
      
      // If no formula, check if cell begins with http (a valid URL value)
      if (!this.dataRange.formulas[i][j]) {
        if (/^http/.test(this.dataRange.values[i][j])) {
          this.output[i][j] = FormulaConverter_._toLinkHtml(this.dataRange.values[i][j]);
        }
        
        continue;
      }
      
      
      this.findFunction(this.dataRange.values[i][j], (this.dataRange.formulas[i][j] || '').slice(1));
    }
  }
  
  return this.output;
};

FormulaConverter_.prototype.findFunction = function(value, formula) {
  var [/*full match*/, funcName, paramString] = formula.match(/^\s*(\w+)\((.+)\)\s*$/) || [];
  
  var params = FormulaConverter_.extractParam(paramString || '');
  
  
  console.log(funcName, params);
  
};
FormulaConverter_.FUNCTIONS = {
  'hyperlink': true,
  'image': true,
  'arrayformula': true,
};
FormulaConverter_.PARAM_EXTRACT = {
  openers: {
    all: {'"': true, "'": true, "(": true},
    '(': {'"': true, "'": true, "(": true},
    '"': {},
    '': {},
  },
  closers: {
    '(': {')': true},
    '"': {'"': true},
    "'": {"'": true},
  },
  isInString: {
    '(': false,
    '"': true,
    "'": true,
  },
  paramSeparator: {
    ',': true,
    ';': true
  }
};

/**
 * Extract a spreadsheet function parameters from a string
 *
 * @param {string} txt
 *
 * @return {Array<string>}
 */
FormulaConverter_.extractParam = function (txt) {
  var group = [];
  var state = {
    openers: FormulaConverter_.PARAM_EXTRACT.openers.all,
    closers: {},
    inString: false,
    token: ''
  };
  var currentParamIndex = 0;
  var params = [];
  
  for (var i = 0; i < txt.length; i++) {
    var char = txt[i];
    
    // No group, and a comma: it's a parameters we can slice
    if (FormulaConverter_.PARAM_EXTRACT.paramSeparator[char] && group.length === 0) {
      params.push(txt.slice(currentParamIndex, i).trim());
      
      currentParamIndex = i + 1;
      continue;
    }
    
    // Manage opener / closer
    if (state.closers[char]) {
      
      // Detect same quote escaping: "bla""bla" or 'bla''bla'
      if (state.inString && char === state.token && txt[i+1] === state.token){
        // Skip next char
        i++;
        continue;
      }
      
      // remove last token
      group.pop();
      
      // update state
      state.token = group[group.length - 1] || '';
      state.closers = FormulaConverter_.PARAM_EXTRACT.closers[state.token] || {};
      state.openers = FormulaConverter_.PARAM_EXTRACT.openers[state.token] || {};
      
      // Reset string status, as when a group close, we are outside a string
      state.inString = false;
    }
    else if (state.openers[char]) {
      group.push(char);
      
      state = {
        closers: FormulaConverter_.PARAM_EXTRACT.closers[char] || {},
        openers: FormulaConverter_.PARAM_EXTRACT.openers[char] || {},
        inString: FormulaConverter_.PARAM_EXTRACT.isInString[char] || false,
        token: char,
      };
    }
  }
  
  // Add last part
  params.push(txt.slice(currentParamIndex, i).trim());
  
  console.log(params);
  return params;
};



/**
 * Get the URL from the given IMAGE or HYPERLINK formula
 * Handle direct link ("https://..."), cell reference (A1) and range (A1:A)
 *
 * @param {object} obj                - An object with 4 keys
 * @param {string} [obj.label]       - the visible part of an HTML link (link text)
 * @param {string} obj.rangeFormula   - the formula for the current cell / value
 * @param {number} obj.row            - the row index of the current cell / value in the given 2D array
 * @param {number} obj.col            - the column index of the current cell / value in the given 2D array
 *
 * @param {Array<Array>} values         - a two-dimensional array of values
 * @param {Array<Array>} output         - a two-dimensional array of values to modify
 */
FormulaConverter_._insertHTML_fromFormula = function (obj, values, output) {
  // Test if formula makes reference to another cell / range
  // eg: =HYPERLINK(C3)
  if (!/^(?:[a-z]+|[a-z]+\d+|\d+)(?::[a-z]+|:[a-z]+\d+|:\d+)?$/i.test(obj.rangeFormula)) {
    
    // formula makes no reference to another cell / range
    // eg: =HYPERLINK("https://www.google.com/")
    output[obj.row][obj.col] = obj.label
      ? FormulaConverter_._toLinkHtml(obj.rangeFormula, obj.label)
      : FormulaConverter_._toImgHtml(obj.rangeFormula);
    
    return;
  }
  
  
  // Test if reference to a single cell or a range
  if (obj.rangeFormula.indexOf(":") < 0) {
    // reference to single cell
    var range = FormulaConverter_._cellA1ToIndex(obj.rangeFormula);
    
    // if label, it's a link, transform to HTML anchor
    // else it's an image, transform to HTML IMG tag
    output[obj.row][obj.col] = obj.label
      ? FormulaConverter_._toLinkHtml(values[range.row][range.col], values[obj.row][obj.col])
      : FormulaConverter_._toImgHtml(values[range.row][range.col]);
    
    return;
  }
  
  /**
   * @type {FormulaConverter_.DataRange}
   */
  var dataRange = {
    rows: values.length,
    cols: values[0].length
  };
  
  // reference to range
  var rangeUrlData = FormulaConverter_.rangeA1ToIndex(obj.rangeFormula, dataRange);
  var rangeLabelData = obj.label && FormulaConverter_.rangeA1ToIndex(obj.label, dataRange);
  
  Logger.log({
    rangeA1: obj.rangeFormula,
    range: rangeUrlData
  });
  Logger.log({
    rangeA1: obj.label,
    range: rangeLabelData
  });
  
  
  for (var i = 0; i < rangeUrlData.nbRows; i++) {
    for (var j = 0; j < rangeUrlData.nbColumns; j++) {
      var formulaValue = values[rangeUrlData.firstRow + i][ rangeUrlData.firstCol + j];
      
      // TODO: check if this skip images (because of no value if there is an image), for HyperLink(Image())
      if (!formulaValue) continue;
      
      // TODO: check image/link selection validity
      output[obj.row + i][obj.col + j] = obj.label
        ? FormulaConverter_._toLinkHtml(formulaValue, values[rangeLabelData.firstRow + i][rangeLabelData.firstCol + j])
        : FormulaConverter_._toImgHtml(formulaValue);
    }
  }
};


/**
 * Convert a cell reference from A1Notation to 0-based indices (for arrays)
 * or 1-based indices (for Spreadsheet Service methods).
 *
 * @param {string}    cellA1   Cell reference to be converted.
 * @param {number}   [index]   (optional, default 0) Indicate 0 or 1 indexing
 *
 * @return {FormulaConverter_.cellA1} 0-based array coordinate.
 *
 * @throws                     Error if invalid parameter
 */
FormulaConverter_._cellA1ToIndex = function (cellA1, index) {
  // Ensure index is (default) 0 or 1, no other values accepted.
  index = index ? 1 : 0;
  
  // Use regex match to find column & row references.
  // Must start with letters, end with numbers.
  // This regex still allows individuals to provide illegal strings like "AB.#%123"
  // Will accept range like : "A2", "2", "A"
  var [colA1, rowA1] = cellA1.match(/(^[A-Z]+)|([0-9]+$)/gm) || [];
  
  if (colA1 === undefined && rowA1 === undefined) throw FormulaConverter_.ERROR.INVALID_CELL_REFERENCE;
  
  
  var output = {};
  
  rowA1 !== undefined && (output.row = FormulaConverter_._rowA1ToIndex(rowA1, index));
  colA1 !== undefined && (output.col = FormulaConverter_._colA1ToIndex(colA1, index));
  
  return output;
};

/**
 * Return a 0-based array index corresponding to a spreadsheet column
 * label, as in A1 notation.
 *
 * @param {string}    colA1    Column label to be converted.
 * @param {number}   [index]   (optional, default 0) Indicate 0 or 1 indexing
 *
 * @return {number}            0-based array index.
 *
 * @throws                     Error if invalid parameter
 */
FormulaConverter_._colA1ToIndex = function (colA1, index) {
  if (typeof colA1 !== 'string' || colA1.length > 2) throw FormulaConverter_.ERROR.EXPECTED_COLUMN_LABEL;
  
  // Ensure index is (default) 0 or 1, no other values accepted.
  index = index ? 1 : 0;
  
  var A = "A".charCodeAt(0);
  var number = colA1.charCodeAt(colA1.length - 1) - A;
  
  if (colA1.length === 2) {
    number += 26 * (colA1.charCodeAt(0) - A + 1);
  }
  
  return number + index;
};

/**
 * Return a 0-based array index corresponding to a spreadsheet row
 * number, as in A1 notation. Almost pointless, really, but maintains
 * symmetry with FormulaConverter_._colA1ToIndex().
 *
 * @param {number | string}    rowA1    Row number to be converted.
 * @param {number}   [index]   (optional, default 0) Indicate 0 or 1 indexing
 *
 * @return {number}            0-based array index.
 */
FormulaConverter_._rowA1ToIndex = function (rowA1, index) {
  // Ensure index is (default) 0 or 1, no other values accepted.
  index = index ? 1 : 0;
  
  // The "+" will convert rowA1 to number if it's a string
  return +rowA1 - 1 + index;
};

/**
 * Get the boundary of the given range in a1notation
 *
 * For example:
 * A1:A will return for a total Number of row of 10:
 * {
 *   firstRow: 0,
 *   firstCol: 0,
 *   numberOfRows: 10,
 *   numberOfColumns: 1
 * }
 *
 * @param {string} range - The range to process in a1 notation (A1:B5, A1:B, 1:2, C:T)
 * @param {FormulaConverter_.CellRange} initialRange - Data accessible in the sheet
 * 
 * @return {FormulaConverter_.CellRange}
 */
FormulaConverter_.prototype.rangeA1ToIndex = function (range, initialRange) {
  var [firstCellA1, secondCellA1] = range.split(":");
  
  var firstCell = FormulaConverter_._cellA1ToIndex(firstCellA1);
  var lastCell = FormulaConverter_._cellA1ToIndex(secondCellA1);
  
  // the first cell of a range is ALWAYS of the A1 format, or range is A:B and must start at first row
  var firstRow = firstCell.row !== undefined
    ? firstCell.row
    : initialRange || this.dataRange.firstRow === 0
      ? 0
      : undefined;
  
  var firstCol = firstCell.col !== undefined
    ? firstCell.col
    : initialRange || this.dataRange.firstCol === 0
      ? 0
      : undefined;
  
  if (firstRow === undefined || firstCol === undefined) throw FormulaConverter_.ERROR.INVALID_RANGE;
  
  
  var lastRow = lastCell.row !== undefined
    ? lastCell.row
    : initialRange
      ? firstCell.row + initialRange.values.length - 1
      : this.dataRange.firstRow + this.dataRange.nbRows - 1;
  
  var lastCol = lastCell.col !== undefined
    ? lastCell.col
    : initialRange
      ? firstCell.col + initialRange.values[0].length - 1
      : this.dataRange.firstCol + this.dataRange.nbColumns - 1;
  
  // Init output
  var output = initialRange && JSON.parse(JSON.stringify(initialRange)) || {};
  
  output.rangeA1 = range;
  
  output.firstRow = Math.min(firstRow, lastRow);
  output.firstCol = Math.min(firstCol, lastCol);
  output.lastRow = Math.max(firstRow, lastRow);
  output.lastCol = Math.max(firstCol, lastCol);
  
  output.nbRows = output.lastRow - output.firstRow + 1;
  output.nbColumns = output.lastCol - output.firstCol + 1;
  
  return output;
};

FormulaConverter_.ERROR = {
  INVALID_RANGE: 'Invalid Range',
  INVALID_CELL_REFERENCE: 'Invalid cell reference',
  EXPECTED_COLUMN_LABEL: 'Expected column label',
  RANGES_DONT_MATCH: 'Ranges do not match'
};


/**
 * Return an html img tag built from the given url
 *
 * @param {string} url
 *
 * @return {string}
 */
FormulaConverter_._toImgHtml = function (url) {
  url = url || '';
  
  // if sheet contains 2 columns, one with raw url to image and one with the image formula placed after
  // the url in first column will be replaced with an HTML anchor
  // And then we will try to add an additional img tag, which will break
  // So check if there's already an HTML anchor and remove it if it's the case
  url = (url.match(/href="(.*?)"/) || [])[1] || url;
  
  return '<img style="max-width:100%" src="'+ url +'"/>';
};

/**
 * Return an html anchor tag built from the given url and label
 *
 * @param {string} url
 * @param {string} [label]
 *
 * @return {string}
 */
FormulaConverter_._toLinkHtml = function (url, label) {
  return '<a href="'+ url +'">'+ (label || url) +'</a>';
};

//</editor-fold>

// LOCAL TEST
function test() {
  console.log('START');
  
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
  
  // test get data bound
  var converter = new FormulaConverter_(param.range, param.values, param.formulas);
  
  var res = converter.process();
  
  console.log(res);
  
}

test();



// extractParam(`"azert", "zaert", "qdsfgh"`);
// extractParam(`"aze,rt", "zaert", "qdsfgh"`);
// extractParam(`IMAGE("aze,rt", "zaert"), "qdsfgh"`);
// FormulaConverter_.extractParam(`IMAGE("az""e,rt"; "zaert"); "qdsfgh", 14`);

