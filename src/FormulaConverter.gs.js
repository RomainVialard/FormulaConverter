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
 * @typedef {{fn: Function, param: Array<{name: string, type: 'CELL' | 'FORMULA'}>}} FormulaConverter_.FUNCTION
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
  this.dataRange = new FormulaConverter_.CellRange(range, {
    values: values,
    formulas: formulas,
  });
  
  
  // init processed cells
  this._processed = [];
  for (var i = 0; i < this.dataRange.nbRows; i++) {
    this._processed[i] = [];
    
    for (var j = 0; j < this.dataRange.nbColumns; j++) {
      this._processed[i][j] = false;
    }
  }
  
  // Clone values to init output
  this.output = JSON.parse(JSON.stringify(this._processed/*values*/));
  
  
  // Prepare quick ignored columns check
  this.columnsIgnored_set = {};
  
  ignoredCols && ignoredCols.forEach(function(col){
    this.columnsIgnored_set[
      typeof col === 'number'
        ? col
        : FormulaConverter_._colA1ToIndex(col) - this.dataRange.firstCol
      ] = true;
  }.bind(this));
  
  // Init SPS functions
  /**
   * @type {Object<FormulaConverter_.FUNCTION>}
   */
  this.FUNCTIONS = {
    hyperlink: {
      fn: this._SPS_FUNCTION_hyperlink,
      param: [
        {name: 'url', type: 'CELL'},
        {name: 'label', type: 'CELL'},
      ]
    },
    image: {
      fn: this._SPS_FUNCTION_image,
      param: [
        {name: 'url', type: 'CELL'},
      ]
    },
    arrayformula: {
      fn: this._SPS_FUNCTION_arrayformula,
      param: [
        {name: 'formula', type: 'FORMULA'},
      ]
    },
  };
  
  // noinspection JSUnusedGlobalSymbols
  this._arrayFormula = null;
  
};

/**
 * Convert all IMAGE / HYPERLINK formulas to HTML value
 */
FormulaConverter_.prototype.process = function () {
  
  // Double loop for 2 dimensions array
  for (var j = 0; j < this.dataRange.nbColumns; j++) {
    // Skip ignored columns
    if (this.columnsIgnored_set[j]) continue;
    
    // process all rows
    for (var i = 0; i < this.dataRange.nbRows; i++) {
      // Skip already processed cells
      if (this._processed[i][j]) continue;
      
      // If no formula, check if cell begins with http (a valid URL value)
      if (!this.dataRange.formulas[i][j]) {
        if (/^http/.test(this.dataRange.values[i][j])) {
          this.output[i][j] = FormulaConverter_._toLinkHtml(this.dataRange.values[i][j]);
        }
        
        this._processed[i][j] = true;
        continue;
      }
      
      var res;
      try {
        // noinspection JSUnusedGlobalSymbols
        this._arrayFormula = null;
        res = this._findFunction((this.dataRange.formulas[i][j] || '').slice(1), this.dataRange.values[i][j]);
      }
      catch (e) {
        res = '#ERROR!';
      }
      
      // Apply Ranged results
      if (Array.isArray(res)){
        
        var nbRows = res.length;
        var nbCols = res[0].length;
        
        for (var offset_j = 0; offset_j < nbCols; offset_j++) {
          var absolute_j = j + offset_j;
          
          // Skip ignored columns
          if (this.columnsIgnored_set[absolute_j]) continue;
          
          for (var offset_i = 0; offset_i < nbRows; offset_i++) {
            var absolute_i = i + offset_i;
            
            // Skip already processed cells
            if (this._processed[absolute_i][absolute_j]) continue;
            
            var result = res[offset_i][offset_j];
            if (result !== undefined) {
              // Test if result is an URL
              /^http/.test(result) && (result = FormulaConverter_._toLinkHtml(result));
              
              this.output[absolute_i][absolute_j] = result;
            }
            
            this._processed[absolute_i][absolute_j] = true;
          }
        }
        
        continue;
      }
      
      if (res) {
        // Test if result is an URL
        /^http/.test(res) && (res = FormulaConverter_._toLinkHtml(res));
        
        // Store result if it is a value
        this.output[i][j] = res;
      }
      
      
      this._processed[i][j] = true;
    }
  }
  
  return this.output;
};


/**
 * Start formula parsing
 *
 * @param {string} formula
 * @param {string} [value]
 *
 * @return {*|boolean|string}
 */
FormulaConverter_.prototype._findFunction = function(formula, value) {
  var [/*full match*/, funcName, paramString] = formula.match(/^\s*(\w+)\((.+)\)\s*$/) || [];
  
  // Clean function and its parameters
  var params = FormulaConverter_._extractParam(paramString || '');
  funcName = (funcName || '').toLowerCase();
  
  // get corresponding Sps function
  var func = this.FUNCTIONS[funcName] || false;
  
  
  // Apply function
  return func
    ? this._applyFunction(func, params)
    : value;
};

/**
 * Return the value for either a quote surrounded string, or a A1 cell reference
 *
 * @param {FormulaConverter_.FUNCTION} func
 * @param {Array<string>} params
 *
 * @private
 */
FormulaConverter_.prototype._applyFunction = function (func, params) {
  var resolvedParams = [];
  var applyArrayFormula_indexes = [];
  var applyArrayFormula = false;
  
  for (var i = 0; i < params.length; i++) {
    var [/*full match*/, value] = params[i].match(/^['"](.*)['"]$/) || [];
    
    // Text value
    if (value !== undefined){
      resolvedParams.push(value);
      continue;
    }
    
    // Is it a cell reference ?
    if (/^[A-Z]+\d+$/.test(params[i])){
      resolvedParams.push(this._getA1CellValue(params[i]));
      continue;
    }
    
    // Pass formula to function if it's the param type
    if (func.param[i].type === 'FORMULA') {
      resolvedParams.push(params[i]);
      continue;
    }
    
    // Test for range / formula
    if (/^[A-Z]+\d+:[A-Z]*\d*$/.test(params[i])){
      
      // Type is a CELL here, so it can not take a range
      if (!this._arrayFormula) throw FormulaConverter_.ERROR.INVALID_CELL_REFERENCE;
      
      var range = new FormulaConverter_.CellRange(params[i]);
      
      // Check if it's the first 'defining' range for this arrayFormula
      !this._arrayFormula.range && (this._arrayFormula.range = range);
      
      // Check that this range are equals to the arrayFormula bounds
      if (!this._arrayFormula.range.hasSameSize(range)) throw FormulaConverter_.ERROR.INVALID_CELL_REFERENCE;
      
      applyArrayFormula_indexes.push(i);
      resolvedParams.push(range);
      applyArrayFormula = true;
      continue;
    }
    
    // It's a formula, process it
    resolvedParams.push(this._findFunction(params[i], undefined));
  }
  
  // Simple resolution
  if (!applyArrayFormula) return func.fn.apply(this, resolvedParams);
  
  // ArrayFormula resolution: return an Array<Array>
  var af_params = resolvedParams.slice(0);
  var output = [];
  
  for (var row = 0; row < this._arrayFormula.range.nbRows; row++) {
    output[row] = [];
    
    for (var col = 0; col < this._arrayFormula.range.nbColumns; col++) {
      
      // Get range current cell
      for (var index = 0; index < applyArrayFormula_indexes.length; index++) {
        af_params[ applyArrayFormula_indexes[index] ] = resolvedParams[ applyArrayFormula_indexes[index] ].getValue(row, col);
      }
      
      // Apply function on current cell
      output[row][col] = func.fn.apply(this, af_params);
    }
  }
  
  console.log('output', output);
  
  return output;
};

/**
 * Get value in the data array by A1 cell notation
 *
 * @param {string} A1
 *
 * @return {*}
 * @private
 */
FormulaConverter_.prototype._getA1CellValue = function (A1) {
  var cellRef = FormulaConverter_._cellA1ToIndex(A1);
  
  if (cellRef.col === undefined && cellRef.row === undefined) throw FormulaConverter_.ERROR.INVALID_CELL_REFERENCE;
  
  return this.dataRange.values[cellRef.row - this.dataRange.firstRow][cellRef.col - this.dataRange.firstCol];
};


//<editor-fold desc="# SPREADSHEET functions">

/**
 * Apply the Hyperlink function
 *
 * @param {string} url
 * @param {string} [label]
 *
 * @private
 */
FormulaConverter_.prototype._SPS_FUNCTION_hyperlink = function (url, label) {
  return FormulaConverter_._toLinkHtml(url, label || '');
};

/**
 * Apply the Image function
 *
 * @param {string} url
 *
 * @private
 */
FormulaConverter_.prototype._SPS_FUNCTION_image = function (url) {
  return FormulaConverter_._toImgHtml(url);
};

/**
 * Apply the ArrayFormula function
 *
 * @param {string} formula
 *
 * @private
 */
FormulaConverter_.prototype._SPS_FUNCTION_arrayformula = function (formula) {
  console.log('ARRAYFORMULA', formula);
  
  /**
   * @type {{range: FormulaConverter_.CellRange}}
   */
  this._arrayFormula = {
    range: undefined,
  };
  
  return this._findFunction(formula, undefined);
};

//</editor-fold>


/**
 * Extract a spreadsheet function parameters from a string
 *
 * @param {string} txt
 *
 * @return {Array<string>}
 */
FormulaConverter_._extractParam = function (txt) {
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
  
  return params;
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


//<editor-fold desc="# CellRange">

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
 * @param {FormulaConverter_.CellRange} [initialRange] - Data accessible in the sheet
 */
FormulaConverter_.CellRange = function (range, initialRange) {
  // Get global range
  if (FormulaConverter_.CellRange.DataRange) {
    this.dataRange = FormulaConverter_.CellRange.DataRange;
  }
  
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
  
  
  if (initialRange) {
    this.values = JSON.parse(JSON.stringify(initialRange.values));
    this.formulas = JSON.parse(JSON.stringify(initialRange.formulas));
    
    // Store global dataRange for later range determination
    FormulaConverter_.CellRange.DataRange = this;
  }
  
  this.rangeA1 = range;
  
  this.firstRow = Math.min(firstRow, lastRow);
  this.firstCol = Math.min(firstCol, lastCol);
  this.lastRow = Math.max(firstRow, lastRow);
  this.lastCol = Math.max(firstCol, lastCol);
  
  this.nbRows = this.lastRow - this.firstRow + 1;
  this.nbColumns = this.lastCol - this.firstCol + 1;
};

/**
 * Test if a CellRange got the same bounds as the current one
 *
 * @param {FormulaConverter_.CellRange} cellRange
 *
 * @return {boolean}
 */
FormulaConverter_.CellRange.prototype.hasSameSize = function(cellRange) {
  return this.nbRows === cellRange.nbRows && this.nbColumns === cellRange.nbColumns;
};

/**
 * Get cell value at given offset
 *
 * @param {number} row - relative 0-based index row
 * @param {number} col - relative 0-based index column
 *
 * @return {*}
 */
FormulaConverter_.CellRange.prototype.getValue = function(row, col) {
  // Sanity check
  if (row < 0 || col < 0 || row >= this.dataRange.nbRows || col >= this.dataRange.nbColumns) {
    throw FormulaConverter_.ERROR.INVALID_CELL_REFERENCE;
  }
  
  return this.dataRange.values[this.dataRange.firstRow + row][this.dataRange.firstCol + col];
};




/**
 * Test if something is a CellRange
 *
 * @param {*} val
 *
 * @return {boolean}
 */
FormulaConverter_.isCellRange = function (val) {
  return val instanceof FormulaConverter_.CellRange;
};

//</editor-fold>





// LOCAL TEST
function test() {
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
      ["", "=ARRAYFORMULA(HYPERLINK(C2:C3, C5:C6))"],
      ["", ""],
      ["", ""],
      ["", ""],
      ["", ""],
      ["", ""],
      ["", ""],
      ["", ""],
      ["", ""],
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

test();



// extractParam(`"azert", "zaert", "qdsfgh"`);
// extractParam(`"aze,rt", "zaert", "qdsfgh"`);
// extractParam(`IMAGE("aze,rt", "zaert"), "qdsfgh"`);
// FormulaConverter_.extractParam(`IMAGE("az""e,rt"; "zaert"); "qdsfgh", 14`);

