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
 * _getLinkFromFormula()
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
 * @param {string[][]} formulas - a two-dimensional array of formulas in string format
 * @param {Array<Array>} values - a two-dimensional array of values
 * @param {number[]} columnsIgnored - an array of indexes of all columns to skip (no conversion)
 */
function convertFormulasToHTML(formulas, values, columnsIgnored) {
  var rowNumber = formulas.length;
  var colNumber = formulas.length;
  
  // Simple sanity check
  if (rowNumber !== values.length || colNumber !== values[0].length) throw new Error("Ranges do not match");
  
  // Prepare quick ignored columns check
  var columnsIgnored_set = {};
  columnsIgnored && columnsIgnored.forEach(function(col){
    columnsIgnored_set[col] = true;
  });
  
  
  // Double loop for 2 dimensions array
  for (var i = 0; i < rowNumber; i++) {
    for (var j = 0; j < colNumber; j++) {
      if (columnsIgnored && columnsIgnored_set[j]) continue;
      
      var formula = formulas[i][j];
      
      // If no formula, check if cell begins with http (a valid URL value)
      if (!formula && /^http/.test(values[i][j].toString())) {
        values[i][j] = FormulaConverter_._toLinkHtml(values[i][j]);
        
        continue;
      }
      
      // TODO: check regex
      var imageFormula = formula.match(/^=(?:arrayformula\(image\((.*?)[,;]?\)\)|image\(['"]?(.*?)[,;]?['"]?\))/i);
      
      if (imageFormula) {
        FormulaConverter_._getLinkFromFormula({row: i, col: j, rangeFormula: imageFormula[1] || imageFormula[2]}, values);
        
        continue;
      }
      
      
      // TODO: check regex
      var hyperLinkFormula = formula.match(/=(?:arrayformula\(HYPERLINK\((.*?)(?:[,;]\s?(.*?))?\)\)|HYPERLINK\(["']*(.*?)["']*(?:[,;]\s?["']*(.*?))?["']*\))/i);
      
      if (hyperLinkFormula) {
        
        // check if it's a simple hyperlink formula (just 2 strings, no cell reference)
        // in that case, process is much more simple, no need to call FormulaConverter_._getLinkFromFormula()
        // TODO: check if we can't use first regex for this as well
        var simpleHyperLink = /=(?:HYPERLINK\(["'](.*?)["'](?:[,;]\s?["'](.*?))?["']\))/i;
        var simple = formula.match(simpleHyperLink);
        
        if (simple) {
          values[i][j] = FormulaConverter_._toLinkHtml(simple[1], simple[2]);
        }
        else {
          FormulaConverter_._getLinkFromFormula({
            row: i,
            col: j,
            rangeFormula: hyperLinkFormula[1] || hyperLinkFormula[3],
            linkText: hyperLinkFormula[2] || hyperLinkFormula[4] || hyperLinkFormula[1] || hyperLinkFormula[3]
          }, values);
        }
      }
      
    }
  }
  
}


// noinspection JSUnusedGlobalSymbols, ThisExpressionReferencesGlobalObjectJS
this['FormulaConverter'] = {
  // Add local alias to run the library as normal code
  convertFormulasToHTML: convertFormulasToHTML
};


//<editor-fold desc="# Private methods">

var FormulaConverter_ = {};


/**
 * Get the URL from the given IMAGE or HYPERLINK formula
 * Handle direct link ("https://..."), cell reference (A1) and range (A1:A)
 *
 * @param {object} obj                - An object with 4 keys
 * @param {string} [obj.linkText]       - the visible part of an HTML link (link text)
 * @param {string} obj.rangeFormula   - the formula for the current cell / value
 * @param {number} obj.row            - the row index of the current cell / value in the given 2D array
 * @param {number} obj.col            - the column index of the current cell / value in the given 2D array
 *
 * @param {Array<Array>} values         - a two-dimensional array of values
 */
FormulaConverter_._getLinkFromFormula = function (obj, values) {
  // Test if formula makes reference to another cell / range
  // eg: =HYPERLINK(C3)
  if (!/^(?:[a-z]+|[a-z]+\d+|\d+)(?::[a-z]+|:[a-z]+\d+|:\d+)?$/i.test(obj.rangeFormula)) {
    
    // formula makes no reference to another cell / range
    // eg: =HYPERLINK("https://www.google.com/")
    values[obj.row][obj.col] = obj.linkText
      ? FormulaConverter_._toLinkHtml(obj.rangeFormula, obj.linkText)
      : FormulaConverter_._toImgHtml(obj.rangeFormula);
    
    return;
  }
  
  
  // Test if reference to a single cell or a range
  if (obj.rangeFormula.indexOf(":") < 0) {
    // reference to single cell
    var range = FormulaConverter_._cellA1ToIndex(obj.rangeFormula);
    
    // if linkText, it's a link, transform to HTML anchor
    // else it's an image, transform to HTML IMG tag
    values[obj.row][obj.col] = obj.linkText
      ? FormulaConverter_._toLinkHtml(values[range.row][range.col], values[obj.row][obj.col])
      : FormulaConverter_._toImgHtml(values[range.row][range.col]);
    
    return;
  }
  
  // reference to range
  var rangeData = FormulaConverter_._getBoundRange(obj.rangeFormula, values.length);
  
  for (var i = rangeData.firstRow; i < rangeData.numberOfRows; i++) {
    for (var j = rangeData.firstCol; j < rangeData.numberOfColumns; j++) {
      if (!values[i][j]) continue;
      
      // TODO: CHECK VALIDITY of LOGIC here. (obj.row + i does not make sense)
      values[obj.row + i][obj.col + j] = obj.linkText
        ? FormulaConverter_._toLinkHtml(values[i][j], values[obj.row + i][obj.col + j])
        : FormulaConverter_._toImgHtml(values[i][j]);
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
 * @return {object}            {row,col}, both 0-based array indices.
 *
 * @throws                     Error if invalid parameter
 */
FormulaConverter_._cellA1ToIndex = function (cellA1, index) {
  // Ensure index is (default) 0 or 1, no other values accepted.
  index = index ? 1 : 0;
  
  // Use regex match to find column & row references.
  // Must start with letters, end with numbers.
  // This regex still allows individuals to provide illegal strings like "AB.#%123"
  var [colA1, rowA1] = cellA1.match(/(^[A-Z]+)|([0-9]+$)/gm) || [];
  
  if (colA1 === undefined && rowA1 === undefined) throw new Error("Invalid cell reference");
  
  return {
    row: FormulaConverter_._rowA1ToIndex(rowA1, index),
    col: FormulaConverter_._colA1ToIndex(colA1, index)
  };
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
  if (typeof colA1 !== 'string' || colA1.length > 2) throw new Error("Expected column label.");
  
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
 * @param {string} range - The range to process in a1 notation
 * @param {number} totalRow - the number of rows of the sheet
 */
FormulaConverter_._getBoundRange = function (range, totalRow) {
  var ranges = range.split(":");
  var firstCell = FormulaConverter_._cellA1ToIndex(ranges[0]);
  
  // TODO: check logic here
  var lastCell = FormulaConverter_._cellA1ToIndex(ranges[1] + (ranges[1].length > 1 ? '' : firstCell.row + 1));
  
  
  return {
    firstRow: firstCell.row,
    firstCol: firstCell.col,
    numberOfRows: totalRow - firstCell.row,
    numberOfColumns: lastCell.col + 1 - firstCell.col
  };
};


/**
 * Return an html img tag built from the given url
 *
 * @param {string} url
 *
 * @return {string}
 */
FormulaConverter_._toImgHtml = function (url) {
  // if sheet contains 2 columns, one with raw url to image and one with the image formula placed after
  // the url in first column will be replaced with an HTML anchor
  // And then we will try to add an additional img tag, which will break
  // So check if there's already an HTML anchor and remove it if it's the case
  url = (url.match(/href="(.*?)"/) || [])[1] || url;
  
  return '<img style="max-width:100%" src="' + url + '"/>';
};

/**
 * Return an html anchor tag built from the given url and linkText
 *
 * @param {string} url
 * @param {string} [linkText]
 *
 * @return {string}
 */
FormulaConverter_._toLinkHtml = function (url, linkText) {
  return '<a href="' + url + '">' + (linkText || url) + '</a>';
};

//</editor-fold>
