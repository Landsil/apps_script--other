/**
 * Calculates MD2 of a string or cell (or concatenation of cells)
 *
 * @param {string} input_string Input string or cell
 * @return Calculates MD2 of a string or cell (or concatenation of cells)
 * @customfunction
 *
 */
function MD2(input_string) {
  var hexstr = '';
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD2, input_string);
  for (i = 0; i < digest.length; i++) {
    var val = (digest[i]+256) % 256;
    hexstr += ('0'+val.toString(16)).slice(-2);
  }
  return hexstr;
}


/**
 * Calculates MD5 of a string or cell (or concatenation of cells)
 *
 * @param {string} input_string Input string or cell
 * @return Calculates MD5 of a string or cell (or concatenation of cells)
 * @customfunction
 *
 */
function MD5(input_string) {
  var hexstr = '';
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, input_string);
  for (i = 0; i < digest.length; i++) {
    var val = (digest[i]+256) % 256;
    hexstr += ('0'+val.toString(16)).slice(-2);
  }
  return hexstr;
}


/**
 * Calculates SHA1 of a string or cell (or concatenation of cells)
 *
 * @param {string} input_string Input string or cell
 * @return Calculates SHA1 of a string or cell (or concatenation of cells)
 * @customfunction
 *
 */
function SHA1(input_string) {
  var hexstr = '';
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_1, input_string)
  for (i = 0; i < digest.length; i++) {
    var val = (digest[i]+256) % 256;
    hexstr += ('0'+val.toString(16)).slice(-2);
  }
  return hexstr;
}


/**
 * Calculates SHA256 of a string or cell (or concatenation of cells)
 *
 * @param {string} input_string Input string or cell
 * @return Calculates SHA256 of a string or cell (or concatenation of cells)
 * @customfunction
 *
 */
function SHA256(input_string) {
  var hexstr = '';
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input_string);
  for (i = 0; i < digest.length; i++) {
    var val = (digest[i]+256) % 256;
    hexstr += ('0'+val.toString(16)).slice(-2);
  }
  return hexstr;
}


/**
 * Calculates SHA384 of a string or cell (or concatenation of cells)
 *
 * @param {string} input_string Input string or cell
 * @return Calculates SHA384 of a string or cell (or concatenation of cells)
 * @customfunction
 *
 */
function SHA384(input_string) {
  var hexstr = '';
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_384, input_string);
  for (i = 0; i < digest.length; i++) {
    var val = (digest[i]+256) % 256;
    hexstr += ('0'+val.toString(16)).slice(-2);
  }
  return hexstr;
}


/**
 * Calculates SHA512 of a string or cell (or concatenation of cells)
 *
 * @param {string} input_string Input string or cell
 * @return Calculates SHA512 of a string or cell (or concatenation of cells)
 * @customfunction
 *
 */
function SHA512(input_string) {
  var hexstr = '';
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_512, input_string);
  for (i = 0; i < digest.length; i++) {
    var val = (digest[i]+256) % 256;
    hexstr += ('0'+val.toString(16)).slice(-2);
  }
  return hexstr;
}


/**
 * Function to calculate percent change
 *
 * @param {number} oldVal Original Value
 * @param {number} newVal New Value
 * @return The percent change between new and old values.
 * @customfunction
 *
 */
function PercentChange(oldVal, newVal) {
  
  return (newVal - oldVal) / oldVal; 
  
}

// Indentation function
var ss = SpreadsheetApp.getActiveSpreadsheet();

function moveText(direction) {
  var values = ss.getActiveRange().getValues();
  var cols = ss.getActiveRange().getNumColumns();
  var rows = ss.getActiveRange().getNumRows();

  var newValues = new Array();

  for (x = 1; x <= rows; x++) {
    for (y = 1; y <= cols; y++) {
      var cell = ss.getActiveRange().getCell(x, y);
      var value = cell.getValue();
      var formula = function() {
        if (direction == ">>>>>") {
          return  '=CONCAT(REPT( CHAR( 160 ), 5),"' + value + '")';
        } else if (direction == "<<<<<") {
          return '=IF(TRIM(LEFT("' + value + '", 5))=CONCAT(REPT( CHAR( 160 ), 5),""), MID("' + value + '", 6, LEN("' + value + '")), TRIM("' + value + '"))';
        } else if (direction == '>') {
          return '=CONCAT(REPT( CHAR( 160 ), 1),"' + value + '")';
        } else if (direction == '<') {
          return '=IF(TRIM(LEFT("' + value + '", 1))=CONCAT(REPT( CHAR( 160 ), 1),""), MID("' + value + '", 2, LEN("' + value + '")), TRIM("' + value + '"))';
        } 
      }
          Logger.log(formula);
      
      if (value != '') {
        cell.setFormula([formula()]);
        cell.setValue(cell.getValue());
      } else {
        cell.setValue(['']);
      }
    }
  }
};

function indentText1() {
  moveText(">");
};

function flushLeft1() {
  moveText("<");
};

function indentText5() {
  moveText(">>>>>");
};

function flushLeft5() {
  moveText("<<<<<");
};

function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();

  var entries = [{
    name : ">",
    functionName : "indentText1"
  },{
    name : ">>>>>",
    functionName : "indentText5"
  },{
    name : "<<<<<",
    functionName : "flushLeft5"
  },{
    name : "<",
    functionName : "flushLeft1"
  }];
  sheet.addMenu("Indent Text", entries);
};
