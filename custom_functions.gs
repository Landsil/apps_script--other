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
