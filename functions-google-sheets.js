/**
 * JSON In Excel - Google Sheets Version
 * 
 * This file contains Google Apps Script implementations of all 25 Excel LAMBDA functions
 * from the original functions.json file, adapted to work in Google Sheets.
 * 
 * Installation:
 * 1. Open your Google Sheet
 * 2. Go to Extensions > Apps Script
 * 3. Delete any existing code
 * 4. Paste this entire file
 * 5. Save (Ctrl+S or Cmd+S)
 * 6. Authorize when prompted
 * 
 * Usage:
 * All functions can be used directly in formulas like built-in functions:
 *   =jsonObject(A1:B10)
 *   =jsonGet(A1, "path/to/value")
 *   =partFill(100, B2:C5)
 * 
 * For more information, see GOOGLE_SHEETS_CHANGES.md
 */

// ============================================================================
// HELPER FUNCTIONS (Not exposed as custom functions)
// ============================================================================

/**
 * Safely test if a string matches a regex pattern
 * Replaces Excel's REGEXTEST
 */
function regexTest_(text, pattern) {
  try {
    return new RegExp(pattern).test(String(text));
  } catch (e) {
    return false;
  }
}

/**
 * Extract all matches from regex
 * Enhanced version of REGEXEXTRACT
 */
function regexExtractAll_(text, pattern, occurrence, count) {
  try {
    var regex = new RegExp(pattern, 'g');
    var matches = String(text).match(regex);
    if (!matches) return [];
    
    if (occurrence !== undefined && count !== undefined) {
      return matches.slice(occurrence - 1, occurrence - 1 + count);
    }
    return matches;
  } catch (e) {
    return [];
  }
}

/**
 * Convert 2D array to flat array (column-major)
 */
function flatten_(arr) {
  if (!Array.isArray(arr)) return [arr];
  var result = [];
  for (var i = 0; i < arr.length; i++) {
    if (Array.isArray(arr[i])) {
      for (var j = 0; j < arr[i].length; j++) {
        if (arr[i][j] !== null && arr[i][j] !== undefined && arr[i][j] !== "") {
          result.push(arr[i][j]);
        }
      }
    } else {
      if (arr[i] !== null && arr[i] !== undefined && arr[i] !== "") {
        result.push(arr[i]);
      }
    }
  }
  return result;
}

/**
 * Get text before delimiter
 */
function textBefore_(text, delimiter) {
  var str = String(text);
  var idx = str.indexOf(delimiter);
  return idx >= 0 ? str.substring(0, idx) : str;
}

/**
 * Get text after delimiter
 */
function textAfter_(text, delimiter) {
  var str = String(text);
  var idx = str.indexOf(delimiter);
  return idx >= 0 ? str.substring(idx + delimiter.length) : "";
}

// ============================================================================
// JSON MANIPULATION FUNCTIONS
// ============================================================================

/**
 * Build a JSON object from a two-column range of keys and values
 * 
 * @param {range} range Two-column range where column 1 is keys and column 2 is values
 * @return {string} JSON object string
 * @customfunction
 */
function jsonObject(range) {
  if (!Array.isArray(range)) return "{}";
  
  var pairs = [];
  for (var i = 0; i < range.length; i++) {
    var key = range[i][0];
    var val = range[i][1];
    
    if (key !== null && key !== undefined && key !== "") {
      var quotedKey = jsonQuote(key);
      var quotedVal = (val === null || val === undefined || val === "") ? '""' : String(val);
      pairs.push(quotedKey + ":" + quotedVal);
    }
  }
  
  if (pairs.length === 0) return "{}";
  return "{" + pairs.join(",") + "}";
}

/**
 * Safely quote or normalize a value for JSON inclusion
 * 
 * @param {string} val Value to quote
 * @return {string} Properly quoted/formatted value for JSON
 * @customfunction
 */
function jsonQuote(val) {
  if (val === null || val === undefined) return '""';
  
  var s = String(val).trim();
  
  // Check if it's already an object or array
  if (regexTest_(s, '^\\s*(\\{.*\\}|\\[.*\\])\\s*$')) {
    return s;
  }
  
  // Check if it's a boolean or null
  if (regexTest_(s, '^(?:true|false|null)$')) {
    return s.toLowerCase();
  }
  
  // Check if it's a number
  if (regexTest_(s, '^-?\\d+(\\.\\d+)?([eE][+-]?\\d+)?$')) {
    return s;
  }
  
  // Check if already properly quoted
  if (regexTest_(s, '^"[^"]*"$')) {
    return s;
  }
  
  // Strip outer quotes and escape internal quotes
  var stripped = s.replace(/^"+|"+$/g, '');
  var core = stripped.replace(/""/g, "'");
  
  return '"' + core + '"';
}

/**
 * Parse a JSON object string and return a two-column array of key/value pairs
 * 
 * @param {string} json JSON object string
 * @return {Array} Two-column array of keys and values
 * @customfunction
 */
function jsonGetKeysAtLevel(json) {
  if (!json || typeof json !== 'string') return [["", ""]];
  
  var content = json.trim();
  if (content.charAt(0) === '{') {
    content = content.substring(1, content.length - 1);
  }
  
  var pairs = [];
  var token = "";
  var curl = 0;
  var square = 0;
  var quotes = false;
  
  for (var i = 0; i < content.length; i++) {
    var ch = content.charAt(i);
    
    if (ch === '"' && (i === 0 || content.charAt(i-1) !== '\\')) {
      quotes = !quotes;
      token += ch;
      continue;
    }
    
    if (!quotes) {
      if (ch === '{') curl++;
      else if (ch === '}') curl--;
      else if (ch === '[') square++;
      else if (ch === ']') square--;
    }
    
    var level = curl + square;
    
    if (ch === ',' && level === 0 && !quotes) {
      if (token.trim() !== "") {
        pairs.push(token.trim());
      }
      token = "";
    } else {
      token += ch;
    }
  }
  
  if (token.trim() !== "") {
    pairs.push(token.trim());
  }
  
  var result = [];
  for (var j = 0; j < pairs.length; j++) {
    var pair = pairs[j];
    var colonIdx = -1;
    var inQuotes = false;
    var depth = 0;
    
    for (var k = 0; k < pair.length; k++) {
      var c = pair.charAt(k);
      if (c === '"' && (k === 0 || pair.charAt(k-1) !== '\\')) {
        inQuotes = !inQuotes;
      }
      if (!inQuotes) {
        if (c === '{' || c === '[') depth++;
        if (c === '}' || c === ']') depth--;
        if (c === ':' && depth === 0 && colonIdx === -1) {
          colonIdx = k;
        }
      }
    }
    
    if (colonIdx >= 0) {
      var key = pair.substring(0, colonIdx).trim().replace(/"/g, '');
      var value = pair.substring(colonIdx + 1).trim();
      result.push([key, value]);
    }
  }
  
  return result.length > 0 ? result : [["", ""]];
}

/**
 * Retrieve a value from a JSON object using a slash-separated path
 * 
 * @param {string} json JSON object string
 * @param {string} path Slash-separated path (e.g., "parent/child")
 * @return {string} Value at path or #N/A if not found
 * @customfunction
 */
function jsonGet(json, path) {
  if (!json || !path) return "#N/A";
  
  var keys = String(path).split('/');
  var current = json;
  
  for (var i = 0; i < keys.length; i++) {
    var pairs = jsonGetKeysAtLevel(current);
    var found = false;
    
    for (var j = 0; j < pairs.length; j++) {
      if (pairs[j][0] === keys[i]) {
        current = pairs[j][1];
        found = true;
        break;
      }
    }
    
    if (!found) return "#N/A";
  }
  
  return current;
}

/**
 * Set or replace a value at a given path inside a JSON object
 * 
 * @param {string} oJson Original JSON object string
 * @param {string} oPath Slash-separated path
 * @param {string} oValue New value to set
 * @return {string} Updated JSON object string
 * @customfunction
 */
function jsonSet(oJson, oPath, oValue) {
  return jsonSetWalk_(oJson, oPath, oValue, 0);
}

function jsonSetWalk_(J, P, V, depth) {
  if (depth > 10) return "#STOP@" + P;
  
  var set = jsonGetKeysAtLevel(J);
  var parts = String(P).split('/');
  var hasTail = parts.length > 1;
  var key = parts[0];
  var tail = hasTail ? parts.slice(1).join('/') : "";
  
  var keypresent = false;
  var curRaw = "";
  
  for (var i = 0; i < set.length; i++) {
    if (set[i][0] === key) {
      keypresent = true;
      curRaw = set[i][1];
      break;
    }
  }
  
  var isObj = keypresent && curRaw.trim().charAt(0) === '{';
  var newValQ = jsonQuote(V);
  
  // Build new set
  var newSet = [];
  var keyReplaced = false;
  
  for (var i = 0; i < set.length; i++) {
    if (set[i][0] === key) {
      keyReplaced = true;
      if (hasTail && isObj) {
        newSet.push([key, jsonSetWalk_(curRaw, tail, V, depth + 1)]);
      } else if (hasTail) {
        newSet.push([key, nestedJsonBuild(tail, V)]);
      } else {
        newSet.push([key, newValQ]);
      }
    } else {
      newSet.push(set[i]);
    }
  }
  
  if (!keyReplaced) {
    if (hasTail) {
      newSet.push([key, nestedJsonBuild(tail, V)]);
    } else {
      newSet.push([key, newValQ]);
    }
  }
  
  return jsonObject(newSet);
}

/**
 * Build nested JSON objects from a slash-separated path and value
 * 
 * @param {string} p Path (e.g., "user/profile/name")
 * @param {string} v Value
 * @return {string} Nested JSON object
 * @customfunction
 */
function nestedJsonBuild(p, v) {
  if (!p || p === "") return jsonQuote(v);
  
  var parts = String(p).split('/');
  if (parts.length === 0) return jsonQuote(v);
  
  var result = jsonQuote(v);
  for (var i = parts.length - 1; i >= 0; i--) {
    result = jsonObject([[parts[i], result]]);
  }
  
  return result;
}

/**
 * Remove a key (or nested key path) from a JSON object
 * 
 * @param {string} oJson Original JSON object
 * @param {string} oPath Path to remove
 * @return {string} Updated JSON object
 * @customfunction
 */
function jsonRemove(oJson, oPath) {
  return jsonRemoveWalk_(oJson, oPath, 0);
}

function jsonRemoveWalk_(J, P, depth) {
  if (depth > 20) return "STOP@" + P;
  
  var set = jsonGetKeysAtLevel(J);
  var parts = String(P).split('/');
  var hasTail = parts.length > 1;
  var key = parts[0];
  var tail = hasTail ? parts.slice(1).join('/') : "";
  
  var newSet = [];
  
  for (var i = 0; i < set.length; i++) {
    if (set[i][0] === key) {
      if (hasTail) {
        var curRaw = set[i][1];
        var isObj = curRaw.trim().charAt(0) === '{';
        if (isObj) {
          newSet.push([key, jsonRemoveWalk_(curRaw, tail, depth + 1)]);
        } else {
          newSet.push(set[i]);
        }
      }
      // If no tail, just skip this key (removing it)
    } else {
      newSet.push(set[i]);
    }
  }
  
  return jsonObject(newSet);
}

/**
 * Merge JSON objects with different modes
 * Note: This is a simplified version. Full Excel version is very complex.
 * 
 * @param {string} json1 Base JSON object
 * @param {string} json2 JSON object to merge
 * @param {number} mode 0=normal, 1=replace, 2=add
 * @return {string} Merged JSON object
 * @customfunction
 */
function jsonJoin(json1, json2, mode) {
  mode = mode || 0;
  
  var set1 = jsonGetKeysAtLevel(json1);
  var set2 = jsonGetKeysAtLevel(json2);
  
  var result = {};
  
  // Add all keys from set1
  for (var i = 0; i < set1.length; i++) {
    if (set1[i][0] !== "") {
      result[set1[i][0]] = set1[i][1];
    }
  }
  
  // Merge keys from set2
  for (var i = 0; i < set2.length; i++) {
    var key = set2[i][0];
    var val = set2[i][1];
    
    if (key === "") continue;
    
    if (result[key] !== undefined) {
      var oldVal = result[key];
      
      // Check if both are objects
      if (oldVal.trim().charAt(0) === '{' && val.trim().charAt(0) === '{') {
        result[key] = jsonJoin(oldVal, val, mode);
      } else if (mode === 2) { // Add mode
        // Try to add numerically
        var oldNum = parseFloat(oldVal);
        var newNum = parseFloat(val);
        if (!isNaN(oldNum) && !isNaN(newNum)) {
          result[key] = String(oldNum + newNum);
        } else {
          result[key] = oldVal + val;
        }
      } else {
        result[key] = val; // Replace
      }
    } else {
      result[key] = val;
    }
  }
  
  // Convert back to array format
  var finalSet = [];
  for (var key in result) {
    finalSet.push([key, result[key]]);
  }
  
  return jsonObject(finalSet);
}

// ============================================================================
// LIST AND ARRAY PROCESSING FUNCTIONS
// ============================================================================

/**
 * Convert a 1-D array to JSON array string
 * 
 * @param {range} arr Array to convert
 * @return {string} JSON array string
 * @customfunction
 */
function listToJson(arr) {
  if (!Array.isArray(arr)) arr = [[arr]];
  
  var flat = flatten_(arr);
  var quoted = flat.map(function(x) {
    return jsonQuote(x);
  });
  
  return "[" + quoted.join(",") + "]";
}

/**
 * Convert a JSON array string to Excel array
 * 
 * @param {string} json JSON array string
 * @return {Array} Vertical array
 * @customfunction
 */
function listFromJson(json) {
  if (!json || typeof json !== 'string') return [[""]];
  
  var inner = json.trim();
  if (inner.charAt(0) === '[') {
    inner = inner.substring(1, inner.length - 1);
  }
  
  if (inner === "") return [[""]];
  
  var parts = inner.split(',');
  var result = [];
  
  for (var i = 0; i < parts.length; i++) {
    var val = parts[i].trim().replace(/^"|"$/g, '');
    result.push([val]);
  }
  
  return result;
}

/**
 * Add or replace a key/value pair in a two-column array
 * Internal helper function
 * 
 * @param {Array} arr Two-column array
 * @param {string} key Key to add/replace
 * @param {string} val Value
 * @return {Array} Updated array
 */
function arrayRepAdd(arr, key, val) {
  if (!Array.isArray(arr) || arr.length === 0) {
    arr = [["", ""]];
  }
  
  var result = [];
  var found = false;
  
  for (var i = 0; i < arr.length; i++) {
    if (arr[i][0] === key) {
      result.push([key, jsonQuote(val)]);
      found = true;
    } else if (arr[i][0] !== "") {
      result.push(arr[i]);
    }
  }
  
  if (!found) {
    result.push([key, jsonQuote(val)]);
  }
  
  return result;
}

/**
 * Count occurrences of each unique value in an array
 * 
 * @param {range} arr Array to analyze
 * @return {Array} Two-column array of unique values and counts
 * @customfunction
 */
function CountUnique(arr) {
  var flat = flatten_(arr);
  var counts = {};
  
  for (var i = 0; i < flat.length; i++) {
    var val = String(flat[i]);
    counts[val] = (counts[val] || 0) + 1;
  }
  
  var result = [];
  for (var key in counts) {
    result.push([key, counts[key]]);
  }
  
  return result.length > 0 ? result : [["", 0]];
}

/**
 * Find the most frequently occurring value in an array
 * 
 * @param {range} arr Array to analyze
 * @return {string} Most frequent value
 * @customfunction
 */
function GiveMostFrequent(arr) {
  var counts = CountUnique(arr);
  
  var maxCount = 0;
  var maxVal = "";
  
  for (var i = 0; i < counts.length; i++) {
    if (counts[i][1] > maxCount) {
      maxCount = counts[i][1];
      maxVal = counts[i][0];
    }
  }
  
  return maxVal;
}

/**
 * Get the last non-empty item from a vertical array
 * 
 * @param {range} array1 Array to search
 * @param {string} emptyvalue Optional value to treat as empty (default: "")
 * @return {string} Last non-empty value or "N/A"
 * @customfunction
 */
function vLastItem(array1, emptyvalue) {
  if (emptyvalue === undefined) emptyvalue = "";
  
  var flat = flatten_(array1);
  
  for (var i = flat.length - 1; i >= 0; i--) {
    if (flat[i] !== emptyvalue && flat[i] !== null && flat[i] !== undefined) {
      return flat[i];
    }
  }
  
  return "N/A";
}

/**
 * Select specific columns and filter rows
 * 
 * @param {range} arrayIn Source array
 * @param {Array} colNums Column numbers to select (1-based)
 * @param {Array} filters Boolean array for row filtering
 * @return {Array} Filtered and column-selected array
 * @customfunction
 */
function SelectFilter(arrayIn, colNums, filters) {
  if (!Array.isArray(arrayIn)) return [[""]];
  if (!Array.isArray(colNums)) colNums = [colNums];
  if (!Array.isArray(filters)) filters = flatten_([[filters]]);
  else filters = flatten_(filters);
  
  var result = [];
  
  for (var i = 0; i < arrayIn.length; i++) {
    if (filters[i]) {
      var row = [];
      for (var j = 0; j < colNums.length; j++) {
        var colIdx = (Array.isArray(colNums[j]) ? colNums[j][0] : colNums[j]) - 1;
        row.push(arrayIn[i][colIdx] || "");
      }
      result.push(row);
    }
  }
  
  return result.length > 0 ? result : [[""]];
}

/**
 * Filter columns based on set specification and repeating pattern
 * 
 * @param {range} range Source range
 * @param {string} setText Set specification (e.g., "1-3,5")
 * @param {number} repeat Pattern repeat period
 * @param {boolean} keepMatch True to keep matching columns, false to drop them
 * @return {Array} Filtered array
 * @customfunction
 */
function dropBySet(range, setText, repeat, keepMatch) {
  if (!Array.isArray(range)) return range;
  
  var numCols = range[0].length;
  var period = repeat || numCols;
  
  var result = [];
  for (var i = 0; i < range.length; i++) {
    var row = [];
    for (var j = 0; j < numCols; j++) {
      var patternPos = (j % period) + 1;
      var inSet = isInSet(patternPos, setText);
      
      if ((keepMatch && inSet) || (!keepMatch && !inSet)) {
        row.push(range[i][j]);
      }
    }
    result.push(row);
  }
  
  return result;
}

// ============================================================================
// SAFETY AND UTILITY FUNCTIONS
// ============================================================================

/**
 * Safely drop rows from an array
 * 
 * @param {range} arr Array
 * @param {number} rows Number of rows to drop from start
 * @return {Array} Array with rows dropped
 * @customfunction
 */
function safeDrop(arr, rows) {
  if (!Array.isArray(arr)) return [[""]];
  
  if (rows >= arr.length) {
    var cols = arr[0].length;
    return [Array(cols).fill("")];
  }
  
  return arr.slice(rows);
}

/**
 * Safely filter an array
 * 
 * @param {range} arr Array to filter
 * @param {Array} include Boolean array for filtering
 * @return {Array} Filtered array
 * @customfunction
 */
function safeFilter(arr, include) {
  if (!Array.isArray(arr)) return [[""]];
  if (!Array.isArray(include)) include = flatten_([[include]]);
  else include = flatten_(include);
  
  var result = [];
  for (var i = 0; i < arr.length; i++) {
    if (include[i]) {
      result.push(arr[i]);
    }
  }
  
  if (result.length === 0) {
    var cols = arr[0] ? arr[0].length : 1;
    return [Array(cols).fill("")];
  }
  
  return result;
}

/**
 * Convert range to dynamic array
 * 
 * @param {range} arr Input range
 * @return {Array} Array
 * @customfunction
 */
function makearr(arr) {
  return arr; // In Google Sheets, ranges are already arrays
}

/**
 * Test if a number is within specified intervals
 * 
 * @param {number} number Number to test
 * @param {string} rangeIn Interval specification (e.g., "[0,10]" or "5-10")
 * @return {boolean} True if number is in range
 * @customfunction
 */
function between(number, rangeIn) {
  if (rangeIn === null || rangeIn === undefined || rangeIn === "") {
    rangeIn = "(0,0)";
  }
  
  var num = parseFloat(number);
  if (isNaN(num)) return false;
  
  var safe = String(rangeIn).replace(/\s/g, '');
  var intervals = safe.split(',');
  
  for (var i = 0; i < intervals.length; i++) {
    var tok = intervals[i].trim();
    var hasBraces = tok.charAt(0) === '(' || tok.charAt(0) === '[';
    
    var leftB, rightB, min, max;
    
    if (hasBraces) {
      leftB = tok.charAt(0);
      rightB = tok.charAt(tok.length - 1);
      var body = tok.substring(1, tok.length - 1);
      var nums = body.split(',');
      min = nums[0] ? parseFloat(nums[0]) : -Infinity;
      max = nums[1] ? parseFloat(nums[1]) : Infinity;
    } else {
      // Dash notation
      if (tok.indexOf('-') > 0) {
        var parts = tok.split('-');
        min = parseFloat(parts[0]);
        max = parseFloat(parts[1]);
        leftB = '[';
        rightB = ']';
      } else {
        continue;
      }
    }
    
    var lowerOK = (leftB === '(') ? num > min : num >= min;
    var upperOK = (rightB === ')') ? num < max : num <= max;
    
    if (lowerOK && upperOK) return true;
  }
  
  return false;
}

/**
 * Test if a number is in a set (similar to between)
 * 
 * @param {number} number Number to test
 * @param {string} rangeIn Range specification
 * @return {boolean} True if number is in set
 * @customfunction
 */
function isInSet(number, rangeIn) {
  return between(number, rangeIn);
}

/**
 * Parse measurement string and convert to inches
 * 
 * @param {string} inval Measurement string (e.g., "5 feet 6 inches")
 * @return {number} Total inches
 * @customfunction
 */
function inches(inval) {
  if (!inval) return 0;
  
  var str = String(inval).replace(/"/g, 'in');
  var units = {
    'yards': 36,
    'yard': 36,
    'yd': 36,
    'feet': 12,
    'foot': 12,
    'ft': 12,
    "'": 12,
    'inches': 1,
    'inch': 1,
    'in': 1,
    '"': 1
  };
  
  var total = 0;
  var pattern = /(\d+(?:\.\d+)?)\s*([a-z'"]+)/gi;
  var match;
  
  while ((match = pattern.exec(str)) !== null) {
    var value = parseFloat(match[1]);
    var unit = match[2].toLowerCase();
    var multiplier = units[unit] || 1;
    total += value * multiplier;
  }
  
  // Check if just a number with no units
  if (total === 0) {
    var justNum = parseFloat(str);
    if (!isNaN(justNum)) total = justNum;
  }
  
  return total;
}

/**
 * Count occurrences of a character in text
 * 
 * @param {string} cellRef Text to search
 * @param {string} char Character to count
 * @return {number} Count of occurrences
 * @customfunction
 */
function countOccurrencesText(cellRef, char) {
  if (!cellRef || String(cellRef).trim() === "") return 0;
  
  var str = String(cellRef);
  var count = 0;
  
  for (var i = 0; i < str.length; i++) {
    if (str.charAt(i) === char) count++;
  }
  
  return count + 1; // Excel version adds 1
}

// ============================================================================
// ALGORITHM FUNCTIONS
// ============================================================================

/**
 * Sequential part allocation algorithm
 * Fill a span with parts, processing sequentially
 * 
 * @param {number} span Target distance to fill
 * @param {range} partarr Two-column array: part names and lengths
 * @param {range} extrahungry Optional additional allocation rules
 * @return {string} JSON object with part counts
 * @customfunction
 */
function partFill(span, partarr, extrahungry) {
  if (!Array.isArray(partarr)) return "{}";
  
  var state = {
    rem: span,
    out: {}
  };
  
  // Process each part
  for (var i = 0; i < partarr.length; i++) {
    var pname = String(partarr[i][0]);
    var plen = parseFloat(partarr[i][1]);
    
    if (isNaN(plen) || plen <= 0) continue;
    
    var pcount = Math.floor(state.rem / plen);
    state.rem = state.rem - (pcount * plen);
    
    if (pcount > 0) {
      state.out[pname] = pcount;
    }
  }
  
  // Convert to JSON
  var pairs = [];
  for (var key in state.out) {
    pairs.push([key, String(state.out[key])]);
  }
  
  return jsonObject(pairs);
}

/**
 * Greedy part allocation algorithm with remainder optimization
 * Fill a span with parts, then optimize to minimize remainder
 * 
 * @param {number} span Target distance to fill
 * @param {range} partarr Two-column array: part names and lengths
 * @param {range} extrahungry Optional additional allocation rules
 * @return {string} JSON object with part counts
 * @customfunction
 */
function greedyPartFill(span, partarr, extrahungry) {
  if (!Array.isArray(partarr)) return "{}";
  
  var state = {
    rem: span,
    out: {}
  };
  
  // Phase 1: Greedy allocation
  for (var i = 0; i < partarr.length; i++) {
    var pname = String(partarr[i][0]);
    var plen = parseFloat(partarr[i][1]);
    
    if (isNaN(plen) || plen <= 0) continue;
    
    var pcount = Math.floor(state.rem / plen);
    state.rem = state.rem - (pcount * plen);
    
    if (pcount > 0) {
      state.out[pname] = pcount;
    }
  }
  
  // Phase 2: Remainder optimization
  if (state.rem > 0) {
    // Find all parts that could fit
    var candidates = [];
    for (var i = 0; i < partarr.length; i++) {
      var plen = parseFloat(partarr[i][1]);
      if (plen >= state.rem) {
        candidates.push({
          name: String(partarr[i][0]),
          len: plen
        });
      }
    }
    
    // Pick smallest part that fits
    if (candidates.length > 0) {
      candidates.sort(function(a, b) { return a.len - b.len; });
      var pick = candidates[0];
      
      state.out[pick.name] = (state.out[pick.name] || 0) + 1;
      state.rem = state.rem - pick.len;
    }
  }
  
  // Convert to JSON
  var pairs = [];
  for (var key in state.out) {
    pairs.push([key, String(state.out[key])]);
  }
  
  return jsonObject(pairs);
}

// ============================================================================
// ADDITIONAL HELPER FUNCTIONS
// ============================================================================

/**
 * Create a safe empty array for returns
 */
function emptyArray_(cols) {
  cols = cols || 1;
  return [Array(cols).fill("")];
}

/**
 * Log message for debugging (appears in Apps Script logs)
 */
function debug_(msg) {
  Logger.log(msg);
}
