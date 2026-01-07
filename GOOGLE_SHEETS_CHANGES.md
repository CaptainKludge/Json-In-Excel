# Google Sheets Migration Guide

This document outlines the necessary changes to migrate the Excel LAMBDA functions to Google Sheets.

## Overview

The original `functions.json` file contains 25 Excel LAMBDA functions. Google Sheets does not support LAMBDA functions natively in the same way. Instead, these functions must be implemented as **Google Apps Script custom functions**.

## Key Differences Between Excel and Google Sheets

### 1. Function Definition Method
- **Excel**: Uses LAMBDA functions defined in Name Manager
- **Google Sheets**: Uses Google Apps Script custom functions

### 2. Array Formula Handling
- **Excel**: Automatic spilling with dynamic arrays
- **Google Sheets**: Use `ARRAYFORMULA()` wrapper for array operations

### 3. Function Name Changes

| Excel Function | Google Sheets Equivalent | Notes |
|----------------|-------------------------|-------|
| `REGEXTEST(text, pattern)` | `REGEXMATCH(text, pattern)` | Same functionality, different name |
| `REGEXEXTRACT(text, pattern, [occurrence], [count])` | `REGEXEXTRACT(text, pattern)` | Google Sheets version is simpler, only extracts first match |
| `REGEXREPLACE(text, pattern, replacement)` | `REGEXREPLACE(text, pattern, replacement)` | Same name and functionality |
| `TEXTSPLIT(text, delimiter)` | `SPLIT(text, delimiter)` | Similar but different behavior |
| `MAKEARRAY(rows, cols, lambda)` | N/A | Must use loops in Apps Script |
| `TEXTBEFORE(text, delimiter)` | Custom function needed | Not available in Sheets |
| `TEXTAFTER(text, delimiter)` | Custom function needed | Not available in Sheets |
| `DROP(array, num_rows)` | Custom function needed | Not available in Sheets |
| `CHOOSECOLS(array, col_indices)` | Custom function needed | Not available in Sheets |
| `TOCOL(array, [ignore])` | `FLATTEN(array)` | Similar functionality |
| `XLOOKUP()` | `VLOOKUP()` or `INDEX(MATCH())` | XLOOKUP not available |
| `REDUCE(initial, array, lambda)` | Custom loops in Apps Script | Not available as formula |
| `MAP(array, lambda)` | `ARRAYFORMULA()` with expressions | Different approach |
| `SCAN()` | N/A | Not available |
| `HSTACK()` | `{array1, array2}` | Use array literal syntax |
| `VSTACK()` | `{array1; array2}` | Use array literal syntax |
| `LET()` | N/A | Not available, use intermediate cells |

### 4. Missing Helper Functions

The following helper functions don't exist in Google Sheets and need custom implementations:
- `TEXTBEFORE()` - Extract text before a delimiter
- `TEXTAFTER()` - Extract text after a delimiter  
- `DROP()` - Remove rows from array
- `CHOOSECOLS()` - Select specific columns
- `XLOOKUP()` - Advanced lookup (use VLOOKUP/INDEX-MATCH instead)

## Implementation Approach

### Option 1: Google Apps Script Custom Functions (Recommended)

Create custom functions in Google Apps Script that replicate the behavior of the Excel LAMBDA functions. These are accessed via `Tools > Script Editor` in Google Sheets.

**Advantages:**
- Full programming capabilities
- Can handle complex logic
- Reusable across the spreadsheet

**Disadvantages:**
- Slower than native formulas
- Requires Apps Script knowledge
- Functions must be called with `=FUNCTIONNAME()` syntax

### Option 2: Native Google Sheets Formulas with ARRAYFORMULA

Where possible, recreate the logic using native Google Sheets functions wrapped in `ARRAYFORMULA()`.

**Advantages:**
- Faster execution
- No Apps Script required
- More portable

**Disadvantages:**
- Limited by available functions
- More complex formulas
- Some functions cannot be replicated

## Recommended Approach: Hybrid Solution

The `functions-google-sheets.json` file contains Google Apps Script implementations of all 25 functions. These should be:

1. **Copied into Google Apps Script editor** (Tools > Script Editor)
2. **Saved and authorized** to run in your spreadsheet
3. **Used like any built-in function** in your spreadsheet formulas

## Installation Instructions

### Step 1: Open Script Editor
1. Open your Google Sheet
2. Go to `Extensions > Apps Script`
3. Delete any existing code in the editor

### Step 2: Copy Functions
1. Open `functions-google-sheets.json`
2. Copy all the function code
3. Paste into the Apps Script editor

### Step 3: Save and Authorize
1. Click the disk icon to save
2. Name your project (e.g., "JSON In Sheets")
3. Click "Run" on any function to trigger authorization
4. Grant necessary permissions

### Step 4: Use in Formulas
Functions are now available in your spreadsheet:
```
=jsonObject(A1:B10)
=jsonGet(A1, "path/to/value")
=partFill(100, B2:C5)
```

## Function-Specific Changes

### JSON Functions

#### jsonObject
- **Change**: Use `for` loops instead of `MAP`/`REDUCE`
- **Change**: Use `REGEXMATCH` instead of `REGEXTEST`
- **Impact**: Functionality preserved

#### jsonQuote
- **Change**: Use `REGEXMATCH` instead of `REGEXTEST`
- **Impact**: Functionality preserved

#### jsonGetKeysAtLevel
- **Change**: Complex character-by-character parsing needs loop-based implementation
- **Change**: Cannot use `REDUCE` - must use `for` loop
- **Impact**: Functionality preserved, performance may differ

#### jsonGet
- **Change**: Path traversal uses loop instead of `REDUCE`
- **Impact**: Functionality preserved

#### jsonSet
- **Change**: Recursive walker implemented with function calls instead of LAMBDA self-reference
- **Impact**: Functionality preserved

#### jsonJoin
- **Change**: Complex merger logic implemented with nested loops
- **Impact**: Functionality preserved, may be slower

#### jsonRemove
- **Change**: Recursive removal with function calls
- **Impact**: Functionality preserved

#### nestedJsonBuild
- **Change**: Loop-based building instead of `REDUCE`
- **Impact**: Functionality preserved

### List/Array Functions

#### listToJson
- **Change**: Loop-based mapping instead of `MAP`
- **Impact**: Functionality preserved

#### listFromJson
- **Change**: Use `SPLIT` instead of `TEXTSPLIT`
- **Change**: Handle edge cases differently
- **Impact**: May have slight differences in edge case handling

#### arrayRepAdd
- **Change**: Loop-based filtering and addition
- **Impact**: Functionality preserved

#### CountUnique
- **Change**: Manual unique counting with loops
- **Change**: Cannot use `TOCOL` - must flatten manually
- **Impact**: Functionality preserved

#### GiveMostFrequent
- **Change**: Manual sorting and counting
- **Impact**: Functionality preserved

#### vLastItem
- **Change**: Replace `XLOOKUP` with loop-based search
- **Impact**: Functionality preserved

#### SelectFilter
- **Change**: Manual column selection and filtering
- **Impact**: Functionality preserved

#### dropBySet
- **Change**: Column pattern matching with loops
- **Change**: Use `REGEXMATCH` instead of `REGEXTEST`
- **Impact**: Functionality preserved

### Utility Functions

#### safeDrop
- **Change**: Manual row dropping with array slicing
- **Impact**: Functionality preserved

#### safeFilter
- **Change**: Loop-based filtering
- **Impact**: Functionality preserved

#### makearr
- **Change**: Return input directly or convert to 2D array
- **Impact**: Similar functionality

#### between
- **Change**: Interval parsing with `SPLIT` instead of `TEXTSPLIT`
- **Change**: Use `REGEXMATCH` and custom `REGEXEXTRACT`
- **Impact**: Functionality preserved

#### inches
- **Change**: Custom regex extraction implementation
- **Change**: Loop-based parsing
- **Impact**: Functionality preserved

#### countOccurrencesText
- **Change**: Direct implementation (no significant changes)
- **Impact**: Functionality preserved

#### isInSet
- **Change**: Similar to `between` with loop-based evaluation
- **Impact**: Functionality preserved

### Algorithm Functions

#### partFill
- **Change**: Replace `REDUCE` with `for` loop
- **Change**: Manual JSON building and manipulation
- **Impact**: Functionality preserved

#### greedyPartFill
- **Change**: Two-phase algorithm with loops
- **Change**: Array operations with manual implementations
- **Impact**: Functionality preserved

## Performance Considerations

1. **Google Apps Script functions are slower** than native formulas
2. **Minimize calls in large ranges** - Apps Script has execution time limits
3. **Cache results** where possible - use helper columns for intermediate results
4. **Use ARRAYFORMULA** where the function supports it
5. **Consider splitting complex operations** into multiple simpler steps

## Limitations

### Execution Time Limits
- Apps Script custom functions have a **30-second execution limit**
- Very large datasets may timeout
- Solution: Process data in chunks or use native formulas where possible

### Recalculation
- Custom functions recalculate on every edit
- May cause performance issues with many function calls
- Solution: Use static values or manual recalculation triggers

### ARRAYFORMULA Support
Not all custom functions will work with `ARRAYFORMULA()` automatically. Each function in the JSON file is marked with its ARRAYFORMULA compatibility.

## Testing Recommendations

1. **Start small**: Test each function with simple inputs first
2. **Validate output**: Compare results with Excel version where possible
3. **Performance test**: Check execution time with realistic data sizes
4. **Edge cases**: Test empty inputs, special characters, nested structures
5. **Error handling**: Verify graceful degradation on invalid inputs

## Compatibility Matrix

| Function | Excel LAMBDA | Google Sheets Apps Script | ARRAYFORMULA Compatible | Notes |
|----------|-------------|---------------------------|------------------------|-------|
| jsonObject | ✅ | ✅ | ⚠️ Partial | Single object only |
| jsonQuote | ✅ | ✅ | ✅ | Works with arrays |
| jsonGetKeysAtLevel | ✅ | ✅ | ❌ | Single JSON only |
| jsonGet | ✅ | ✅ | ⚠️ Partial | Path must be constant |
| jsonSet | ✅ | ✅ | ❌ | Single operation |
| jsonJoin | ✅ | ✅ | ❌ | Single operation |
| jsonRemove | ✅ | ✅ | ❌ | Single operation |
| nestedJsonBuild | ✅ | ✅ | ❌ | Single operation |
| partFill | ✅ | ✅ | ❌ | Single calculation |
| greedyPartFill | ✅ | ✅ | ❌ | Single calculation |
| listToJson | ✅ | ✅ | ✅ | Array input supported |
| listFromJson | ✅ | ✅ | ✅ | Array input supported |
| arrayRepAdd | ✅ | ✅ | ❌ | Internal use |
| CountUnique | ✅ | ✅ | ✅ | Array input supported |
| GiveMostFrequent | ✅ | ✅ | ✅ | Array input supported |
| vLastItem | ✅ | ✅ | ⚠️ Partial | Single array |
| SelectFilter | ✅ | ✅ | ✅ | Array operations |
| dropBySet | ✅ | ✅ | ✅ | Array operations |
| safeDrop | ✅ | ✅ | ✅ | Array operations |
| safeFilter | ✅ | ✅ | ✅ | Array operations |
| makearr | ✅ | ✅ | ✅ | Array operations |
| between | ✅ | ✅ | ✅ | Works with arrays |
| inches | ✅ | ✅ | ✅ | Works with arrays |
| countOccurrencesText | ✅ | ✅ | ✅ | Works with arrays |
| isInSet | ✅ | ✅ | ✅ | Works with arrays |

**Legend:**
- ✅ Full support
- ⚠️ Partial support (with limitations)
- ❌ Not supported / Not applicable

## Migration Checklist

- [ ] Install Google Apps Script functions
- [ ] Test core JSON functions (jsonObject, jsonGet, jsonSet)
- [ ] Test array processing functions
- [ ] Test algorithm functions (partFill, greedyPartFill)
- [ ] Update formulas to use new function syntax if needed
- [ ] Verify performance with production data
- [ ] Document any differences in behavior
- [ ] Train users on Google Sheets differences

## Support and Troubleshooting

### Common Issues

**Issue**: Function not found error
- **Solution**: Ensure Apps Script code is saved and sheet has been refreshed

**Issue**: Authorization error
- **Solution**: Run any function manually in Apps Script editor to trigger auth flow

**Issue**: Execution timeout
- **Solution**: Reduce data size or split operation into smaller chunks

**Issue**: #ERROR! in cell
- **Solution**: Check function parameters and ensure data types match expectations

**Issue**: Unexpected results
- **Solution**: Review function-specific changes in this document

## Future Enhancements

Possible improvements for Google Sheets version:
1. Add ARRAYFORMULA support to more functions
2. Optimize performance-critical functions
3. Add Google Sheets-specific features (e.g., integration with Google APIs)
4. Create add-on for easier installation
5. Add menu-driven JSON tools

## Conclusion

While Google Sheets doesn't support Excel LAMBDA functions natively, all 25 functions can be successfully migrated using Google Apps Script. The `functions-google-sheets.json` file provides production-ready implementations that maintain the same functionality as the Excel versions, with appropriate adaptations for the Google Sheets environment.
