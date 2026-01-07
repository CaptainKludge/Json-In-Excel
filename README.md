# Json In Excel

A comprehensive collection of Excel LAMBDA functions for JSON manipulation, list/array processing, and advanced mathematical algorithms. This project extends Excel's functionality with 25 custom functions that enable native JSON data object manipulations and sophisticated problem-solving capabilities.

**Now available for Google Sheets!** See [Google Sheets Installation Guide](#google-sheets-support) below.

## Quick Start

### Excel

Use the `jsonexcelexctraction.cmd` PowerShell script to import/export functions:

```cmd
# Import all functions from functions.json to Excel Name Manager
jsonexcelexctraction.cmd

# Export current Excel functions to functions.json (backup)
jsonexcelexctraction.cmd -export
```

### Google Sheets

1. Open your Google Sheet
2. Go to **Extensions > Apps Script**
3. Copy the code from `functions-google-sheets.js`
4. Save and authorize
5. Use functions like any built-in formula: `=jsonObject(A1:B10)`

See [GOOGLE_SHEETS_INSTALL.md](GOOGLE_SHEETS_INSTALL.md) for detailed instructions.

## Function Categories

This library contains **25 functions** organized into four main categories:

| Category | Function Count | Description |
|----------|----------------|-------------|
| [JSON Manipulation](docs/json.md) | 8 functions | Create, parse, modify, and navigate JSON objects |
| [List & Array Processing](docs/list-and-array.md) | 8 functions | Convert, filter, and manipulate arrays and lists |
| [Utility & Safety](docs/safety-and-utils.md) | 7 functions | Helper functions for safe operations and data validation |
| [Algorithm Solutions](docs/algorithms.md) | 2 functions | Advanced mathematical algorithms including money change solutions |

### Key Features

- **Native JSON Support**: Create and manipulate JSON objects directly in Excel formulas
- **Money Change Algorithms**: Solve optimization problems like part fitting and resource allocation
- **Safe Operations**: Error-resistant functions that handle edge cases gracefully
- **Array Processing**: Advanced list manipulation and filtering capabilities
- **Measurement Conversions**: Handle complex unit conversions and text parsing

### Money Change Algorithm Highlights

Two sophisticated algorithms solve the classic "money change making" problem adapted for physical parts:

- **`partFill`**: Sequential part allocation with remainder tracking
- **`greedyPartFill`**: Optimized greedy algorithm that minimizes remainder by prioritizing largest fitting parts

These functions take a target span (distance/amount) and an array of parts (with names and lengths) to calculate optimal allocation strategies.

## Google Sheets Support

All 25 functions are now available for Google Sheets! The functions have been converted from Excel LAMBDA functions to Google Apps Script custom functions.

### Key Files

- **`functions-google-sheets.js`** - Google Apps Script implementation of all 25 functions
- **`GOOGLE_SHEETS_INSTALL.md`** - Complete installation guide with examples
- **`GOOGLE_SHEETS_CHANGES.md`** - Detailed migration guide and compatibility matrix

### Installation Overview

1. Open your Google Sheet
2. Navigate to **Extensions > Apps Script**
3. Paste the code from `functions-google-sheets.js`
4. Save and authorize the script
5. Start using functions in your formulas!

Example usage:
```
=jsonObject(A1:B10)          // Create JSON from range
=jsonGet(A1, "user/name")    // Extract value from JSON
=partFill(100, B2:C5)        // Optimal part allocation
=listToJson(A1:A5)           // Convert array to JSON
```

### Differences from Excel

- Functions are implemented as Google Apps Script custom functions
- Uses `REGEXMATCH` instead of `REGEXTEST`
- Some functions have performance differences
- ARRAYFORMULA support varies by function

See [GOOGLE_SHEETS_CHANGES.md](GOOGLE_SHEETS_CHANGES.md) for a complete compatibility matrix and detailed migration notes.

### Function Compatibility

✅ **Fully Compatible** (21 functions): All JSON functions, most array and utility functions
⚠️ **Partial Support** (4 functions): Some ARRAYFORMULA limitations
❌ **See Documentation**: Performance considerations for large datasets

All 25 functions preserve the same functionality and behavior as the Excel versions.

