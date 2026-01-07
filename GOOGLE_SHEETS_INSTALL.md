# Google Sheets Installation Guide

This guide walks you through installing and using the JSON In Excel functions in Google Sheets.

## Prerequisites

- A Google account with access to Google Sheets
- Basic familiarity with Google Sheets formulas
- Understanding of JSON concepts (helpful but not required)

## Installation Steps

### Step 1: Open Google Apps Script Editor

1. Open your Google Sheet (or create a new one)
2. Click on **Extensions** in the menu bar
3. Select **Apps Script**
4. This will open the Apps Script editor in a new tab

### Step 2: Clear Default Code

1. In the Apps Script editor, you'll see a default `myFunction()` function
2. Select all the code (Ctrl+A or Cmd+A)
3. Delete it

### Step 3: Copy Function Code

1. Open the `functions-google-sheets.js` file from this repository
2. Copy the entire contents (Ctrl+A, then Ctrl+C)
3. Return to the Apps Script editor
4. Paste the code (Ctrl+V or Cmd+V)

### Step 4: Save the Project

1. Click the disk icon or press Ctrl+S (Cmd+S on Mac)
2. Give your project a name (e.g., "JSON In Sheets" or "JSON Functions")
3. Click **Save**

### Step 5: Authorize the Script

Before you can use the functions, you need to authorize them:

1. In the Apps Script editor, select any function from the dropdown (e.g., `jsonObject`)
2. Click the **Run** button (▶️ play icon)
3. A dialog will appear: "Authorization required"
4. Click **Review permissions**
5. Select your Google account
6. Click **Advanced** (if you see a warning)
7. Click **Go to [Your Project Name] (unsafe)**
8. Click **Allow**

**Note**: Google shows a warning because this is a custom script. The functions only access your spreadsheet data and don't share it externally.

### Step 6: Test the Installation

1. Return to your Google Sheet
2. In any cell, type: `=jsonObject(A1:B5)`
3. Press Enter
4. If you see a result (even if it's `{}`), the installation was successful!

## Basic Usage Examples

### Example 1: Create a JSON Object

**Setup:**
In cells A1:B3, enter:
```
| A        | B      |
|----------|--------|
| name     | "Alice"|
| age      | 30     |
| city     | "NYC"  |
```

**Formula:**
```
=jsonObject(A1:B3)
```

**Result:**
```
{"name":"Alice","age":30,"city":"NYC"}
```

### Example 2: Get a Value from JSON

**Setup:**
In cell A1, enter:
```
{"person":{"name":"Bob","age":25}}
```

**Formula:**
```
=jsonGet(A1, "person/name")
```

**Result:**
```
Bob
```

### Example 3: Set a Value in JSON

**Setup:**
In cell A1, enter:
```
{"user":{"name":"Charlie"}}
```

**Formula:**
```
=jsonSet(A1, "user/email", "charlie@example.com")
```

**Result:**
```
{"user":{"name":"Charlie","email":"charlie@example.com"}}
```

### Example 4: Part Allocation Algorithm

**Setup:**
In cells A2:B4, enter part specifications:
```
| A       | B  |
|---------|----| 
| Pipe A  | 30 |
| Pipe B  | 20 |
| Pipe C  | 5  |
```

**Formula:**
```
=partFill(100, A2:B4)
```

**Result:**
```
{"Pipe A":"3","Pipe B":"0","Pipe C":"2"}
```

This fills a 100-unit span: 3×30=90, remainder 10; 0×20=0, remainder 10; 2×5=10, remainder 0.

### Example 5: Convert Array to JSON

**Setup:**
In cells A1:A4, enter:
```
apple
banana
cherry
```

**Formula:**
```
=listToJson(A1:A4)
```

**Result:**
```
["apple","banana","cherry"]
```

## Common Issues and Solutions

### Issue 1: "Unknown function"
**Cause**: The script hasn't been saved or the spreadsheet hasn't been refreshed.
**Solution**: 
- Save the Apps Script project
- Refresh your Google Sheet (F5 or Cmd+R)
- Wait a few seconds for Google to recognize the new functions

### Issue 2: "Authorization required"
**Cause**: You haven't authorized the script yet.
**Solution**: Follow Step 5 above to authorize the script.

### Issue 3: "#ERROR!" in the cell
**Cause**: The function encountered an error (wrong parameters, invalid data, etc.)
**Solution**: 
- Check that you're using the correct parameter types
- Verify your data format
- Look at the Apps Script logs: Extensions > Apps Script > Execution log

### Issue 4: Slow performance
**Cause**: Apps Script custom functions are slower than native functions.
**Solution**: 
- Minimize the number of function calls
- Use helper columns for intermediate results
- Consider processing smaller datasets
- Cache results in cells instead of recalculating

### Issue 5: "Exceeded maximum execution time"
**Cause**: The function is taking too long (>30 seconds).
**Solution**: 
- Reduce the dataset size
- Split the operation into smaller steps
- Use native Google Sheets functions where possible

## ARRAYFORMULA Support

Some functions work well with `ARRAYFORMULA`, allowing you to apply them to entire ranges:

### Works Well:
```
=ARRAYFORMULA(jsonQuote(A1:A10))
=ARRAYFORMULA(between(B1:B10, "[0,100]"))
=ARRAYFORMULA(inches(C1:C10))
```

### Limited Support:
```
=ARRAYFORMULA(jsonObject(A1:B10))  // May not work as expected
=ARRAYFORMULA(jsonGet(A1:A10, "key"))  // Path must be constant
```

### Not Supported:
Most complex JSON manipulation functions process one object at a time.

## Performance Tips

1. **Use Helper Columns**: Break complex operations into steps
   ```
   A1: =jsonObject(Data!A1:B10)
   B1: =jsonSet(A1, "new/path", "value")
   C1: =jsonGet(B1, "result")
   ```

2. **Cache Results**: Copy and paste values to avoid recalculation
   - Select cells with formulas
   - Copy (Ctrl+C)
   - Right-click > Paste special > Values only

3. **Limit Function Calls**: Avoid using custom functions in large ranges
   - Instead of calling a function 1000 times, process data in batches

4. **Use Native Functions When Possible**: 
   - For simple text operations, use built-in SPLIT, CONCATENATE, etc.
   - Only use custom functions for JSON-specific operations

## Advanced Usage

### Chaining Operations

Combine multiple functions for complex workflows:

```
=jsonGet(
  jsonSet(
    jsonObject(A1:B10),
    "metadata/timestamp",
    NOW()
  ),
  "metadata/timestamp"
)
```

### Dynamic Paths

Use cell references for dynamic paths:

```
// A1 contains: {"user":{"name":"Alice","age":30}}
// B1 contains: user/name

=jsonGet(A1, B1)  // Returns: Alice
```

### Error Handling

Wrap functions in IFERROR for graceful failures:

```
=IFERROR(jsonGet(A1, "path/to/value"), "Not found")
```

## Next Steps

1. **Read the Documentation**: See `GOOGLE_SHEETS_CHANGES.md` for detailed information about each function
2. **Experiment**: Try different functions with your own data
3. **Build Workflows**: Combine functions to solve real problems
4. **Share**: The Apps Script can be shared by sharing the Google Sheet

## Getting Help

- **Function Reference**: See the original documentation in the `docs/` folder
- **Migration Guide**: Read `GOOGLE_SHEETS_CHANGES.md` for Excel vs. Google Sheets differences
- **Compatibility Matrix**: Check which functions support ARRAYFORMULA
- **Code Comments**: The `functions-google-sheets.js` file includes JSDoc comments explaining each function

## Uninstalling

To remove the functions:

1. Open **Extensions > Apps Script**
2. Delete all code from the editor
3. Save the project
4. The functions will no longer be available in your sheet

## License

These functions are provided under the same license as the original JSON In Excel project. See the LICENSE file for details.
