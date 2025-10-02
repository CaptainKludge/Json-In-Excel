## Json-In-Excel function docs

This folder contains generated documentation for the Excel LAMBDA helper functions stored in the project's functions JSON file. Each file groups related functions by purpose and shows the function's formula (as an Excel formula) with notes.

Files:

| File | Purpose |
|---|---|
| `json.md` | Functions for JSON creation, quoting, getting and setting values |
| `list-and-array.md` | Helpers for lists, arrays and conversions |
| `safety-and-utils.md` | Utility and safe-wrapping functions (safeDrop, safeFilter, makearr, etc.) |

Notes:

- The code blocks contain Excel formulas intended to be copy-pasted into the Excel Name Manager or used as LAMBDA definitions. They do not include comments inside the formulas themselves (Excel formulas don't support comments). Additional explanatory notes are provided outside the code blocks.

---

Generated on: 2025-10-02

## Using the command-line helper (`jsonexcelexctraction.cmd`)

There is a small bundled helper script in the repo root named `jsonexcelexctraction.cmd`. It launches a simple PowerShell GUI to extract Excel LAMBDA Name Manager entries to a JSON file or insert functions from a JSON file back into an Excel workbook.

Usage summary:

- Double-click `jsonexcelexctraction.cmd` to open the GUI. You can also run it from PowerShell or CMD. When launched it will allow selecting an Excel file and a target JSON file.
- Modes:
	- Extract Mode (default): scans the selected workbook's defined names and exports any `=LAMBDA(...)` formulas to a JSON file named `<workbook name> - functions.json` by default.
	- Insert Mode (check the "Insert Mode" box): reads a JSON functions file and inserts each entry as a defined name in the selected workbook (overwriting existing names with the same name).

Command-line invocation (optional):

You may pass an Excel file path as the first argument to pre-select it when the GUI opens. Example from PowerShell (pwsh.exe):

```powershell
& .\jsonexcelexctraction.cmd "C:\path\to\workbook.xlsx"
```

Notes & safety:

- The script uses COM automation to open and edit the workbook. It saves the workbook automatically when inserting names. Make a backup before running Insert Mode on important workbooks.
- The GUI also allows picking a different JSON file path if you don't want the default `<workbook name> - functions.json` file.

