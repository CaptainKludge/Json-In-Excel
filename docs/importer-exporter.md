# Importer And Exporter

This project includes two separate ways to move Excel LAMBDA names between a workbook and a JSON file.

| Tool | Location | Role |
|---|---|---|
| GUI import/export tool | [../jsonexcelexctraction.cmd](../jsonexcelexctraction.cmd) | Main workflow for exporting workbook names to JSON or importing JSON names back into a workbook |
| VBA module | [../tools/excelnamespaceimport-export.vb](../tools/excelnamespaceimport-export.vb) | Workbook-side implementation of similar import/export behavior |

## 1. CMD And PowerShell Tool

Role: a Windows launcher that opens a small GUI for syncing Excel named LAMBDA formulas with a JSON file.

### What it is

- The file is named `jsonexcelexctraction.cmd`.
- The first line is a CMD wrapper.
- The rest of the file is embedded PowerShell.
- The PowerShell code builds a Windows Forms interface and talks to Excel through COM automation.

### What it does

The tool runs in two modes.

| Mode | What happens |
|---|---|
| Extract mode | Opens the workbook, scans defined names, keeps only names whose formulas start with `=` plus optional whitespace and `LAMBDA`, and writes them to JSON |
| Insert mode | Reads the JSON file, deletes existing workbook names with the same names, recreates them, restores Name Manager comments when present, and saves the workbook |

### Typical export workflow

1. Run [../jsonexcelexctraction.cmd](../jsonexcelexctraction.cmd).
2. Pick an Excel workbook.
3. Accept the suggested JSON path or choose another one.
4. Leave `Insert Mode` unchecked.
5. Click `Run Operation`.

### Typical import workflow

1. Run [../jsonexcelexctraction.cmd](../jsonexcelexctraction.cmd).
2. Pick the target workbook.
3. Pick the source JSON file.
4. Check `Insert Mode`.
5. Click `Run Operation`.

### Default file naming

When you choose a workbook, the tool proposes a JSON path in the same folder:

```text
<workbook name> - functions.json
```

If the tool is launched with an Excel file path as its first argument, it pre-fills the workbook path and the matching JSON path.

### How the export works internally

- Opens the workbook in read-only mode.
- Loops through `Workbook.Names`.
- Keeps only names where `RefersTo` starts with `=` plus optional whitespace and `LAMBDA`.
- Serializes the result as a JSON object of `name: formula` pairs.
- When a defined name has a Name Manager comment, stores it under the reserved metadata key `__nameManagerComments`.

### How the import works internally

- Reads the JSON file into memory.
- Opens the workbook for editing.
- Deletes any existing defined name with the same name.
- Adds a new workbook name whose `RefersTo` is the imported formula.
- Restores Name Manager comments from the optional `__nameManagerComments` metadata object.
- Saves the workbook.

### Safety notes

- Insert mode overwrites defined names with matching names.
- Insert mode saves the workbook automatically.
- Files exported before comment support still import correctly because formula entries remain backward compatible.
- The tool depends on Excel COM automation, so it is Windows-specific in practice.

## 2. VBA Module

Role: a VBA version of the same idea that can be imported into Excel and run from inside the workbook environment.

### File

See [../tools/excelnamespaceimport-export.vb](../tools/excelnamespaceimport-export.vb).

### What it provides

- A `ShowFunctionSyncForm` entry point.
- File pickers for choosing the workbook and JSON file.
- `ExtractFunctionsToJson` to export workbook LAMBDA names.
- `InsertFunctionsFromJson` to import names back into a workbook.

### How it differs from the CMD tool

| Area | CMD/PowerShell tool | VBA module |
|---|---|---|
| Primary environment | Windows desktop script | Excel VBA project |
| UI | Windows Forms GUI | UserForm-style workbook UI |
| Export logic | Filters workbook names by `=LAMBDA` | Same general idea |
| Import logic | Recreates workbook names and saves | Same general idea |
| JSON handling | PowerShell JSON serialization | Simple custom dictionary-to-JSON and JSON-to-dictionary code |

### Practical note

The VBA module uses a simplified JSON parser and serializer. That is fine for the `functions.json` shape used by this project, but it is not a general-purpose JSON engine.

The reserved `__nameManagerComments` metadata object used by the CMD/PowerShell exporter is not described as part of the VBA module workflow here. If you want comment round-tripping in the VBA path too, that module should be updated separately.

## Which One To Use

- Use [../jsonexcelexctraction.cmd](../jsonexcelexctraction.cmd) when you want the main repo workflow.
- Use [../tools/excelnamespaceimport-export.vb](../tools/excelnamespaceimport-export.vb) when you want the import/export logic inside an Excel VBA project.