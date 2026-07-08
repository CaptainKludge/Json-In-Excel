# Json In Excel

Json In Excel is a workbook function library built around Excel LAMBDA formulas stored in [functions.json](functions.json). The current export contains 27 named functions focused on JSON strings, list and array work, validation helpers, and one allocation algorithm.

## Documentation Map

Start here, then jump to the section you need.

| Section | What it covers |
|---|---|
| [docs/README.md](docs/README.md) | Full documentation index and implementation notes |
| [docs/json.md](docs/json.md) | JSON creation, reading, updating, merging, and removal |
| [docs/list-and-array.md](docs/list-and-array.md) | Array conversion, frequency analysis, filtering, and permutations |
| [docs/safety-and-utils.md](docs/safety-and-utils.md) | Range checks, text parsing, edge detection, and utility helpers |
| [docs/algorithms.md](docs/algorithms.md) | The `partFill` span allocation algorithm |
| [docs/importer-exporter.md](docs/importer-exporter.md) | How the import/export tools work, including the CMD/PowerShell tool and the VBA module |

## Function Groups

The exported functions are easiest to use when grouped by purpose rather than by raw file order.

| Group | Count | Functions |
|---|---:|---|
| JSON object operations | 8 | `jsonQuote`, `jsonObject`, `jsonGetKeysAtLevel`, `jsonGet`, `nestedJsonBuild`, `jsonSet`, `jsonRemove`, `jsonJoin` |
| List and array operations | 8 | `listToJson`, `listFromJson`, `arrayRepAdd`, `CountUnique`, `GiveMostFrequent`, `vLastItem`, `SelectFilter`, `permutate` |
| Utilities and analysis | 10 | `countOccurancesText`, `isInSet`, `COMMENT`, `dropBySet`, `EdgeDetect`, `between`, `safeFilter`, `makearr`, `safeDrop`, `inches` |
| Allocation algorithm | 1 | `partFill` |

## Quick Start

The root tool is [jsonexcelexctraction.cmd](jsonexcelexctraction.cmd). It is a CMD wrapper around an embedded PowerShell GUI.

1. Run `jsonexcelexctraction.cmd`.
2. Select an Excel workbook.
3. Accept the auto-generated JSON path or choose another JSON file.
4. Leave `Insert Mode` unchecked to export workbook LAMBDAs to JSON.
5. Check `Insert Mode` to import functions from JSON into the workbook's defined names.

The default export target is `<workbook name> - functions.json` in the same folder as the workbook.

## What This Repo Is Good At

- Building JSON object strings directly in Excel formulas.
- Reading and updating nested JSON-like structures with slash-delimited paths.
- Converting between Excel arrays and JSON arrays.
- Counting, selecting, and reshaping tabular data with reusable LAMBDAs.
- Parsing compact range syntax such as `[0,10)` and converting feet/inches text into total inches.
- Allocating parts against a target span with `partFill`.

## Source Of Truth

The function definitions themselves live in [functions.json](functions.json). The docs in [docs/json.md](docs/json.md), [docs/list-and-array.md](docs/list-and-array.md), [docs/safety-and-utils.md](docs/safety-and-utils.md), and [docs/algorithms.md](docs/algorithms.md) are written from that file, with human-readable explanations and examples.

## Current Implementation Notes

The current exported set is useful, and the formulas in [functions.json](functions.json) also expose a few implementation details that matter when you import into a clean workbook.

- `dropBySet` calls `inInSet` in the saved formula, which appears to be a typo or an alias for `isInSet`.
- Name Manager comments are stored under the reserved metadata key `__nameManagerComments`.

Those notes matter if you are importing this exact `functions.json` into a clean workbook.
