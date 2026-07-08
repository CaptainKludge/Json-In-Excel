# Json-In-Excel Docs

This folder is the readable guide to the functions exported in [../functions.json](../functions.json). The formulas themselves are the source of truth. Each function section includes a linted formula block, implementation notes that explain the actual mechanics, real-world examples, and notation notes.

## Read By Purpose

| File | Purpose |
|---|---|
| [json.md](json.md) | Build, inspect, update, merge, and remove JSON object content |
| [list-and-array.md](list-and-array.md) | Convert arrays, count values, project columns, and generate pairings |
| [safety-and-utils.md](safety-and-utils.md) | Parse ranges, handle measurements, reshape data safely, and annotate formulas |
| [algorithms.md](algorithms.md) | Use `partFill` to allocate parts against a target span |
| [importer-exporter.md](importer-exporter.md) | Import/export workflows for workbook names and JSON files |

## Coverage

The current `functions.json` export contains 27 functions.

| Group | Functions |
|---|---|
| JSON object operations | `jsonQuote`, `jsonObject`, `jsonGetKeysAtLevel`, `jsonGet`, `nestedJsonBuild`, `jsonSet`, `jsonRemove`, `jsonJoin` |
| List and array operations | `listToJson`, `listFromJson`, `arrayRepAdd`, `CountUnique`, `GiveMostFrequent`, `vLastItem`, `SelectFilter`, `permutate` |
| Utilities and analysis | `countOccurancesText`, `isInSet`, `COMMENT`, `dropBySet`, `EdgeDetect`, `between`, `safeFilter`, `makearr`, `safeDrop`, `inches` |
| Algorithm | `partFill` |

## Important Implementation Notes

The docs match the names that are actually exported. A few implementation details are still worth knowing when you import into a clean workbook.

- Name Manager comments can be exported in the reserved metadata object `__nameManagerComments`.

If you are importing this file set into a brand-new workbook, verify those names before relying on the dependent functions.

