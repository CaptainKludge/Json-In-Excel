# Json In Excel

A comprehensive collection of Excel LAMBDA functions for JSON manipulation, list/array processing, and advanced mathematical algorithms. This project extends Excel's functionality with 22 custom functions that enable native JSON data object manipulations and sophisticated problem-solving capabilities.

## Quick Start

Use the `jsonexcelexctraction.cmd` PowerShell script to import/export functions:

```cmd
# Import all functions from functions.json to Excel Name Manager
jsonexcelexctraction.cmd

# Export current Excel functions to functions.json (backup)
jsonexcelexctraction.cmd -export
```

## Function Categories

This library contains **22 functions** organized into four main categories:

| Category | Function Count | Description |
|----------|----------------|-------------|
| [JSON Manipulation](docs/json.md) | 8 functions | Create, parse, modify, and navigate JSON objects |
| [List & Array Processing](docs/list-and-array.md) | 6 functions | Convert, filter, and manipulate arrays and lists |
| [Utility & Safety](docs/safety-and-utils.md) | 6 functions | Helper functions for safe operations and data validation |
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
