format_docs.py — simple formatter for docs Excel formula blocks

Usage

    python tools/format_docs.py docs/

What it does

- Replaces tabs with 4 spaces inside fenced ```excel code blocks
- Adds conservative newlines after top-level commas to improve readability
- Saves a backup of edited files with a `.bak` suffix

Notes

- This is intentionally conservative. It avoids parsing or rewriting the formulas beyond whitespace/newline changes.
- Review diffs after running to confirm no unintended edits.
