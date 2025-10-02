"""
Simple docs formatter for Excel LAMBDA code blocks.

This script is conservative: it replaces tab characters with 4 spaces inside fenced ```excel blocks and
applies a basic top-level comma -> newline rule to improve readability while preserving tokens.

Run:
    python tools/format_docs.py path/to/docs

It edits files in-place and creates a .bak copy alongside each edited file.
"""
from pathlib import Path
import re
import sys

FENCE_RE = re.compile(r"^```excel\s*$", flags=re.IGNORECASE)
END_FENCE_RE = re.compile(r"^```\s*$")

def process_block(lines):
    """Process lines inside an ```excel block.

    - Replace tabs with 4 spaces
    - Normalize trailing spaces
    - Insert a newline after top-level commas when not inside quotes or brackets (conservative)
    """
    text = "\n".join(line.replace('\t', '    ') for line in lines)

    # Conservative approach: add newline after commas at top-level (not inside parentheses/brackets/quotes)
    out = []
    depth = 0
    in_quote = False
    escape = False
    current = []

    for ch in text:
        current.append(ch)
        if ch == '"' and not escape:
            in_quote = not in_quote
        if ch == '\\' and not escape:
            escape = True
            continue
        escape = False
        if in_quote:
            continue
        if ch in '({[':
            depth += 1
        elif ch in ')}]':
            depth = max(0, depth-1)
        elif ch == ',' and depth == 0:
            # insert newline after comma
            current.append('\n')
    out_text = ''.join(current)

    # Trim trailing spaces on each line
    out_lines = [ln.rstrip() for ln in out_text.split('\n')]
    return out_lines


def process_file(path: Path):
    s = path.read_text(encoding='utf-8')
    lines = s.splitlines()
    out_lines = []
    i = 0
    changed = False
    while i < len(lines):
        line = lines[i]
        if FENCE_RE.match(line):
            out_lines.append(line)
            i += 1
            block = []
            # collect until closing fence
            while i < len(lines) and not END_FENCE_RE.match(lines[i]):
                block.append(lines[i])
                i += 1
            if i >= len(lines):
                # unterminated fence; append as-is
                out_lines.extend(block)
                break
            # process block
            new_block = process_block(block)
            if new_block != block:
                changed = True
            out_lines.extend(new_block)
            # append end fence
            out_lines.append(lines[i])
            i += 1
        else:
            out_lines.append(line)
            i += 1

    if changed:
        bak = path.with_suffix(path.suffix + '.bak')
        path.rename(bak)
        path.write_text('\n'.join(out_lines) + '\n', encoding='utf-8')
        print(f"Formatted {path} -> backup saved as {bak}")
    else:
        print(f"No changes for {path}")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python tools/format_docs.py path/to/docs_dir_or_file [...]")
        sys.exit(1)
    for p in sys.argv[1:]:
        pth = Path(p)
        if pth.is_dir():
            for f in pth.rglob('*.md'):
                process_file(f)
        elif pth.is_file():
            process_file(pth)
        else:
            print("Not found:", p)
