# Safety And Utility Functions

These functions support defensive array work, interval parsing, measurement parsing, and inline formula annotation. Several of them use a pattern of expanding text or arrays into intermediate structures, then reducing those structures back into a Boolean or scalar result.

## safeDrop

Purpose: drop rows safely without throwing an error when the requested drop is too large.

### Linted Formula

```excel
=LAMBDA(
  arr,
  rows,
  LET(
    dropped, IFERROR(DROP(arr, rows), MAKEARRAY(1, COLUMNS(arr), LAMBDA(r, c, ""))),
    out, dropped,
    out
  )
)
```

### How It Works

- Calls `DROP` directly.
- Catches failure with `IFERROR`.
- Synthesizes a blank array with the same width when the drop fails.

### Real-World Examples

```excel
=safeDrop(A1:C20,1)
=safeDrop(A1:C2,10)
```

### Notation Notes

- Useful when the number of header rows is dynamic.

## makearr

Purpose: force a range or array-like input into a dynamic array created with `MAKEARRAY`.

### Linted Formula

```excel
=LAMBDA(
  arr,
  MAKEARRAY(ROWS(arr), COLUMNS(arr), LAMBDA(a, b, INDEX(arr, a, b)))
)
```

### How It Works

- Reconstructs every cell by coordinate.
- Returns a fresh spilled array regardless of the original input source.

### Real-World Examples

```excel
=makearr(A1:C5)
```

### Notation Notes

- Mostly a compatibility helper for array pipelines.

## safeFilter

Purpose: filter an array safely and return a blank-shaped array when no rows match.

### Linted Formula

```excel
=LAMBDA(
  arr,
  include,
  LET(
    cols, COLUMNS(arr),
    safe, IFERROR(FILTER(arr, include), MAKEARRAY(1, cols, LAMBDA(r, c, ""))),
    safe
  )
)
```

### How It Works

- Calls `FILTER` directly on the input.
- If `FILTER` would error because nothing matched, returns a generated blank array with the same width.

This is a defensive wrapper used by the JSON helpers so they can rebuild objects without hard failures when a filtered key set becomes empty.

### Real-World Examples

```excel
=safeFilter(A1:C10,B1:B10="Open")
=safeFilter(A1:C10,B1:B10="Missing")
```

### Notation Notes

- Useful when downstream formulas expect an array result even when there are no matches.

## between

Purpose: test whether numbers fall inside text-defined intervals.

### Linted Formula

```excel
=LAMBDA(
  numberinpreclean,
  rangein,
  LET(
    parseset,
      LAMBDA(
        singleset,
        LET(
          raw, TRIM(singleset),
          numberin, VALUE(TRIM(numberinpreclean)),
          safe, SUBSTITUTE(REGEXREPLACE(raw, "([(\[].*?),(.*?[\])])", "$1|$2"), " ", ""),
          intervals, TEXTSPLIT(safe, ","),
          parseInterval,
            LAMBDA(
              tok,
              LET(
                t, TRIM(SUBSTITUTE(tok, "|", ",")),
                hasBraces, OR(LEFT(t, 1) = "(", LEFT(t, 1) = "["),
                dashParts, IF(hasBraces, "", TEXTSPLIT(t, "-")),
                leftB, IF(hasBraces, LEFT(t, 1), "["),
                rightB, IF(hasBraces, RIGHT(t, 1), "]"),
                body, IF(hasBraces, MID(t, 2, LEN(t) - 2), t),
                nums, IF(hasBraces, REGEXEXTRACT(body, "([-0-9]+(?:\.[-0-9]+)?)?,([-0-9]+(?:\.[-0-9]+)?)?", 2), ""),
                minRaw, IF(hasBraces, IFERROR(INDEX(nums, 1), -1E+99), IF(COUNTA(dashParts) >= 1, TRIM(INDEX(dashParts, 1)), "")),
                maxRaw, IF(hasBraces, IFERROR(INDEX(nums, 2), 1E+99), IF(COUNTA(dashParts) = 2, TRIM(INDEX(dashParts, 2)), "")),
                min, IF(OR(minRaw = "", minRaw = "-∞"), -1E+99, VALUE(minRaw)),
                max, IF(OR(maxRaw = "", maxRaw = "∞", maxRaw = "+∞"), 1E+99, VALUE(maxRaw)),
                lowerOK, IF(leftB = "(", numberin > min, numberin >= min),
                upperOK, IF(rightB = ")", numberin < max, numberin <= max),
                AND(lowerOK, upperOK)
              )
            ),
          result, OR(MAP(intervals, parseInterval)),
          result
        )
      ),
    MAP(rangein, parseset)
  )
)
```

### How It Works

- Preprocesses interval text so commas inside bracket pairs do not split too early.
- Expands the text into interval tokens.
- Parses bracket semantics and dash-range semantics separately.
- Maps each interval to a Boolean and reduces them with `OR`.

### Real-World Examples

```excel
=between(7,"[0,10]")
=between(12,"(0,10],[12,20)")
=between(18,"10-20")
```

### Notation Notes

- `[` and `]` are inclusive bounds.
- `(` and `)` are exclusive bounds.
- Multiple intervals can be supplied in one text string.

## isInSet

Purpose: evaluate set membership with interval syntax and array-aware broadcasting.

### Linted Formula

```excel
=LAMBDA(
  numberinpreclean,
  singleset,
  LET(
    n, numberinpreclean,
    s, singleset,
    nRows, IFERROR(ROWS(n), 1),
    sRows, IFERROR(ROWS(s), 1),
    len, MAX(nRows, sRows),
    nArr, IF(nRows = 1, MAKEARRAY(len, 1, LAMBDA(r, c, n)), n),
    sArr, IF(sRows = 1, MAKEARRAY(len, 1, LAMBDA(r, c, s)), s),
    MAP(
      nArr,
      sArr,
      LAMBDA(
        numberin,
        singleset,
        LET(
          raw, TRIM(IF(singleset & "" = "", "(0,0)", singleset & "")),
          numberval, VALUE(TRIM(numberin)),
          intervals, TOCOL(IFERROR(REGEXEXTRACT(raw, "(\[[^\]]*\]|\([^\)]*\))"), ""), 1),
          parseInterval,
            LAMBDA(
              tok,
              IF(
                TRIM(tok) = "",
                FALSE,
                LET(
                  t, TRIM(tok),
                  leftB, LEFT(t, 1),
                  rightB, RIGHT(t, 1),
                  body, MID(t, 2, LEN(t) - 2),
                  hasComma, ISNUMBER(SEARCH(",", body)),
                  singleValue, IF(NOT(hasComma), IFERROR(VALUE(body), ""), ""),
                  min, IF(singleValue <> "", singleValue, IFERROR(VALUE(TEXTBEFORE(body, ",")), -1E+99)),
                  max, IF(singleValue <> "", singleValue, IFERROR(VALUE(TEXTAFTER(body, ",")), 1E+99)),
                  validInterval, OR(hasComma, singleValue <> ""),
                  IF(
                    NOT(validInterval),
                    FALSE,
                    AND(
                      IF(leftB = "(", numberval > min, numberval >= min),
                      IF(rightB = ")", numberval < max, numberval <= max)
                    )
                  )
                )
              )
            ),
          OR(TOCOL(MAP(intervals, parseInterval), 1))
        )
      )
    )
  )
)
```

### How It Works

- Broadcasts scalar inputs across array inputs when needed.
- Extracts only bracket-style interval tokens.
- Maps each candidate interval into a Boolean membership test.
- Reduces the mapped result with `OR`.

### Real-World Examples

```excel
=isInSet(25,"[20,30]")
=isInSet({5;15;25},{"[0,10]";"[10,20)";"[20,30]"})
```

### Notation Notes

- This function is better suited than `between` when you need array broadcasting.

## dropBySet

Purpose: keep or remove columns based on repeating set membership.

### Linted Formula

```excel
=LAMBDA(
  range,
  setText,
  repeat,
  keepMatch,
  LET(
    dataRange, range,
    setString, setText,
    repeatPattern, repeat,
    keepMatched, keepMatch,
    firstCol, COLUMN(INDEX(dataRange, 1, 1)),
    colOffsets, COLUMN(dataRange) - firstCol,
    numsRawText, IFERROR(TEXTSPLIT(REGEXREPLACE(setString, "[^\d]+", " "), " "), ""),
    numsFiltered, IF(COUNTA(numsRawText) = 0, "", FILTER(numsRawText, numsRawText <> "")),
    maxNum, IF(COUNTA(numsFiltered) = 0, COLUMNS(dataRange), MAX(VALUE(numsFiltered))),
    period, IF(maxNum <= 0, COLUMNS(dataRange), maxNum),
    patternPos, MOD(colOffsets, period) + 1,
    labelArray, patternPos,
    matchedMask, MAP(labelArray, LAMBDA(lbl, inInSet(lbl, setString))),
    includeMask, IF(keepMatched, matchedMask, NOT(matchedMask)),
    colIndices, FILTER(SEQUENCE(1, COLUMNS(dataRange)), includeMask),
    CHOOSECOLS(dataRange, colIndices)
  )
)
```

### How It Works

- Computes each column's position inside a repeating cycle.
- Evaluates each cycle position against the requested set.
- Builds a column include mask and feeds that into `CHOOSECOLS`.

### Real-World Examples

```excel
=dropBySet(A1:J20,"3",3,TRUE)
=dropBySet(A1:Z10,"[1,2]",5,FALSE)
```

### Notation Notes

- Intended for patterned imports where every `n` columns repeat the same structure.
- The stored formula calls `inInSet`, which appears to be a typo or local alias for `isInSet`.

## countOccurancesText

Purpose: count the number of times a substring appears in text.

### Linted Formula

```excel
=LAMBDA(
  cell_ref,
  char,
  IF(LEN(TRIM(cell_ref)) = 0, 0, LEN(cell_ref) - LEN(SUBSTITUTE(cell_ref, char, "")) + 1)
)
```

### How It Works

- Removes the target substring.
- Compares the original and reduced lengths.
- Returns zero for blank text.

### Real-World Examples

```excel
=countOccurancesText("M8,M10,M8,M8",",")
=countOccurancesText("banana","a")
```

### Notation Notes

- The stored formula adds `+1`, so it behaves more like delimiter-count-to-item-count for single-character separators than a strict raw occurrence counter.

## inches

Purpose: parse mixed measurement text and convert the result to total inches.

### Linted Formula

```excel
=LAMBDA(
  inval,
  LET(
    rng, IFERROR(REGEXEXTRACT(SUBSTITUTE(inval, """", "in"), "(\d+|yards|feet|inches|""""|'|in|ft|yd)", 1, 1), ),
    vals,
      LET(
        n, COUNTA(rng),
        pairs, IF(MOD(n, 2) = 0, SEQUENCE(n / 2, 2, 1, 1), SEQUENCE((n - 1) / 2, 2, 1, 1)),
        INDEX(rng, pairs)
      ),
    sumation,
      SUM(
        BYROW(
          vals,
          LAMBDA(arr, LET(a, INDEX(arr, , 1), a * XLOOKUP(INDEX(arr, 1, 2), INDEX(units, , 1), INDEX(units, , 2))))
        )
      ),
    sumation
  )
)
```

### How It Works

- Extracts alternating number and unit tokens with regex.
- Regroups the token stream into two-column number/unit pairs.
- Uses `BYROW` and `XLOOKUP` against a `units` table to convert each pair.
- Sums the converted rows.

### Real-World Examples

```excel
=inches("3 ft 6 in")
=inches("2 yd 1 ft")
=inches("5'6""")
```

### Notation Notes

- This function assumes a workbook-level `units` table exists.
- It is best suited for compact shop-floor or construction notation.

## EdgeDetect

Purpose: compare each row with its previous row, next row, or both using a supplied detector lambda.

### Linted Formula

```excel
=LAMBDA(
  data,
  mode,
  detector,
  LET(
    n, ROWS(data),
    prev, VSTACK(TAKE(data, -1), TAKE(data, n - 1)),
    next, VSTACK(DROP(data, 1), TAKE(data, 1)),
    SWITCH(
      mode,
      -1, MAP(data, prev, LAMBDA(cur, p, detector(cur, p))),
      1, MAP(data, next, LAMBDA(cur, nxt, detector(cur, nxt))),
      0, MAP(data, prev, next, LAMBDA(cur, p, nxt, detector(cur, p, nxt)))
    )
  )
)
```

### How It Works

- Creates a shifted previous-row array and next-row array.
- Uses `MAP` to apply the detector row by row.
- Wraps at the boundaries using `TAKE` and `DROP`.

### Real-World Examples

```excel
=EdgeDetect({1;1;2;2;3},-1,LAMBDA(cur,prev,cur<>prev))
=EdgeDetect(B2:B20,1,LAMBDA(cur,nxt,cur<>nxt))
```

### Notation Notes

- `mode = 0` expects a detector that accepts three arguments.

## COMMENT

Purpose: attach a note to a formula segment without changing the returned value.

### Linted Formula

```excel
=LAMBDA(
  FEATURE,
  NOTE,
  FEATURE
)
```

### How It Works

- Returns the first argument unchanged.
- Ignores the second argument completely.

### Real-World Examples

```excel
=LET(
  cleaned, COMMENT(TRIM(A2), "Normalize operator input"),
  UPPER(cleaned)
)
```

### Notation Notes

- This is purely a readability helper for long formulas.
