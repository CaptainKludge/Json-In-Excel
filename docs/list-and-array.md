# List and Array Processing Functions

This document covers the **8 list and array processing functions** that handle conversions between Excel arrays, JSON arrays, and provide specialized array manipulation capabilities.

## listToJson

**Purpose**: Convert a 1-D Excel array (range) into a JSON array string.

**Syntax**: 
```excel
=listToJson(arr)
```

```excel
=LAMBDA(arr,
    LET(
        quoted, MAP(arr, LAMBDA(x, jsonQuote(x))),
        joined, TEXTJOIN(",",TRUE,quoted),
        "[" & joined & "]"
    )
)
```

**Example**:
```excel
=listToJson({"apple";"banana";"cherry"})
    -> ["apple","banana","cherry"]

=listToJson({1;2;3})
    -> [1,2,3]
```

---

## listFromJson

**Purpose**: Convert a JSON array string into an Excel vertical array of unquoted values.

**Syntax**: 
```excel
=listFromJson(json)
```

```excel
=LAMBDA(json,
    LET(
        inner, MID(TRIM(json),2,LEN(TRIM(json))-2),
        parts, IF(inner="", MAKEARRAY(0,1,LAMBDA(r,c,"")), TEXTSPLIT(inner, ",",,TRUE)),
        TRIM(SUBSTITUTE(parts, """", ""))
    )
)
```

**Example**:
```excel
=listFromJson("[\"apple\",\"banana\",\"cherry\"]")
    -> {"apple";"banana";"cherry"}

=listFromJson("[1,2,3]")
    -> {1;2;3}
```

**Note**: This parser assumes simple values without nested arrays/objects containing commas.

---

## arrayRepAdd

**Purpose**: Internal helper function that adds or replaces a key/value pair in a two-column array while ensuring no duplicate keys remain.

**Syntax**: 
```excel
=arrayRepAdd(arr, key, val)
```

```excel
=LAMBDA(arr,key,val,
    LET(
        safeArr, IF(ROWS(arr)=0, HSTACK("", ""), arr),
        kcol, INDEX(safeArr,,1),
        vcol, INDEX(safeArr,,2),
        newarr, safeFilter(safeArr, kcol<>key),
        cleanArr, IF(ROWS(newarr)=0, HSTACK("", ""), newarr),
        VSTACK(cleanArr, HSTACK(key, jsonQuote(val)))
    )
)
```

**Usage**: Used internally by `jsonSet`, `jsonJoin`, and `jsonRemove` functions. Generally not called directly by users.

---

## CountUnique

**Purpose**: Count occurrences of each unique value in an array and return a two-column summary.

**Syntax**: 
```excel
=CountUnique(arr)
```

```excel
=LAMBDA(arr,
    LET(
        flat, TOCOL(arr,1),
        uniques, UNIQUE(flat),
        counts, MAP(uniques, LAMBDA(u, SUM((flat=u)*1))),
        HSTACK(uniques, counts)
    )
)
```

**Example**:
```excel
=CountUnique({"apple";"banana";"apple";"cherry";"banana";"apple"})
    -> {"apple",3;"banana",2;"cherry",1}
```

**Use Cases**:
- Data analysis and frequency counting
- Survey response analysis
- Inventory summarization

---

## GiveMostFrequent

**Purpose**: Find and return the most frequently occurring value in an array.

**Syntax**: 
```excel
=GiveMostFrequent(arr)
```

```excel
=LAMBDA(arr,
    INDEX(
        SORT(
            LET(
                array, arr,
                uniques, UNIQUE(array),
                counts, COUNTIF(array, uniques),
                HSTACK(uniques, counts)
            ),
            2,
            -1
        ),
        1,
        1
    )
)
```

**Example**:
```excel
=GiveMostFrequent({"apple";"banana";"apple";"cherry";"banana";"apple"})
    -> "apple"  (appears 3 times)

=GiveMostFrequent({1;2;2;3;3;3;4})
    -> 3  (appears 3 times)
```

**How it works**:
1. Creates a unique list of values
2. Counts occurrences of each value using COUNTIF
3. Sorts by count in descending order
4. Returns the first (most frequent) value

**Use Cases**:
- Statistical mode calculation
- Finding dominant category in datasets
- Identifying most common response in surveys
- Data quality analysis (finding most prevalent value)

---

## vLastItem

**Purpose**: Get the last non-empty item from a vertical array, with optional default value for empty arrays.

**Syntax**: 
```excel
=vLastItem(array1, [emptyvalue])
```

```excel
=LAMBDA(array1,[emptyvalue],
    XLOOKUP(TRUE,(array1<>IF(ISOMITTED(emptyvalue),"",emptyvalue)),array1,"N/A",0,-1)
)
```

**Example**:
```excel
=vLastItem({"first";"middle";"";"last"})
    -> "last"

=vLastItem({"";"";""},"default")
    -> "default"
```

**Use Cases**:
- Finding the most recent non-empty entry in a list
- Getting final valid values from data series
- Default value handling for empty datasets

---

## SelectFilter

**Purpose**: Advanced filtering that combines column selection with row filtering in a single operation.

**Syntax**: 
```excel
=SelectFilter(ArrayIn, ComSepColNumsInBraces, Filters)
```

```excel
=LAMBDA(ArrayIn,ComSepColNumsInBraces,Filters,
    FILTER(INDEX(ArrayIn,SEQUENCE(ROWS(ArrayIn)),ComSepColNumsInBraces),Filters,NA())
)
```

**Parameters**:
- `ArrayIn`: Source array to filter
- `ComSepColNumsInBraces`: Array of column numbers to select (e.g., {1;3;5})
- `Filters`: Boolean array for row filtering

**Example**:
```excel
// Select columns 1 and 3,

 filter rows where column 2 > 10
=SelectFilter(A1:C10, {1;3}, B1:B10>10)
```

**Use Cases**:
- Complex data filtering with column projection
- Report generation with specific field selection
- Database-like operations in Excel

---

## dropBySet

**Purpose**: Advanced column filtering that removes or keeps columns based on a set specification and repeating pattern.

**Syntax**: 
```excel
=dropBySet(range, setText, repeat, keepMatch)
```

```excel
=LAMBDA(range,setText,repeat,keepMatch,
    LET(
        dataRange, range,
        setString, setText,
        repeatPattern, repeat,
        keepMatched, keepMatch,
        firstCol, COLUMN(INDEX(dataRange,1,1)),
        colOffsets, COLUMN(dataRange) - firstCol,
        numsRawText, IFERROR(TEXTSPLIT(REGEXREPLACE(setString,"[^\d]+"," ")," "), ""),
        numsFiltered, IF(COUNTA(numsRawText)=0, "", FILTER(numsRawText, numsRawText<>"")),
        maxNum, IF(COUNTA(numsFiltered)=0, COLUMNS(dataRange), MAX(VALUE(numsFiltered))),
        period, IF(maxNum<=0, COLUMNS(dataRange), maxNum),
        patternPos, MOD(colOffsets, period) + 1,
        labelArray, patternPos,
        matchedMask, MAP(labelArray, LAMBDA(lbl, isInSet(lbl,setString))),
        includeMask, IF(keepMatched, matchedMask, NOT(matchedMask)),
        colIndices, FILTER(SEQUENCE(1,COLUMNS(dataRange)), includeMask),
        CHOOSECOLS(dataRange, colIndices)
    )
)
```

**Parameters**:
- `range`: Source data range to process
- `setText`: Set specification string defining which columns to match (uses interval notation)
- `repeat`: Repeat pattern period for column matching
- `keepMatch`: TRUE to keep matching columns, FALSE to drop them

**How it works**:
1. Determines column positions relative to the first column
2. Applies a repeating pattern with the specified period
3. Evaluates which columns match the set specification using `isInSet`
4. Keeps or drops matching columns based on `keepMatch` parameter

**Example**:
```excel
// Keep every 3rd column (columns 3, 6, 9, etc.)
=dropBySet(A1:J10, "3", 3, TRUE)

// Drop columns 1-2 in a repeating pattern of 5
=dropBySet(A1:Z10, "[1,2]", 5, FALSE)
```

**Use Cases**:
- Removing repeating column patterns from imported data
- Extracting specific columns from structured datasets
- Data cleanup with periodic column structures
- Report formatting with column pattern filtering

## Function Integration

These array processing functions integrate seamlessly with the JSON system:

- **`listToJson`** and **`listFromJson`** provide bidirectional conversion between Excel arrays and JSON arrays
- **`arrayRepAdd`** powers the JSON manipulation functions by managing key-value collections
- **`CountUnique`**, **`GiveMostFrequent`**, **`vLastItem`**, **`SelectFilter`**, and **`dropBySet`** provide advanced data processing capabilities that complement JSON workflows

This combination enables sophisticated data transformation pipelines that can process arrays, convert to JSON for complex manipulation, and convert back to Excel arrays for final presentation.
