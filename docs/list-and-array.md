# List And Array Functions

These functions convert between Excel arrays and JSON list text, count values, select slices of tables, and generate pairwise outputs. The common pattern is to expand arrays with `MAP`, `MAKEARRAY`, or `TOCOL`, then contract them back into a summary result.

## listToJson

Purpose: convert an Excel array into JSON list text.

### Linted Formula

```excel
=LAMBDA(
	arr,
	LET(
		quoted, MAP(arr, LAMBDA(x, jsonQuote(x))),
		joined, TEXTJOIN(",", TRUE, quoted),
		"[" & joined & "]"
	)
)
```

### How It Works

- Expands across the input with `MAP`.
- Normalizes every element through `jsonQuote`.
- Contracts the mapped values into one comma-separated string.

### Real-World Examples

```excel
=listToJson({"red";"green";"blue"})
=listToJson({5;10;15})
```

### Notation Notes

- Output is JSON text, not an Excel spilled array.
- Mixed text and numeric values are allowed.

## listFromJson

Purpose: split a flat JSON list into a vertical Excel array.

### Linted Formula

```excel
=LAMBDA(
	json,
	LET(
		inner, MID(TRIM(json), 2, LEN(TRIM(json)) - 2),
		parts, IF(inner = "", MAKEARRAY(0, 1, LAMBDA(r, c, "")), TEXTSPLIT(inner, ",", , TRUE)),
		TRIM(SUBSTITUTE(parts, """", ""))
	)
)
```

### How It Works

- Removes the outer brackets.
- Splits on commas.
- Dequotes each token.

### Real-World Examples

```excel
=listFromJson("[""north"",""south"",""west""]")
```

### Notation Notes

- Intended for simple flat lists.
- Nested lists or embedded commas are not preserved as grouped substructures.

## arrayRepAdd

Purpose: replace or append one key/value row in a two-column array.

### Linted Formula

```excel
=LAMBDA(
	arr,
	key,
	val,
	LET(
		safeArr, IF(ROWS(arr) = 0, HSTACK("", ""), arr),
		kcol, INDEX(safeArr, , 1),
		vcol, INDEX(safeArr, , 2),
		newarr, safeFilter(safeArr, kcol <> key),
		cleanArr, IF(ROWS(newarr) = 0, HSTACK("", ""), newarr),
		VSTACK(cleanArr, HSTACK(key, jsonQuote(val)))
	)
)
```

### How It Works

- Expands the two-column input into separate key and value columns.
- Removes any existing key match.
- Appends a normalized replacement row.

### Real-World Examples

If `arr` is:

| Key | Value |
|---|---|
| status | "open" |
| qty | 2 |

then:

```excel
=arrayRepAdd(arr,"qty",5)
```

### Notation Notes

- Used mainly as an internal helper by the JSON functions.
- Depends on `safeFilter`, which is not exported in the current `functions.json`.

## CountUnique

Purpose: produce a frequency table for the unique values in an array.

### Linted Formula

```excel
=LAMBDA(
	arr,
	LET(
		flat, TOCOL(arr, 1),
		uniques, UNIQUE(flat),
		counts, MAP(uniques, LAMBDA(u, SUM((flat = u) * 1))),
		HSTACK(uniques, counts)
	)
)
```

### How It Works

- Flattens the source with `TOCOL`.
- Uses `UNIQUE` to expand distinct values.
- Uses `MAP` to count each value against the flattened source.
- Returns a two-column summary.

### Real-World Examples

```excel
=CountUnique({"M8";"M8";"M10";"M8";"M10"})
```

### Notation Notes

- The result order follows Excel's `UNIQUE` output order.
- Useful for quick bill-of-materials summaries.

## GiveMostFrequent

Purpose: return the mode-like most common item in an array.

### Linted Formula

```excel
=LAMBDA(
	arr,
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

### How It Works

- Builds a unique list.
- Counts each candidate with `COUNTIF`.
- Sorts descending by count.
- Returns the first value.

### Real-World Examples

```excel
=GiveMostFrequent({"A";"B";"A";"C";"A"})
```

### Notation Notes

- If counts tie, the sort order determines which value is returned first.

## vLastItem

Purpose: return the last non-empty item from a vertical array.

### Linted Formula

```excel
=LAMBDA(
	array1,
	[emptyvalue],
	XLOOKUP(
		TRUE,
		array1 <> IF(ISOMITTED(emptyvalue), "", emptyvalue),
		array1,
		"N/A",
		0,
		-1
	)
)
```

### How It Works

- Builds a Boolean mask of acceptable cells.
- Uses `XLOOKUP` in reverse search mode.
- Returns the last matching row.

### Real-World Examples

```excel
=vLastItem({"cut";"drill";"";"ship"})
```

### Notation Notes

- The optional second argument changes what counts as empty.

## SelectFilter

Purpose: select a subset of columns and then filter rows in one call.

### Linted Formula

```excel
=LAMBDA(
	ArrayIn,
	ComSepColNumsInBraces,
	Filters,
	FILTER(
		INDEX(ArrayIn, SEQUENCE(ROWS(ArrayIn)), ComSepColNumsInBraces),
		Filters,
		NA()
	)
)
```

### How It Works

- Uses `INDEX` to project only the requested columns.
- Feeds that projection into `FILTER`.

### Real-World Examples

```excel
=SelectFilter(A1:D20,{1,4},C1:C20="Open")
```

### Notation Notes

- This is effectively a compact table projection plus row predicate.

## permutate

Purpose: evaluate a custom LAMBDA over every ordered pair from two arrays, then repeat in reverse order.

### Linted Formula

```excel
=LAMBDA(
	A,
	B,
	f,
	LET(
		nA, ROWS(A),
		nB, ROWS(B),
		forward,
			MAKEARRAY(
				nA * nB,
				1,
				LAMBDA(r, c, f(INDEX(A, MOD(r - 1, nA) + 1), INDEX(B, INT((r - 1) / nA) + 1)))
			),
		reverse,
			MAKEARRAY(
				nA * nB,
				1,
				LAMBDA(r, c, f(INDEX(B, MOD(r - 1, nB) + 1), INDEX(A, INT((r - 1) / nB) + 1)))
			),
		VSTACK(forward, reverse)
	)
)
```

### How It Works

- Uses `MAKEARRAY` twice to synthesize ordered pair positions.
- The forward pass enumerates `(A_i, B_j)`.
- The reverse pass enumerates `(B_j, A_i)`.
- Both blocks are vertically stacked.

### Real-World Examples

```excel
=permutate({"Bolt";"Nut"},{"Zinc";"Black"},LAMBDA(x,y,x&" - "&y))
```

### Notation Notes

- This returns ordered permutations, not unique combinations.
