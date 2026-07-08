# JSON Functions

These functions treat JSON as text and then use Excel array operations to parse, traverse, rebuild, and merge that text. Every section below includes the stored function rewritten into a readable multi-line form, followed by notes on the actual implementation technique.

## jsonQuote

Purpose: normalize one Excel value into valid JSON literal text.

### Linted Formula

```excel
=LAMBDA(
    val,
    LET(
        s, TRIM(val),
        isObjOrArr, REGEXTEST(s, "^\s*(\{.*\}|\[.*\])\s*$"),
        isBoolNull, REGEXTEST(s, "^(?i:true|false|null)$"),
        isNum, REGEXTEST(s, "^-?\d+(\.\d+)?([eE][+-]?\d+)?$"),
        isProperQuoted, REGEXTEST(s, "^""[^""]*""$"),
        stripOuter, REGEXREPLACE(s, "^""+|""+$", ""),
        core, REGEXREPLACE(stripOuter, "("""")", "'"),
        result,
            IF(
                isObjOrArr,
                s,
                IF(
                    isNum,
                    s,
                    IF(
                        isBoolNull,
                        LOWER(s),
                        IF(isProperQuoted, s, """" & core & """")
                    )
                )
            ),
        result
    )
)
```

### How It Works

- Uses regex guards to classify the value before doing any quoting.
- Leaves existing objects, arrays, numbers, and JSON booleans/null intact.
- Strips outer quotes from ordinary text before rebuilding a clean quoted value.
- Replaces internal triple-quote patterns with single quotes to avoid Excel quoting collisions.

### Real-World Examples

```excel
=jsonQuote("Steel")
=jsonQuote("12.5")
=jsonQuote("false")
=jsonQuote("{""a"":1}")
```

Results:

```text
"Steel"
12.5
false
{"a":1}
```

### Notation Notes

- Text results come back with JSON quotes included.
- This function is a normalizer, not a full JSON escape engine.

## jsonObject

Purpose: build one JSON object from a two-column key/value array.

### Linted Formula

```excel
=LAMBDA(
    range,
    LET(
        keys, INDEX(range, , 1),
        vals, INDEX(range, , 2),
        nonEmptyRows, safeFilter(SEQUENCE(ROWS(keys)), keys <> ""),
        safeKeys, IFERROR(INDEX(keys, nonEmptyRows), ""),
        safeVals, IFERROR(INDEX(vals, nonEmptyRows), ""),
        pairStrings,
            MAP(
                safeKeys,
                safeVals,
                LAMBDA(k, v, IF(k <> "", jsonQuote(k) & ":" & v, ""))
            ),
        joined, TEXTJOIN(",", TRUE, safeFilter(pairStrings, pairStrings <> "")),
        IF(joined = "", "{}", "{" & joined & "}")
    )
)
```

### How It Works

- Treats column 1 as keys and column 2 as already-prepared JSON values.
- Filters out blank keys before assembly.
- Quotes keys with `jsonQuote` and joins each key/value pair with `:`.
- Uses `TEXTJOIN` to contract the row set into one object string.

The main trick is that the function does not quote values itself beyond what is already in column 2. That lets you pass raw numbers, nested objects, or arrays without flattening them into strings.

### Real-World Examples

```excel
=jsonObject(HSTACK({"sku";"qty";"active"},{"""A-100""";12;TRUE}))
```

Result:

```text
{"sku":"A-100","qty":12,"active":TRUE}
```

```excel
=jsonObject(HSTACK({"name";"tags"},{"""Bolt""";"[""steel"",""metric""]"}))
```

### Notation Notes

- Values must already be JSON-ready.

## jsonGetKeysAtLevel

Purpose: tokenize the top level of a JSON object into a two-column key/value array.

### Linted Formula

```excel
=LAMBDA(
    json,
    LET(
        content, IF(LEFT(json, 1) = "{", MID(json, 2, LEN(json) - 2), json),
        chars, MID(content, SEQUENCE(LEN(content)), 1),
        initialState, VSTACK("", 0, 0, FALSE, "-0-0-FALSE", ""),
        processed,
            REDUCE(
                initialState,
                chars,
                LAMBDA(
                    state,
                    ch,
                    LET(
                        token, INDEX(state, 1),
                        curl, INDEX(state, 2),
                        square, INDEX(state, 3),
                        quotes, INDEX(state, 4),
                        info, INDEX(state, 5),
                        tail, IF(ROWS(state) > 5, DROP(state, 5), ""),
                        isQuote, ch = CHAR(34),
                        newQuotes, IF(isQuote, NOT(quotes), quotes),
                        newCurl, IF(quotes, curl, IF(ch = "{", curl + 1, IF(ch = "}", curl - 1, curl))),
                        newSquare, IF(quotes, square, IF(ch = "[", square + 1, IF(ch = "]", square - 1, square))),
                        level, newCurl + newSquare,
                        isSplit, AND(ch = ",", level = 0, NOT(newQuotes)),
                        newToken, IF(isSplit, "", token & ch),
                        newTail, IF(isSplit, VSTACK(tail, token), tail),
                        newInfo, TEXTJOIN("-", TRUE, info, ";", isSplit, newToken, newCurl, newSquare, newQuotes),
                        VSTACK(newToken, newCurl, newSquare, newQuotes, newInfo, newTail)
                    )
                )
            ),
        rawPairs, IF(ROWS(processed) > 5, DROP(processed, 5), ""),
        lastToken, INDEX(processed, 1),
        combined, IF(lastToken <> "", VSTACK(rawPairs, lastToken), rawPairs),
        allPairs, IFERROR(FILTER(combined, combined <> ""), ""),
        keypairs,
            REDUCE(
                {"", ""},
                allPairs,
                LAMBDA(
                    acc,
                    pair,
                    LET(
                        hasColon, ISNUMBER(SEARCH(":", pair)),
                        sep, IF(hasColon, TEXTAFTER(pair, ":"), ""),
                        key, IF(hasColon, TEXTBEFORE(pair, ":"), pair),
                        cleanKey, TRIM(SUBSTITUTE(key, """", "")),
                        cleanVal, TRIM(sep),
                        VSTACK(acc, HSTACK(cleanKey, cleanVal))
                    )
                )
            ),
        IF(ROWS(keypairs) > 1, DROP(keypairs, 1), keypairs)
    )
)
```

### How It Works

- Implements a small lexer with `REDUCE` over every character in the JSON text.
- Tracks object depth, array depth, and quote state so commas inside nested structures do not split the token stream.
- Accumulates finished top-level pairs into the tail of the state stack.
- Runs a second reduction to split each pair on the first colon.

This is the core parser for the rest of the JSON helpers. The important trick is the expanding state vector carried through `REDUCE`, which simulates a parser without VBA or script code.

### Real-World Examples

```excel
=jsonGetKeysAtLevel("{""item"":{""sku"":""A1""},""qty"":5,""tags"":""[1,2]""}")
```

Conceptual result:

| Key | Raw value |
|---|---|
| item | {"sku":"A1"} |
| qty | 5 |
| tags | "[1,2]" |

### Notation Notes

- Returned values are raw JSON fragments, not automatically dequoted Excel values.
- The parser is object-oriented, not a full general JSON document parser.

## jsonGet

Purpose: read a nested JSON value by slash-delimited path.

### Linted Formula

```excel
=LAMBDA(
    json,
    path,
    LET(
        keys, TEXTSPLIT(path, "/"),
        REDUCE(
            json,
            keys,
            LAMBDA(
                j,
                k,
                LET(
                    pairs, jsonGetKeysAtLevel(j),
                    vals, IFERROR(FILTER(pairs, INDEX(pairs, , 1) = k), ""),
                    IF(OR(vals = "", ROWS(vals) = 0), NA(), TEXTJOIN(",", TRUE, INDEX(vals, , 2)))
                )
            )
        )
    )
)
```

### How It Works

- Splits the path into segments.
- Uses `REDUCE` to walk segment by segment.
- Re-parses the current JSON fragment at each step with `jsonGetKeysAtLevel`.
- Contracts any matching value rows into a single text result with `TEXTJOIN`.

### Real-World Examples

```excel
=jsonGet("{""order"":{""customer"":{""name"":""Ana""}}}","order/customer/name")
```

Result:

```text
"Ana"
```

```excel
=jsonGet("{""config"":{""retry"":3}}","config/retry")
```

### Notation Notes

- Paths use `/` as the level separator.
- Missing paths return `#N/A`.

## jsonSet

## nestedJsonBuild

Purpose: build a nested JSON object chain from a slash-delimited path and a final value.

### Linted Formula

```excel
=LAMBDA(
    p,
    v,
    LET(
        parts, TEXTSPLIT(p, , "/", TRUE),
        IF(
            OR(p = "", ROWS(parts) = 0),
            jsonQuote(v),
            LET(
                rev, INDEX(parts, SEQUENCE(ROWS(parts), 1, ROWS(parts), -1)),
                REDUCE(
                    jsonQuote(v),
                    rev,
                    LAMBDA(acc, layer, jsonObject(HSTACK(layer, acc)))
                )
            )
        )
    )
)
```

### How It Works

- Splits the input path into segments.
- Reverses the segment order.
- Starts with the quoted final value.
- Uses `REDUCE` to wrap that value one layer at a time from the inside out.

This is the object-construction counterpart to `jsonGet`. Instead of walking downward through an existing object, it synthesizes the missing branch upward by repeatedly nesting the current accumulator.

### Real-World Examples

```excel
=nestedJsonBuild("user/profile/email","""ana@example.com""")
```

Result:

```text
{"user":{"profile":{"email":"ana@example.com"}}}
```

### Notation Notes

- Used mainly by `jsonSet` when intermediate objects do not already exist.
- An empty path returns `jsonQuote(v)` directly.

## jsonSet

Purpose: write or replace a nested JSON value at a slash-delimited path.

### Linted Formula

```excel
=LAMBDA(
    oJson,
    oPath,
    oValue,
    LET(
        MAX_DEPTH, 10,
        walk,
            LAMBDA(
                J,
                P,
                V,
                depth,
                self,
                IF(
                    depth > MAX_DEPTH,
                    "#STOP@" & P,
                    LET(
                        set, jsonGetKeysAtLevel(J),
                        parts, TEXTSPLIT(P, , "/", TRUE),
                        n, ROWS(parts),
                        hasTail, n > 1,
                        key, INDEX(parts, 1),
                        tail, IF(hasTail, TEXTJOIN("/", , DROP(parts, 1)), ""),
                        keys, IFERROR(INDEX(set, , 1), MAKEARRAY(1, 1, LAMBDA(r, c, ""))),
                        keypresent, SUM(--(keys = key)) > 0,
                        curRaw, IF(keypresent, jsonGet(J, key), ""),
                        isObj, AND(keypresent, LEFT(TRIM(curRaw), 1) = "{"),
                        newValQ, jsonQuote(V),
                        result,
                            SWITCH(
                                TRUE,
                                AND(hasTail, keypresent, isObj), jsonObject(arrayRepAdd(set, key, self(curRaw, tail, V, depth + 1, self))),
                                AND(hasTail, keypresent, NOT(isObj)), jsonObject(arrayRepAdd(set, key, nestedJsonBuild(tail, V))),
                                hasTail, jsonObject(arrayRepAdd(set, key, nestedJsonBuild(tail, V))),
                                AND(NOT(hasTail), keypresent), jsonObject(arrayRepAdd(set, key, newValQ)),
                                AND(NOT(hasTail), NOT(keypresent)), jsonObject(VSTACK(set, HSTACK(key, newValQ))),
                                ISNA(set), "#SETERR:notObject@" & P,
                                TRUE, "#SETERR:unhandled@" & P
                            ),
                        result
                    )
                )
            ),
        walk(oJson, oPath, oValue, 0, walk)
    )
)
```

### How It Works

- Defines a recursive worker `walk` inside `LET`.
- Splits the path into head and tail on every recursion.
- Re-enters itself when the current branch already contains a nested object.
- Falls back to `nestedJsonBuild` when a missing branch has to be created.
- Uses `arrayRepAdd` and `jsonObject` to rebuild the object after each recursive mutation.

The trick here is recursive object rebuilding. Excel has no mutable JSON object, so every write is actually a remove-and-reconstruct pass on the key/value table for the current level.

### Real-World Examples

```excel
=jsonSet("{""settings"":{""theme"":""light""}}","settings/theme","""dark""")
```

Result:

```text
{"settings":{"theme":"dark"}}
```

```excel
=jsonSet("{}","user/email","""ana@example.com""")
```

### Notation Notes

- Recursion is capped at depth 10.

## jsonRemove

Purpose: delete a key or nested key from a JSON object.

### Linted Formula

```excel
=LAMBDA(
    oJson,
    oPath,
    LET(
        walk,
            LAMBDA(
                J,
                P,
                self,
                depth,
                IF(
                    depth > 20,
                    "STOP@" & P,
                    LET(
                        set, jsonGetKeysAtLevel(J),
                        parts, TEXTSPLIT(P, , "/", TRUE),
                        n, ROWS(parts),
                        hasTail, n > 1,
                        key, INDEX(parts, 1),
                        tail, IF(hasTail, TEXTJOIN("/", , DROP(parts, 1)), ""),
                        keys, IFERROR(INDEX(set, , 1), MAKEARRAY(1, 1, LAMBDA(r, c, ""))),
                        keypresent, SUM(--(keys = key)) > 0,
                        curRaw, IF(keypresent, jsonGet(J, key), NA()),
                        isObj, AND(NOT(ISNA(curRaw)), LEFT(TRIM(curRaw), 1) = "{"),
                        result,
                            SWITCH(
                                TRUE,
                                AND(hasTail, keypresent, isObj), jsonObject(arrayRepAdd(set, key, self(curRaw, tail, self, depth + 1))),
                                NOT(hasTail), jsonObject(safeFilter(set, keys <> key)),
                                J
                            ),
                        result
                    )
                )
            ),
        walk(oJson, oPath, walk, 0)
    )
)
```

### How It Works

- Uses the same recursive head/tail path walk pattern as `jsonSet`.
- When the target is nested, removes inside the child object first and then reconstructs the parent.
- When the target is local, filters the key out of the current level and rebuilds the object.

### Real-World Examples

```excel
=jsonRemove("{""part"":{""sku"":""A1"",""obsolete"":true}}","part/obsolete")
```

Result:

```text
{"part":{"sku":"A1"}}
```

### Notation Notes

- Recursion is capped at depth 20.

## jsonJoin

Purpose: merge JSON objects using replace, recursive merge, append, or additive behavior depending on mode and value types.

### Linted Formula

```excel
=LAMBDA(
    StartingJson,
    Additions,
    MODE,
    LET(
        solver,
            LAMBDA(
                base,
                new,
                LET(
                    WALK,
                        LAMBDA(
                            JSONE,
                            JSTWO,
                            SUBMODE,
                            DEPTH,
                            SELF,
                            IF(
                                DEPTH > 10,
                                "#STOP:recursiontoodeep",
                                LET(
                                    JsonSet1, jsonGetKeysAtLevel(JSONE),
                                    JsonSet2, jsonGetKeysAtLevel(JSTWO),
                                    ReduceSET,
                                        MAP(
                                            INDEX(JsonSet2, , 1),
                                            INDEX(JsonSet2, , 2),
                                            LAMBDA(k, v, jsonObject(HSTACK(k, v)))
                                        ),
                                    RESULT,
                                        REDUCE(
                                            JsonSet1,
                                            ReduceSET,
                                            LAMBDA(
                                                returnset,
                                                testsetobj,
                                                LET(
                                                    objkeyset, jsonGetKeysAtLevel(testsetobj),
                                                    key, INDEX(objkeyset, 1, 1),
                                                    value, INDEX(objkeyset, 1, 2),
                                                    keys, INDEX(returnset, , 1),
                                                    keypresent, SUM(--(keys = key)) > 0,
                                                    oldvalue, IF(keypresent, INDEX(FILTER(returnset, INDEX(returnset, , 1) = key), 1, 2), NA()),
                                                    test1, IF(ISNA(oldvalue), "NA", LEFT(TRIM(oldvalue), 1)),
                                                    test2, IF(ISNA(value), "NA", LEFT(TRIM(value), 1)),
                                                    isObjectPair, IFERROR(AND(test1 = "{", test2 = "{"), FALSE),
                                                    isListPair, IFERROR(AND(test1 = "[", test2 = "["), FALSE),
                                                    repMode, SUBMODE = 1,
                                                    addMode, SUBMODE = 2,
                                                    SWITCH(
                                                        TRUE,
                                                        AND(keypresent, NOT(repMode), isObjectPair), arrayRepAdd(returnset, key, SELF(oldvalue, value, SUBMODE, DEPTH + 1, SELF)),
                                                        AND(keypresent, NOT(repMode), isListPair), arrayRepAdd(returnset, key, listToJson(VSTACK(listFromJson(oldvalue), listFromJson(value)))),
                                                        AND(addMode, XOR(test1 = "[", test2 = "[")),
                                                            IF(
                                                                test1 = "[",
                                                                arrayRepAdd(returnset, key, listToJson(VSTACK(listFromJson(oldvalue), value))),
                                                                arrayRepAdd(returnset, key, listToJson(VSTACK(oldvalue, listFromJson(value))))
                                                            ),
                                                        AND(keypresent, addMode), arrayRepAdd(returnset, key, IF(AND(NOT(ISERROR(VALUE(oldvalue))), NOT(ISERROR(VALUE(value)))), VALUE(oldvalue) + VALUE(value), oldvalue & value)),
                                                        arrayRepAdd(returnset, key, value)
                                                    )
                                                )
                                            )
                                        ),
                                    jsonObject(RESULT)
                                )
                            )
                        ),
                    WALK(base, new, MODE, 0, WALK)
                )
            ),
        REDUCE(StartingJson, Additions, solver)
    )
)
```

### How It Works

- Builds a recursive merge engine `WALK`.
- Converts the incoming object into a stream of one-key test objects.
- Uses `REDUCE` to fold each incoming key into the accumulated result set.
- Detects object-vs-object and array-vs-array cases by looking at the first non-space character.
- In add mode, performs arithmetic addition for numeric pairs and concatenation for text pairs.

This is the most algorithmic JSON helper in the set. It relies on repeated expansion to key/value tables and repeated contraction back to object strings.

### Real-World Examples

```excel
=jsonJoin("{""A"":2,""B"":1}","{""A"":3}",2)
```

Result:

```text
{"A":5,"B":1}
```

```excel
=jsonJoin("{""cfg"":{""x"":1}}","{""cfg"":{""y"":2}}",0)
```

Result:

```text
{"cfg":{"x":1,"y":2}}
```

### Notation Notes

- `MODE = 0` does recursive merge and list append.
- `MODE = 1` forces replace behavior.
- `MODE = 2` enables additive or append-like behavior.

```
jsonRemove("{"person":{"name":"Alice","age":30}}", "person/age")
    -> "{"person":{"name":"Alice"}}"
```

## nestedJsonBuild

Description: Build nested JSON objects from a slash-separated path and a value. Used internally by `jsonSet` to create deeply nested structures.

```excel
=LAMBDA(p,v,
    LET(
        parts, TEXTSPLIT(p,,"/",TRUE),
        IF(
            OR(p="", ROWS(parts)=0),
            jsonQuote(v),
            LET(
                rev, INDEX(parts, SEQUENCE(ROWS(parts),1,ROWS(parts),-1)),
                REDUCE(
                    jsonQuote(v),
                    rev,
                    LAMBDA(acc,layer, jsonObject(HSTACK(layer, acc)))
                )
            )
        )
    )
)
```

Notes: This function works by reversing the path components and building nested objects from the inside out. Essential for creating deep object structures in a single operation.

### Example

```
nestedJsonBuild("user/profile/settings", "enabled")
    -> "{"user":{"profile":{"settings":"enabled"}}}"
```

## Function Relationships

These JSON functions work together as an integrated system:

- **`jsonQuote`** → Safely formats values for JSON inclusion
- **`jsonObject`** → Assembles key-value pairs into JSON objects  
- **`jsonGetKeysAtLevel`** → Parses objects for navigation and manipulation
- **`jsonGet`** → Retrieves values using path notation
- **`jsonSet`** → Modifies or creates nested values (uses `nestedJsonBuild`)
- **`jsonJoin`** → Merges multiple JSON objects with conflict resolution
- **`jsonRemove`** → Deletes keys while preserving structure
- **`nestedJsonBuild`** → Creates deep object hierarchies from paths

This comprehensive toolkit enables sophisticated JSON manipulation directly within Excel formulas, supporting complex data structures and operations that would traditionally require external tools.
