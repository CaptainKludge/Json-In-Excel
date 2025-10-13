# JSON helper functions

This document groups functions that create, manipulate, and query JSON strings in Excel formulas.

## jsonObject

Description: Build a JSON object from a two-column range of keys and values. Empty keys are ignored.

Formula (paste into Excel Name Manager as the LAMBDA body):

```excel
=LAMBDA(range,
    LET(
        keys, INDEX(range,,1),
        vals, INDEX(range,,2),
        nonEmptyRows, safeFilter(SEQUENCE(ROWS(keys)), keys<>""),
        safeKeys, IFERROR(INDEX(keys, nonEmptyRows), ""),
        safeVals, IFERROR(INDEX(vals, nonEmptyRows), ""),
        pairStrings,
            MAP(
                safeKeys,
                safeVals,
                LAMBDA(k,v,
                    IF(k<>":", jsonQuote(k) & ":" & v, "")
                )
            ),
        joined, TEXTJOIN(",", TRUE, safeFilter(pairStrings, pairStrings<>"")),
        IF(joined="", "{}", "{" & joined & "}")
    )
)
```

Notes: Accepts a two-column input where the first column is the key and the second column is the JSON-ready value (already quoted or a literal JSON fragment). Uses `jsonQuote` to quote keys.

### Example

Inputs (two-column range):

| Key  | Value  |
|---|---|
| name | "Alice" |
| age  | 30 |

Result (JSON):

```
{"name":"Alice","age":30}
```

## jsonQuote

Description: Safely quote or normalize a value for inclusion in JSON. Detects objects, arrays, booleans/null and numbers and leaves them unquoted. Wraps other values in double-quotes and escapes internal quotes.

```excel
=LAMBDA(val,
    LET(
        s, TRIM(val),
        isObjOrArr, REGEXTEST(s, "^\\s*(\\{.*\\}|\\[.*\\])\\s*$"),
        isBoolNull, REGEXTEST(s, "^(?i:true|false|null)$"),
        isNum, REGEXTEST(s, "^-?\\d+(\\.\\d+)?([eE][+-]?\\d+)?$"),
        isProperQuoted, REGEXTEST(s, "^""[^""]*""$"),
        stripOuter, REGEXREPLACE(s, "^""+|""+$", ""),
        core, REGEXREPLACE(stripOuter, "("""")", "'"),
        result,
            IF(
                isObjOrArr, s,
                IF(
                    isNum, s,
                    IF(isBoolNull, LOWER(s), IF(isProperQuoted, s, """" & core & """"))
                )
            ),
        result
    )
)
```

Notes: This function keeps JSON arrays/objects and numbers/boolean/null unmodified, surrounds text with quotes and converts interior triple double-quotes to single quotes to avoid breaking Excel string quoting.

### Example

```
jsonQuote("Bob")    -> "\"Bob\""
jsonQuote("123")    -> "123"
jsonQuote("{\"a\":1}") -> "{\"a\":1}"
```

## jsonGetKeysAtLevel

Description: Parse a JSON object string and return a two-column array of key/value raw strings at the top level of that object. Useful internal helper for many other JSON operations.

```excel
=LAMBDA(json,
    LET(
        content, IF(LEFT(json,1)="{", MID(json,2,LEN(json)-2), json),
        chars, MID(content, SEQUENCE(LEN(content)), 1),
        initialState, VSTACK("", 0, 0, FALSE, "-0-0-FALSE",""),
        processed,
            REDUCE(
                initialState,
                chars,
                LAMBDA(state,ch,
                    LET(
                        token, INDEX(state,1),
                        curl, INDEX(state,2),
                        square, INDEX(state,3),
                        quotes, INDEX(state,4),
                        info, INDEX(state,5),
                        tail, IF(ROWS(state)>5, DROP(state,5), ""),
                        isQuote, ch=CHAR(34),
                        newQuotes, IF(isQuote, NOT(quotes), quotes),
                        newCurl, IF(quotes, curl, IF(ch="{", curl+1, IF(ch="}", curl-1, curl))),
                        newSquare, IF(quotes, square, IF(ch="[", square+1, IF(ch="]", square-1, square))),
                        level, newCurl+newSquare,
                        isSplit, AND(ch=",", level=0, NOT(newQuotes)),
                        newToken, IF(isSplit, "", token&ch),
                        newTail, IF(isSplit, VSTACK(tail, token), tail),
                        newInfo, TEXTJOIN("-",TRUE,info,";",isSplit,newToken,newCurl,newSquare,newQuotes),
                        VSTACK(newToken,newCurl,newSquare,newQuotes,newInfo,newTail)
                    )
                )
            ),
        rawPairs, IF(ROWS(processed)>5, DROP(processed,5), ""),
        lastToken, INDEX(processed,1),
        combined, IF(lastToken<>"", VSTACK(rawPairs,lastToken), rawPairs),
        allPairs, IFERROR(FILTER(combined, combined<>""), ""),
        keypairs,
            REDUCE(
                {"",""},
                allPairs,
                LAMBDA(acc,pair,
                    LET(
                        hasColon, ISNUMBER(SEARCH(":",pair)),
                        sep, IF(hasColon, TEXTAFTER(pair,":"), ""),
                        key, IF(hasColon, TEXTBEFORE(pair,":"), pair),
                        cleanKey, TRIM(SUBSTITUTE(key, """","")),
                        cleanVal, TRIM(sep),
                        VSTACK(acc,HSTACK(cleanKey,cleanVal))
                    )
                )
            ),
        IF(ROWS(keypairs)>1, DROP(keypairs,1), keypairs)
    )
)
```

Notes: This is a robust lexer for simple JSON key/value extraction at a single object level. It handles nested brackets and quoted strings.

## jsonGet

Description: Retrieve a value from a JSON object string given a slash-separated path (e.g. "parent/child"). Returns NA() when the path cannot be resolved.

```excel
=LAMBDA(json,path,
    LET(
        keys, TEXTSPLIT(path,"/"),
        REDUCE(
            json,
            keys,
            LAMBDA(j,k,
                LET(
                    pairs, jsonGetKeysAtLevel(j),
                    vals, IFERROR(FILTER(pairs, INDEX(pairs,,1)=k), ""),
                    IF(OR(vals="", ROWS(vals)=0), NA(), TEXTJOIN(",",TRUE,INDEX(vals,,2)))
                )
            )
        )
    )
)
```

Notes: Uses `jsonGetKeysAtLevel` repeatedly to traverse nested objects.

### Example

```
jsonGet("{"person":{"name":"Alice","age":30}}","person/name")
	-> "Alice"
```

## jsonSet

Description: Set or replace a value at a given slash-separated path inside a JSON object string. Creates nested objects as needed. Returns an updated JSON string or an error token beginning with "#SETERR:" on failure.

```excel
=LAMBDA(oJson,oPath,oValue,
    LET(
        MAX_DEPTH, 10,
        walk,
            LAMBDA(J,P,V,depth,self,
                IF(
                    depth>MAX_DEPTH,
                    "#STOP@"&P,
                    LET(
                        set, jsonGetKeysAtLevel(J),
                        parts, TEXTSPLIT(P,,"/",TRUE),
                        n, ROWS(parts),
                        hasTail, n>1,
                        key, INDEX(parts,1),
                        tail, IF(hasTail, TEXTJOIN("/",,DROP(parts,1)), ""),
                        keys, IFERROR(INDEX(set,,1), MAKEARRAY(1,1,LAMBDA(r,c,""))),
                        keypresent, SUM(--(keys=key))>0,
                        curRaw, IF(keypresent, jsonGet(J,key), ""),
                        isObj, AND(keypresent, LEFT(TRIM(curRaw),1)="{"),
                        newValQ, jsonQuote(V),
                        result,
                            SWITCH(
                                TRUE,
                                AND(hasTail, keypresent, isObj),
                                    jsonObject(arrayRepAdd(set,key,self(curRaw,tail,V,depth+1,self))),
                                AND(hasTail, keypresent, NOT(isObj)),
                                    jsonObject(arrayRepAdd(set,key,nestedJsonBuild(tail,V))),
                                hasTail,
                                    jsonObject(arrayRepAdd(set,key,nestedJsonBuild(tail,V))),
                                AND(NOT(hasTail), keypresent),
                                    jsonObject(arrayRepAdd(set,key,newValQ)),
                                AND(NOT(hasTail), NOT(keypresent)),
                                    jsonObject(VSTACK(set,HSTACK(key,newValQ))),
                                ISNA(set),
                                    "#SETERR:notObject@"&P,
                                TRUE,
                                    "#SETERR:unhandled@"&P
                            ),
                        result
                    )
                )
            )
    ),
    walk(oJson,oPath,oValue,0,walk)
)
```

Notes: This function is recursive and builds objects as required. `nestedJsonBuild` is used to construct nested value structures.

### Example

Starting JSON:

```
{"person":{"name":"Alice","age":30}}
```

Insert an email:

```
jsonSet("{""person"":{""name"":""Alice"",""age"":30}}","person/email","""alice@example.com""")
```

Result:

```
{"person":{"name":"Alice","age":30,"email":"alice@example.com"}}
```

## jsonJoin

Description: Merge one or more JSON objects or arrays into another with several modes (replace, add, append). This is a higher-level merge routine and is intentionally complex.

```excel
=LAMBDA(
    JSON1, JSONNEW, MODE,
    LET(
        solver,
            LAMBDA(JSON2,
                LET(
                    WALK,
                        LAMBDA(JSONE,JSTWO,SUBMODE,DEPTH,SELF,
                            IF(
                                DEPTH>10,
                                "#STOP:recursiontoodeep",
                                LET(
                                    JsonSet1, jsonGetKeysAtLevel(JSONE),
                                    JsonSet2, jsonGetKeysAtLevel(JSTWO),
                                    ReduceSET, MAP(INDEX(JsonSet2,,1), INDEX(JsonSet2,,2), LAMBDA(ARRAYONE,ARRAYTWO, jsonObject(HSTACK(ARRAYONE,ARRAYTWO)))),
                                    RESULT,
                                        REDUCE(
                                            JsonSet1,
                                            ReduceSET,
                                            LAMBDA(returnset,testsetobj,
                                                LET(
                                                    objkeyset, jsonGetKeysAtLevel(testsetobj),
                                                    key, INDEX(objkeyset,1,1),
                                                    value, INDEX(objkeyset,1,2),
                                                    keys, INDEX(returnset,,1),
                                                    values, INDEX(returnset,,2),
                                                    keypresent, SUM(--(keys=key))>0,
                                                    oldvalue, IF(keypresent, INDEX(FILTER(returnset, INDEX(returnset,,1)=key),1,2), NA()),
                                                    test1, IF(ISNA(oldvalue), "NA", LEFT(TRIM(oldvalue),1)),
                                                    test2, IF(ISNA(value), "NA", LEFT(TRIM(value),1)),
                                                    isObjectPair, IFERROR(AND(test1="{", test2="{"), FALSE),
                                                    isListPair, IFERROR(AND(test1="[", test2="["), FALSE),
                                                    normmode, SUBMODE=0,
                                                    repMode, SUBMODE=1,
                                                    addMode, SUBMODE=2,
                                                    SWITCH(
                                                        TRUE,
                                                        AND(keypresent, NOT(repMode), isObjectPair), arrayRepAdd(returnset, key, SELF(oldvalue, value, SUBMODE, DEPTH+1, SELF)),
                                                        AND(keypresent, NOT(repMode), isListPair), arrayRepAdd(returnset, key, listToJson(VSTACK(listFromJson(oldvalue), listFromJson(value)))),
                                                        AND(addMode, XOR(test1="[", test2="[")), IF(test1="[", arrayRepAdd(returnset, key, listToJson(VSTACK(listFromJson(oldvalue), value))), arrayRepAdd(returnset, key, listToJson(VSTACK(oldvalue, listFromJson(value))))),
                                                        AND(keypresent, addMode), arrayRepAdd(returnset, key, IF(AND(NOT(ISERROR(VALUE(oldvalue))), NOT(ISERROR(VALUE(value)))), VALUE(oldvalue)+VALUE(value), oldvalue & value)),
                                                        arrayRepAdd(returnset, key, value)
                                                    )
                                                )
                                            )
                                        ),
                                    jsonObject(RESULT)
                                )
                            )
                        )
                ),
                MAP(JSONNEW, solver)
            )
    )
)
```

Notes: Because of its length and complexity, review before using. Mode determines how values combine (0/1/2). Use conservatively and test on sample JSON.

## jsonRemove

Description: Remove a key (or nested key path) from a JSON object string. Returns an updated JSON string.

```excel
=LAMBDA(oJson,oPath,
    LET(
        walk,
            LAMBDA(J,P,self,depth,
                IF(
                    depth>20,
                    "STOP@"&P,
                    LET(
                        set, jsonGetKeysAtLevel(J),
                        parts, TEXTSPLIT(P,,"/",TRUE),
                        n, ROWS(parts),
                        hasTail, n>1,
                        key, INDEX(parts,1),
                        tail, IF(hasTail, TEXTJOIN("/",,DROP(parts,1)), ""),
                        keys, IFERROR(INDEX(set,,1), MAKEARRAY(1,1,LAMBDA(r,c,""))),
                        keypresent, SUM(--(keys=key))>0,
                        curRaw, IF(keypresent, jsonGet(J,key), NA()),
                        isObj, AND(NOT(ISNA(curRaw)), LEFT(TRIM(curRaw),1)="{"),
                        result,
                            SWITCH(
                                TRUE,
                                AND(hasTail, keypresent, isObj), jsonObject(arrayRepAdd(set, key, self(curRaw, tail, self, depth+1))),
                                NOT(hasTail), jsonObject(safeFilter(set, keys<>key)),
                                J
                            ),
                        result
                    )
                )
            )
    ),
    walk(oJson,oPath,walk,0)
)
```

Notes: Recursive remover that preserves other keys.
