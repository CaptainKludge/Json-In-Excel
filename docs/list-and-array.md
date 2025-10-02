# List and array helpers

Functions to convert between lists, arrays and JSON list representations.

## listToJson

Description: Convert a 1-D array (range) into a JSON array string. Each element is quoted via `jsonQuote`.

```excel
=LAMBDA(arr,
    LET(
        quoted, MAP(arr, LAMBDA(x, jsonQuote(x))),
        joined, TEXTJOIN(",", TRUE, quoted),
        "[" & joined & "]"
    )
)
```

Notes: Useful to build JSON arrays from Excel ranges.

## listFromJson

Description: Convert a JSON array string into an Excel vertical array of unquoted values. Returns an empty array for an empty JSON array.

```excel
=LAMBDA(json,
    LET(
        inner, MID(TRIM(json),2,LEN(TRIM(json))-2),
        parts, IF(inner="", MAKEARRAY(0,1,LAMBDA(r,c,"")), TEXTSPLIT(inner, ",",,TRUE)),
        TRIM(SUBSTITUTE(parts, """", ""))
    )
)
```

Notes: This is a simple parser that assumes there are no nested arrays/objects containing commas in the values. Use with arrays of simple values.

## arrayRepAdd

Description: Helper used by JSON builders to add or replace a key/value pair in a two-column key/value array. Ensures no duplicate keys remain.

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

Notes: Used internally by `jsonSet`, `jsonJoin`, and `jsonRemove`.

## nestedJsonBuild

Description: Build nested JSON object structure from a slash-separated path and a value. Example: path "a/b" and value 1 -> {"a":{"b":1}}

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

Notes: Used by `jsonSet` to construct nested values when intermediate keys are missing.

## GREEDYPARTFILL and partFill

Description: Helpers to compute how many of each "part" fit into a target length. `GREEDYPARTFILL` is a more compact implementation; `partFill` uses the project's JSON helpers to return state objects.

```excel
=LAMBDA(targetLength,partArray,[pad],[extraHungry],
    LET(
        padVal, IF(ISOMITTED(pad), 0, pad),
        isExtraHungry, IF(ISOMITTED(extraHungry), FALSE, extraHungry),
        totalLength, targetLength + padVal,
        partNames, INDEX(partArray,,1),
        partLengths, INDEX(partArray,,2),
        numParts, ROWS(partArray),
        fillLoop,
            LAMBDA(index,state,
                LET(
                    remaining, INDEX(state,1),
                    used, INDEX(state,2),
                    name, INDEX(partNames, index),
                    length, INDEX(partLengths, index),
                    count, IF(length > 0, INT(remaining / length), 0),
                    newRemain, IF(length > 0, remaining - count * length, remaining),
                    newUsed, IF(count > 0, HSTACK(used, HSTACK(name, count)), used),
                    VSTACK(newRemain, newUsed)
                )
            ),
        reduceResult, REDUCE(VSTACK(totalLength, ""), SEQUENCE(numParts), fillLoop),
        leftover, INDEX(reduceResult, 1),
        usedRaw, INDEX(reduceResult, 2),
        lastName, INDEX(partNames, numParts),
        lastLength, INDEX(partLengths, numParts),
        adjustedUsed,
            IF(
                isExtraHungry * (leftover > 0) * (lastLength > leftover),
                LET(
                    matchIndex, XMATCH(lastName, INDEX(usedRaw,,1), 0),
                    oldCount, IF(ISNUMBER(matchIndex), INDEX(usedRaw,,2, matchIndex), 0),
                    newCount, oldCount + 1,
                    namesFiltered, IFERROR(FILTER(INDEX(usedRaw,,1), INDEX(usedRaw,,1)<>lastName), ""),
                    countsFiltered, IFERROR(FILTER(INDEX(usedRaw,,2), INDEX(usedRaw,,1)<>lastName), ""),
                    finalNames, IF(namesFiltered="", lastName, VSTACK(namesFiltered, lastName)),
                    finalCounts, IF(countsFiltered="", newCount, VSTACK(countsFiltered, newCount)),
                    HSTACK(finalNames, finalCounts)
                ),
                usedRaw
            ),
        resultText, TEXTJOIN(",", TRUE, MAP(INDEX(adjustedUsed,,1), INDEX(adjustedUsed,,2), LAMBDA(n,c, TEXTJOIN(",", TRUE, SEQUENCE(c,,n))))),
        resultText
    )
)
```

```excel
=LAMBDA(span,partarr,[extrahungry],
    LET(
        num, ROWS(partarr),
        initialState,
            jsonObject(HSTACK(VSTACK("rem","out","row"), VSTACK(span, "{ }", "0"))),
        rowarray,
            MAKEARRAY(
                num, 1,
                LAMBDA(rown,ind,
                    jsonObject(
                        HSTACK(
                            VSTACK("name","len","row"),
                            VSTACK(INDEX(partarr,rown,1), INDEX(partarr,rown,2), rown)
                        )
                    )
                )
            ),
        finalState,
            REDUCE(
                initialState,
                rowarray,
                LAMBDA(state,partJson,
                    LET(
                        rem, VALUE(IFERROR(jsonGet(state,"rem"), "#BROKEN:rem")),
                        outJson, IFERROR(jsonGet(state,"out"), "#MISSING:out"),
                        rownum, VALUE(IFERROR(jsonGet(state,"row"), "#BROKEN:row")),
                        pname, IFERROR(jsonGet(partJson,"name"), "#BROKEN:name"),
                        plen, VALUE(IFERROR(jsonGet(partJson,"len"), "#BROKEN:len")),
                        lastpart, IFERROR(jsonGet(state,"lastpart"), "#BROKEN:lastpart"),
                        lenlpart, IFERROR(jsonGet(state,"lenlpart"), "#BROKEN:lenlpart"),
                        pcount, INT(rem/plen),
                        newRem, rem - pcount*plen,
                        fcount, IF(AND(NOT(ISNA(additionalconfig)), newRem<0, pcount>0), pcount-1, pcount),
                        finalRem, rem - fcount*plen,
                        newOut,
                            IF(NUMBERVALUE(fcount)>0, IFERROR(jsonSet(outJson,pname,fcount), "#SETERR:newOut:" & pname & ":" & fcount), outJson),
                        IFERROR(
                            jsonSet(IFERROR(jsonSet(state,"rem",finalRem), "#SETERR:rem"), "out", newOut),
                            "#SETERR:out"
                        )
                    )
                )
            ),
        additionalconfig,
            IF(ISOMITTED(extrahungry), jsonGet(finalState,"out"), jsonJoin(jsonGet(finalState,"out"), FILTER(INDEX(extrahungry,,1), between(jsonGet(finalState,"rem"), INDEX(extrahungry,,2))), 2)),
        additionalconfig
    )
)
```
