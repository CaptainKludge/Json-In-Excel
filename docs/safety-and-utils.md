# Safety and utility helpers

General-purpose helpers and safety wrappers used across the JSON functions.

## safeDrop

Description: Drop rows from an array and return a safe default empty array when the drop would fail.

```excel
=LAMBDA(arr,rows,
    LET(
        dropped, IFERROR(DROP(arr, rows), MAKEARRAY(1, COLUMNS(arr), LAMBDA(r,c, ""))),
        out, dropped,
        out
    )
)
```

Notes: Prevents errors when DROP would remove too many rows.

## safeFilter

Description: Wrapper around FILTER that returns a safe empty array with correct column count instead of raising an error when no rows match.

```excel
=LAMBDA(arr,include,
    LET(
        cols, COLUMNS(arr),
        safe, IFERROR(FILTER(arr, include), MAKEARRAY(1, cols, LAMBDA(r,c, ""))),
        safe
    )
)
```

Notes: Keeps downstream code simpler by ensuring a consistent array shape.

## makearr

Description: Convert an input range/reference into a new array using MAKEARRAY to copy values. This can be helpful to ensure a dynamic array (not a range reference) is returned.

```excel
=LAMBDA(arr,
    MAKEARRAY(ROWS(arr), COLUMNS(arr), LAMBDA(a,b, INDEX(arr,a,b)))
)
```

## between

Description: Test whether a number falls into one or more intervals described in a string. Supports interval notation like "[0,10), (5,20]" and simple dash ranges like "0-10".

```excel
=LAMBDA(numberinpreclean,rangein,
    LET(
        parseset,
            LAMBDA(singleset,
                LET(
                    raw, TRIM(singleset),
                    numberin, VALUE(TRIM(numberinpreclean)),
                    safe, SUBSTITUTE(REGEXREPLACE(raw, "([(\\[].*?),(.*?[\\])])", "$1|$2"), " ", ""),
                    intervals, TEXTSPLIT(safe, ","),
                    parseInterval,
                        LAMBDA(tok,
                            LET(
                                t, TRIM(SUBSTITUTE(tok, "|", ",")),
                                hasBraces, OR(LEFT(t,1)="(", LEFT(t,1)="["),
                                dashParts, IF(hasBraces, "", TEXTSPLIT(t, "-")),
                                leftB, IF(hasBraces, LEFT(t,1), "["),
                                rightB, IF(hasBraces, RIGHT(t,1), "]"),
                                body, IF(hasBraces, MID(t,2,LEN(t)-2), t),
                                nums, IF(hasBraces, REGEXEXTRACT(body, "([-0-9]+(?:\\.[-0-9]+)?)?,([-0-9]+(?:\\.[-0-9]+)?)?", 2), ""),
                                minRaw, IF(hasBraces, IFERROR(INDEX(nums,1), -1E+99), IF(COUNTA(dashParts)>=1, TRIM(INDEX(dashParts,1)), "")),
                                maxRaw, IF(hasBraces, IFERROR(INDEX(nums,2), 1E+99), IF(COUNTA(dashParts)=2, TRIM(INDEX(dashParts,2)), "")),
                                min, IF(OR(minRaw="", minRaw="-∞"), -1E+99, VALUE(minRaw)),
                                max, IF(OR(maxRaw="", maxRaw="∞", maxRaw="+∞"), 1E+99, VALUE(maxRaw)),
                                lowerOK, IF(leftB="(", numberin>min, numberin>=min),
                                upperOK, IF(rightB=")", numberin<max, numberin<=max),
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

Notes: Useful for deciding configurable thresholds stored as strings.
