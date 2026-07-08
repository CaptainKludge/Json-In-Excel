# Algorithms

The current `functions.json` export contains one algorithmic function. Unlike the JSON helpers, this function uses JSON as a state container while it runs a sequential fill algorithm.

## partFill

Purpose: allocate named parts against a target span in row order, then optionally apply remainder-based post adjustments.

### Linted Formula

```excel
=LAMBDA(
  span,
  partarr,
  [additionalconfig],
  LET(
    num, ROWS(partarr),
    usingExtra, NOT(ISOMITTED(additionalconfig)),
    initialState,
      jsonObject(
        HSTACK(
          VSTACK("rem", "out"),
          VSTACK(span, "{}")
        )
      ),
    finalState,
      REDUCE(
        initialState,
        IF(usingExtra, SEQUENCE(1), SEQUENCE(num)),
        LAMBDA(
          state,
          r,
          LET(
            rem, VALUE(jsonGet(state, "rem")),
            outJson, jsonGet(state, "out"),
            pname, INDEX(partarr, r, 1),
            plen, VALUE(INDEX(partarr, r, 2)),
            isLastPart, r = num,
            baseCount, QUOTIENT(rem, plen),
            needsExtra, MOD(rem, plen) > 0,
            fcount,
              IF(
                isLastPart,
                IF(usingExtra, baseCount, IF(needsExtra, baseCount + 1, baseCount)),
                baseCount
              ),
            finalRem, rem - fcount * plen,
            newOut, IF(fcount > 0, jsonSet(outJson, pname, fcount), outJson),
            jsonSet(jsonSet(state, "rem", finalRem), "out", newOut)
          )
        )
      ),
    IF(
      usingExtra,
      jsonJoin(
        jsonGet(finalState, "out"),
        FILTER(INDEX(additionalconfig, , 1), between(jsonGet(finalState, "rem"), INDEX(additionalconfig, , 2))),
        2
      ),
      jsonGet(finalState, "out")
    )
  )
)
```

### How It Works

- Builds a JSON state object with two keys: remaining span and current output.
- Uses `REDUCE` to carry that state through the part list.
- For each row, calculates `QUOTIENT(rem, plen)` as the base count.
- Writes non-zero counts into the output JSON with `jsonSet`.
- Updates the remainder and passes the new state forward.
- If `additionalconfig` exists, filters post-processing rules with `between` and merges matching JSON fragments using `jsonJoin(..., 2)`.

The main implementation trick is stateful reduction. Instead of separate running variables, the algorithm packs everything into JSON and mutates that JSON indirectly at each reduction step.

### Real-World Examples

Example 1: fill stock lengths against a required span.

```excel
=partFill(100,HSTACK({"Pipe A";"Pipe B";"Pipe C"},{30;20;5}))
```

Result:

```text
{"Pipe A":3,"Pipe C":2}
```

Example 2: apply a remainder rule after the main pass.

```excel
=partFill(
  101,
  HSTACK({"A";"B";"C"},{30;20;5}),
  HSTACK({"{""C"":1}";"{""A"":1}"},{"[1,1]";"[0,0]"})
)
```

Walkthrough:

- Main pass yields `{"A":3,"C":2}` and remainder `1`.
- The rule `[1,1]` matches the remainder.
- The adjustment `{"C":1}` is added in merge mode `2`.

### Notation Notes

- `partarr` is always a two-column array: name, then part length.
- Row order matters because the reduction is sequential.
- When `additionalconfig` is omitted, the last-part branch may overfill by one if there is leftover remainder.
