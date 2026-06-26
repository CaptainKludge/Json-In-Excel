# Algorithm Solutions

This category contains allocation and optimization functions built on top of the JSON helpers.

## Functions

### `partFill` - Sequential Part Allocation

**Purpose**: Allocate part counts in input order using integer division of the remaining span.

**Syntax**:
```excel
=partFill(span, partarr, [extrahungry])
```

**Parameters**:
- `span`: Target span/amount (number)
- `partarr`: Two-column array: part name (col 1), part length (col 2)
- `extrahungry` (optional): Two-column post-processing rules
	- Column 1: JSON adjustments (for example `{"Pipe A":2}`)
	- Column 2: remainder interval string consumed by `between` (for example `[0,4]`)

**How it works**:
1. Iterate each part row in order.
2. Compute `pcount = INT(rem / part_length)`.
3. Write non-zero counts into an output JSON object.
4. Keep reducing remainder through the loop.
5. If `extrahungry` is provided, filter matching rule rows where `between(finalRemainder, interval)` is TRUE.
6. Merge matching JSON adjustments into output using `jsonJoin(..., 2)` (numeric fields are added).

**Important behavior**:
- `extrahungry` is post-processing. It does not alter the per-row `INT(rem / len)` calculation while iterating.
- Return value is the allocation JSON object.

**Example 1**:
```excel
=partFill(100, HSTACK({"Pipe A";"Pipe B";"Pipe C"}, {30;20;5}))
```

Result:
```text
{"Pipe A":3,"Pipe C":2}
```

**Example 2 (extrahungry adjustment)**:
```excel
=partFill(
	101,
	HSTACK({"A";"B";"C"}, {30;20;5}),
	HSTACK({"{""C"":1}";"{""A"":1}"}, {"[1,1]";"[0,0]"})
)
```

Explanation:
- Base pass yields remainder `1` and output `{"A":3,"C":2}`.
- Rule `[1,1]` applies, so `{"C":1}` is added.

Result:
```text
{"A":3,"C":3}
```

---

### `greedyPartFill` - Two-Phase Greedy Allocation

**Purpose**: Run the same sequential pass, then attempt a one-part remainder correction.

**Syntax**:
```excel
=greedyPartFill(span, partarr, [extrahungry])
```

**Parameters**:
- `span`: Target span/amount (number)
- `partarr`: Two-column array: part name and part length
- `extrahungry` (optional): Same shape/behavior as `partFill`

**How it works**:
1. Phase 1 is equivalent to `partFill` base counting.
2. Compute `remAfterGreedy`.
3. Pick one additional length:
	 - Prefer `MIN(partLens where partLen >= remAfterGreedy)`.
	 - If none exist, use `MIN(partLens)`.
4. Add one count to that part.
5. Recompute remainder as `remAfterGreedy - pickLen` (can be negative).
6. Apply `extrahungry` post-processing exactly like `partFill` using final remainder.

**Important behavior**:
- The selection expression uses `MIN` on candidates `>= remainder`, which picks the smallest overshooting part.
- If no part is large enough, the smallest part is still added, which may produce negative remainder.

**Example 1 (overshoot case)**:
```excel
=greedyPartFill(85, HSTACK({"A";"B";"C"}, {30;20;7}))
```

Explanation:
- Phase 1 gives `{"A":2,"B":1}` with remainder `5`.
- Candidate lengths `>=5` are `30,20,7`; `MIN` picks `7`.
- Add one `C`: remainder becomes `-2`.

Result:
```text
{"A":2,"B":1,"C":1}
```

**Example 2 (already exact)**:
```excel
=greedyPartFill(100, HSTACK({"Pipe A";"Pipe B";"Pipe C"}, {30;20;5}))
```

Result:
```text
{"Pipe A":3,"Pipe C":2}
```

## Comparison

| Aspect | partFill | greedyPartFill |
|---|---|---|
| Base pass | Sequential integer allocation | Same base pass |
| Extra phase | None | Adds one extra chosen part |
| Remainder | Non-negative in normal cases | Can become negative after phase 2 |
| extrahungry | Post-merge by final remainder | Same post-merge by final remainder |
| Typical use | Deterministic ordered allocation | Prefer closer fit via one-step correction |
