# Smoke / Parity Tests (Authoring Expectations)

Load when **adding or reviewing** test expectations, or when any test fails ("Handling a test mismatch"). Not needed when gates only run and pass (build-gates skill). For complex-domain cases, also read complex-support.

## Coverage in `SmokeTest_MathParser.bas`

- Normal behavior; invalid arity/types; array/scalar combos when arrays apply; variadic min + extra args when applicable.
- 1-arg or variadic functions: one normal single-arg case, one single-arg edge, one overflow/range case when the domain has limits.
- Regression cases for fixed defects (fix-behavior flow).
- Match existing naming and assertion style. Mirror equivalent cases into `cpp/Tests_MathParser.cpp` (parity skill).
- The Basic main `tests(1 to N)` table and C++ `expectEval` **skip** a case whose signature duplicates an earlier one. Bump `N` when appending, and check each run's "Detected duplicate cases" count so a new case is not silently skipped.
- C++ parity rows construct a fresh `MathParser` per case; multi-step state checks (assign, then read) need a dedicated test.
- Feature flags: time values and lambdas are **on** by default, so ordinary duration and lambda cases belong in the main tables; their option suites (`RunTimeValuesSupportOptionTests`, `RunLambdaFunctionsSupportOptionTests` and the C++ equivalents) are for behavior that depends on the flag. Only complex support is off by default; complex cases follow complex-support isolation. In C++, any case that needs a feature also sits inside that feature's `#if SMARTMATH_*` guard (complex-support flow).

## External reference for expected values (Required)

Applies to every expectation: `SmokeTest_MathParser.bas`, `cpp/Tests_MathParser.cpp`, generated tables, coverage-audit batches.

- **Never** seed `expected` from current parser output (`Parser_TryEvaluateEx`, `parseAndEvaluate`, `calc.exe`, probes) — that locks bugs in.
- Derive expectations from:
  1. **Python** (`math` / `cmath` / `statistics` / project rules) — prefer `tools/verify_smoke_expected_vs_python.py` when it translates the expression; else a short script.
  2. `USAGE_AND_SYNTAX.md` for intentional project semantics Python does not model (percent postfix, time values, exact-int promotion, error policy). When the docs and Python disagree on plain math, verify by hand and fix whichever is wrong.
  3. IEEE / exact-integer rules when both are silent (e.g. `nan==nan` → `0`, int64 overflow promotion).
- After setting expectations, run both harnesses and handle every failure with "Handling a test mismatch" below.

### Expected string format

- **Floats:** C++ prints `'%.16g'` with the exponent padded to at least 3 digits. Convert Python accordingly: `'%.16g' % x` gives `1e-05` → write `1e-005`; `1.5e+20` → `1.5e+020`. Basic (FreeBASIC `str()`) matches except on exact decimal ties (parity skill), which is a floating-point display difference under "Handling a test mismatch"; keep tie inputs such as `1e15+0.5` out of exact-string tables.
- **Exact integers** (int64/uint64 metadata) print every digit, not `%.16g` (`2**63` → `9223372036854775808`).
- **Non-finite:** `nan`, `inf`, `-inf`.
- **Complex:** real and imaginary parts each formatted as above, joined by `+`/`-`, suffix `i`; a zero part is omitted (`2i`, `3`); a unit coefficient prints as `i` / `-i` (`1+i`, `-i`).
- **Special display rules** override `%.16g` where documented. Rules that differ between Basic and C++ (the `±2^63` cases) are listed in the parity skill's known differences and use tolerant comparison.

### Comparison mode per harness

| Harness | Float comparison |
|---|---|
| Basic `ResultCloseEnough` (main table `RunCase`, band batches, most sub-suites, `cxLogBase`) and C++ `smokeResultCloseEnough` (parity rows, `expectEval`, the complex `ln`/`log` table) | same rule on both sides: equal or adjacent doubles, or within `16 * DBL_EPSILON * max(abs(a), abs(e))` (relative at every magnitude) |
| C++ hand-written `kRows` tables compared with `p.getResult() != row.expect`, Basic tables compared with `rt <>` (e.g. `cxFmtOk`) | **exact string** |

Both tolerant comparators treat an expected `0`, `±1`, or `±inf` strictly, compare complex results (last character `i`) part by part, and compare array results element by element with the same rules. Keep the two comparators identical; a change to one is a mirrored harness change. Limits to keep in mind when choosing where a case goes:

- The tolerance compares numeric values only. Near `2^53` it spans about ±32, and an exact integer equals a float with the same double value (`9223372036854775808` matches `9.223372036854776e+018`). Exact-integer results that must be checked digit for digit (int64/uint64 metadata, values above `2^53`) belong in an exact-string table or are checked through `hex` / `uhex`.
- The tolerance is relative, so tiny values are checked as strictly as large ones (`1e-20` does not match `5e-20`, and `1e-20` does not match `0`); they belong in tolerant tables like any float.
- Exact-string tables are only for display/format checks (exponent width, base prefix, separators, duration, rational form) and digit-exact integers. Do not put general float results there: a last-digit floating-point difference would fail them (see "Handling a test mismatch").

### Generated tables

Both test suites include generated tables: `tools/*_generated*.bas` in `SmokeTest_MathParser.bas`, and `tools/*_generated*.cpp.inc` in `cpp/Tests_MathParser.cpp`. Any `tools/gen_*.py` that calls `NegBandProbe.exe` or `Tests_MathParser.exe --eval` (currently `gen_neg_band_cases.py` and `gen_new_coverage_tests.py`) takes its expectations from **parser output**.

- The generators do not build anything: before regenerating, rebuild `tools/NegBandProbe.exe` (`tools/build_neg_band_probe.bat`) and `cpp/Tests_MathParser.exe` (C++ test build) from the current sources. A missing or stale probe silently gives old-parser output, or falls back to C++.
- Treat a regeneration as a draft: diff it, and confirm every changed row against Python before committing.
- `gen_neg_band_cases.py` prefers the Basic probe (32-bit when built by `tools/build_neg_band_probe.bat`, while smoke runs 64-bit) and falls back to C++ `--eval`. Existing band tables may come from C++ output, so a regeneration can change rows where Basic differs (for example the `±2^63` cases in the parity skill) or where the two bitness builds differ. Keep the reference value for those rows.
- After a behavior fix, prefer editing the individual stale rows by hand over regenerating.
- Most tables exist in both languages (same base name, `.bas` and `.cpp.inc`); edit both copies of a row together. A few are one-sided (`naninf_mirror_generated.bas`, `pos_band_cases_generated*.cpp.inc`): mirror their cases into the other harness by hand only when the reflection gate requires it.

### Strange Python / principal-branch results

- Prefer a **real** value under project real-only semantics when applicable.
- Confirm with the **inverse** (`(-5)**3` = `-125` ⇒ `(-125)**(1/3)` = `-5`, not a complex root).
- Comment the test when diverging from Python's default branch.

### Errors

Prefer substrings from `USAGE_AND_SYNTAX.md` or documented error policy; do not copy ad-hoc parser text unchecked. C++ lowercases identifiers in unknown-name errors; use lowercase names in tests shared by both sides.

## Handling a test mismatch (Required)

Applies to every failing expectation in smoke, regression (formatter/copy), band/generated, and C++ tests.

1. **Re-derive the expected value independently** in Python (or the reference order above), from the expression, not from either parser's output. Show the computation in the progress note. Formatter and copy regression tests check formatting and copy normalization, which Python does not model: their reference is the formatter/copy rules and `USAGE_AND_SYNTAX.md`, and they always compare exact strings, so step 3 does not apply to them.
2. **The expected value was wrong:** correct it in every copy (Basic, C++, generated `.bas` / `.cpp.inc`) and list it in the report as old → new with the Python evidence.
3. **The expected value is right and the difference is only floating-point inaccuracy:** allowed, not a defect. That means the same mathematical result, rounded differently because a double has a limited number of fractional bits: a difference of a few ulps relative to the larger magnitude (accumulated rounding, `log(hypot(...))` versus Python's `log1p` form, a 16-digit decimal tie). Keep the Python value as the expected value and compare it with a tolerant comparator: move the case out of an exact-string table rather than copying the parser value into it. Note the allowed difference in a test comment when it is larger than one ulp.
4. **Anything else is a defect** (fix-behavior flow), even when the digits look close:
   - any difference in an exact integer or exact-int metadata, or a value that should be an exact integer printed as a float (or the reverse), except the parity skill's known differences;
   - wrong sign, magnitude, or digit count (a dropped or extra digit is not rounding);
   - a missing, extra, or snapped-to-`0` Cartesian part; `nan` / `inf` versus a finite value;
   - a value versus an error, or different error text;
   - display format (exponent width, base prefix, separators, duration format, rational form);
   - a difference larger than the harness tolerance. If it can still be explained by rounding (cancellation, ill-conditioned input), report it as a precision issue; do not widen a tolerance to make it pass.
5. **The user accepts a defect instead of fixing it:** record it in the parity skill's known differences (inputs, each side's value, the reference value). Keep the reference value as the expected value, with a comparator that accepts the difference; never copy the parser value into an exact-string table.
6. Do not change a comparator's tolerance as part of fixing a mismatch. A tolerance change is a separate, reported test-harness change, mirrored in Basic and C++.

## Value preservation in tests (Required)

Expectations encode correct math and structure; step 4 above lists what counts as a defect. In addition:

- Aggregates / structure-only builtins (`sum`, `product`, `reverse`, `unique`, `unpack`, `sort`, aliases): operands survive except for the operation itself.
- Operands with `inf`/`-inf`/`nan`: expected strings still show every nonzero Cartesian part, exact-int metadata when applicable, and duration formatting when applicable.
- `nan` is valid only when IEEE or documented domain rules require it — not from spurious `0*Inf` or display shortcuts. Wrapped integer overflow is never a valid expectation.
