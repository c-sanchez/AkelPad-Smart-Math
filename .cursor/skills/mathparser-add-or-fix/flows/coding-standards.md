# Coding Standards (Parser Implementation)

Load when writing or editing Basic/C++ parser implementation (skip for docs-only or gate-only work). This file is the single home for the rules below; other flows link here instead of restating them.

## Helper reuse and optimization

- Search for an existing helper that covers or nearly covers the need before adding one.
- If close: extend it and migrate old and new call sites; keep it hot-path efficient.
- Otherwise add a **narrow** new helper (no overlapping responsibilities).
- Reuse existing locals/state before adding temporaries.
- Keep project indentation and clean nested blocks.

## C++ class shape

C++ helper placement follows `.cursor/skills/mathparser-add-or-fix/flows/cpp-class-design.md`: a `MathParser` method only when it reads or writes session state. Basic stays procedural; do not restructure `MathParser.bas` to match C++ types.

## Non-finite values

For a new operator or builtin, or a touched path whose behavior changes: decide whether `inf` / `-inf` / `nan` are accepted or rejected (standard math and typical calculator behavior), keep Basic/C++ parity, and lock the decision with tests (including error text when rejecting).

## Undefined behavior

No UB in C++ or Basic. Reuse the existing helpers below instead of writing new checks:

| Need | C++ | Basic |
|---|---|---|
| Signed overflow checked **before** the operation (never by inspecting a wrapped result) | `tryAddInt64Checked`, `checkedSubLL`, `tryMulInt64Checked` | `TryAddULongChecked`, `TryMulULongChecked` (unsigned) |
| Sign-magnitude form, so `INT64_MIN` is never negated as a signed value | `uint64MagnitudeOfInt64`, `int64FromNegatedMagnitude`, `tryGetExactSignMagnitudeFromScalar`, `makeScalarFromSignMagnitude` | `UInt64MagnitudeOfInt64`, `TryGetExactSignMagnitudeScalar`, `ValueSetFromSignMagnitude` |
| Double → integer with a range check first | `tryRoundDoubleToInt64`, `tryRoundDoubleToUInt64`, `tryRoundHalfUpDoubleToLongLong` | `TryRoundHalfUpDoubleToLongInt` |

Range-check every double before converting it to an integer: C++ casts out of range are UB, and FreeBASIC `Int` / `CLngInt` wrap through `LongInt` for magnitudes `>= 2^63`.

## Integer metadata preservation (Required)

`MAX_MAGNITUDE` = `2^53`: the largest double magnitude at which every integer is exact. Code names: Basic `FB_MAX_EXACT_INT_FROM_DOUBLE`; C++ `smartmath_internal::kMaxExactIntFromDouble`, aliased as `K_MAX_EXACT_INT_FROM_DOUBLE` inside `MathParser.cpp` (search for both).

Exact signed/unsigned int64 metadata (`exactInt64` / `exactUInt64` or equivalents) is first-class in Basic and C++:

| Where the integer comes from | exact-int | exact-uint |
|---|---|---|
| **Preserved:** produced by an integer operation (sum, difference, product, integer-valued builtin) or already-valid metadata | `INT64_MIN`…`INT64_MAX` | `0`…`UINT64_MAX` |
| **Restored from a double:** only if the double is an exact integer with magnitude `<= MAX_MAGNITUDE` | `-MAX_MAGNITUDE`…`MAX_MAGNITUDE` | `0`…`MAX_MAGNITUDE` |

Both flags: preserved integers in `0`…`INT64_MAX`, restored integers in `0`…`MAX_MAGNITUDE`. Exact-int only: any negative value. Exact-uint only: preserved integers above `INT64_MAX`.

Rules:

- Never drop valid metadata, including when the stored result is floating-point, and including structure-only transforms (flatten, reshape, sort, reverse, dedupe, format, unpack). Prefer metadata-carrying scalar walkers over raw `double` buffers.
- Never restore from a double with magnitude `> MAX_MAGNITUDE` (its ulp exceeds 1); do not cast it. Return it as a float.
- Never round a non-integer into an integer, and never use a tolerance helper (such as `nearlyInt`) to stamp exact metadata.
- Transcendental, variance, stddev, and similar paths may return a non-integral float without metadata; justify with tests.
- Applies when adding a builtin/operator or intentionally changing numeric behavior. A behavior-preserving cleanup keeps the metadata behavior it found.
- Keep Basic/C++ boundaries and errors aligned; lock display kind (`RSK_INT64` vs floating) with tests where affected.

## String constants naming

- String literals longer than one character → named constant; one-character literals may stay inline.
- Reuse an existing constant for the same literal.
- Names are semantic (`STR_*` in C++, `FB_STR_*` in Basic); no opaque `STR_0001` names.
