# Coding Standards (Parser Implementation)

Load when writing or editing Basic/C++ parser implementation (skip for docs-only or gate-only work). This file is the single home for the rules below; other flows link here instead of restating them.

## Helper reuse and optimization

- Search for an existing helper that covers or nearly covers the need before adding one.
- User-function and lambda body, parameter, and definition rules already share one helper per rule (user-functions-and-lambdas flow). Extend that helper. Do not add a second copy for the other surface.
- If close: extend it and migrate old and new call sites; keep it hot-path efficient.
- Otherwise add a **narrow** new helper (no overlapping responsibilities).
- Reuse existing locals/state before adding temporaries.
- Keep project indentation and clean nested blocks.

## C++ class shape

C++ helper placement follows `.cursor/skills/mathparser-add-or-fix/flows/cpp-class-design.md`: a `MathParser` method only when it reads or writes session state. Basic stays procedural; do not restructure `MathParser.bas` to match C++ types.

## Non-finite values

For a new operator or builtin, or a touched path whose behavior changes: whether `inf` / `-inf` / `nan` are accepted or rejected is a raise-concerns decision (standard math and typical calculator behavior). After the user decides, keep Basic/C++ parity and lock the decision with tests (including error text when rejecting).

## Undefined behavior

No UB in C++ or Basic. Reuse the existing helpers below instead of writing new checks:

| Need | C++ | Basic |
|---|---|---|
| Signed overflow checked **before** the operation (never by inspecting a wrapped result) | `tryAddSIntChecked`, `checkedSubLL`, `tryMulSIntChecked` | `TryAddULongChecked`, `TryMulULongChecked` (unsigned) |
| Sign-magnitude form, so the most negative `SmInt` is never negated as a signed value | `uintMagnitudeOfSInt`, `sintFromNegatedMagnitude`, `tryGetExactSignMagnitudeFromScalar`, `makeScalarFromSignMagnitude` | `uintMagnitudeOfSInt`, `TryGetExactSignMagnitudeScalar`, `ValueSetFromSignMagnitude` |
| Double → integer with a range check first | `tryRoundFloatToSInt`, `tryRoundFloatToUInt`, `tryRoundHalfUpFloatToSInt` | `TryRoundHalfUpFloatToSInt` |

Range-check every double before converting it to an integer: C++ casts out of range are UB, and FreeBASIC `Int` / `CLngInt` wrap through `LongInt` for magnitudes `>= 2^63`.

## Numeric types and limits (Required)

The parser is retargetable: only `SmFloat`, `SmInt` and `SmUInt` name the numeric types (C++ `smartmath::SmFloat/SmInt/SmUInt`, Basic `SmFloat/SmInt/SmUInt`), and every limit, bound, tolerance and table derived from them lives in one file per language: `cpp/MathParserTypes.hpp` and `Inc/MathParserTypes.bi` (`kIntMax`/`SM_INT_MAX`, `kFloatSIntBound`/`SM_FLOAT_SINT_BOUND`, `kMaxExactIntFromFloat`/`SM_MAX_EXACT_INT_FROM_FLOAT`, `kTolLost*`/`SM_TOL_LOST*`, `pow10UInt`, `factorialInt`, the range helpers, ...).

- Never write `double`, `long long`, `int64_t`, `Double`, `LongInt`, `ULongInt`, `9223372036854775807`, `2^53`-style literals or `1ull << n` in parser code; use the types and the derived names above. Add a new derived limit to the type file, with a baseline `static_assert` / `SM_STATIC_ASSERT` that pins the default value.
- The float-to-integer range checks use the exclusive bounds (`isFloatInSIntRange`, `isFloatInUIntRange`); never compare against `double(LLONG_MAX)` (that value rounds up to `2^63`).
- Code that must stay bound to a width (the binary64 `ddMulExact`, the 64-bit factorint algorithm, the libc extern block, binary64 test helpers) is wrapped in `SM_FIXED_WIDTH_BEGIN` / `SM_FIXED_WIDTH_END` (or a trailing `SM_FIXED_WIDTH`) and explains why.
- Tests that expect 64-bit / binary64 text are built only when `SMARTMATH_BASELINE_WIDTHS` (C++) / `SM_BASELINE_WIDTHS` (Basic) is set; width-generic expectations are derived by decimal-string arithmetic from the type widths (`widthgeneric` in `Tests_MathParser.cpp`, `tests/WidthGenericDecimal.bi`), never from the parser output or its own tables.
- `tools/check_type_literals.ps1 -Strict` must report zero hits outside the type files and the marked fixed-width regions.

## Integer metadata preservation (Required)

`MAX_MAGNITUDE` = `2^53`: the largest double magnitude at which every integer is exact. Code names: Basic `SM_MAX_EXACT_INT_FROM_FLOAT` (`Inc/MathParserTypes.bi`); C++ `smartmath::kMaxExactIntFromFloat` (`cpp/MathParserTypes.hpp`); it is derived from the float type, so the value is `2^mantissa` for the configured `SmFloat`.

Exact signed/unsigned SmInt metadata (`exactSInt` / `exactUInt` or equivalents) is first-class in Basic and C++:

| Where the integer comes from | exact-int | exact-uint |
|---|---|---|
| **Preserved:** produced by an integer operation (sum, difference, product, integer-valued builtin) or already-valid metadata | `SmInt` min…max (`kIntMin`…`kIntMax` / `SM_INT_MIN`…`SM_INT_MAX`) | `0`…`SmUInt` max (`kUIntMax` / `SM_UINT_MAX`) |
| **Restored from a double:** only if the double is an exact integer with magnitude `<= MAX_MAGNITUDE` | `-MAX_MAGNITUDE`…`MAX_MAGNITUDE` | `0`…`MAX_MAGNITUDE` |

Both flags: preserved integers in `0`…`SmInt` max, restored integers in `0`…`MAX_MAGNITUDE`. Exact-int only: any negative value. Exact-uint only: preserved integers above `SmInt` max.

Rules:

- Never drop valid metadata, including when the stored result is floating-point, and including structure-only transforms (flatten, reshape, sort, reverse, dedupe, format, unpack). Prefer metadata-carrying scalar walkers over raw `double` buffers.
- Never restore from a double with magnitude `> MAX_MAGNITUDE` (its ulp exceeds 1); do not cast it. Return it as a float.
- Never round a non-integer into an integer, and never use a tolerance helper (such as `nearlyInt`) to stamp exact metadata.
- Transcendental, variance, stddev, and similar paths may return a non-integral float without metadata; justify with tests.
- Applies when adding a builtin/operator or intentionally changing numeric behavior. A behavior-preserving cleanup keeps the metadata behavior it found.
- Keep Basic/C++ boundaries and errors aligned; lock display kind (`RSK_SINT` vs floating) with tests where affected.

## String constants naming

- String literals longer than one character → named constant; one-character literals may stay inline.
- Reuse an existing constant for the same literal.
- Names are semantic (`STR_*` in C++, `FB_STR_*` in Basic); no opaque `STR_0001` names.
