# Complex Number Support And Real/Complex Parity

Load when the task mentions complex numbers, the parser-wide complex support flag, Cartesian builtins, or changes that touch real-only and complex-enabled operator/function paths.

## Parser-wide option

Global switch: Basic `Parser_SetSupportComplexNumbers` / `Parser_GetSupportComplexNumbers`; C++ `MathParser::setSupportComplexNumbers` / `getSupportComplexNumbers`.

- **Default off:** keep historical real-only semantics (non-real domains → `NaN` or errors as today) unless the task explicitly extends complex support.
- **When on:** every path that produces or consumes complex values consults the flag (or a thin helper over it), so real-only users never see semantic surprises.
- **Scalars and arrays:** with the switch on, builtins accept one complex scalar or an array of complex values per argument, consistent with existing scalar/array conventions.
- **Exact int metadata:** apply the coding-standards exact-int rules to **both** Cartesian parts, consistently across construction, binary ops, and display.
- **Tiny nonzero parts:** never snap a tiny nonzero real or imaginary part to `0` — not in extraction (`real` / `imag` / `conj`), arithmetic, rational display (`ratio`), or base formatting (`hex` / `oct` / `bin`). `1e-15+2i` displays as `1e-015+2i`.

## Compile-time feature macros (C++ only)

C++ also compiles features in or out: `SMARTMATH_COMPLEX_NUMBERS`, `SMARTMATH_TIME_VALUES`, `SMARTMATH_LAMBDA_FUNCTIONS`, `SMARTMATH_FACTORINT`, each defaulting to `1` in `cpp/MathParserInternal.hpp`. Basic has no equivalent (parity skill known differences).

- Code that only exists with a feature (complex builtins, `i`, Cartesian paths, time constants) goes inside the matching `#if SMARTMATH_*` block, next to the existing code of that feature; its C++ tests go inside the same guard.
- The runtime flag is still checked inside the block; the macro only removes the code.
- The build must pass with every macro off: the feature matrix gate (build-gates) covers all combinations.

## Tests (strict separation)

- Basic: complex-domain expectations go only in **`RunComplexNumberSupportOptionTests()`** (or a dedicated sub), which enables `Parser_SetSupportComplexNumbers(TRUE)` at its start and restores the flag at its end when other suites need isolation. Never in the main `SmokeCase` table.
- C++: only in **`buildComplexNumberSupportOptionCases()`**, run via `runSuite("Complex number support (parser option)", ...)`. Each case calls `setSupportComplexNumbers(true)` at the start of its own lambda; no suite-level setup.
- Purely real inputs of the same shape go in the main tables, so both modes stay covered.

## Real ↔ complex operator/function parity (Required)

Treat real-only and complex-enabled behavior as one design surface. Apply both directions unless semantics truly differ (document why in code/tests).

- **Real → complex:** fixes, policies, and helpers on a real-only path (exact-int verification after `sqr`, float vs int64 export, non-finite handling, domain errors such as `log(0, b)`) are reused or mirrored on the complex path when complex mode can hit the same case. Prefer one shared helper; complex-only wrappers must not weaken rules the real path enforces.
- **Complex → real:** fixes on a complex path are checked on the real-only path for purely real operands (imag = 0). Do not leave real-only code on an older divergent branch.

### Before finishing

- Name the real dispatch and the complex dispatch (if any), and the shared helpers or the documented asymmetry.
- Exact int64 on one side and float on the other for the same numeric situation is a bug unless tests lock different semantics. A part whose magnitude is greater than `MAX_MAGNITUDE` is not that situation: do not promote the double so the sides match.
- Mirror display-kind expectations (`RSK_INT64` vs `RSK_FLOATING`) in Basic, C++, and formatter/raw-export tests when affected.
