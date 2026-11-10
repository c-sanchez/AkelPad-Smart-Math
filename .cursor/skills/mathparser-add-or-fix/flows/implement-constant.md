# Implement Builtin Constant (Reserved Names)

Load when introducing a new identifier parsed as a **constant** (like `pi` / `e`), not a function call. Tests, docs, and gates follow the `SKILL.md` workflow.

Collect first: value (and exactness: exact int or float), whether it depends on a feature flag (time values, complex), and case-sensitivity (match `pi` / `e` unless the task says otherwise).

## Basic: canonical table in `MathParser.bas`

Same pattern as `FunctionNames`:

- **`BuiltinConstId`** — `CONST_*` enum + sentinel `CONST__COUNT`
- **`ConstNames(0 to CONST__COUNT - 1)`** — the only name table, filled in **`EnsureConstNames()`** (`ConstNames(CONST_PI) = FB_STR_PI`)
- **`TryGetConstant`** — resolve via `TryFindBuiltinConstId` / `ConstNames`, then `select case` on id for the value
- **`IsBuiltinConstantName`** — true when `TryFindBuiltinConstId` finds an index (assignment / UDF name / parameter reserved checks)

When adding: extend the enum, one line in `EnsureConstNames`, one `case` in `TryGetConstant`. Never register the spelling only in `TryGetConstant`, or reserved-name checks miss it. Do not add a second constant-name list in Basic.

Flag-gated constants (time units, `i`) are in the same table; their gating is in `TryGetConstant` and the reserved checks. Follow the existing gated constant of the same feature.

## C++: no name table (`cpp/MathParser.cpp`)

C++ spells the reserved constants out at two sites; update **both**, with the same feature gating:

- `isReservedBuiltinConstantName(name, supportTimeValues)` — UDF parameter checks
- `MathParser::getReservedIdentifierError` — assignment and UDF name checks; the complex `i` check sits here and next to the parameter check under `#if SMARTMATH_COMPLEX_NUMBERS`

Then register the value in the `MathParser` constructor next to `addConst(STR_PI, kPi)` (time units use `setVariable` under `#if SMARTMATH_TIME_VALUES`), and add a `STR_*` constant for the spelling. This duplication is a known C++ structure difference (parity skill); do not add a third list.

## Constant-specific tests

Mirror `pi` / `e` coverage (expectation rules: smoke-tests flow); for a flag-gated constant, test it both with the flag on and off:

- Assignment LHS rejected (`expectedErrContains` matching `reserved constant name`)
- Case-insensitive variant if resolution is case-insensitive
- UDF parameter name and UDF function name rejected when equal to the new constant

## Docs

usage-docs flow, constant section: list the new constant everywhere `pi` / `e` appear.
