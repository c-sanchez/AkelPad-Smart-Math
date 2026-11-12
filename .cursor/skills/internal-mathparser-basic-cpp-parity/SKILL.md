---
name: internal-mathparser-basic-cpp-parity
description: >-
  Internal Basic↔C++ MathParser porting and reflection rules, plus the list of
  known intentional differences. Not a user entry point — load only when
  referenced by mathparser-add-or-fix or mathparser-refactor.
disable-model-invocation: true
---

# Internal: Basic/C++ Parity

Load when Basic or C++ parser/test sources change and mirroring may be required, or when `cpp/calc.cpp` or `tools/calc.bas` changes.

## Porting rules

All C++ artifacts (sources, test batches, tools) stay under `cpp/`.

- Port in **both** directions when either parser changes: fixes, refactors, cleanup, hardening, and behavior-preserving rewrites that can still affect edges.
- Mirror test expectations and cases in each language's own harness.
- Preserve intent, not only syntax.
- Implement equivalent logic in the other language; do not call or depend on the other language's runtime, binaries, or test harness.
- **Independent test sources:** C++ tests build and run without any Basic test source, and Basic tests without C++ sources or binaries. A convenience helper that reads the other side's files must be optional and must not fail when they are missing.
- No UB (coding-standards).

## Structure may diverge

C++ may move functions off `MathParser` per `.cursor/skills/mathparser-add-or-fix/flows/cpp-class-design.md`; Basic stays procedural. Such a move is a **no-op** on the Basic side when behavior, error text, and exact-int metadata are unchanged — say so in the reflection note. A behavior or error-text change is still mirrored both ways.

Known structural differences (not behavior; nothing to mirror):

- C++ compiles features in or out with `SMARTMATH_*` macros (complex-support flow); Basic always contains every feature and gates it only at runtime.
- Basic has one constant-name table (`ConstNames`); C++ spells reserved constants at two sites (implement-constant flow).

Language-specific workarounds are not mirrored when the other side has no such defect (for example FreeBASIC constant-folding or `Int()` wrap workarounds); record them as no-ops.

## Known intentional differences

Do not "fix" these unless the user asks, and do not write shared tests that require them to match; record new ones here when a task leaves a non-floating-point difference in place. Floating-point inaccuracy needs no entry; smoke-tests "Handling a test mismatch" covers it.

- Unknown-name errors: C++ lowercases identifiers and lists every unknown name in the input; Basic keeps case and stops at the first.
- Float formatting: C++ prints correctly rounded `%.16g`; Basic uses FreeBASIC `str()`, which can differ on exact decimal ties (`1e15+0.5`: Basic `1000000000000001.0`, C++ `1000000000000000.0`). Both append `.0` when that general form has no fraction and no exponent, and both use scientific form for a non-exact value with magnitude above `2^53`. The plugin auto formatter uses the same rule and then strips leading zeros in the exponent (`9.907919180215092e+15`).
- `±2^63` results (accepted within the test tolerance; do not write exact-string tests that require them to match):
  - `-9223372036854775808-1` and `9223372036854775807+0.5`: C++ displays `±9.223372036854778e+018` (deliberate `DecScientificPow63High` display rule, not `%.16g`); Basic displays `±9.223372036854776e+018` (its `decScientificPow63High` property is never set). The two are adjacent doubles.
  - `-(-9223372036854775808)` and `h(x)=-x; h(-9223372036854775808)`: C++ returns the exact `9223372036854775808` (as Python does); Basic returns the float `9.223372036854776e+018`, the same double value.
  - `0-(-9223372036854775808)`: both return the float `9.223372036854776e+018`; Python gives the exact integer.
- Complex `ln`/`log` real part is `log(hypot(re, im))`, which can be one ulp from Python near `|z| = 1` (`log(1+i,2)` → `0.5000000000000001+...`, Python `0.5+...`). This is allowed floating-point inaccuracy: the tests expect the Python value with tolerant comparison.

Empty-array results follow `.cursor/skills/mathparser-add-or-fix/flows/empty-array-policy.md`. A Basic/C++ mismatch with that policy is a defect, not an entry in this list.

## Calc tools

`cpp/calc.cpp` (C++ parser) and `tools/calc.bas` (Basic parser) are command-line calculators for testing each parser. They are always synchronized:

- **Every change to one is made to the other in the same task**: options, argument parsing, help text, output and error formats, exit codes, enabled features, and stdin/terminal handling. Neither is ever left ahead of the other.
- Allowed divergence is language-level only: runtime workarounds with no counterpart (for example `_dowildcard` in `calc.bas`, which stops the MinGW runtime from expanding `*` in arguments) and idiomatic structure. Record such items as no-ops.
- After editing either, rebuild both (build-gates "Behavior probes") and run the same command lines through both, covering the changed behavior. Outputs must match except for the parser differences listed below.
- **Cross-language verification:** an issue found while testing with either calc tool (wrong result, wrong error, crash, hang) is checked on the other with the same command line before it is reported or fixed. Report both outputs. If the other side has it too, fix both (fix-behavior flow); if not, say so and name the side that matches the reference.

## Reflection gate

Single home for the gate-rerun rule. Whenever parser/tests change in either language, for each affected item record either:

- **Mirrored change** — equivalent update on the other side, or
- **No-op with justification** — the other side already has *provably* equivalent behavior. Prefer a regression test that proves it (expectations per `.cursor/skills/mathparser-add-or-fix/flows/smoke-tests.md`).

If gates already ran for the pre-reflection diff, rerun **only** the gates required for files the reflection changed. If gates have not run yet, run the gates required for the final diff (build-gates skill).

Parity mismatch, build failure, or test failure is a blocker: fix, rebuild, rerun until required sides pass.
