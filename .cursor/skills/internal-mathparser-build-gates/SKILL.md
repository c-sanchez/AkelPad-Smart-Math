---
name: internal-mathparser-build-gates
description: >-
  Internal MathParser build/test gate rules (sequential Basic then C++, compile
  error discipline, change-detection inventories, invocation, feature matrix,
  factorint and docs index checks, behavior probes, bitness, performance).
  Not a user entry point — load only when referenced by mathparser-add-or-fix
  or mathparser-refactor.
disable-model-invocation: true
---

# Internal: MathParser Build And Test Gates

Load when deciding which gates apply, immediately before any compile/test run, or before probing behavior.

## Source inventories (single source of truth)

**Basic implementation sources** (trigger `Compile.bat` and formatter/copy regression):

- `MathParser.bas` and the modules it includes: `MathParserFactorInt.bas`, `MathParserRawResult.bas`, `Inc/*.bi`
- `SmartMath.bas`, `SmartMath_About.bas`, `SmartMath_Config.bas`, `SmartMath_Format.bas`, `SmartMath_CopyNormalize.bas`, `SmartMath_Globals.bi`, `SmartMath_Menu.bas`
- plus any other Basic implementation module

**Formatter / copy test sources** (also trigger formatter/copy regression):

- `FormatterRegressionTests.bas`, `FormatterTest_Globals.bas`, `CopyRegressionTests.bas`

**Basic smoke triggers** (`RunSmokeTests.bat`):

- `SmokeTest_MathParser.bas` or other Basic test modules
- `MathParser.bas` or any module it includes (above)
- generated Basic test data included by the smoke test: `tools/*_generated*.bas`
- precaution: `SmartMath.bas` / `SmartMath_Format.bas`

**C++ triggers** (`cpp/BuildTests_vc2022_x64.bat` then `cpp/Tests_MathParser.exe`):

- `cpp/MathParser.cpp`, `cpp/MathParser.hpp`, `cpp/MathParserInternal.hpp`, `cpp/MathParserFactorInt.cpp`, `cpp/MathParserRawResult.cpp`, `cpp/Tests_MathParser.cpp`
- generated C++ test data included by the tests: `tools/*_generated*.cpp.inc`
- plus any other C++ parser/test source or header under `cpp/`

**Feature matrix triggers** (`cpp/RunFeatureMatrixTests_vc2022_x64.bat`, instead of the single C++ build/run):

- adding, removing, or moving code across an `#if SMARTMATH_*` boundary (`SMARTMATH_COMPLEX_NUMBERS`, `_TIME_VALUES`, `_LAMBDA_FUNCTIONS`, `_FACTORINT`) in C++ parser or test sources
- a new builtin, constant, or test that exists only with one feature

**Factorint triggers** (`RunFactorintTimeoutTests.bat`, in addition to smoke):

- `MathParserFactorInt.bas`, `SmokeTest_Factorint.bas`

**Docs index triggers** (`cpp/CheckDocsFunctionIndex.bat`):

- `USAGE_AND_SYNTAX.md` "Quick index", or any change to `FunctionNames` (added, renamed, or removed builtin)

**Non-triggers alone:** `USAGE_AND_SYNTAX.md` outside the Quick index, `.cursor/` skill/rule/process-doc edits, unrelated notes, probe tools (`tools/NegBandProbe.bas`, `cpp/calc.cpp`, `tools/calc.bas`). An edit to either calc tool instead requires rebuilding both calc tools and comparing them (parity skill "Calc tools").

When unsure whether a file affects runtime behavior, treat it as a trigger.

## Invocation

The batch files pause at the end; run them non-interactively:

| Gate | Working directory | Command |
|---|---|---|
| Compile | repo root | `cmd /c "Compile.bat < nul"` |
| Formatter | repo root | `cmd /c "RunFormatterTests.bat < nul"` |
| Copy | repo root | `cmd /c "RunCopyRegressionTests.bat < nul"` |
| Smoke | repo root | `cmd /c "RunSmokeTests.bat < nul"` (compiles `SmokeTest_MathParser.exe` itself) |
| C++ build | `cpp/` | `cmd /c "BuildTests_vc2022_x64.bat < nul"` |
| C++ run | `cpp/` | `.\Tests_MathParser.exe` (exit code 0 = pass) |
| Feature matrix | `cpp/` | `cmd /c "RunFeatureMatrixTests_vc2022_x64.bat < nul"` (builds and runs all 8 flag combinations; replaces C++ build + run) |
| Factorint | repo root | `cmd /c "RunFactorintTimeoutTests.bat < nul"` (29 cases, 20 s limit each) |
| Docs index | repo root | `cmd /c "cpp\CheckDocsFunctionIndex.bat < nul"` |

Output is long: redirect to a file and search it for `FAIL` and the final totals, rather than reading it whole. The C++ build is MSVC with the default language standard (C++14); do not use C++17 library features.

## Behavior probes (not gates)

Use probes to reproduce and re-verify behavior; never to seed test expectations (smoke-tests flow).

- **Calc tools (preferred for probing either parser):** `cpp/calc.cpp` tests the C++ parser and `tools/calc.bas` tests the Basic parser. They are synchronized twins (parity skill "Calc tools"): the same command line gives the same output on both, apart from the parser differences listed in the parity skill. Both enable complex, time, and lambda support. Run `calc -h` for usage. Common forms: `calc "<expr>"`, `calc -f setup.txt "<expr>"` (setup lines first, so variables and `ans` carry over), `calc x=5 -f cases.txt`, and expressions piped on stdin, one per line. Errors in files are prefixed `file:line:`; the exit code is 1 if any evaluation failed.
  - Build C++: `cpp/BuildCalc_vc2022_x64.bat` (from `cpp/`) → `cpp/calc.exe`. Build Basic: `tools/build_calc.bat` (from anywhere) → `tools/calc.exe`, 64-bit when FreeBASIC `bin\win64` exists. Rebuild after parser changes.
  - **Cross-check rule:** whenever a calc run on one side shows a wrong result, error, or crash, run the same command line on the other side before concluding anything, and report both outputs. A defect on one side is a candidate defect on the other until verified.
- **C++, one expression:** `cpp/Tests_MathParser.exe --eval "<expr>"` (or `--eval --complex "<expr>"`) prints the result with default flags (complex off unless `--complex`); rebuild with the C++ test build after parser changes.
- **Basic, one expression:** `tools/NegBandProbe.bas` evaluates one expression from the command line (`--complex` first to enable complex mode; time and lambdas keep parser defaults). `tools/build_neg_band_probe.bat` (from the repo root) builds the 32-bit `tools/NegBandProbe.exe`, which the `tools/gen_*.py` generators also use; rebuild after parser changes. For a 64-bit probe, from the repo root with `C:\tools\FreeBASIC` and its `bin\win64` on `PATH`: `fbc -arch x86_64 -s console -O 2 tools\NegBandProbe.bas MathParser.bas -x <temp>\NegBandProbe64.exe` (output outside the repo, or delete it when done). For a 32-bit calc, run `tools/build_calc.bat` on a machine without `bin\win64`, or call `fbc` without `-arch x86_64` the same way.
- **Probing `HEAD`:** to tell whether a difference predates the current diff, build the probe from a temporary `git worktree` outside the repo (protect-workspace); never `git checkout` over the working tree. If the probe source differs from `HEAD` or is missing there (`git diff HEAD -- tools/NegBandProbe.bas tools/calc.bas cpp/calc.cpp`, `git ls-files`), copy the current probe source into the worktree first; an older copy may not print results.
- **Bitness:** `RunSmokeTests.bat` builds 64-bit when FreeBASIC `bin\win64` exists, otherwise 32-bit, so smoke covers one architecture. The plugin ships both (`Compile32.bat` / `Compile64.bat`). For numeric, literal-parsing, shift, or overflow changes, probe Basic built with and without `-arch x86_64`: 32-bit x87 extended precision can hide double rounding and overflow, and FreeBASIC 64-bit constant handling can differ. For the same changes on the C++ side, also run `cpp/BuildTests_vc2022_win32.bat` then `cpp/Tests_MathParser.exe`.
- **Performance:** for refactors of hot paths (evaluation loop, scalar arithmetic, array walks), compare `cpp/RunPerfBaseline_vc2022_x64.bat` timings before and after the change, run sequentially on an idle machine. Report the numbers; treat a consistent slowdown beyond run-to-run noise as a regression.

## Compile and build error discipline (Required)

On any build/compile failure (`Compile.bat`, `Compile32.bat`, `Compile64.bat`, formatter/copy/smoke/factorint batches, `cpp/BuildTests_vc2022_x64.bat` / `_win32.bat`, the feature matrix, related toolchains):

- **Toolchain stop:** if the batch is missing or the failure is the environment (FreeBASIC path, MSVC toolset), stop immediately. Report the blocker and the last output. Do not edit sources to work around it, and do not spend the 3-iteration loop on it.
- Assume the toolchain is right; "compiler bug" is last resort after the checklist.
- Fix the **first meaningful error**; ignore cascading diagnostics until it is gone.
- Smallest direct fix only; no drive-by refactors.
- One hypothesis → one edit → one compile/build. If diagnostics do not improve, **roll back** that edit before the next hypothesis.
- Checklist before blaming the compiler, in order:
  1. Touched files saved, and they are the files the build compiles
  2. Includes/modules (`Compile.bat` ↔ `Compile32.bat`/`Compile64.bat` parity, C++ headers)
  3. Types, signatures, and declaration order (Basic `byref`/`byval`; C++ const/member/static, arity, a type used before its `using` declaration)
  4. Generated or copied code (stale `.obj`/`.exe`, wrong output path)
  5. Arch/flags (`-arch`, MSVC config, bitness)
  6. Clean rebuild (incremental artifacts)
  7. Toolchain versions (FreeBASIC path, MSVC toolset)
  8. API/parity drift (renamed helpers, moved declarations, Basic vs C++)
- **Hard loop limit:** after **3** failed fix-compile iterations on the same primary error, stop, summarize facts (first error text, files changed, hypotheses tried, rollback status), and ask for human review — do not invent workarounds or disable checks.

## Test and guard execution order (Required)

- Always run sequentially — **never in parallel**. Each gate must exit 0 before the next starts, except for the `Compile.bat` exception below.
- **Any test failure** (formatter, copy, smoke, factorint, C++, feature matrix): handle each failing value with mathparser-add-or-fix `flows/smoke-tests.md` "Handling a test mismatch" before changing code or expectations (re-derive the expected value first; floating-point inaccuracy is allowed). Some suites stop at their first failing row; after a fix, rerun until the suite is clean rather than assuming one failure was the only one.
- **Exception:** if `Compile.bat` failed, still run `RunFormatterTests.bat` then `RunCopyRegressionTests.bat` so those regressions stay visible. After fixing compile, rerun `Compile.bat` and both regression batches.
- When Basic and C++ gates are both required, complete **all** required Basic gates first; a failed Basic gate blocks C++:
  1. `Compile.bat`
  2. `RunFormatterTests.bat` then `RunCopyRegressionTests.bat`
  3. `RunSmokeTests.bat`, then `RunFactorintTimeoutTests.bat` when triggered
  4. `cpp/BuildTests_vc2022_x64.bat` then `cpp/Tests_MathParser.exe`, or the feature matrix when triggered
  5. Docs index check, when triggered (no build dependency; may also run alone for docs-only changes)

## Gate: Compile (`Compile.bat`)

- Run only when a **Basic implementation source** changed.
- If you edit `Compile.bat`'s module list, keep `Compile32.bat` and `Compile64.bat` in lockstep (same sources, same order; only `-arch` / `-x` / arch flags may differ).
- On failure: toolchain stop, otherwise compile-error discipline.

## Gate: Formatter + copy regression

- Run **both** when a **Basic implementation source** or a **formatter/copy test source** changed.
- `RunFormatterTests.bat` → `FormatterRegressionTests.exe`; then `RunCopyRegressionTests.bat` → `CopyRegressionTests.exe`.
- On failure: toolchain stop, otherwise mismatch handling (execution order section), then fix format/copy sources or tests until exit 0.

## Gate: Basic smoke (`RunSmokeTests.bat`)

- Run only for **Basic smoke triggers**. Must pass before any required C++ gate.
- Failures are blockers: toolchain stop, otherwise mismatch handling (execution order section), then fix and rerun until pass.

## Gate: C++ build/test

- Run only for **C++ triggers**, after every required Basic gate passed.
- Build failure: toolchain stop, otherwise compile-error discipline.
- Test failure: mismatch handling (execution order section), then fix → rebuild → rerun until pass.
- When the feature matrix is triggered, it replaces the single build/run; a failure names the flag combination to reproduce.

## Gate: Factorint (`RunFactorintTimeoutTests.bat`)

- Run only for **factorint triggers**, after smoke passed. Exit 124 is a timeout (a performance regression, not a mismatch); any other non-zero exit is a failing case handled like a smoke failure.

## Gate: Docs index (`cpp/CheckDocsFunctionIndex.bat`)

- Compares the `USAGE_AND_SYNTAX.md` Quick index with `FunctionNames` in `MathParser.bas`. It does not resolve aliases, so it already fails on the clean tree, listing aliases (`arcsin`, `prod`, `sorted`, ...) as "Unknown in docs".
- Run it before and after the change and compare: only lines the change added are findings. Fix those in the docs or the name table; do not edit the checker in the same task unless asked.
