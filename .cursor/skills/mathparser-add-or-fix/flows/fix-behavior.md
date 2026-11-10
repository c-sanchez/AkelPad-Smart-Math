# Fix Parser Behavior (Defects and Intentional Semantics Changes)

Load when fixing a reported defect or otherwise changing existing results or error text.

## 1) Reproduce on every side

- Probe Basic and C++ with the same inputs before editing; `tools/calc.exe` and `cpp/calc.exe` take identical command lines (probe tools: build-gates skill). A defect reported or found on one side is checked on the other (parity skill "Calc tools"). For numeric or literal-parsing issues, probe both Basic bitness builds (build-gates "Bitness").
- Record the reference value for each input (smoke-tests flow: docs, Python, IEEE/exact-int rules). A case is a defect only if a side disagrees with the reference by more than floating-point inaccuracy (smoke-tests "Handling a test mismatch").
- When the working tree already has uncommitted parser changes, probe `HEAD` too (build-gates "Probing `HEAD`") and say in the report whether the difference already existed or came from the current diff.

## 2) Fix

- Fix the root cause in both parsers (parity skill), following coding-standards. Prefer one shared helper for the same rule used at several call sites.
- Keep each site's existing error ids and texts unless the error text is the defect.
- When a fix makes a previously ignored case reachable (overflow, non-finite, `INT64_MIN`), decide its accepted/rejected behavior per coding-standards and lock it with tests.

## 3) Re-verify

- Rerun the same probes on all sides (including the other bitness when step 1 required it) and confirm they match the reference.

## 4) Tests and docs

- Add a regression case for each fixed input (smoke-tests flow), mirrored in C++.
- Existing expectations that fail after the fix: handle each with smoke-tests "Handling a test mismatch". Update an expectation only when Python (or the reference) shows it encoded the defect, and list each correction in the report. If it already matched the reference (within floating-point inaccuracy), the fix is wrong.
- Update `USAGE_AND_SYNTAX.md` when documented semantics change, or when the docs were wrong (usage-docs flow).

## 5) Report what stays different

List intentional or out-of-scope differences that remain (Basic vs C++, or parser vs Python), so they are not rediscovered as defects. When the user accepts a difference instead of a fix, record it per smoke-tests "Handling a test mismatch" step 5; known ones live in the parity skill.
