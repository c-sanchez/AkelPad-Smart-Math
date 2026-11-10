---
name: mathparser-add-or-fix
description: >-
  Add or change AkelPad Smart Math parser behavior end-to-end in Basic and C++:
  new builtin functions, operators, or constants; signature hints; reserved
  names; bug fixes that change results or errors; smoke/parity test
  expectations (from Python/docs, never from live parser output) and
  investigating failing tests; the complex-number support flag;
  USAGE_AND_SYNTAX documentation; and the build/test gates.
---

# Add or Fix MathParser Behavior

## Purpose

Implement a builtin function, operator, constant, or behavior fix consistently across parser code, signature hints (functions), reserved-name rules (constants), tests, and docs. Basic and C++ always stay mirrored.

For behavior-preserving refactors or dead-code removal, use `mathparser-refactor` instead.

Also triggered by: добавить новую функцию / добавить сигнатуру/подсказку функции / исправить ошибку парсера / обновить тесты и документацию.

## Agent loading rules

- Do not read every linked file at the start. Read a file only when its row in the table below applies.
- Use the repo-relative paths below with `Read`.
- Internal skills (`internal-mathparser-*`) are dependencies, not entry points.
- This table is the only routing table; flows do not repeat it.

## Inputs To Collect First

- Name (function, operator, or constant) and purpose, or the reported defect and its reproduction
- Functions: arity (fixed or variadic minimum), scalar/array per argument, error/edge behavior (invalid args, empty arrays, divide by zero), and the feature it depends on, if any (complex, time, lambdas)
- Operators and constants: the extra inputs listed at the top of their flow

Ask concise clarifying questions if ambiguous.

## Classify Then Load

| If the task… | Read now |
|---|---|
| Edits any source file or script (parser, test, generated test data, `.bat`, `.py`); not docs | `.cursor/skills/internal-mathparser-protect-workspace/SKILL.md` **before first edit** |
| Adds/changes a builtin **function** or an **operator** | `.cursor/skills/mathparser-add-or-fix/flows/implement-function.md` |
| Adds a builtin **constant** | `.cursor/skills/mathparser-add-or-fix/flows/implement-constant.md` |
| Fixes a defect / changes existing results or errors | `.cursor/skills/mathparser-add-or-fix/flows/fix-behavior.md` |
| Edits Basic/C++ parser **implementation** | `.cursor/skills/mathparser-add-or-fix/flows/coding-standards.md` |
| Edits `cpp/MathParser.hpp` / `.cpp` shape, helpers, or builtin routing | `.cursor/skills/mathparser-add-or-fix/flows/cpp-class-design.md` |
| Touches complex flag / Cartesian builtins / real↔complex paths, or code under `#if SMARTMATH_*` | `.cursor/skills/mathparser-add-or-fix/flows/complex-support.md` |
| Authors or reviews smoke/parity **expectations**, or any test fails | `.cursor/skills/mathparser-add-or-fix/flows/smoke-tests.md` |
| Edits `USAGE_AND_SYNTAX.md` | `.cursor/skills/mathparser-add-or-fix/flows/usage-docs.md` |
| Changes Basic or C++ parser/tests (mirroring), or `cpp/calc.cpp` / `tools/calc.bas` | `.cursor/skills/internal-mathparser-basic-cpp-parity/SKILL.md` |
| Is about to compile/run gates, probe behavior, or decide which gates apply | `.cursor/skills/internal-mathparser-build-gates/SKILL.md` |

## Required Workflow

Follow in order. Skip a step only when its trigger does not apply.

1. **Protect** — before the first source/script edit.
2. **Implement** — every table row that applies to the edit.
3. **Tests** — for any new or changed function, operator, constant, or behavior, mirrored in C++.
4. **Docs** — when documented behavior is added or changes.
5. **Gates** — before finishing any change; build-gates decides which apply (a docs-only change that touches the Quick index still runs the docs index check).
6. **Reflect** — when either parser or its tests changed: parity skill reflection gate (it owns the gate-rerun rule).

## Validation Checklist

- Implementation, metadata rows, hints, tests, and docs agree.
- Required gates from build-gates ran; Basic passed before any required C++.
- Parity complete (parity skill reflection gate).
- `FunctionNames`, `OperatorNames`, and `ConstNames` stay the only Basic name tables; C++ name lists follow implement-function / implement-constant.
- Complex work is gated and its tests are isolated (complex-support flow).
- Expected values come from references, not parser output; every mismatch was handled per smoke-tests "Handling a test mismatch".
- Integer metadata, helpers, string constants, and non-finite handling follow coding-standards.
- No unrelated behavior changes; C++ artifacts under `cpp/` only.
- C++ additions follow cpp-class-design; Basic structure is unchanged.

## Response Format For Completion

Report:

- files changed;
- concise behavior summary;
- test coverage added, and any existing expectations corrected (with the reference that justified each correction);
- one pass/fail line for each gate that ran (gate names as in build-gates "Invocation");
- remaining Basic/C++ or Python differences (noting which are allowed floating-point inaccuracy) and assumptions from missing requirements.
