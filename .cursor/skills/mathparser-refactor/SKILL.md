---
name: mathparser-refactor
description: >-
  Review and refactor Basic/C++ parser code for reusability and cleanup. Use
  when the user asks to extract duplicate code, reuse/update helpers, simplify
  without behavior or performance regressions, improve maintainability, or
  remove legacy/unused/dead parser code.
---

# Refactor MathParser (Behavior-Preserving)

## Purpose

Behavior-preserving, efficiency-safe cleanup/refactor of `MathParser.bas` (and the modules it includes), `cpp/MathParser.cpp`, `cpp/MathParser.hpp`, and `cpp/MathParserInternal.hpp`.

Maintenance only. For new builtins/constants or behavior fixes, use `mathparser-add-or-fix`. If a review finds a correctness defect, report it; fix it only when asked, through that skill's fix-behavior flow.

## Agent loading rules

Read on demand:

| When | Read |
|---|---|
| Before first edit | `.cursor/skills/internal-mathparser-protect-workspace/SKILL.md` |
| Editing helpers / int metadata / string constants / non-finite handling | `.cursor/skills/mathparser-add-or-fix/flows/coding-standards.md` |
| Editing `cpp/MathParser.hpp` / `.cpp`, or moving C++ functions off `MathParser` | `.cursor/skills/mathparser-add-or-fix/flows/cpp-class-design.md` |
| Changing real/complex operator or builtin paths | `.cursor/skills/mathparser-add-or-fix/flows/complex-support.md` |
| Adding or changing test expectations, or a test fails | `.cursor/skills/mathparser-add-or-fix/flows/smoke-tests.md` |
| Mirroring Basic↔C++, or editing `cpp/calc.cpp` / `tools/calc.bas` | `.cursor/skills/internal-mathparser-basic-cpp-parity/SKILL.md` |
| Before compile/test or choosing gates | `.cursor/skills/internal-mathparser-build-gates/SKILL.md` |

Do not load the full `mathparser-add-or-fix` skill unless the task becomes a feature add or fix.

## What to optimize

1. Extract/reuse helpers for duplicate logic (repeated loops, branches, validation, error text); extend existing helpers before adding new ones.
2. Simplify control flow; cut maintenance burden.
3. Remove dead/unused code only with evidence.

All of it within the hard constraints below.

## Hard constraints

- No behavior change unless explicitly requested. Keep the exact-int/exact-uint behavior the path already had; do not add metadata it did not keep.
- No user-visible error message drift unless explicitly requested.
- C++: follow cpp-class-design, including its cleanup order (one cluster at a time). Leave Basic's procedure layout as it is.
- No UB; keep Basic/C++ parity; string constants per coding-standards.
- No runtime slowdown: avoid extra allocations/copies on hot paths. Do not replace iterator-style walks with eager flattening unless semantics require it.
- If the user approved a behavior change on a touched operator/builtin path, apply the coding-standards non-finite rule and lock it with tests.

## Refactor workflow

1. Protect-workspace before the first edit.
2. Inventory duplicates (impact / risk / testability).
3. Choose helper strategy (reuse → extend → narrow new helper) per coding-standards and cpp-class-design.
4. Preserve efficiency: single-pass walks, pre-count + one allocation, scratch-buffer reuse in C++, no extra hot-path temps. For hot-path clusters, record a perf baseline before the first edit and compare after (build-gates "Performance").
5. Apply one logical cluster at a time. Between clusters, compile the side you touched when practical (build-gates compile discipline on failure).
6. Remove dead code only when unused/unreachable/superseded is confirmed; otherwise list it as a candidate.
7. Mirror and reflect via the parity skill, then run the required build-gates once for the final diff; all of them must pass before completion.

## Dead/unused code criteria

Removable only if:

- No references/call sites in the repo and not part of the required external API, or
- Provably unreachable under enclosing guards, or
- Fully superseded helper with no remaining use.

Record a short justification per removal.

## Completion report format

Report: files changed; duplicates eliminated; helper reuse/extraction; dead code removed + reason; dead-code candidates kept (and why); efficiency notes (why no slowdown, with perf baseline numbers when measured); one pass/fail line per gate that ran; defects found but not fixed.
