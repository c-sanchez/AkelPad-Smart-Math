# Update `USAGE_AND_SYNTAX.md`

Load when editing `USAGE_AND_SYNTAX.md`: function or constant documentation, or operator/numeric/error semantics that a change added, changed, or that the docs got wrong.

## Shared style

- Scan the nearest headings, bullets, and examples first; match structure, terms, list depth, and example shape (`expr` -> `result` vs inline prose).
- Prefer ASCII in new/changed lines: `-` for dashes, `->` for results, straight quotes, unless extending an existing non-ASCII paragraph in place.
- One author voice: no one-off style for a single symbol.
- Example results use the parser's display format (3-digit exponents, e.g. `1e+020`) and must be verified against a reference (smoke-tests flow), not copied from parser output.
- Typo pass on changed lines (watch split words like `implemen ted`).

## Function documentation

- Add/update the function in the appropriate section.
- Update **"Quick index (alphabetical)"** under **"Built-in Functions"**: alphabetical names; aliases on the same row as the canonical name (e.g. `asin/arcsin(value)`); category labels consistent with neighbors. Then run the docs index gate (build-gates).
- Document purpose, signature, each argument/constraints, return behavior (arrays/variadic if relevant).
- At least 2–3 examples: basic scalar; edge/interesting; array when supported.

## Semantics changes

- Put the rule next to the related operator or function (for example integer accuracy, NaN/Inf, complex-mode notes), with one example.
- Fix wrong existing examples in place; do not add a contradicting note elsewhere.
- Complex-only behavior goes in the matching "... With Complex Numbers" subsection.

## Constant documentation

List the new constant **everywhere** older constants (`pi`, `e`) appear for builtins and reserved names. At minimum scan:

- **Built-in constants** list
- Reserved-name / assignment / UDF-parameter restriction sections
- Examples that constants resolve **before** variables (extend in the same tone, e.g. `(e=3)` notes)
