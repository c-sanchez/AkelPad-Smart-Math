# Implement Builtin Function Or Operator

Load when adding or changing a **builtin function** or an **operator**. Tests, docs, and gates follow the `SKILL.md` workflow.

## Implementation files

- `MathParser.bas`
- C++ under `cpp/`: `cpp/MathParserInternal.hpp` (enums), `cpp/MathParser.cpp` (tables, handlers), mirrored tests

## 1) Register the function in `MathParser.bas`

Each builtin is one id plus one row in each table; all tables are indexed by the id, so order must match the enum exactly.

- `BuiltinFunctionId`: add a `FUNC_*` enumerator.
- `FunctionNames(FUNC_*) = FB_STR_*`: the only spelling (string constant per coding-standards). Elsewhere (dispatch, hints, keywords, reserved-name checks) use `FunctionNames(FUNC_*)` / `OperatorNames(OP_*)` or their wrappers, never name literals. UDFs then cannot reuse the name.
- `K_BUILTIN_META_ROWS`: one row at the enum position with flags (`BUILTIN_FLAG_*`), min/max arity (`BUILTIN_META_ARITY_UNBOUNDED` for variadic), hint kind (`BHK_*`), unary scalar kind (`USK_*`, `USK_NONE` when not unary), and category (`BC_*`, which selects the dispatch family).
## 2) Implement in `MathParser.bas`

- Route through the handler for the row's `BC_*` category; follow existing helper patterns.
- Validate arity and type/support rules beyond what the metadata row enforces.
- Arrays: preserve broadcasting, size checks, and return-shape conventions.
- Variadic: enforce minimum arity; process extra args deterministically.
- Decide non-finite handling and follow the other coding-standards rules.

## 3) Signature hint

Hints come from the metadata row's hint kind: `TryGetBuiltinSignatureHint` (Basic) and `getBuiltinFunctionMissingCallHint` (C++) switch on it.

- Reuse an existing `BHK_*` / `BuiltinHintKind` whose text matches the signature (spelling, fixed/variadic arity, argument meaning/optionals).
- Only for a new argument shape: add the `BHK_*` / `BuiltinHintKind` enumerator, its `FB_STR_PAR_*` / `STR_*` text constant, and a `case` in both hint functions. Match neighboring hint formatting.

## 4) Implement in C++

- One `BuiltinFunctionId` enumerator (`cpp/MathParserInternal.hpp`), one `functionNames()` entry, one `kBuiltinMeta` row, and one `kUnaryScalarKind` entry, all at the same position. Update the `static_assert` count on `kBuiltinMeta`.
- Mirror behavior, error text, and exact-int metadata, placing code per cpp-class-design ("Adding a builtin or operator"). Do not mirror Basic's procedure layout onto `MathParser`.
- If the builtin only exists with a feature (complex, time, lambdas, factorint), wrap its C++ code and tests in the matching `#if SMARTMATH_*` block (complex-support flow, build-gates "Feature matrix").

## Operators

Collect before editing: token spelling, precedence and associativity relative to neighbors, unary/binary/postfix form, and conflicts with existing tokens, identifiers, or number literals.

- Register keyword operators (`and`, `or`, `not`, `mod`) only in `OperatorNames` with an `OP_*` index. Symbol operators have no name table; find how a neighbor at the same precedence level is lexed and parsed, and extend that code.
- Place the operator at the same precedence level in both parsers, and test it against each neighboring level and against both associativity directions.
- Follow the real ↔ complex parity rules (complex-support flow) when the operator accepts complex operands.
