# C++ MathParser Class Shape

Load when editing `cpp/MathParser.hpp` or `cpp/MathParser.cpp`, adding a C++ helper, or moving C++ parser code. Basic stays procedural. This flow does not apply to `MathParser.bas`.

Behavior, error text, and exact-int metadata stay mirrored with Basic (parity skill). Type and file layout do not.

## What MathParser is

`MathParser` is the session facade. It owns:

- compile / evaluate lifecycle (`compile`, `evaluate`, `parseAndEvaluate`, `parseAndEvaluateRaw`)
- last result and last error
- variables, constants, and user functions
- the compiled program
- feature flags (`supportComplexNumbers_`, `supportTimeValues_`, `supportLambdaFunctions_`)

Public method names and `MathParser::RawResult` field layout stay stable unless the task asks for an API change.

It orchestrates parse, eval, and builtin dispatch. It does not own value algebra, formatting, lexing utilities, or per-builtin math.

## Member or not (Core Guidelines C.4)

A function is a `MathParser` member only when it reads or writes session state: `variables_`, `userFunctions_`, `compiled_`, `lastResult_`, `lastError_`, or a feature flag.

| The function uses | Put it |
|---|---|
| Only `EvalValue` / `ScalarValue` / integers / doubles | On that value type, or in `smartmath_internal` (`cpp/MathParserInternal.hpp`) |
| Only `EvalContext` error text or the cursor | On `EvalContext` |
| Only `Expr` / `AstStatement` structure | On that AST type, or a file-local function in the `.cpp` |
| A feature flag and otherwise pure math | Pure function taking the flag or a small options value; `MathParser` passes the flag in |
| Session state | `MathParser` method |

Do not add a `static` private `MathParser` method. The existing ones (the `static` declarations in `cpp/MathParser.hpp`) are the extraction set for the cleanup order below.

Do not pass `const MathParser&` into a helper to reach one lookup or one flag. Pass that fact.

A `const` method whose only effect is writing `EvalContext` is not session logic. `const` means the session object is unchanged.

## Where operations live

Move operations off `MathParser` as functions first. Leave the nested types in place while doing that. `EvalValue` stays internal. `RawResult` stays a public nested type. A namespace-scope alias is only for a later move that would otherwise break `MathParser::EvalValue` spellings inside the parser.

- **`EvalValue` / `ScalarValue`** — factories (`makeScalar`, `makeScalarInt`, `makeScalarUInt`, time and complex constructors), exact-metadata flags, negate, numeric real/imag, array walks that only read values. One `ScalarValue` layout stays; `vector<ScalarValue>` stays contiguous.
- **`RawResult`** — public snapshot. Stays a nested public type; field layout stays stable. Conversion from `EvalValue` is an internal function called by `getRawResult()`. Do not add `RawResult::from(EvalValue)`: that would publish `EvalValue`.
- **Formatting** — free functions of `(const EvalValue&, RenderBase)`. `getResult` / `getResultAsHex` / `getResultAsDec` / `getResultAsOct` / `getResultAsBin` stay one-liners on the session.
- **`EvalContext`** — parse cursor and errors. Errors go through `EvalContext::fail(ParseErrorId)`, or `fail(msg)` for the few messages that need a name or a count. Do not add per-error setters (`setFooError`).
- **`Expr`** — node constructors and structural queries (`makeBinaryExpr`, scalar-only checks).
- **Builtin handlers** — file-local functions or an existing family function (`builtinUnaryMath`, `builtinAggregateFamily`, `builtinScalarBinaryFamily`, and the other family routers). Indexed from `kBuiltinMeta` / `BuiltinFunctionId`.

`smartmath_internal` is the home for checked integer arithmetic and exact-int range tests. Extend it before adding a parser method.

Real / complex / time differences are a policy or a function chosen from metadata. They are not derived `MathParser` types. Virtual dispatch on `evalExpr` is out of scope.

## Adding a builtin or operator (C++)

1. Register per implement-function step 4 (`BuiltinFunctionId`, `functionNames()`, `kBuiltinMeta`, `kUnaryScalarKind`). Name tables stay the only spellings.
2. Math in a file-local function or on `ScalarValue` / `smartmath_internal`.
3. Route through the existing family handler when the category already exists.
4. Reuse `ParseErrorId`. A new id is a data change, not a new method.
5. Add a `MathParser::builtin…` method only when the handler must read session state and no family router can host it. The method stays a thin call into the function from step 2.

## Cleanup order

Single home for the extraction order (`mathparser-refactor` links here). Moving functions is cleanup work; feature and fix tasks leave existing class shape alone.

1. Error setters onto `EvalContext::fail`.
2. Static exact-int and scalar factories onto `ScalarValue` / `smartmath_internal`.
3. `toRawResult` and formatting off the session.
4. Lexer helpers (`skipSpaces`, identifier tests, `toLower`, `trim`) to file-local functions.
5. Static builtin helpers off the class.

Leave `compile`, `parseProgram`, `evalExpr`, and `runCompiledProgram` on `MathParser` until those clusters are done. Splitting `MathParser.cpp` into more files while every function remains a class method does not change the design.

During a cleanup task that already edits a static `MathParser` helper, move that helper off the class in the same edit when the declaration, definition, and call sites are in that cluster. Leave unrelated methods where they are. Feature and fix tasks only place **new** helpers per this flow.

## Do not

- Subclass `MathParser`, or split `ScalarValue` into a type per kind.
- Replace the `CartesianScalar` union with `std::variant` on a hot path without a measured reason. Keep tagged accessors as the only readers.
- Pimpl the whole parser as a substitute for moving functions.
- Restructure `MathParser.bas` so it matches the C++ types.
- Change public method names, result strings, or `RawResult` fields as part of a move.
