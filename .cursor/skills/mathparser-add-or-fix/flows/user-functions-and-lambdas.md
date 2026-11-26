# User Functions and Lambdas

Load when the task analyzes, verifies, tests, corrects, optimizes, refactors, or changes a user-defined function or a lambda. Either surface is the whole callable. Tests, docs, and gates still follow the `SKILL.md` workflow.

## Rule

- Treat user-defined functions and lambdas as two surfaces of the same callable.
- When user-defined functions are analyzed, verified, tested, corrected, optimized, refactored, or changed, do the same with lambdas.
- When lambdas are analyzed, verified, tested, corrected, optimized, refactored, or changed, do the same with user-defined functions.
- Reuse one implementation of each shared rule. A change on one surface can change the other, including when the task named only one.

## One callable

A user-defined function and a lambda are a parameter list plus a body expression.

- Named, classic: `f(x)=x*x+1`, `mix(a,b)=a*2+b`.
- Named, lambda-style: `f=x:x*x+1`, `f=(x,y):(x+y)`, `f=():42`. This stores the same user function as the classic form. It is not a variable assignment.
- Anonymous: `sortby((3,1,2), x:-x)`, `filter((1,0,2), x: x)`, including wrapped `(x):(1/x)`. This is a value (`VK_INLINE_LAMBDA` / an inline-lambda payload), not a table entry.

The lambda support flag gates lambda syntax only (named lambda-style definitions and anonymous lambdas). Classic `f(x)=...`, and a named function passed to `sortby` or `filter`, stay available when the flag is off.

## Both surfaces, every time

A body, parameter, arity, error, empty-array, non-finite, or exact-int change on one surface is the same change on the others that can express it.

Name-only rules stay on named functions, and still on both definition spellings: a reserved function, constant, or `ans` as the function name; replacing a variable of the same name; late binding of a referenced name; direct recursion of that name. Parameter rules (duplicate names, a reserved constant as a parameter, shadowing a global during the body check) apply to anonymous lambdas too.

`sortby` and `filter` require one parameter for a named function and for an anonymous lambda. A zero-parameter lambda is rejected there. A zero-parameter named function is still callable as `f()`.

After the edit, probe the surfaces the rule can express, not only the spelling in the request. Say in the report which surfaces were checked and which shared helper owns the rule.

## Reuse the shared code

Do not copy a rule onto the other surface. Extend the helper both surfaces already call. If the two call sites have diverged on a shared rule, collapse that divergence into the helper.

Shared today:

| Rule | Basic | C++ |
|---|---|---|
| Bind parameters and evaluate the body | `EvaluateBoundBody`, called by `EvaluateUserFunction` and `EvaluateInlineLambda` | `tryCompileExprToProgram` + `runBoundProgram` (`runCompiledProgram`), called by `evalUserFunctionCall` and `evalInlineLambdaCall` |
| Syntax-check the body | `SyntaxCheckBodyWithParams` (`ParseSyntaxCheckedBody`), called by `TryValidateUserFunctionBodyExpression` and `SyntaxCheckInlineLambdaBody` | `tryParseSyntaxCheckedExpression` / `getUserFunctionDefinitionErrorText` |
| Empty body, including `()` | `UdfBodyIsEmptyTupleLiteral`; `InlineLambdaBodyIsEffectivelyEmpty` is that check | `bodyIsEffectivelyEmpty` in `getUserFunctionDefinitionErrorText`, `makeUnaryInlineLambdaExpr`, and `evalInlineLambdaCall` |
| Reject a definition (names, empty body, direct recursion, body syntax) | `TryValidateUserFunctionDefinition` for both `f(x)=...` and `f=x:...` | `getUserFunctionDefinitionErrorText` / `trySetUserFunctionDefinitionError` for both; lambda-style assignment stores a `UserFunction` |
| Parameter names during the body check and while the body is compiled | `functionVariableNames` while the body is checked; the same names are variables when the body runs | `boundBodyParamNames` on the parse context in `getUserFunctionDefinitionErrorText`, `makeUnaryInlineLambdaExpr`, and `tryCompileExprToProgram`. A nested lambda binds its own parameters for that check |

These call-site differences stay. Do not merge them away:

- An anonymous call has no function name: arity errors say `lambda`, bare-name tail hints are suppressed, and error columns are remapped onto the lambda argument.
- A named call tracks the call stack and rejects recursion of that name.
- C++ caches a named function's compiled body and compiles an anonymous body on each call.

Any other divergence between the two surfaces is a candidate to collapse into the shared helper. Basic and C++ still mirror each other (parity skill); sharing code inside one language does not replace that mirror.

## Tests and docs

- For a body or parameter rule, add the classic definition, the lambda-style definition of the same body, and the anonymous `sortby` or `filter` form when that form can express the rule.
- For a name-only rule, add both definition spellings.
- Flag-off cases stay in the lambda option suites (smoke-tests). Ordinary cases stay in the main tables; lambda support is on by default.
- Document the rule once. Show both spellings where the reader needs to see that they match (usage-docs, User-Defined Functions).
- Empty arrays follow empty-array-policy on every surface: the body runs with `()` bound to the parameter, with no extra empty-array check.
