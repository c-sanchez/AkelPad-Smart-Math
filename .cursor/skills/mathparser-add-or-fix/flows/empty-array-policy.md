# Empty Array Policy

Load when adding or changing a builtin function or operator, or when changing array / `()` results, errors, tests, or docs. Basic and C++ match. Do not add an "empty array is not supported" error.

## No scalars left

When the inputs flatten to no scalars, the result is `()`. The operation does not invent a number (`0` for `sum`, `1` for `prod`).

This covers:

- Elementwise functions: trig and hyperbolic, `exp` / `ln` / `log` / `log10`, `pow` / `sqrt` / `sqr` / `hypot`, numeric utilities (`abs`, `sign`, rounding, `frac`, `clamp`, `gcd`, `lcm`, `mod`, `ncr`, `npr`, `fact` / `factorial`), `ratio`, `deg` / `rad`, base formatters (`hex`, `oct`, `bin`, `uhex`, `uoct`, `ubin`), and time converters (`days`, `hours`, `minutes`, `seconds`, `milliseconds`).
- Elementwise operators: `+`, `-`, `*`, `/`, `**`, modulo `%`, bitwise `&` `^` `|` `<<` `>>` `~`, unary `+` and unary `-`. `sin(())`, `()+1`, and `-()` are `()`.
- Structural functions: `sort` / `sorted`, `sortby`, `filter`, `reverse` / `reversed`, `unique`, `flat`, `unpack`, `repeat`.
- Reductions: `sum`, `prod` / `product`, `avg` / `mean`, `median`, `variance`, `stddev`, `min`, `max`. `product(())`, `max(())`, and `stddev((), (), ())` are `()`.

A reduction that still has a value ignores empty arrays and folds the rest: `sum((), 1)` is `1`, `min((), (3,1))` is `1`. Python identities (`sum([]) == 0`, `math.prod([]) == 1`) are not the reference.

`sum((1,2))` is a number and `sum(())` is an array. That return-kind change is part of the rule.

## Still an error, or still a scalar

- A missing call is an arity error (`sum()` expects at least one argument). `filter` and `sortby` still require a one-parameter function. `repeat((), 0)` still expects a positive integer.
- `len(())` / `length(())` is `0`. `len()` with no arguments is `0`.
- An empty array is falsy: `!()` and `not ()` are `1`, `() && 1` is `0`, `() || 0` is `0`.
- Comparisons use length: `() == ()` is `1`, `() == 0` is `0`.
- `()+(1,2)` and other length mismatches stay `incompatible operands`.
- `range` with an empty integer span returns `()` (`range(5, 1)`, `range(1, 1)`, `range(1, 5, -1)`). Non-integers stay `expects integer values`. Whole-valued floats such as `1.0` stay integers. Step `0` stays `step must not be zero`. The 1000-item cap stays.
- `range((), 5)` stays `expects integer values`.
- `factorint` and `random` reject every array with `expects scalar values`.
- `cart` and `polar` with the wrong shape stay `incompatible operands`.
- Postfix `%` stays `percentage requires a scalar value` for every array.
- `()[0]` stays `array index is out of range`. A slice that selects nothing stays `()`.
- `clamp(5, (), 9)` stays `expects scalar min/max`. `clamp((), 0, 9)` is `()`.
- An array literal still flattens, so `((), ())` is `()`.

## User functions and lambdas

No extra empty-array check. The body runs with `()` bound to the parameter. `f(x)=x+1; f(())` is `()`. `f(x)=sum(x); f(())` is `()`. `f(a)=a[0]; f(())` stays the index error.
