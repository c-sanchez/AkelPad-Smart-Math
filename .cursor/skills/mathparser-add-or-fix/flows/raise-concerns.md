# Raise Limits and Mathematical Concerns

Load for planning and for implementation, including while investigating a different question. The trigger is the agent's own reasoning, not a user question about the case.

## When to raise

As soon as thinking identifies any of the following, raise that concern to the user explicitly:

- some numbers or operations are rejected
- the functionality is limited
- the behavior is mathematically incorrect, limited, or unexplained
- the functionality contradicts mathematical rules, common usage, or common sense

This is required for visibility and an explicit user decision. A concern found beside the requested task counts. `1^18446744073709551615` was valid unsigned XOR and was rejected on one C++ path; that belonged in the user-visible message when it was noticed, not only after the user asked.

Say it in the message the user sees. Do not leave it only in private reasoning, a plan footnote, or the completion report's leftover list.

State the input, the result or error, which side does it (Basic, C++, or a particular evaluation path), and the mathematical, common-usage, or common-sense expectation it breaks.

## Propose a fix with the report

In the same user-visible message, propose a fix or change for each concern, using your best judgement. Do this while planning and while implementing. The proposal is a recommendation, not a decision already made.

For every proposed fix or change, give examples of the behavior before and the behavior after the change. Each example names the input and both results (the value or the error text). Cover the case that is wrong today and one nearby case that should stay as it is when that contrast is what makes the proposal clear.

One proposal can cover several inputs of the same kind. Separate concerns get separate proposals.

## Wait for the decision

Until the user decides, do not:

- choose whether the case is accepted or rejected
- encode the rejection or limit as a test expectation
- call it out of scope, intentional, or a known difference
- finish the plan, or finish the implementation, as though the limit were accepted

Planning stops on these items and asks. The plan includes the proposed fix and its before/after examples, and it does not treat that proposal as accepted or record a "leave it".

Implementation stops at the concern, including one found mid-task from reading code or from a probe. Continue only the part the user already decided.

## Already decided

These are not new concerns:

- behavior the user specified in this task
- behavior specified by `USAGE_AND_SYNTAX.md`, by empty-array-policy, or by another rule in these skills
- a difference the user already accepted that is recorded in the parity skill
- floating-point inaccuracy as defined by smoke-tests "Handling a test mismatch"

A new rejection or limit noticed while applying one of those rules is still raised.

When the user decides to keep a rejection or limit, record it per fix-behavior "Report what stays different" and the parity skill's known differences. Tests may lock it only after that decision.
