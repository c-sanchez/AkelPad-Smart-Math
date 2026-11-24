'' Isolated factorint smoke cases with optional single-case mode for timeout harness.
'' Usage:
''   SmokeTest_Factorint.exe           - run all cases (may hang without external timeout)
''   SmokeTest_Factorint.exe <index>   - run one case (1-based), exit 0 on success else 1

#include once "..\Inc\MathParser.bi"

const FACTORINT_TEST_COUNT as Integer = 29

type FactorintCase
  expr as String
  expected as String
  expectedErrContains as String
end type

dim shared g_cases(1 to FACTORINT_TEST_COUNT) as FactorintCase

private sub InitFactorintCases()
  g_cases(1).expr = "factorint(33)": g_cases(1).expected = "(3, 11)"
  g_cases(2).expr = "factorint(12)": g_cases(2).expected = "(2**2, 3)"
  g_cases(3).expr = "factorint(13)": g_cases(3).expected = "(13)"
  g_cases(4).expr = "factorint(-33)": g_cases(4).expected = "(-3, 11)"
  g_cases(5).expr = "factorint(-13)": g_cases(5).expected = "(-13)"
  g_cases(6).expr = "factorint(-12)": g_cases(6).expected = "(-2**2, 3)"
  g_cases(7).expr = "factorint(0)": g_cases(7).expected = "(0)"
  g_cases(8).expr = "factorint(1)": g_cases(8).expected = "(1)"
  g_cases(9).expr = "factorint(-1)": g_cases(9).expected = "(-1)"
  g_cases(10).expr = "factorint(2**52)": g_cases(10).expected = "(2**52)"
  g_cases(11).expr = "factorint(2**63-1)": g_cases(11).expected = "(7**2, 73, 127, 337, 92737, 649657)"
  g_cases(12).expr = "factorint((33,12))": g_cases(12).expectedErrContains = "expects scalar values"
  g_cases(13).expr = "factorint(2**64)": g_cases(13).expectedErrContains = "expects integer values"
  g_cases(14).expr = "factorint(18446744073709551615)": g_cases(14).expected = "(3, 5, 17, 257, 641, 65537, 6700417)"
  g_cases(15).expr = "prod(factorint(12))": g_cases(15).expected = "12"
  g_cases(16).expr = "prod(factorint(-33))": g_cases(16).expected = "-33"
  g_cases(17).expr = "prod(factorint(2**52))": g_cases(17).expected = "4503599627370496"
  g_cases(18).expr = "factorint(33.0)": g_cases(18).expected = "(3, 11)"
  g_cases(19).expr = "factorint(33.5)": g_cases(19).expectedErrContains = "expects integer values"
  g_cases(20).expr = "factorint(90)": g_cases(20).expected = "(2, 3**2, 5)"
  g_cases(21).expr = "factorint(9007)": g_cases(21).expected = "(9007)"
  g_cases(22).expr = "factorint(900719)": g_cases(22).expected = "(900719)"
  g_cases(23).expr = "factorint(90071992)": g_cases(23).expected = "(2**3, 11258999)"
  g_cases(24).expr = "factorint(9007199254)": g_cases(24).expected = "(2, 89, 50602243)"
  g_cases(25).expr = "factorint(900719925474)": g_cases(25).expected = "(2, 3, 12907, 11630897)"
  g_cases(26).expr = "factorint(76568758722)": g_cases(26).expected = "(2, 3**2, 47, 101, 896107)"
  '' Perf regression: 113 * large prime (sqrt(n) > 10^7); was ~10s before MR/trial fixes.
  g_cases(27).expr = "factorint(76568758722112367)": g_cases(27).expected = "(113, 677599634708959)"
  g_cases(28).expr = "factorint(-3333*9)": g_cases(28).expected = "(-3**3, 11, 101)"
  g_cases(29).expr = "factorint(-9999)": g_cases(29).expected = "(-3**2, 11, 101)"
end sub

private function CasePassed(byref c as FactorintCase, byval ok as Boolean, byref actual as String, byref errText as String) as Boolean
  if len(c.expectedErrContains) > 0 then
    return (len(errText) > 0) andalso (instr(lcase(errText), lcase(c.expectedErrContains)) > 0)
  end if
  return ok andalso (len(errText) = 0) andalso (trim(actual) = trim(c.expected))
end function

private function RunCaseIndex(byval idx as Integer, byref elapsedSec as Double) as Integer
  dim c as FactorintCase = g_cases(idx)
  dim t0 as Double = timer
  dim result as Double
  dim resultText as String
  dim isArray as Boolean
  dim ok as Boolean = Parser_TryEvaluateEx(c.expr, result, resultText, isArray)
  elapsedSec = timer - t0

  dim errText as String = Parser_GetLastError()
  dim actual as String
  if ok then
    actual = resultText
  else
    actual = "ERR: " & errText
  end if

  print "CASE " & idx & "/" & FACTORINT_TEST_COUNT & " (" & elapsedSec & "s): " & c.expr
  if CasePassed(c, ok, actual, errText) then
    print "  PASS -> " & actual
    return 0
  end if
  print "  FAIL -> " & actual
  if len(c.expected) > 0 then print "  WANT -> " & c.expected
  if len(c.expectedErrContains) > 0 then print "  WANT ERR CONTAINS -> " & c.expectedErrContains
  return 1
end function

sub Main()
  InitFactorintCases()
  Parser_ClearVariables()

  dim arg1 as String = trim(command(1))
  if len(arg1) > 0 then
    dim onlyIdx as Integer = valint(arg1)
    if onlyIdx < 1 orelse onlyIdx > FACTORINT_TEST_COUNT then
      print "Invalid case index: " & arg1 & " (use 1.." & FACTORINT_TEST_COUNT & ")"
      end 2
    end if
    dim elapsed as Double
    dim rc as Integer = RunCaseIndex(onlyIdx, elapsed)
    end rc
  end if

  print "=== factorint isolated tests ==="
  dim failCount as Integer = 0
  dim i as Integer
  for i = 1 to FACTORINT_TEST_COUNT
    dim elapsed as Double
    if RunCaseIndex(i, elapsed) <> 0 then failCount += 1
  next i
  print "=== Result: failed " & failCount & " / " & FACTORINT_TEST_COUNT & " ==="
  if failCount > 0 then end 1
  end 0
end sub

Main()
