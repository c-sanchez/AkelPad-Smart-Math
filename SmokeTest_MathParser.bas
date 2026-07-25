#include once "Inc\MathParser.bi"
#include "tools\neg_band_cases_generated.bas"
#include "tools\float_magnitude_cases_generated.bas"
#include "tools\naninf_mirror_generated.bas"
#include "tools\complex_coverage_generated.bas"
#include "tools\smoke_parity_append_generated.bas"

type SmokeCase
  expr as String
  setup as String
  expected as String
  expectedErrContains as String
  expectNoResult as Boolean
end type

dim shared g_total as Integer = 0
dim shared g_passed as Integer = 0
dim shared g_failed as Integer = 0
dim shared g_idx as Integer = 0

private function IsDecimalNumberToken(byref sText as String) as Boolean
  dim s as String = trim(sText)
  if len(s) = 0 then return FALSE
  dim i as Integer = 1
  if mid(s, i, 1) = "+" orelse mid(s, i, 1) = "-" then i += 1
  if i > len(s) then return FALSE

  dim hasDigits as Boolean = FALSE
  dim hasDot as Boolean = FALSE
  dim hasExp as Boolean = FALSE
  dim expHasDigits as Boolean = FALSE

  while i <= len(s)
    dim ch as String = mid(s, i, 1)
    if ch >= "0" andalso ch <= "9" then
      hasDigits = TRUE
      if hasExp then expHasDigits = TRUE
      i += 1
      continue while
    end if
    if ch = "." then
      if hasDot orelse hasExp then return FALSE
      hasDot = TRUE
      i += 1
      continue while
    end if
    if ch = "e" orelse ch = "E" then
      if hasExp orelse hasDigits = FALSE then return FALSE
      hasExp = TRUE
      i += 1
      if i <= len(s) then
        dim signCh as String = mid(s, i, 1)
        if signCh = "+" orelse signCh = "-" then i += 1
      end if
      if i > len(s) then return FALSE
      continue while
    end if
    return FALSE
  wend

  if hasDigits = FALSE then return FALSE
  if hasExp andalso expHasDigits = FALSE then return FALSE
  return TRUE
end function

private function TryParseDecimalToken(byref sText as String, byref outV as Double) as Boolean
  if IsDecimalNumberToken(sText) = FALSE then return FALSE
  outV = val(trim(sText))
  return TRUE
end function

private function SmokeTrimScalarToken(byref s as String) as String
  return trim(s)
end function

'' -1 = -inf, 0 = not inf, 1 = +inf (case-insensitive).
private function SmokeInfSignFromLiteral(byref s as String) as Integer
  dim t as String = lcase(SmokeTrimScalarToken(s))
  if t = "inf" orelse t = "+inf" then return 1
  if t = "-inf" then return -1
  return 0
end function

private function SmokeIsExactZeroLiteral(byref s as String) as Boolean
  dim t as String = SmokeTrimScalarToken(s)
  return (t = "0") orelse (t = "0.0") orelse (t = "-0") orelse (t = "-0.0") orelse (t = "+0") orelse (t = "+0.0")
end function

private function SmokeIsExactOneLiteral(byref s as String) as Boolean
  dim t as String = SmokeTrimScalarToken(s)
  return (t = "1") orelse (t = "-1") orelse (t = "1.0") orelse (t = "-1.0") orelse (t = "+1") orelse (t = "+1.0")
end function

private function SmokeOneSignFromLiteral(byref s as String) as Integer
  dim t as String = SmokeTrimScalarToken(s)
  if t = "1" orelse t = "1.0" orelse t = "+1" orelse t = "+1.0" then return 1
  if t = "-1" orelse t = "-1.0" then return -1
  return 0
end function

private function SmokeExpectedRequiresStrictScalarMatch(byref expected as String) as Boolean
  return SmokeIsExactZeroLiteral(expected) orelse SmokeIsExactOneLiteral(expected) orelse (SmokeInfSignFromLiteral(expected) <> 0)
end function

'' Expected is 0/0.0, +/-1/1.0, or +/-Inf: actual must use the same literal family (no near-miss floats).
private function SmokeStrictScalarMatch(byref actual as String, byref expected as String) as Boolean
  if SmokeIsExactZeroLiteral(expected) then
    if SmokeIsExactZeroLiteral(actual) = FALSE then return FALSE
    dim da as Double
    dim de as Double
    if TryParseDecimalToken(actual, da) = FALSE then return FALSE
    if TryParseDecimalToken(expected, de) = FALSE then return FALSE
    return (da = 0.0) andalso (de = 0.0)
  end if
  if SmokeIsExactOneLiteral(expected) then
    if SmokeIsExactOneLiteral(actual) = FALSE then return FALSE
    if SmokeOneSignFromLiteral(actual) <> SmokeOneSignFromLiteral(expected) then return FALSE
    dim da as Double
    dim de as Double
    if TryParseDecimalToken(actual, da) = FALSE then return FALSE
    if TryParseDecimalToken(expected, de) = FALSE then return FALSE
    if SmokeOneSignFromLiteral(expected) = 1 then
      return (da = 1.0) andalso (de = 1.0)
    end if
    return (da = -1.0) andalso (de = -1.0)
  end if
  dim eInf as Integer = SmokeInfSignFromLiteral(expected)
  if eInf <> 0 then
    return (SmokeInfSignFromLiteral(actual) = eInf)
  end if
  return FALSE
end function

private function ScalarCloseEnough(byref actual as String, byref expected as String) as Boolean
  dim ea as String = SmokeTrimScalarToken(actual)
  dim ee as String = SmokeTrimScalarToken(expected)
  if ea = ee then return TRUE
  if SmokeExpectedRequiresStrictScalarMatch(ee) then
    return SmokeStrictScalarMatch(ea, ee)
  end if
  dim da as Double
  dim de as Double
  if TryParseDecimalToken(ea, da) = FALSE then return FALSE
  if TryParseDecimalToken(ee, de) = FALSE then return FALSE
  if da = de then return TRUE
  dim scale as Double = abs(da)
  if abs(de) > scale then scale = abs(de)
  if scale < 1 then scale = 1
  dim tol as Double = 256.0 * 2.2204460492503131e-16 * scale
  if tol < 1e-14 then tol = 1e-14
  return abs(da - de) <= tol
end function

private function SplitComplexText(byref s as String, byref rePart as String, byref imPart as String) as Boolean
  dim t as String = trim(s)
  if len(t) = 0 then
    rePart = "0"
    imPart = "0"
    return TRUE
  end if
  if (right(t, 1) <> "i") andalso (right(t, 1) <> "I") then return FALSE
  dim splitAt as Integer = 0
  dim i as Integer
  for i = len(t) to 2 step -1
    dim ch as String = mid(t, i, 1)
    if (ch = "+") orelse (ch = "-") then
      if i > 1 then
        dim prevCh as String = mid(t, i - 1, 1)
        if (prevCh <> "e") andalso (prevCh <> "E") then
          splitAt = i
          exit for
        end if
      elseif ch = "-" then
        splitAt = i
        exit for
      end if
    end if
  next i
  if splitAt = 0 then
    rePart = "0"
    imPart = left(t, len(t) - 1)
    return TRUE
  end if
  rePart = trim(left(t, splitAt - 1))
  if len(rePart) = 0 then rePart = "0"
  imPart = trim(mid(t, splitAt))
  if left(imPart, 1) = "+" then imPart = mid(imPart, 2)
  if len(imPart) > 0 then
    dim imLast as String = right(imPart, 1)
    if (imLast = "i") orelse (imLast = "I") then imPart = left(imPart, len(imPart) - 1)
  end if
  return TRUE
end function

private function ComplexResultCloseEnough(byref actual as String, byref expected as String) as Boolean
  dim aRe as String, aIm as String, eRe as String, eIm as String
  if SplitComplexText(actual, aRe, aIm) = FALSE then return FALSE
  if SplitComplexText(expected, eRe, eIm) = FALSE then return FALSE
  if ScalarCloseEnough(aRe, eRe) = FALSE then return FALSE
  return ScalarCloseEnough(aIm, eIm)
end function

private function SplitArrayElems(byref sText as String, elems() as String) as Integer
  erase elems
  dim s as String = trim(sText)
  if len(s) < 2 then return 0
  if left(s, 1) <> "(" orelse right(s, 1) <> ")" then return 0
  s = mid(s, 2, len(s) - 2)
  if len(trim(s)) = 0 then return 0

  dim count as Integer = 1
  for i as Integer = 1 to len(s)
    if mid(s, i, 1) = "," then count += 1
  next i
  redim elems(0 to count - 1)
  dim startPos as Integer = 1
  dim outIdx as Integer = 0
  for i as Integer = 1 to len(s)
    if mid(s, i, 1) = "," then
      elems(outIdx) = trim(mid(s, startPos, i - startPos))
      outIdx += 1
      startPos = i + 1
    end if
  next i
  elems(outIdx) = trim(mid(s, startPos))
  return count
end function

private function ResultCloseEnough(byref actual as String, byref expected as String) as Boolean
  if actual = expected then return TRUE
  dim aElems() as String, eElems() as String
  dim ac as Integer = SplitArrayElems(actual, aElems())
  dim ec as Integer = SplitArrayElems(expected, eElems())
  if ac > 0 orelse ec > 0 then
    if ac <> ec then return FALSE
    for i as Integer = 0 to ac - 1
      if ResultCloseEnough(aElems(i), eElems(i)) = FALSE then return FALSE
    next i
    return TRUE
  end if
  if (instr(lcase(actual), "i") > 0) orelse (instr(lcase(expected), "i") > 0) then
    return ComplexResultCloseEnough(actual, expected)
  end if
  return ScalarCloseEnough(actual, expected)
end function

private function SmokeCaseSignature(byref c as SmokeCase) as String
  dim kind as String
  dim payload as String
  if c.expectNoResult then
    kind = "noResult"
    payload = ""
  elseif len(c.expectedErrContains) > 0 then
    kind = "errorContains"
    payload = c.expectedErrContains
  else
    kind = "expected"
    payload = c.expected
  end if
  SmokeCaseSignature = kind & "|" & c.setup & "|" & c.expr & "|" & payload
end function

sub RunCase(byref c as SmokeCase)
  g_idx += 1
  if len(c.setup) > 0 then
    print "[" & g_idx & "/" & g_total & "] RUN  : " & c.setup & "  ->  " & c.expr
  else
    print "[" & g_idx & "/" & g_total & "] RUN  : " & c.expr
  end if

  dim result as Double
  dim resultText as String
  dim isArray as Boolean
  dim ok as Boolean = FALSE
  if len(c.setup) > 0 then
    dim setupOk as Boolean = Parser_TryEvaluateEx(c.setup, result, resultText, isArray)
    if setupOk = FALSE then
      g_failed += 1
      print "          FAIL : setup """ & c.setup & """ failed: """ & Parser_GetLastError() & """"
      print ""
      exit sub
    end if
  end if
  ok = Parser_TryEvaluateEx(c.expr, result, resultText, isArray)

  dim actual as String
  dim errText as String
  errText = Parser_GetLastError()
  if ok then
    actual = resultText
  else
    actual = "ERR: " & errText
  end if

  dim passCase as Boolean = FALSE
  if c.expectNoResult then
    if (ok = FALSE) andalso (len(errText) = 0) then
      passCase = TRUE
    end if
  elseif len(c.expectedErrContains) > 0 then
    if (len(errText) > 0) andalso (instr(lcase(errText), lcase(c.expectedErrContains)) > 0) then
      passCase = TRUE
    end if
  else
    if (len(errText) = 0) andalso ResultCloseEnough(actual, c.expected) then passCase = TRUE
  end if

  if passCase then
    g_passed += 1
    if c.expectNoResult then
      print "          PASS : no result (comment/empty line)"
    elseif len(c.expectedErrContains) > 0 then
      print "          PASS : got expected error """ & c.expectedErrContains & """"
    else
      print "          PASS : got """ & actual & """"
    end if
  else
    g_failed += 1
    print "          FAIL : got """ & actual & """"
    if c.expectNoResult then
      print "                 expected no result and no error"
    elseif len(c.expectedErrContains) > 0 then
      print "                 expected error containing """ & c.expectedErrContains & """"
    else
      print "                 expected """ & c.expected & """"
    end if
  end if
  print ""
end sub

private sub RunRawResultApiTests()
  print "=== RawResult API (parser last-eval snapshot) ==="
  dim subPass as Integer = 0
  dim subFail as Integer = 0
  dim r as Double
  dim rt as String
  dim ia as Boolean
  dim raw as RawResult

  Parser_ClearVariables()
  if Parser_TryEvaluateEx("ratio(0.5)", r, rt, ia) = FALSE then
    print "[raw] FAIL: ratio(0.5) eval"
    subFail += 1
  elseif Parser_GetLastRawResult(raw) = FALSE orelse raw.kind <> RRK_SCALAR orelse raw.scalar.kind <> RSK_RATIONAL _
    orelse raw.scalar.real.ratNum <> 1 orelse raw.scalar.real.ratDen <> 2 then
    print "[raw] FAIL: ratio(0.5) raw rational"
    subFail += 1
  else
    print "[raw] PASS: ratio(0.5) -> rational 1/2"
    subPass += 1
  end if

  if Parser_TryEvaluateEx("2+3", r, rt, ia) = FALSE then
    print "[raw] FAIL: 2+3 eval"
    subFail += 1
  elseif Parser_GetLastRawResult(raw) = FALSE orelse raw.kind <> RRK_SCALAR orelse raw.scalar.real.kind <> RSK_INT64 orelse raw.scalar.real.intValue <> 5 then
    print "[raw] FAIL: 2+3 raw int64"
    subFail += 1
  else
    print "[raw] PASS: 2+3 -> int64 5"
    subPass += 1
  end if

  if Parser_TryEvaluateEx("ratio((1,2,3))", r, rt, ia) = FALSE orelse ia = FALSE then
    print "[raw] FAIL: ratio array eval"
    subFail += 1
  elseif Parser_GetLastRawResult(raw) = FALSE orelse raw.kind <> RRK_ARRAY orelse ubound(raw.arr) <> 2 then
    print "[raw] FAIL: ratio array shape"
    subFail += 1
  else
    dim badArrKind as Boolean = FALSE
    if raw.arr(0).kind <> RSK_INT64 then badArrKind = TRUE
    if raw.arr(1).kind <> RSK_INT64 then badArrKind = TRUE
    if raw.arr(2).kind <> RSK_INT64 then badArrKind = TRUE
    if badArrKind then
      print "[raw] FAIL: ratio array element kinds"
      subFail += 1
    else
      print "[raw] PASS: ratio((1,2,3)) array of integers"
      subPass += 1
    end if
  end if

  if Parser_TryEvaluateEx("f(x)=1", r, rt, ia) = FALSE then
    print "[raw] FAIL: UDF defined eval"
    subFail += 1
  elseif Parser_GetLastRawResult(raw) then
    print "[raw] FAIL: defined should not expose raw result"
    subFail += 1
  else
    print "[raw] PASS: defined clears raw snapshot"
    subPass += 1
  end if

  Parser_SetSupportComplexNumbers(TRUE)
  if Parser_TryEvaluateEx("ratio(0.5+0.25i)", r, rt, ia) = FALSE then
    print "[raw] FAIL: complex ratio eval"
    subFail += 1
  elseif Parser_GetLastRawResult(raw) = FALSE orelse raw.scalar.kind <> RSK_COMPLEX orelse raw.scalar.real.kind <> RSK_RATIONAL _
    orelse raw.scalar.imag.kind <> RSK_RATIONAL then
    print "[raw] FAIL: complex ratio raw parts"
    subFail += 1
  else
    print "[raw] PASS: ratio(0.5+0.25i) complex+rational parts"
    subPass += 1
  end if
  Parser_SetSupportComplexNumbers(FALSE)

  if Parser_TryEvaluateEx("sqrt(81)", r, rt, ia) = FALSE then
    print "[raw] FAIL: sqrt(81) eval"
    subFail += 1
  elseif Parser_GetLastRawResult(raw) = FALSE orelse raw.kind <> RRK_SCALAR orelse raw.scalar.real.kind <> RSK_INT64 orelse raw.scalar.real.intValue <> 9 then
    print "[raw] FAIL: sqrt(81) raw int64 9"
    subFail += 1
  else
    print "[raw] PASS: sqrt(81) -> verified exact int64"
    subPass += 1
  end if

  if Parser_TryEvaluateEx("sqrt(4611686014132420611)", r, rt, ia) = FALSE then
    print "[raw] FAIL: sqrt(4611686014132420611) eval"
    subFail += 1
  elseif Parser_GetLastRawResult(raw) = FALSE orelse raw.kind <> RRK_SCALAR orelse raw.scalar.real.kind <> RSK_FLOATING then
    print "[raw] FAIL: sqrt(4611686014132420611) must be floating (not fake exact int)"
    subFail += 1
  else
    print "[raw] PASS: sqrt(4611686014132420611) -> floating (no verified square)"
    subPass += 1
  end if

  if Parser_TryEvaluateExRaw("ratio(0.5)", raw) = FALSE then
    print "[raw] FAIL: Parser_TryEvaluateExRaw ratio(0.5)"
    subFail += 1
  elseif raw.kind <> RRK_SCALAR orelse raw.scalar.real.kind <> RSK_RATIONAL then
    print "[raw] FAIL: Parser_TryEvaluateExRaw rational kind"
    subFail += 1
  else
    print "[raw] PASS: Parser_TryEvaluateExRaw"
    subPass += 1
  end if

  if Parser_TryEvaluateExRaw("f(x)=2", raw) = FALSE then
    print "[raw] FAIL: Parser_TryEvaluateExRaw UDF defined should succeed"
    subFail += 1
  elseif RawResultHasValue(raw) then
    print "[raw] FAIL: Parser_TryEvaluateExRaw defined should leave empty raw"
    subFail += 1
  else
    print "[raw] PASS: Parser_TryEvaluateExRaw defined (no raw value)"
    subPass += 1
  end if

  g_passed += subPass
  g_failed += subFail
  print "RawResult sub-tests: passed " & str(subPass) & ", failed " & str(subFail)
  print ""
end sub

private sub RunNegBandRealBatch(byref subPass as Integer, byref subFail as Integer)
  dim r as Double
  dim rt as String
  dim ia as Boolean
  dim i as Integer
  for i = 1 to NEG_BAND_REAL_COUNT
    if negBandRealIsErr(i) then
      if Parser_TryEvaluateEx(negBandRealExpr(i), r, rt, ia) then
        print "[neg-band-real] FAIL: "; negBandRealLabel(i); " expected error, got """; rt; """"
        subFail += 1
      elseif instr(lcase(Parser_GetLastError()), lcase(negBandRealErr(i))) = 0 then
        print "[neg-band-real] FAIL: "; negBandRealLabel(i); " err="; Parser_GetLastError()
        subFail += 1
      else
        subPass += 1
      end if
    else
      if Parser_TryEvaluateEx(negBandRealExpr(i), r, rt, ia) = FALSE orelse ResultCloseEnough(rt, negBandRealExpect(i)) = FALSE then
        print "[neg-band-real] FAIL: "; negBandRealLabel(i); " "; negBandRealExpr(i); " -> """; rt; """ want """; negBandRealExpect(i); """"
        subFail += 1
      else
        subPass += 1
      end if
    end if
  next i
end sub

private sub RunNegBandCxBatch(byref subPass as Integer, byref subFail as Integer)
  dim r as Double
  dim rt as String
  dim ia as Boolean
  dim i as Integer
  for i = 1 to NEG_BAND_CX_COUNT
    if negBandCxIsErr(i) then
      if Parser_TryEvaluateEx(negBandCxExpr(i), r, rt, ia) then
        print "[neg-band-cx] FAIL: "; negBandCxLabel(i); " expected error, got """; rt; """"
        subFail += 1
      elseif instr(lcase(Parser_GetLastError()), lcase(negBandCxErr(i))) = 0 then
        print "[neg-band-cx] FAIL: "; negBandCxLabel(i); " err="; Parser_GetLastError()
        subFail += 1
      else
        subPass += 1
      end if
    else
      if Parser_TryEvaluateEx(negBandCxExpr(i), r, rt, ia) = FALSE orelse ResultCloseEnough(rt, negBandCxExpect(i)) = FALSE then
        print "[neg-band-cx] FAIL: "; negBandCxLabel(i); " "; negBandCxExpr(i); " -> """; rt; """ want """; negBandCxExpect(i); """"
        subFail += 1
      else
        subPass += 1
      end if
    end if
  next i
end sub

private sub RunPosBandRealBatch(byref subPass as Integer, byref subFail as Integer)
  dim r as Double
  dim rt as String
  dim ia as Boolean
  dim i as Integer
  for i = 1 to POS_BAND_REAL_COUNT
    if posBandRealIsErr(i) then
      if Parser_TryEvaluateEx(posBandRealExpr(i), r, rt, ia) then
        print "[pos-band-real] FAIL: "; posBandRealLabel(i); " expected error, got """; rt; """"
        subFail += 1
      elseif instr(lcase(Parser_GetLastError()), lcase(posBandRealErr(i))) = 0 then
        print "[pos-band-real] FAIL: "; posBandRealLabel(i); " err="; Parser_GetLastError()
        subFail += 1
      else
        subPass += 1
      end if
    else
      if Parser_TryEvaluateEx(posBandRealExpr(i), r, rt, ia) = FALSE orelse ResultCloseEnough(rt, posBandRealExpect(i)) = FALSE then
        print "[pos-band-real] FAIL: "; posBandRealLabel(i); " "; posBandRealExpr(i); " -> """; rt; """ want """; posBandRealExpect(i); """"
        subFail += 1
      else
        subPass += 1
      end if
    end if
  next i
end sub

private sub RunPosBandCxBatch(byref subPass as Integer, byref subFail as Integer)
  dim r as Double
  dim rt as String
  dim ia as Boolean
  dim i as Integer
  for i = 1 to POS_BAND_CX_COUNT
    if posBandCxIsErr(i) then
      if Parser_TryEvaluateEx(posBandCxExpr(i), r, rt, ia) then
        print "[pos-band-cx] FAIL: "; posBandCxLabel(i); " expected error, got """; rt; """"
        subFail += 1
      elseif instr(lcase(Parser_GetLastError()), lcase(posBandCxErr(i))) = 0 then
        print "[pos-band-cx] FAIL: "; posBandCxLabel(i); " err="; Parser_GetLastError()
        subFail += 1
      else
        subPass += 1
      end if
    else
      if Parser_TryEvaluateEx(posBandCxExpr(i), r, rt, ia) = FALSE orelse ResultCloseEnough(rt, posBandCxExpect(i)) = FALSE then
        print "[pos-band-cx] FAIL: "; posBandCxLabel(i); " "; posBandCxExpr(i); " -> """; rt; """ want """; posBandCxExpect(i); """"
        subFail += 1
      else
        subPass += 1
      end if
    end if
  next i
end sub

'' USAGE_AND_SYNTAX.md precedence (high -> low): **, unary, postfix %, */%, +- ...
private function TrigEvalAbs(byval expr as String) as Double
  dim r as Double
  dim rt as String
  dim ia as Boolean
  if Parser_TryEvaluateEx(expr, r, rt, ia) = FALSE then return -1.0
  return abs(r)
end function

'' sin/cos/tan: large integers must not hit the false N*pi shortcut (IsMultipleOf fmod fix).
private sub RunTrigAngleReductionTests()
  print "=== Trig angle-reduction (large integers vs exact multiples) ==="
  dim subPass as Integer = 0
  dim subFail as Integer = 0
  dim nonzeroExpr(1 to 12) as String
  nonzeroExpr(1) = "sin(2**52)"
  nonzeroExpr(2) = "sin(2**51)"
  nonzeroExpr(3) = "sin(2**50)"
  nonzeroExpr(4) = "sin(2**49)"
  nonzeroExpr(5) = "sin(2**48)"
  nonzeroExpr(6) = "sin(2**47)"
  nonzeroExpr(7) = "sin(3**32)"
  nonzeroExpr(8) = "sin(7**18)"
  nonzeroExpr(9) = "cos(2**52)"
  nonzeroExpr(10) = "cos(2**51)"
  nonzeroExpr(11) = "tan(2**52)"
  nonzeroExpr(12) = "tan(3**32)"
  dim nearZeroExpr(1 to 6) as String
  nearZeroExpr(1) = "sin(pi)"
  nearZeroExpr(2) = "sin(2*pi)"
  nearZeroExpr(3) = "cos(pi/2)"
  nearZeroExpr(4) = "cos(3*pi/2)"
  nearZeroExpr(5) = "tan(pi)"
  nearZeroExpr(6) = "tan(2*pi)"
  dim exactOneExpr(1 to 4) as String
  dim exactOneExpect(1 to 4) as String
  exactOneExpr(1) = "tan(pi/4)": exactOneExpect(1) = "1"
  exactOneExpr(2) = "tan(3*pi/4)": exactOneExpect(2) = "-1"
  exactOneExpr(3) = "tan(-pi/4)": exactOneExpect(3) = "-1"
  exactOneExpr(4) = "tan(-3*pi/4)": exactOneExpect(4) = "1"
  dim ti as Integer
  dim mag as Double
  for ti = 1 to 12
    mag = TrigEvalAbs(nonzeroExpr(ti))
    if mag < 1e-6 then
      print "[trig-reduction] FAIL: """ & nonzeroExpr(ti) & """ |result|=" & str(mag) & " (expected non-zero)"
      subFail += 1
    else
      print "[trig-reduction] PASS: """ & nonzeroExpr(ti) & """ |result|=" & str(mag)
      subPass += 1
    end if
  next ti
  for ti = 1 to 6
    mag = TrigEvalAbs(nearZeroExpr(ti))
    if mag >= 1e-9 then
      print "[trig-reduction] FAIL: """ & nearZeroExpr(ti) & """ |result|=" & str(mag) & " (expected ~0)"
      subFail += 1
    else
      print "[trig-reduction] PASS: """ & nearZeroExpr(ti) & """ |result|=" & str(mag)
      subPass += 1
    end if
  next ti
  for ti = 1 to 4
    dim tr as Double
    dim trt as String
    dim tia as Boolean
    if Parser_TryEvaluateEx(exactOneExpr(ti), tr, trt, tia) = FALSE orelse trt <> exactOneExpect(ti) then
      print "[trig-reduction] FAIL: """ & exactOneExpr(ti) & """ -> """ & trt & """ (expected """ & exactOneExpect(ti) & """)"
      subFail += 1
    else
      print "[trig-reduction] PASS: """ & exactOneExpr(ti) & """ -> """ & trt & """"
      subPass += 1
    end if
  next ti
  g_passed += subPass
  g_failed += subFail
  print "Trig angle-reduction sub-tests: passed " & str(subPass) & ", failed " & str(subFail)
  print ""
end sub

private sub RunOperatorPrecedenceDocTests()
  print "=== Operator precedence (USAGE_AND_SYNTAX.md) ==="
  dim subPass as Integer = 0
  dim subFail as Integer = 0
  dim precOk(1 to 24) as String
  dim precExpect(1 to 24) as String
  precOk(1) = "-2**2": precExpect(1) = "-4"
  precOk(2) = "-(2**2)": precExpect(2) = "-4"
  precOk(3) = "2**-2": precExpect(3) = "0.25"
  precOk(4) = "2**3**2": precExpect(4) = "512"
  precOk(5) = "2**3*4": precExpect(5) = "32"
  precOk(6) = "2*3**4": precExpect(6) = "162"
  precOk(7) = "3*-2**2": precExpect(7) = "-12"
  precOk(8) = "-2%": precExpect(8) = "-0.02"
  precOk(9) = "2*3%": precExpect(9) = "0.06"
  precOk(10) = "5+2*3%": precExpect(10) = "5.06"
  precOk(11) = "2+3*4": precExpect(11) = "14"
  precOk(12) = "2*3+4": precExpect(12) = "10"
  precOk(13) = "2+3<<1": precExpect(13) = "10"
  precOk(14) = "1|2^3&6<<1": precExpect(14) = "3"
  precOk(15) = "!2==1": precExpect(15) = "0"
  precOk(16) = "not 2==1": precExpect(16) = "1"
  precOk(17) = "1||0&&0": precExpect(17) = "1"
  precOk(18) = "0&&1||1": precExpect(18) = "1"
  precOk(19) = "16**-0.5": precExpect(19) = "0.25"
  precOk(20) = "2(3+4)**2": precExpect(20) = "98"
  precOk(21) = "~2**2": precExpect(21) = "-5"
  precOk(22) = "!!3": precExpect(22) = "1"
  precOk(23) = "5+(2*3)%": precExpect(23) = "5.3"
  precOk(24) = "200+15%": precExpect(24) = "230"
  dim r as Double
  dim rt as String
  dim ia as Boolean
  dim pi as Integer
  for pi = 1 to 24
    if Parser_TryEvaluateEx(precOk(pi), r, rt, ia) = FALSE orelse ResultCloseEnough(rt, precExpect(pi)) = FALSE then
      print "[precedence-doc] FAIL: """ & precOk(pi) & """ -> """ & rt & """ want """ & precExpect(pi) & """ err=" & Parser_GetLastError()
      subFail += 1
    else
      print "[precedence-doc] PASS: """ & precOk(pi) & """ -> """ & rt & """"
      subPass += 1
    end if
  next pi
  g_passed += subPass
  g_failed += subFail
  print "Operator-precedence doc sub-tests: passed " & str(subPass) & ", failed " & str(subFail)
  print ""
end sub

private sub RunNegativeArgumentMagnitudeBandTests()
  print "=== Negative argument magnitude bands (real, 3 ranges) ==="
  dim subPass as Integer = 0
  dim subFail as Integer = 0
  RunNegBandRealBatch subPass, subFail
  g_passed += subPass
  g_failed += subFail
  print "Negative-argument band (real) sub-tests: passed " & str(subPass) & ", failed " & str(subFail)
  print ""
end sub

private sub RunPositiveArgumentMagnitudeBandTests()
  print "=== Positive argument magnitude bands (real, 3 ranges) ==="
  dim subPass as Integer = 0
  dim subFail as Integer = 0
  RunPosBandRealBatch subPass, subFail
  g_passed += subPass
  g_failed += subFail
  print "Positive-argument band (real) sub-tests: passed " & str(subPass) & ", failed " & str(subFail)
  print ""
end sub

' Complex-number tests: keep separate from the main scalar/real suite. This sub begins by enabling
' the parser-wide complex support flag (see USAGE_AND_SYNTAX.md); future complex tests belong here.
private sub RunComplexNumberSupportOptionTests()
  print "=== Complex number support (parser-wide option) ==="

  Parser_SetSupportComplexNumbers(TRUE)

  dim subPass as Integer = 0
  dim subFail as Integer = 0

  if Parser_GetSupportComplexNumbers() = FALSE then
    print "[complex-opt] FAIL: expected support flag ON after enabling"
    subFail += 1
  else
    print "[complex-opt] PASS: getter reports enabled after Parser_SetSupportComplexNumbers(TRUE)"
    subPass += 1
  end if

  dim r as Double
  dim rt as String
  dim ia as Boolean
  if Parser_TryEvaluateEx("1+1", r, rt, ia) = FALSE orelse rt <> "2" then
    print "[complex-opt] FAIL: simple eval with flag ON, got """ & rt & """ err=" & Parser_GetLastError()
    subFail += 1
  else
    print "[complex-opt] PASS: 1+1 -> 2 with support flag ON"
    subPass += 1
  end if

  if Parser_TryEvaluateEx("abs(0x7FFFFFFFFFFFFFFF+20)", r, rt, ia) = FALSE orelse rt <> "9223372036854775827" then
    print "[complex-opt] FAIL: abs exact int with flag ON, got """ & rt & """ err=" & Parser_GetLastError()
    subFail += 1
  else
    print "[complex-opt] PASS: abs(0x7FFFFFFFFFFFFFFF+20) exact int with flag ON"
    subPass += 1
  end if
  if Parser_TryEvaluateEx("hex(abs(0x7FFFFFFFFFFFFFFF+20))", r, rt, ia) = FALSE orelse rt <> "0x8000000000000013" then
    print "[complex-opt] FAIL: hex(abs(...)) with flag ON, got """ & rt & """ err=" & Parser_GetLastError()
    subFail += 1
  else
    print "[complex-opt] PASS: hex(abs(0x7FFFFFFFFFFFFFFF+20)) with flag ON"
    subPass += 1
  end if

  if Parser_TryEvaluateEx("a=range(1,101); a[7+3i]", r, rt, ia) then
    print "[complex-opt] FAIL: a[7+3i] expected integer index error, got """ & rt & """"
    subFail += 1
  elseif instr(lcase(Parser_GetLastError()), "array index must be an integer") > 0 then
    print "[complex-opt] PASS: complex index bound rejected"
    subPass += 1
  else
    print "[complex-opt] FAIL: a[7+3i] expected integer index error, got """ & Parser_GetLastError() & """"
    subFail += 1
  end if

  dim cxMaxLit as String = "0x7FFFFFFFFFFFFFFF+0x7FFFFFFFFFFFFFFFi"
  dim cxMaxExactExpr(1 to 8) as String
  dim cxMaxExactExpect(1 to 8) as String
  cxMaxExactExpr(1) = "real(" & cxMaxLit & ")": cxMaxExactExpect(1) = "9223372036854775807"
  cxMaxExactExpr(2) = "imag(" & cxMaxLit & ")": cxMaxExactExpect(2) = "9223372036854775807"
  cxMaxExactExpr(3) = "cart(" & cxMaxLit & ")": cxMaxExactExpect(3) = "9223372036854775807+9223372036854775807i"
  cxMaxExactExpr(4) = "conj(" & cxMaxLit & ")": cxMaxExactExpect(4) = "9223372036854775807-9223372036854775807i"
  cxMaxExactExpr(5) = "hex(real(" & cxMaxLit & "))": cxMaxExactExpect(5) = "0x7FFFFFFFFFFFFFFF"
  cxMaxExactExpr(6) = "hex(imag(" & cxMaxLit & "))": cxMaxExactExpect(6) = "0x7FFFFFFFFFFFFFFF"
  cxMaxExactExpr(7) = "hex(cart(" & cxMaxLit & "))": cxMaxExactExpect(7) = "0x7FFFFFFFFFFFFFFF+0x7FFFFFFFFFFFFFFFi"
  cxMaxExactExpr(8) = "hex(conj(" & cxMaxLit & "))": cxMaxExactExpect(8) = "0x7FFFFFFFFFFFFFFF-0x7FFFFFFFFFFFFFFFi"
  dim rawConj as RawResult
  if Parser_TryEvaluateExRaw(cxMaxExactExpr(4), rawConj) = FALSE orelse _
     rawConj.scalar.real.kind <> RSK_INT64 orelse rawConj.scalar.imag.kind <> RSK_INT64 then
    print "[complex-opt] FAIL: conj raw export not exact int64"
    subFail += 1
  else
    print "[complex-opt] PASS: conj raw export preserves exact int64 components"
    subPass += 1
  end if
  dim cmxi as Integer
  for cmxi = 1 to 8
    if Parser_TryEvaluateEx(cxMaxExactExpr(cmxi), r, rt, ia) = FALSE orelse rt <> cxMaxExactExpect(cmxi) then
      print "[complex-opt] FAIL: """ & cxMaxExactExpr(cmxi) & """ -> """ & rt & """ err=" & Parser_GetLastError()
      subFail += 1
    else
      print "[complex-opt] PASS: """ & cxMaxExactExpr(cmxi) & """ -> """ & rt & """"
      subPass += 1
    end if
  next cmxi

  dim rawNeg as RawResult
  dim cxNegLit as String = "0x7FFFFFFFFFFFFFFF-0x7FFFFFFFFFFFFFFFi"
  dim cxNegLitExpect as String = "9223372036854775807-9223372036854775807i"
  if Parser_TryEvaluateEx(cxNegLit, r, rt, ia) = FALSE orelse rt <> cxNegLitExpect then
    print "[complex-opt] FAIL: bare """ & cxNegLit & """ -> """ & rt & """ err=" & Parser_GetLastError()
    subFail += 1
  else
    print "[complex-opt] PASS: bare """ & cxNegLit & """"
    subPass += 1
  end if
  if Parser_TryEvaluateExRaw(cxNegLit, rawNeg) = FALSE orelse rawNeg.scalar.kind <> RSK_COMPLEX orelse _
     rawNeg.scalar.real.kind <> RSK_INT64 orelse rawNeg.scalar.imag.kind <> RSK_INT64 then
    print "[complex-opt] FAIL: bare raw """ & cxNegLit & """"
    subFail += 1
  else
    print "[complex-opt] PASS: bare raw """ & cxNegLit & """ int64 complex"
    subPass += 1
  end if

  dim cxNegLit2 as String = "-0x7FFFFFFFFFFFFFFF-0x7FFFFFFFFFFFFFFFi"
  dim cxNegLit2Expect as String = "-9223372036854775807-9223372036854775807i"
  if Parser_TryEvaluateEx(cxNegLit2, r, rt, ia) = FALSE orelse rt <> cxNegLit2Expect then
    print "[complex-opt] FAIL: bare """ & cxNegLit2 & """ -> """ & rt & """ err=" & Parser_GetLastError()
    subFail += 1
  else
    print "[complex-opt] PASS: bare """ & cxNegLit2 & """"
    subPass += 1
  end if
  if Parser_TryEvaluateExRaw(cxNegLit2, rawNeg) = FALSE orelse rawNeg.scalar.kind <> RSK_COMPLEX orelse _
     rawNeg.scalar.real.kind <> RSK_INT64 orelse rawNeg.scalar.imag.kind <> RSK_INT64 then
    print "[complex-opt] FAIL: bare raw """ & cxNegLit2 & """"
    subFail += 1
  else
    print "[complex-opt] PASS: bare raw """ & cxNegLit2 & """ int64 complex"
    subPass += 1
  end if

  dim cxNegExpr(1 to 8) as String
  dim cxNegExpect(1 to 8) as String
  cxNegExpr(1) = "real(" & cxNegLit & ")": cxNegExpect(1) = "9223372036854775807"
  cxNegExpr(2) = "imag(" & cxNegLit & ")": cxNegExpect(2) = "-9223372036854775807"
  cxNegExpr(3) = "cart(" & cxNegLit & ")": cxNegExpect(3) = cxNegLitExpect
  cxNegExpr(4) = "conj(" & cxNegLit & ")": cxNegExpect(4) = "9223372036854775807+9223372036854775807i"
  cxNegExpr(5) = "real(" & cxNegLit2 & ")": cxNegExpect(5) = "-9223372036854775807"
  cxNegExpr(6) = "imag(" & cxNegLit2 & ")": cxNegExpect(6) = "-9223372036854775807"
  cxNegExpr(7) = "cart(" & cxNegLit2 & ")": cxNegExpect(7) = cxNegLit2Expect
  cxNegExpr(8) = "conj(" & cxNegLit2 & ")": cxNegExpect(8) = "-9223372036854775807+9223372036854775807i"
  dim cni as Integer
  for cni = 1 to 8
    if Parser_TryEvaluateEx(cxNegExpr(cni), r, rt, ia) = FALSE orelse rt <> cxNegExpect(cni) then
      print "[complex-opt] FAIL: """ & cxNegExpr(cni) & """ -> """ & rt & """ err=" & Parser_GetLastError()
      subFail += 1
    else
      print "[complex-opt] PASS: """ & cxNegExpr(cni) & """ -> """ & rt & """"
      subPass += 1
    end if
    if Parser_TryEvaluateExRaw(cxNegExpr(cni), rawNeg) = FALSE then
      print "[complex-opt] FAIL: raw """ & cxNegExpr(cni) & """"
      subFail += 1
    elseif cni = 1 orelse cni = 2 orelse cni = 5 orelse cni = 6 then
      if rawNeg.scalar.kind <> RSK_INT64 then
        print "[complex-opt] FAIL: raw scalar """ & cxNegExpr(cni) & """ kind=" & rawNeg.scalar.kind
        subFail += 1
      end if
    elseif rawNeg.scalar.kind <> RSK_COMPLEX orelse rawNeg.scalar.real.kind <> RSK_INT64 orelse rawNeg.scalar.imag.kind <> RSK_INT64 then
      print "[complex-opt] FAIL: raw complex """ & cxNegExpr(cni) & """"
      subFail += 1
    end if
  next cni

  dim complexCases(1 to 7) as String
  dim complexExpect(1 to 7) as String
  complexCases(1) = "10+5i": complexExpect(1) = "10+5i"
  complexCases(2) = "-1+3i": complexExpect(2) = "-1+3i"
  complexCases(3) = "2-3*i": complexExpect(3) = "2-3i"
  complexCases(4) = "-i+5": complexExpect(4) = "5-i"
  complexCases(5) = "(1+2i)*(3+4i)": complexExpect(5) = "-5+10i"
  complexCases(6) = "10+5i-10-5i": complexExpect(6) = "0"
  complexCases(7) = "(1+2i)/i": complexExpect(7) = "2-i"

  dim cxSuffixOk(1 to 25) as String
  dim cxSuffixExpect(1 to 25) as String
  cxSuffixOk(1) = "cart(3i)": cxSuffixExpect(1) = "3i"
  cxSuffixOk(2) = "cart(0x10i)": cxSuffixExpect(2) = "16i"
  cxSuffixOk(3) = "cart(0b101i)": cxSuffixExpect(3) = "5i"
  cxSuffixOk(4) = "cart(0o7i)": cxSuffixExpect(4) = "7i"
  cxSuffixOk(5) = "cart(-3i)": cxSuffixExpect(5) = "-3i"
  cxSuffixOk(6) = "cart(2i)": cxSuffixExpect(6) = "2i"
  cxSuffixOk(7) = "cart(2*i)": cxSuffixExpect(7) = "2i"
  cxSuffixOk(8) = "cart((2**3)i)": cxSuffixExpect(8) = "8i"
  cxSuffixOk(9) = "cart((2*3)i)": cxSuffixExpect(9) = "6i"
  cxSuffixOk(10) = "cart((2/3)i)": cxSuffixExpect(10) = "0.6666666666666666i"
  cxSuffixOk(11) = "cart((2%3)i)": cxSuffixExpect(11) = "2i"
  cxSuffixOk(12) = "cart((1+2)i)": cxSuffixExpect(12) = "3i"
  cxSuffixOk(13) = "cart((1+2)*i)": cxSuffixExpect(13) = "3i"
  cxSuffixOk(14) = "cart(2*3i)": cxSuffixExpect(14) = "6i"
  cxSuffixOk(15) = "cart(2/3i)": cxSuffixExpect(15) = "-0.6666666666666666i"
  cxSuffixOk(16) = "cart(1/2i)": cxSuffixExpect(16) = "-0.5i"
  cxSuffixOk(17) = "cart(1/2*i)": cxSuffixExpect(17) = "0.5i"
  cxSuffixOk(18) = "cart(1/(1+2)i)": cxSuffixExpect(18) = "-0.3333333333333333i"
  cxSuffixOk(19) = "cart(1/(1+2)*i)": cxSuffixExpect(19) = "0.3333333333333333i"
  cxSuffixOk(20) = "cart(1/2/3i)": cxSuffixExpect(20) = "-0.1666666666666667i"
  cxSuffixOk(21) = "ratio(1/2i)": cxSuffixExpect(21) = "-1/2*i"
  cxSuffixOk(22) = "cart(-2**57/3**30+2**55/7**20*i)": cxSuffixExpect(22) = "-699.9582090286702+0.4515324440664775i"
  cxSuffixOk(23) = "cart(2**3i)": cxSuffixExpect(23) = "-0.4869944179657813+0.8734050817748715i"
  cxSuffixOk(24) = "cart(7**20i)": cxSuffixExpect(24) = "0.3444991159598805+0.9387866419495224i"
  cxSuffixOk(25) = "cart(7**2i)": cxSuffixExpect(25) = "-0.7315336785874539-0.681805307321898i"

  dim cxSuffixErr(1 to 2) as String
  dim cxSuffixErrSub(1 to 2) as String
  cxSuffixErr(1) = "3 i"
  cxSuffixErrSub(1) = "unexpected"
  cxSuffixErr(2) = "2%3i"
  cxSuffixErrSub(2) = "incompatible operands"

  dim ci as Integer
  for ci = 1 to 7
    if Parser_TryEvaluateEx(complexCases(ci), r, rt, ia) = FALSE orelse rt <> complexExpect(ci) then
      print "[complex-opt] FAIL: """ & complexCases(ci) & """ -> """ & rt & """ err=" & Parser_GetLastError()
      subFail += 1
    else
      print "[complex-opt] PASS: """ & complexCases(ci) & """ -> """ & rt & """"
      subPass += 1
    end if
  next ci

  dim si as Integer
  for si = 1 to 25
    if Parser_TryEvaluateEx(cxSuffixOk(si), r, rt, ia) = FALSE orelse ResultCloseEnough(rt, cxSuffixExpect(si)) = FALSE then
      print "[complex-opt] FAIL: """ & cxSuffixOk(si) & """ -> """ & rt & """ err=" & Parser_GetLastError()
      print "[complex-opt]      want: """ & cxSuffixExpect(si) & """"
      subFail += 1
    else
      print "[complex-opt] PASS: """ & cxSuffixOk(si) & """ -> """ & rt & """"
      subPass += 1
    end if
  next si

  dim se as Integer
  for se = 1 to 2
    if Parser_TryEvaluateEx(cxSuffixErr(se), r, rt, ia) then
      print "[complex-opt] FAIL: """ & cxSuffixErr(se) & """ expected error, got """ & rt & """"
      subFail += 1
    elseif instr(lcase(Parser_GetLastError()), lcase(cxSuffixErrSub(se))) = 0 then
      print "[complex-opt] FAIL: """ & cxSuffixErr(se) & """ err=" & Parser_GetLastError()
      subFail += 1
    else
      print "[complex-opt] PASS: """ & cxSuffixErr(se) & """ -> error as expected"
      subPass += 1
    end if
  next se

  ' Mixed real/complex arrays: element-wise broadcast (+ - * /) for array+scalar, scalar+array, array+array.
  dim arrCases(1 to 12) as String
  dim arrExpect(1 to 12) as String
  arrCases(1) = "(1,1+2i)+10": arrExpect(1) = "(11, 11+2i)"
  arrCases(2) = "10+(1,1+2*i)": arrExpect(2) = "(11, 11+2i)"
  arrCases(3) = "(5,1+2i)-(1,i)": arrExpect(3) = "(4, 1+i)"
  arrCases(4) = "(1,2)-i": arrExpect(4) = "(1-i, 2-i)"
  arrCases(5) = "(2,1+2i)*3": arrExpect(5) = "(6, 3+6i)"
  arrCases(6) = "3*(2,1+2*i)": arrExpect(6) = "(6, 3+6i)"
  arrCases(7) = "(3,4)*(1,i)": arrExpect(7) = "(3, 4i)"
  arrCases(8) = "(1,2)*(1+i,0)": arrExpect(8) = "(1+i, 0)"
  arrCases(9) = "(8,6+8i)/2": arrExpect(9) = "(4, 3+4i)"
  arrCases(10) = "(8,6)/(2,3)": arrExpect(10) = "(4, 2)"
  arrCases(11) = "(4+2i,6)/(2,3)": arrExpect(11) = "(2+i, 2)"
  arrCases(12) = "i+(1,2)": arrExpect(12) = "(1+i, 2+i)"

  dim ai as Integer
  for ai = 1 to 12
    if Parser_TryEvaluateEx(arrCases(ai), r, rt, ia) = FALSE orelse rt <> arrExpect(ai) then
      print "[complex-opt] FAIL: """ & arrCases(ai) & """ -> """ & rt & """ err=" & Parser_GetLastError()
      subFail += 1
    else
      print "[complex-opt] PASS: """ & arrCases(ai) & """ -> """ & rt & """"
      subPass += 1
    end if
  next ai

  dim powCases(1 to 17) as String
  dim powExpect(1 to 17) as String
  powCases(1) = "(1+i)**2": powExpect(1) = "2i"
  powCases(2) = "(3+4i)**0.5": powExpect(2) = "2+i"
  powCases(3) = "pow(1+i,2)": powExpect(3) = "2i"
  powCases(4) = "(1+i,2)**(2,2)": powExpect(4) = "(2i, 4)"
  powCases(5) = "sqr(1+2i)": powExpect(5) = "-3+4i"
  powCases(6) = "sqrt(3+4i)": powExpect(6) = "2+i"
  powCases(7) = "hypot(3*i,4)": powExpect(7) = "5"
  powCases(8) = "hypot(1+i,1-i)": powExpect(8) = "2"
  powCases(9) = "sqrt(-4)": powExpect(9) = "2i"
  powCases(10) = "(-2)**(1/2)": powExpect(10) = "1.414213562373095i"
  powCases(11) = "(-7)**(3/2)": powExpect(11) = "-18.52025917745213i"
  powCases(12) = "(-5)**(1/3)": powExpect(12) = "-1.709975946676697" ' odd real root, not principal complex
  powCases(13) = "64**(1/3)": powExpect(13) = "4"
  powCases(14) = "pow(-27,1/3)": powExpect(14) = "-3"
  powCases(15) = "pow(2**64,1/2)": powExpect(15) = "4294967296"
  powCases(16) = "(-8)**(1/3)": powExpect(16) = "-2"
  powCases(17) = "pow(-1-0i,1/2)": powExpect(17) = "i"

  dim sqrtCases(1 to 3) as String
  dim sqrtExpect(1 to 3) as String
  sqrtCases(1) = "sqrt(81)": sqrtExpect(1) = "9"
  sqrtCases(2) = "sqrt(3+4i)": sqrtExpect(2) = "2+i"
  sqrtCases(3) = "pow(2**64,1/2)": sqrtExpect(3) = "4294967296"
  dim swi as Integer
  for swi = 1 to 3
    if Parser_TryEvaluateEx(sqrtCases(swi), r, rt, ia) = FALSE orelse rt <> sqrtExpect(swi) then
      print "[complex-opt] FAIL: """ & sqrtCases(swi) & """ -> """ & rt & """ err=" & Parser_GetLastError()
      subFail += 1
    else
      print "[complex-opt] PASS: """ & sqrtCases(swi) & """ -> """ & rt & """"
      subPass += 1
    end if
  next swi

  dim pwi as Integer
  for pwi = 1 to 17
    if Parser_TryEvaluateEx(powCases(pwi), r, rt, ia) = FALSE orelse rt <> powExpect(pwi) then
      print "[complex-opt] FAIL: """ & powCases(pwi) & """ -> """ & rt & """ err=" & Parser_GetLastError()
      subFail += 1
    else
      print "[complex-opt] PASS: """ & powCases(pwi) & """ -> """ & rt & """"
      subPass += 1
    end if
  next pwi

  dim cxExpLnCases(1 to 11) as String
  dim cxExpLnExpect(1 to 11) as String
  cxExpLnCases(1) = "ln(-1)": cxExpLnExpect(1) = "3.141592653589793i"
  cxExpLnCases(2) = "ln(-2)": cxExpLnExpect(2) = "0.6931471805599453+3.141592653589793i"
  cxExpLnCases(3) = "log10(-10)": cxExpLnExpect(3) = "0.9999999999999999+1.364376353841841i"
  cxExpLnCases(4) = "exp(i)": cxExpLnExpect(4) = "0.5403023058681398+0.8414709848078965i"
  cxExpLnCases(5) = "exp(pi*i)": cxExpLnExpect(5) = "-1"
  cxExpLnCases(6) = "ln((-1, -4))": cxExpLnExpect(6) = "(3.141592653589793i, 1.386294361119891+3.141592653589793i)"
  cxExpLnCases(7) = "exp((0, pi*i))": cxExpLnExpect(7) = "(1, -1)"
  cxExpLnCases(8) = "log10((-10, -100))": cxExpLnExpect(8) = "(0.9999999999999999+1.364376353841841i, 2+1.364376353841841i)"
  cxExpLnCases(9) = "ln(2)": cxExpLnExpect(9) = "0.6931471805599453"
  cxExpLnCases(10) = "ln(0)": cxExpLnExpect(10) = "-inf"
  cxExpLnCases(11) = "ln(-1-0i)": cxExpLnExpect(11) = "3.141592653589793i"

  dim exi as Integer
  for exi = 1 to 11
    if Parser_TryEvaluateEx(cxExpLnCases(exi), r, rt, ia) = FALSE orelse ResultCloseEnough(rt, cxExpLnExpect(exi)) = FALSE then
      print "[complex-opt] FAIL: """ & cxExpLnCases(exi) & """ -> """ & rt & """ err=" & Parser_GetLastError()
      subFail += 1
    else
      print "[complex-opt] PASS: """ & cxExpLnCases(exi) & """ -> """ & rt & """"
      subPass += 1
    end if
  next exi

  dim cxTrigOk(1 to 12) as String
  dim cxTrigExpect(1 to 12) as String
  cxTrigOk(1) = "sin(i)": cxTrigExpect(1) = "1.175201193643801i"
  cxTrigOk(2) = "cos(i)": cxTrigExpect(2) = "1.543080634815244"
  cxTrigOk(3) = "sin(1+i)": cxTrigExpect(3) = "1.298457581415977+0.6349639147847361i"
  cxTrigOk(4) = "tan(i)": cxTrigExpect(4) = "0.7615941559557649i"
  cxTrigOk(5) = "sinh(i)": cxTrigExpect(5) = "0.8414709848078965i"
  cxTrigOk(6) = "cosh(1+i)": cxTrigExpect(6) = "0.833730025131149+0.9888977057628651i"
  cxTrigOk(7) = "asinh(1+i)": cxTrigExpect(7) = "1.061275061905036+0.6662394324925153i"
  cxTrigOk(8) = "atan(1+i)": cxTrigExpect(8) = "1.017221967897851+0.4023594781085251i"
  cxTrigOk(9) = "sin((1+i,2))": cxTrigExpect(9) = "(1.298457581415977+0.6349639147847361i, 0.9092974268256817)"
  cxTrigOk(10) = "atanh(i)": cxTrigExpect(10) = "0.7853981633974483i"
  cxTrigOk(11) = "acos(1)": cxTrigExpect(11) = "0"
  cxTrigOk(12) = "asin(i)": cxTrigExpect(12) = "0.8813735870195428i"

  dim tri as Integer
  for tri = 1 to 12
    if Parser_TryEvaluateEx(cxTrigOk(tri), r, rt, ia) = FALSE orelse ResultCloseEnough(rt, cxTrigExpect(tri)) = FALSE then
      print "[complex-opt] FAIL: """ & cxTrigOk(tri) & """ -> """ & rt & """ err=" & Parser_GetLastError()
      subFail += 1
    else
      print "[complex-opt] PASS: """ & cxTrigOk(tri) & """ -> """ & rt & """"
      subPass += 1
    end if
  next tri

  if Parser_TryEvaluateEx("atan2(1+2i,1)", r, rt, ia) then
    print "[complex-opt] FAIL: atan2(1+2i,1) expected error, got """ & rt & """"
    subFail += 1
  else
    dim errA2 as String = lcase(Parser_GetLastError())
    if instr(errA2, "incompatible operands") > 0 then
      print "[complex-opt] PASS: atan2(1+2i,1) -> incompatible operands"
      subPass += 1
    else
      print "[complex-opt] FAIL: atan2(1+2i,1) got """ & Parser_GetLastError() & """"
      subFail += 1
    end if
  end if

  if Parser_TryEvaluateEx("deg(1+2i)", r, rt, ia) then
    print "[complex-opt] FAIL: deg(1+2i) expected error, got """ & rt & """"
    subFail += 1
  else
    dim errDeg as String = lcase(Parser_GetLastError())
    if instr(errDeg, "incompatible operands") > 0 then
      print "[complex-opt] PASS: deg(1+2i) -> incompatible operands"
      subPass += 1
    else
      print "[complex-opt] FAIL: deg(1+2i) got """ & Parser_GetLastError() & """"
      subFail += 1
    end if
  end if

  if Parser_TryEvaluateEx("rad(1+i)", r, rt, ia) then
    print "[complex-opt] FAIL: rad(1+i) expected error, got """ & rt & """"
    subFail += 1
  else
    dim errRad as String = lcase(Parser_GetLastError())
    if instr(errRad, "incompatible operands") > 0 then
      print "[complex-opt] PASS: rad(1+i) -> incompatible operands"
      subPass += 1
    else
      print "[complex-opt] FAIL: rad(1+i) got """ & Parser_GetLastError() & """"
      subFail += 1
    end if
  end if

  dim cxLogBaseCases(1 to 5) as String
  dim cxLogBaseExpect(1 to 5) as String
  cxLogBaseCases(1) = "log(8,2)": cxLogBaseExpect(1) = "3"
  cxLogBaseCases(2) = "log(-9,7)": cxLogBaseExpect(2) = "1.129150068107159+1.614459257080781i"
  cxLogBaseCases(3) = "log((-9, 16), (7, 2))": cxLogBaseExpect(3) = "(1.129150068107159+1.614459257080781i, 4)"
  cxLogBaseCases(4) = "log(1+i, 2)": cxLogBaseExpect(4) = "0.5000000000000001+1.133090035456799i"
  cxLogBaseCases(5) = "log((8, 16), (2, 4))": cxLogBaseExpect(5) = "(3, 2)"

  dim lgbi as Integer
  for lgbi = 1 to 5
    if Parser_TryEvaluateEx(cxLogBaseCases(lgbi), r, rt, ia) = FALSE orelse rt <> cxLogBaseExpect(lgbi) then
      print "[complex-opt] FAIL: """ & cxLogBaseCases(lgbi) & """ -> """ & rt & """ err=" & Parser_GetLastError()
      subFail += 1
    else
      print "[complex-opt] PASS: """ & cxLogBaseCases(lgbi) & """ -> """ & rt & """"
      subPass += 1
    end if
  next lgbi

  ' Comparisons: `=`/`==` and `<>`/`!=` are defined for complex scalars and arrays (including mixed real/complex);
  ' ordering comparisons (`<` `<=` `>` `>=`) are errors when any operand has a nonzero imaginary part;
  ' complex mixed with time in `=`/`==`/`<>`/`!=` is an error.
  dim cmpCases(1 to 10) as String
  dim cmpExpect(1 to 10) as String
  cmpCases(1) = "1+2i = 1+2i": cmpExpect(1) = "1"
  cmpCases(2) = "1+2i == 3+4i": cmpExpect(2) = "0"
  cmpCases(3) = "3 <> 1+2i": cmpExpect(3) = "1"
  cmpCases(4) = "1+0*i = 1": cmpExpect(4) = "1"
  cmpCases(5) = "(1+1i, 2) = (1+1i, 2)": cmpExpect(5) = "1"
  cmpCases(6) = "(1+1i, 2) <> (1+1i, 3)": cmpExpect(6) = "1"
  cmpCases(7) = "(1+1i, 2) = (1, 2)": cmpExpect(7) = "0"
  cmpCases(8) = "(1, 2+1i) != (1, 2+1i)": cmpExpect(8) = "0"
  cmpCases(9) = "2 = 1+1i": cmpExpect(9) = "0"
  cmpCases(10) = "(2, 0) = (1+1i, 0)": cmpExpect(10) = "0"

  dim cmi as Integer
  for cmi = 1 to 10
    if Parser_TryEvaluateEx(cmpCases(cmi), r, rt, ia) = FALSE orelse rt <> cmpExpect(cmi) then
      print "[complex-opt] FAIL: """ & cmpCases(cmi) & """ -> """ & rt & """ err=" & Parser_GetLastError()
      subFail += 1
    else
      print "[complex-opt] PASS: """ & cmpCases(cmi) & """ -> """ & rt & """"
      subPass += 1
    end if
  next cmi

  dim cmpErrExpr(1 to 8) as String
  cmpErrExpr(1) = "1+2i > 0"
  cmpErrExpr(2) = "1+2i >= 1+2i"
  cmpErrExpr(3) = "(1, 2+1i) < (1, 3)"
  cmpErrExpr(4) = "1+2i = 1s"
  cmpErrExpr(5) = "1+2i <> 0:01"
  cmpErrExpr(6) = "1+2i + 1s"
  cmpErrExpr(7) = "1+2i - 1s"
  cmpErrExpr(8) = "1+2i * 1s"

  dim cei as Integer
  for cei = 1 to 8
    if Parser_TryEvaluateEx(cmpErrExpr(cei), r, rt, ia) then
      print "[complex-opt] FAIL: expected error for """ & cmpErrExpr(cei) & """ but got """ & rt & """"
      subFail += 1
    else
      dim errCmp as String = lcase(Parser_GetLastError())
      if instr(errCmp, "incompatible operands") > 0 then
        print "[complex-opt] PASS: """ & cmpErrExpr(cei) & """ -> incompatible operands"
        subPass += 1
      else
        print "[complex-opt] FAIL: """ & cmpErrExpr(cei) & """ expected incompatible operands, got """ & Parser_GetLastError() & """"
        subFail += 1
      end if
    end if
  next cei

  ' `~`, `%` / `mod`, bitwise ops: error on non-zero imaginary (scalar or array element); `!` / `not` truthiness on both Cartesian parts.
  dim cxNotTruthCases(1 to 5) as String
  dim cxNotTruthExpect(1 to 5) as String
  cxNotTruthCases(1) = "!(5+5i)": cxNotTruthExpect(1) = "0"
  cxNotTruthCases(2) = "!(-3*i)": cxNotTruthExpect(2) = "0"
  cxNotTruthCases(3) = "!(0+0*i)": cxNotTruthExpect(3) = "1"
  cxNotTruthCases(4) = "not (1+1i)": cxNotTruthExpect(4) = "0"
  cxNotTruthCases(5) = "not (0)": cxNotTruthExpect(5) = "1"

  dim li as Integer
  for li = 1 to 5
    if Parser_TryEvaluateEx(cxNotTruthCases(li), r, rt, ia) = FALSE orelse rt <> cxNotTruthExpect(li) then
      print "[complex-opt] FAIL: """ & cxNotTruthCases(li) & """ -> """ & rt & """ err=" & Parser_GetLastError()
      subFail += 1
    else
      print "[complex-opt] PASS: """ & cxNotTruthCases(li) & """ -> """ & rt & """"
      subPass += 1
    end if
  next li

  dim cxIntErr(1 to 9) as String
  cxIntErr(1) = "~(1+2i)"
  cxIntErr(2) = "~(1, 2+1i)"
  cxIntErr(3) = "(1+2i) % 2"
  cxIntErr(4) = "mod(1+2i, 3)"
  cxIntErr(5) = "(1+2i) & 1"
  cxIntErr(6) = "(1+2i) | 1"
  cxIntErr(7) = "(1+2i) ^ 1"
  cxIntErr(8) = "(1+2i) << 1"
  cxIntErr(9) = "(1+2i) >> 1"

  dim xi as Integer
  for xi = 1 to 9
    if Parser_TryEvaluateEx(cxIntErr(xi), r, rt, ia) then
      print "[complex-opt] FAIL: expected error for """ & cxIntErr(xi) & """ but got """ & rt & """"
      subFail += 1
    else
      dim errXi as String = lcase(Parser_GetLastError())
      if instr(errXi, "incompatible operands") > 0 orelse instr(errXi, "modulo operands must be integer values") > 0 orelse instr(errXi, "bitwise operands must be integer values") > 0 then
        print "[complex-opt] PASS: """ & cxIntErr(xi) & """ -> error"
        subPass += 1
      else
        print "[complex-opt] FAIL: """ & cxIntErr(xi) & """ expected int/mod/bit error, got """ & Parser_GetLastError() & """"
        subFail += 1
      end if
    end if
  next xi

  ' Aggregation/array utilities with complex operands (allowed vs rejected builtins).
  dim cxAggOk(1 to 36) as String
  dim cxAggExpect(1 to 36) as String
  cxAggOk(1) = "sum(1+2i, 3)": cxAggExpect(1) = "4+2i"
  cxAggOk(2) = "sum((1+2i, 3+4i))": cxAggExpect(2) = "4+6i"
  cxAggOk(3) = "prod(1+i, 2)": cxAggExpect(3) = "2+2i"
  cxAggOk(4) = "product((2, 1+i))": cxAggExpect(4) = "2+2i"
  cxAggOk(5) = "avg(2+2i, 4+4i)": cxAggExpect(5) = "3+3i"
  cxAggOk(6) = "mean((1+2i, 3))": cxAggExpect(6) = "2+i"
  cxAggOk(7) = "reverse((1+2i, 3))": cxAggExpect(7) = "(3, 1+2i)"
  cxAggOk(8) = "unique((1+2i, 1+2i, 3))": cxAggExpect(8) = "(1+2i, 3)"
  cxAggOk(9) = "unique((1+2i, 1+3i))": cxAggExpect(9) = "(1+2i, 1+3i)"
  cxAggOk(10) = "unpack((1+2i, 3))": cxAggExpect(10) = "(1+2i, 3)"
  cxAggOk(11) = "sum(unpack((1+2i, 2+2i)))": cxAggExpect(11) = "3+4i"
  cxAggOk(12) = "f(x,y)=x*y; f(unpack((1+i, 2)))": cxAggExpect(12) = "2+2i"
  cxAggOk(13) = "sum(1i, 2i)": cxAggExpect(13) = "3i"
  cxAggOk(14) = "prod(2+i, 2-i)": cxAggExpect(14) = "5"
  cxAggOk(15) = "product(1+i, 1+i)": cxAggExpect(15) = "2i"
  cxAggOk(16) = "mean(1+2i, 3+4i, 5+6i)": cxAggExpect(16) = "3+4i"
  cxAggOk(17) = "avg(0, 2i)": cxAggExpect(17) = "i"
  cxAggOk(18) = "sum(2+i, -2+i)": cxAggExpect(18) = "2i"
  cxAggOk(19) = "sum(1, Inf+2i, 3)": cxAggExpect(19) = "inf+2i"
  cxAggOk(20) = "prod(1, Inf+2i, 3)": cxAggExpect(20) = "inf+6i"
  cxAggOk(21) = "avg(1, Inf+2i, 3)": cxAggExpect(21) = "inf+0.6666666666666666i"
  cxAggOk(22) = "reverse(1, Inf+2i, 3)": cxAggExpect(22) = "(3, inf+2i, 1)"
  cxAggOk(23) = "unique(1, Inf+2i, 3)": cxAggExpect(23) = "(1, inf+2i, 3)"
  cxAggOk(24) = "unpack(Inf+2i, 3)": cxAggExpect(24) = "(inf+2i, 3)"
  cxAggOk(25) = "sum(1, 2+Inf*i, 3)": cxAggExpect(25) = "6+inf*i"
  cxAggOk(26) = "prod(1, 2+Inf*i, 3)": cxAggExpect(26) = "6+inf*i"
  cxAggOk(27) = "avg(1, 2+Inf*i, 3)": cxAggExpect(27) = "2+inf*i"
  cxAggOk(28) = "reverse(1, 2+Inf*i, 3)": cxAggExpect(28) = "(3, 2+inf*i, 1)"
  cxAggOk(29) = "unique(1, 2+Inf*i, 3)": cxAggExpect(29) = "(1, 2+inf*i, 3)"
  cxAggOk(30) = "unpack(2+Inf*i, 3)": cxAggExpect(30) = "(2+inf*i, 3)"
  cxAggOk(31) = "prod(5, Inf+2i, 3)": cxAggExpect(31) = "inf+30i"
  cxAggOk(32) = "prod(5, 2+Inf*i, 3)": cxAggExpect(32) = "30+inf*i"
  cxAggOk(33) = "repeat((1, Inf, i), 2)": cxAggExpect(33) = "(1, inf, i, 1, inf, i)"
  cxAggOk(34) = "flat((1+2i, 3), 4)": cxAggExpect(34) = "(1+2i, 3, 4)"
  cxAggOk(35) = "flat(1, Inf+2i, 3)": cxAggExpect(35) = "(1, inf+2i, 3)"
  cxAggOk(36) = "flat(1, 2+Inf*i, 3)": cxAggExpect(36) = "(1, 2+inf*i, 3)"

  dim agi as Integer
  for agi = 1 to 36
    if Parser_TryEvaluateEx(cxAggOk(agi), r, rt, ia) = FALSE orelse rt <> cxAggExpect(agi) then
      print "[complex-opt] FAIL: """ & cxAggOk(agi) & """ -> """ & rt & """ err=" & Parser_GetLastError()
      subFail += 1
    else
      print "[complex-opt] PASS: """ & cxAggOk(agi) & """ -> """ & rt & """"
      subPass += 1
    end if
  next agi

  if Parser_TryEvaluateEx("repeat((1,2), i)", r, rt, ia) then
    print "[complex-opt] FAIL: expected error for ""repeat((1,2), i)"" but got """ & rt & """"
    subFail += 1
  elseif instr(lcase(Parser_GetLastError()), "repeat() expects integer values") > 0 then
    print "[complex-opt] PASS: ""repeat((1,2), i)"" -> repeat() expects integer values"
    subPass += 1
  else
    print "[complex-opt] FAIL: ""repeat((1,2), i)"" expected repeat() expects integer values, got """ & Parser_GetLastError() & """"
    subFail += 1
  end if

  if Parser_TryEvaluateEx("range(1, 5, i)", r, rt, ia) then
    print "[complex-opt] FAIL: expected error for ""range(1, 5, i)"" but got """ & rt & """"
    subFail += 1
  elseif instr(lcase(Parser_GetLastError()), "range() expects integer values") > 0 then
    print "[complex-opt] PASS: ""range(1, 5, i)"" -> range() expects integer values"
    subPass += 1
  else
    print "[complex-opt] FAIL: ""range(1, 5, i)"" expected range() expects integer values, got """ & Parser_GetLastError() & """"
    subFail += 1
  end if

  dim cxAggErr(1 to 12) as String
  cxAggErr(1) = "min(1+2i, 3)"
  cxAggErr(2) = "max((1+2i, 2))"
  cxAggErr(3) = "sort(3+4i, 1+2i)"
  cxAggErr(4) = "median(1+2i, 5)"
  cxAggErr(5) = "variance(1+2i, 2)"
  cxAggErr(6) = "stddev((1+2i, 2))"
  cxAggErr(7) = "min(1, Inf+2i, 3)"
  cxAggErr(8) = "max(1, 2+Inf*i, 3)"
  cxAggErr(9) = "sort(Inf+2i, 1+2i)"
  cxAggErr(10) = "median(Inf+2i, 1)"
  cxAggErr(11) = "variance(2+Inf*i, 1)"
  cxAggErr(12) = "stddev((Inf+2i, 1))"

  dim aei as Integer
  for aei = 1 to 12
    if Parser_TryEvaluateEx(cxAggErr(aei), r, rt, ia) then
      print "[complex-opt] FAIL: expected error for """ & cxAggErr(aei) & """ but got """ & rt & """"
      subFail += 1
    else
      dim errAgg as String = lcase(Parser_GetLastError())
      if instr(errAgg, "incompatible operands") > 0 then
        print "[complex-opt] PASS: """ & cxAggErr(aei) & """ -> incompatible operands"
        subPass += 1
      else
        print "[complex-opt] FAIL: """ & cxAggErr(aei) & """ expected incompatible operands, got """ & Parser_GetLastError() & """"
        subFail += 1
      end if
    end if
  next aei

  dim cxUniOk(1 to 18) as String
  dim cxUniExpect(1 to 18) as String
  cxUniOk(1) = "int(2.7+3.2i)": cxUniExpect(1) = "2+3i"
  cxUniOk(2) = "frac(2.5+0.5i)": cxUniExpect(2) = "0.5+0.5i"
  cxUniOk(3) = "round(2.4+3.6i)": cxUniExpect(3) = "2+4i"
  cxUniOk(4) = "floor(2.9+3.1i)": cxUniExpect(4) = "2+3i"
  cxUniOk(5) = "ceil(2.1+3.1i)": cxUniExpect(5) = "3+4i"
  cxUniOk(6) = "abs(3+4i)": cxUniExpect(6) = "5"
  cxUniOk(7) = "sign(3+4i)": cxUniExpect(7) = "0.6+0.8i"
  cxUniOk(8) = "real(1+2i)": cxUniExpect(8) = "1"
  cxUniOk(9) = "imag(1+2i)": cxUniExpect(9) = "2"
  cxUniOk(10) = "conj(1+2i)": cxUniExpect(10) = "1-2i"
  cxUniOk(11) = "phase(1)": cxUniExpect(11) = "0"
  cxUniOk(12) = "polar(1+1i)": cxUniExpect(12) = "(1.414213562373095, 0.7853981633974483)"
  cxUniOk(13) = "cart(polar(1+1i))": cxUniExpect(13) = "1+i"
  cxUniOk(14) = "fact(5)": cxUniExpect(14) = "120"
  cxUniOk(15) = "int((1+2.7i, 4.2+5.8i))": cxUniExpect(15) = "(1+2i, 4+5i)"
  cxUniOk(16) = "abs((3+4i, 0))": cxUniExpect(16) = "(5, 0)"
  cxUniOk(17) = "phase(-1-0i)": cxUniExpect(17) = "3.141592653589793"
  cxUniOk(18) = "polar(-1-0i)": cxUniExpect(18) = "(1, 3.141592653589793)"

  dim ui as Integer
  for ui = 1 to 18
    if Parser_TryEvaluateEx(cxUniOk(ui), r, rt, ia) = FALSE orelse ResultCloseEnough(rt, cxUniExpect(ui)) = FALSE then
      print "[complex-opt] FAIL: """ & cxUniOk(ui) & """ -> """ & rt & """ err=" & Parser_GetLastError()
      subFail += 1
    else
      print "[complex-opt] PASS: """ & cxUniOk(ui) & """ -> """ & rt & """"
      subPass += 1
    end if
  next ui

  dim cxPolCartOk(1 to 8) as String
  dim cxPolCartExpect(1 to 8) as String
  cxPolCartOk(1) = "polar(3)": cxPolCartExpect(1) = "(3, 0)"
  cxPolCartOk(2) = "polar(4i)": cxPolCartExpect(2) = "(4, 1.570796326794897)"
  cxPolCartOk(3) = "polar(3+4i)": cxPolCartExpect(3) = "(5, 0.9272952180016122)"
  cxPolCartOk(4) = "polar((3+4i))": cxPolCartExpect(4) = "(5, 0.9272952180016122)"
  cxPolCartOk(5) = "cart(2)": cxPolCartExpect(5) = "2"
  cxPolCartOk(6) = "cart((2))": cxPolCartExpect(6) = "2"
  cxPolCartOk(7) = "cart((2,pi/3))": cxPolCartExpect(7) = "1+1.7320508075688772i"
  cxPolCartOk(8) = "cart(2,pi/3)": cxPolCartExpect(8) = "1+1.7320508075688772i"
  dim pci as Integer
  for pci = 1 to 8
    if Parser_TryEvaluateEx(cxPolCartOk(pci), r, rt, ia) = FALSE orelse ResultCloseEnough(rt, cxPolCartExpect(pci)) = FALSE then
      print "[complex-opt] FAIL: """ & cxPolCartOk(pci) & """ -> """ & rt & """ err=" & Parser_GetLastError()
      print "[complex-opt]      want: """ & cxPolCartExpect(pci) & """"
      subFail += 1
    else
      print "[complex-opt] PASS: """ & cxPolCartOk(pci) & """ -> """ & rt & """"
      subPass += 1
    end if
  next pci

  dim cxPolCartErr(1 to 7) as String
  cxPolCartErr(1) = "polar(3+4i, 2)"
  cxPolCartErr(2) = "polar((3+4i, 2))"
  cxPolCartErr(3) = "cart((2,pi/3,1))"
  cxPolCartErr(4) = "cart(2,pi/3,1)"
  cxPolCartErr(5) = "cart((2,pi/3),(1))"
  cxPolCartErr(6) = "polar(3+4i, 2i)"
  cxPolCartErr(7) = "polar(3+4i, 1+2i)"
  dim pcei as Integer
  for pcei = 1 to 7
    if Parser_TryEvaluateEx(cxPolCartErr(pcei), r, rt, ia) then
      print "[complex-opt] FAIL: expected error for """ & cxPolCartErr(pcei) & """ but got """ & rt & """"
      subFail += 1
    else
      dim errPolCart as String = lcase(Parser_GetLastError())
      if instr(errPolCart, "expects") > 0 orelse instr(errPolCart, "incompatible operands") > 0 then
        print "[complex-opt] PASS: """ & cxPolCartErr(pcei) & """ -> " & Parser_GetLastError()
        subPass += 1
      else
        print "[complex-opt] FAIL: """ & cxPolCartErr(pcei) & """ expected arity/incompatible error, got """ & Parser_GetLastError() & """"
        subFail += 1
      end if
    end if
  next pcei

  '' cart/polar/real/imag/phase/ln/exp and complex sin/cos/tan/sinh/cosh at angles where
  '' CalcSin/Cos/Tan/Atan2 return exact 0, 1, -1, inf, or pi multiples (integer-like floats).
  dim cxTrigAuxOk(1 to 39) as String
  dim cxTrigAuxExpect(1 to 39) as String
  cxTrigAuxOk(1) = "cart(6,0)": cxTrigAuxExpect(1) = "6"
  cxTrigAuxOk(2) = "cart(5,pi/2)": cxTrigAuxExpect(2) = "5i"
  cxTrigAuxOk(3) = "cart(2,pi)": cxTrigAuxExpect(3) = "-2"
  cxTrigAuxOk(4) = "cart(3,3*pi/2)": cxTrigAuxExpect(4) = "-3i"
  cxTrigAuxOk(5) = "real(cart(8,pi/2))": cxTrigAuxExpect(5) = "0"
  cxTrigAuxOk(6) = "imag(cart(8,pi/2))": cxTrigAuxExpect(6) = "8"
  cxTrigAuxOk(7) = "real(cart(7,0))": cxTrigAuxExpect(7) = "7"
  cxTrigAuxOk(8) = "imag(cart(7,0))": cxTrigAuxExpect(8) = "0"
  cxTrigAuxOk(9) = "real(cart(5,pi))": cxTrigAuxExpect(9) = "-5"
  cxTrigAuxOk(10) = "imag(cart(5,3*pi/2))": cxTrigAuxExpect(10) = "-5"
  cxTrigAuxOk(11) = "polar(9)": cxTrigAuxExpect(11) = "(9, 0)"
  cxTrigAuxOk(12) = "polar(9i)": cxTrigAuxExpect(12) = "(9, 1.570796326794897)"
  cxTrigAuxOk(13) = "polar(-9)": cxTrigAuxExpect(13) = "(9, 3.141592653589793)"
  cxTrigAuxOk(14) = "polar(-9i)": cxTrigAuxExpect(14) = "(9, -1.570796326794897)"
  cxTrigAuxOk(15) = "phase(9)": cxTrigAuxExpect(15) = "0"
  cxTrigAuxOk(16) = "phase(9i)": cxTrigAuxExpect(16) = "1.570796326794897"
  cxTrigAuxOk(17) = "phase(-9)": cxTrigAuxExpect(17) = "3.141592653589793"
  cxTrigAuxOk(18) = "phase(-9i)": cxTrigAuxExpect(18) = "-1.570796326794897"
  cxTrigAuxOk(19) = "phase(0)": cxTrigAuxExpect(19) = "0"
  cxTrigAuxOk(20) = "sin(pi/2)": cxTrigAuxExpect(20) = "1"
  cxTrigAuxOk(21) = "cos(pi/2)": cxTrigAuxExpect(21) = "0"
  cxTrigAuxOk(22) = "tan(pi/2)": cxTrigAuxExpect(22) = "inf"
  cxTrigAuxOk(23) = "tan(-3*pi/2)": cxTrigAuxExpect(23) = "-inf"
  cxTrigAuxOk(24) = "sin(pi)": cxTrigAuxExpect(24) = "0"
  cxTrigAuxOk(25) = "tan(pi/4)": cxTrigAuxExpect(25) = "1"
  cxTrigAuxOk(26) = "tan(-pi/4)": cxTrigAuxExpect(26) = "-1"
  cxTrigAuxOk(27) = "real(sin(pi/2))": cxTrigAuxExpect(27) = "1"
  cxTrigAuxOk(28) = "imag(sin(pi/2))": cxTrigAuxExpect(28) = "0"
  cxTrigAuxOk(29) = "sinh(i*pi/2)": cxTrigAuxExpect(29) = "i"
  cxTrigAuxOk(30) = "cosh(i*pi/2)": cxTrigAuxExpect(30) = "0"
  cxTrigAuxOk(31) = "cosh(i*pi)": cxTrigAuxExpect(31) = "-1"
  cxTrigAuxOk(32) = "exp(pi/2*i)": cxTrigAuxExpect(32) = "i"
  cxTrigAuxOk(33) = "exp(2*pi*i)": cxTrigAuxExpect(33) = "1"
  cxTrigAuxOk(34) = "sin(3*pi/2)": cxTrigAuxExpect(34) = "-1"
  cxTrigAuxOk(35) = "real(ln(-1))": cxTrigAuxExpect(35) = "0"
  cxTrigAuxOk(36) = "imag(ln(-1))": cxTrigAuxExpect(36) = "3.141592653589793"
  cxTrigAuxOk(37) = "imag(ln(-i))": cxTrigAuxExpect(37) = "-1.570796326794897"
  cxTrigAuxOk(38) = "cart(polar(6i))": cxTrigAuxExpect(38) = "6i"
  cxTrigAuxOk(39) = "cart(polar(-4))": cxTrigAuxExpect(39) = "-4"
  dim txi as Integer
  for txi = 1 to 39
    if Parser_TryEvaluateEx(cxTrigAuxOk(txi), r, rt, ia) = FALSE orelse ResultCloseEnough(rt, cxTrigAuxExpect(txi)) = FALSE then
      print "[complex-opt] FAIL: """ & cxTrigAuxOk(txi) & """ -> """ & rt & """ err=" & Parser_GetLastError()
      print "[complex-opt]      want: """ & cxTrigAuxExpect(txi) & """"
      subFail += 1
    else
      print "[complex-opt] PASS: """ & cxTrigAuxOk(txi) & """ -> """ & rt & """"
      subPass += 1
    end if
  next txi

  dim cxFmtOk(1 to 7) as String
  dim cxFmtExpect(1 to 7) as String
  cxFmtOk(1) = "hex(10+15i)": cxFmtExpect(1) = "0xA+0xFi"
  cxFmtOk(2) = "bin(1+i)": cxFmtExpect(2) = "0b1+i"
  cxFmtOk(3) = "oct(8i)": cxFmtExpect(3) = "0o10i"
  cxFmtOk(4) = "hex((1+2i,10+11i))": cxFmtExpect(4) = "(0x1+0x2i, 0xA+0xBi)"
  cxFmtOk(5) = "uhex(-1+2i)": cxFmtExpect(5) = "0xFFFFFFFFFFFFFFFF+0x2i"
  cxFmtOk(6) = "uoct(-1+i)": cxFmtExpect(6) = "0o1777777777777777777777+i"
  cxFmtOk(7) = "ubin(5+10i)": cxFmtExpect(7) = "0b101+0b1010i"

  if Parser_TryEvaluateEx("(0xFFFFFFFFFFFFFF + 0x7FFFFFFFFFFFFFi);uhex", r, rt, ia) = FALSE orelse rt <> "0xFFFFFFFFFFFFFF+0x7FFFFFFFFFFFFFi" then
    print "[complex-opt] FAIL: large hex complex uhex -> """ & rt & """ err=" & Parser_GetLastError()
    subFail += 1
  else
    print "[complex-opt] PASS: large hex complex uhex preserves exact integers"
    subPass += 1
  end if

  if Parser_TryEvaluateEx("(0xFFFFFFFFFFFFFF + 0x7FFFFFFFFFFFFFi)", r, rt, ia) = FALSE orelse rt <> "72057594037927935+36028797018963967i" then
    print "[complex-opt] FAIL: large hex complex sum -> """ & rt & """ err=" & Parser_GetLastError()
    subFail += 1
  else
    print "[complex-opt] PASS: large hex complex sum exact"
    subPass += 1
  end if

  if Parser_TryEvaluateEx("sum(0xFFFFFFFFFFFFFF, 0x7FFFFFFFFFFFFFi)", r, rt, ia) = FALSE orelse rt <> "72057594037927935+36028797018963967i" then
    print "[complex-opt] FAIL: sum exact complex scalars -> """ & rt & """ err=" & Parser_GetLastError()
    subFail += 1
  else
    print "[complex-opt] PASS: sum() preserves exact large complex"
    subPass += 1
  end if

  if Parser_TryEvaluateEx("conj(0xFFFFFFFFFFFFFF+0x7FFFFFFFFFFFFFi);uhex", r, rt, ia) = FALSE orelse rt <> "0xFFFFFFFFFFFFFF-0x7FFFFFFFFFFFFFi" then
    print "[complex-opt] FAIL: conj exact uhex -> """ & rt & """ err=" & Parser_GetLastError()
    subFail += 1
  else
    print "[complex-opt] PASS: conj() preserves exact integers"
    subPass += 1
  end if

  if Parser_TryEvaluateEx("inf+i*inf", r, rt, ia) = FALSE orelse rt <> "inf+inf*i" then
    print "[complex-opt] FAIL: inf+i*inf -> """ & rt & """ err=" & Parser_GetLastError()
    subFail += 1
  else
    print "[complex-opt] PASS: inf+i*inf non-finite complex multiply/add"
    subPass += 1
  end if

  if Parser_TryEvaluateEx("inf+i*nan", r, rt, ia) = FALSE orelse rt <> "nan" then
    print "[complex-opt] FAIL: inf+i*nan -> """ & rt & """ err=" & Parser_GetLastError()
    subFail += 1
  else
    print "[complex-opt] PASS: inf+i*nan collapses to scalar NaN"
    subPass += 1
  end if

  if Parser_TryEvaluateEx("(1+2i)/0", r, rt, ia) = FALSE orelse rt <> "nan" then
    print "[complex-opt] FAIL: (1+2i)/0 -> """ & rt & """ err=" & Parser_GetLastError()
    subFail += 1
  else
    print "[complex-opt] PASS: complex divide-by-zero collapses to scalar NaN"
    subPass += 1
  end if

  dim ffi as Integer
  for ffi = 1 to 7
    if Parser_TryEvaluateEx(cxFmtOk(ffi), r, rt, ia) = FALSE orelse rt <> cxFmtExpect(ffi) then
      print "[complex-opt] FAIL: """ & cxFmtOk(ffi) & """ -> """ & rt & """ err=" & Parser_GetLastError()
      print "[complex-opt]      want: """ & cxFmtExpect(ffi) & """"
      subFail += 1
    else
      print "[complex-opt] PASS: """ & cxFmtOk(ffi) & """ -> """ & rt & """"
      subPass += 1
    end if
  next ffi

  if Parser_TryEvaluateEx("hex(1+2.5i)", r, rt, ia) then
    print "[complex-opt] FAIL: hex(1+2.5i) expected error, got """ & rt & """"
    subFail += 1
  else
    dim errHx as String = lcase(Parser_GetLastError())
    if instr(errHx, "hex() expects integer values") > 0 then
      print "[complex-opt] PASS: hex(1+2.5i) -> hex() expects integer values"
      subPass += 1
    else
      print "[complex-opt] FAIL: hex(1+2.5i) expected integer error, got """ & Parser_GetLastError() & """"
      subFail += 1
    end if
  end if

  dim cxUniErr(1 to 5) as String
  cxUniErr(1) = "clamp(1+2i, 0, 5)"
  cxUniErr(2) = "gcd(1+2i, 2)"
  cxUniErr(3) = "mod(1+2i, 2)"
  cxUniErr(4) = "ncr(5+5i, 2)"
  cxUniErr(5) = "lcm(2, 1+2i)"

  dim uei as Integer
  for uei = 1 to 5
    if Parser_TryEvaluateEx(cxUniErr(uei), r, rt, ia) then
      print "[complex-opt] FAIL: expected error for """ & cxUniErr(uei) & """ but got """ & rt & """"
      subFail += 1
    else
      dim errUni as String = lcase(Parser_GetLastError())
      if instr(errUni, "incompatible operands") > 0 then
        print "[complex-opt] PASS: """ & cxUniErr(uei) & """ -> incompatible operands"
        subPass += 1
      else
        print "[complex-opt] FAIL: """ & cxUniErr(uei) & """ expected incompatible operands, got """ & Parser_GetLastError() & """"
        subFail += 1
      end if
    end if
  next uei

  dim cxSortRatioOk(1 to 24) as String
  dim cxSortRatioExpect(1 to 24) as String
  cxSortRatioOk(1) = "sortby((3+4i, 1+2i), abs)": cxSortRatioExpect(1) = "(1+2i, 3+4i)"
  cxSortRatioOk(6) = "sortby((3+4i, 1+2i), polar)": cxSortRatioExpect(6) = "(1+2i, 3+4i)"
  cxSortRatioOk(2) = "ratio(1+2i)": cxSortRatioExpect(2) = "1+2i"
  cxSortRatioOk(3) = "ratio(0.5+0.25i)": cxSortRatioExpect(3) = "1/2+1/4*i"
  cxSortRatioOk(4) = "ratio(2+3i)": cxSortRatioExpect(4) = "2+3i"
  cxSortRatioOk(5) = "ratio(e+10i)": cxSortRatioExpect(5) = "14665106/5394991+10i"
  cxSortRatioOk(7) = "ratio(0.25i)": cxSortRatioExpect(7) = "1/4*i"
  cxSortRatioOk(8) = "ratio(0.5+0.25i)+1": cxSortRatioExpect(8) = "1.5+0.25i"
  cxSortRatioOk(9) = "ratio(-0.5+0.25i)": cxSortRatioExpect(9) = "-1/2+1/4*i"
  cxSortRatioOk(10) = "ratio(3+0.25i)": cxSortRatioExpect(10) = "3+1/4*i"
  cxSortRatioOk(11) = "ratio(-2**57/3**30+2**55/7**20*i)": cxSortRatioExpect(11) = "-113792906/162571+1293463/2864607*i"
  cxSortRatioOk(12) = "ratio(-2**57/3**30+2**55/7**20i)": cxSortRatioExpect(12) = "12411888722130364-33823353366914148i"
  cxSortRatioOk(13) = "ratio((1+2i)/(3+4i))": cxSortRatioExpect(13) = "11/25+2/25*i"
  cxSortRatioOk(14) = "ratio((1.11+2.33i)*(3.37+4.71i))": cxSortRatioExpect(14) = "-4521/625+65401/5000*i"
  cxSortRatioOk(15) = "ratio((1+i)**(2+3i))": cxSortRatioExpect(15) = "-740749/4531935+329494/3432051*i"
  cxSortRatioOk(16) = "ratio(1/(1+i))": cxSortRatioExpect(16) = "1/2-1/2*i"
  cxSortRatioOk(17) = "ratio((1+3i)/2)": cxSortRatioExpect(17) = "1/2+3/2*i"
  cxSortRatioOk(18) = "ratio((1+2i)/2)": cxSortRatioExpect(18) = "1/2+i"
  cxSortRatioOk(19) = "ratio(2/(1+3i))": cxSortRatioExpect(19) = "1/5-3/5*i"
  cxSortRatioOk(20) = "ratio((2+3i)**2)": cxSortRatioExpect(20) = "-5+12i"
  cxSortRatioOk(21) = "ratio((1+i)/(1-i))": cxSortRatioExpect(21) = "i"
  cxSortRatioOk(22) = "ratio(3i/2)": cxSortRatioExpect(22) = "3/2*i"
  cxSortRatioOk(23) = "ratio(1/2+3i/4)": cxSortRatioExpect(23) = "1/2+3/4*i"
  cxSortRatioOk(24) = "3+ratio(0.2)i": cxSortRatioExpect(24) = "3+0.2i"
  dim sri as Integer
  for sri = 1 to 24
    if Parser_TryEvaluateEx(cxSortRatioOk(sri), r, rt, ia) = FALSE orelse rt <> cxSortRatioExpect(sri) then
      print "[complex-opt] FAIL: """ & cxSortRatioOk(sri) & """ -> """ & rt & """ err=" & Parser_GetLastError()
      print "[complex-opt]      want: """ & cxSortRatioExpect(sri) & """"
      subFail += 1
    else
      print "[complex-opt] PASS: """ & cxSortRatioOk(sri) & """ -> """ & rt & """"
      subPass += 1
    end if
  next sri

  print "=== Negative argument magnitude bands (complex, 3 ranges) ==="
  dim negCxPass as Integer = 0
  dim negCxFail as Integer = 0
  RunNegBandCxBatch negCxPass, negCxFail
  subPass += negCxPass
  subFail += negCxFail
  print "Negative-argument band (complex) sub-tests: passed " & str(negCxPass) & ", failed " & str(negCxFail)

  print "=== Positive argument magnitude bands (complex, 3 ranges) ==="
  dim posCxPass as Integer = 0
  dim posCxFail as Integer = 0
  RunPosBandCxBatch posCxPass, posCxFail
  subPass += posCxPass
  subFail += posCxFail
  print "Positive-argument band (complex) sub-tests: passed " & str(posCxPass) & ", failed " & str(posCxFail)

  dim cxi as Integer
  for cxi = 1 to CX_COVERAGE_COUNT
    if cxCovIsErr(cxi) then
      if Parser_TryEvaluateEx(cxCovExpr(cxi), r, rt, ia) then
        print "[complex-cov] FAIL: "; cxCovLabel(cxi); " expected error, got """; rt; """"
        subFail += 1
      elseif instr(lcase(Parser_GetLastError()), lcase(cxCovErr(cxi))) = 0 then
        print "[complex-cov] FAIL: "; cxCovLabel(cxi); " err="; Parser_GetLastError()
        subFail += 1
      else
        subPass += 1
      end if
    else
      if Parser_TryEvaluateEx(cxCovExpr(cxi), r, rt, ia) = FALSE orelse ResultCloseEnough(rt, cxCovExpect(cxi)) = FALSE then
        print "[complex-cov] FAIL: "; cxCovLabel(cxi); " "; cxCovExpr(cxi); " -> """; rt; """ want """; cxCovExpect(cxi); """"
        subFail += 1
      else
        subPass += 1
      end if
    end if
  next cxi

  Parser_SetSupportComplexNumbers(FALSE)
  if Parser_GetSupportComplexNumbers() <> FALSE then
    print "[complex-opt] FAIL: expected support flag OFF after disable"
    subFail += 1
  else
    print "[complex-opt] PASS: getter reports disabled after Parser_SetSupportComplexNumbers(FALSE)"
    subPass += 1
  end if

  dim cxRealOffOk(1 to 3) as String
  dim cxRealOffExpect(1 to 3) as String
  cxRealOffOk(1) = "real(5)": cxRealOffExpect(1) = "5"
  cxRealOffOk(2) = "cart(polar(2))": cxRealOffExpect(2) = "2"
  cxRealOffOk(3) = "cart((5,0))": cxRealOffExpect(3) = "5"
  dim cro as Integer
  for cro = 1 to 3
    if Parser_TryEvaluateEx(cxRealOffOk(cro), r, rt, ia) = FALSE orelse rt <> cxRealOffExpect(cro) then
      print "[complex-opt] FAIL: """ & cxRealOffOk(cro) & """ with support OFF -> """ & rt & """ err=" & Parser_GetLastError()
      subFail += 1
    else
      print "[complex-opt] PASS: """ & cxRealOffOk(cro) & """ with support OFF -> """ & rt & """"
      subPass += 1
    end if
  next cro

  if Parser_TryEvaluateEx("cart((5,1))", r, rt, ia) then
    print "[complex-opt] FAIL: expected error for cart((5,1)) with support OFF but got """ & rt & """"
    subFail += 1
  else
    dim errCart as String = lcase(Parser_GetLastError())
    if instr(errCart, "incompatible operands") > 0 then
      print "[complex-opt] PASS: cart((5,1)) with support OFF -> incompatible operands"
      subPass += 1
    else
      print "[complex-opt] FAIL: cart((5,1)) with support OFF expected incompatible operands, got """ & Parser_GetLastError() & """"
      subFail += 1
    end if
  end if

  if Parser_TryEvaluateEx("10+5i", r, rt, ia) then
    print "[complex-opt] FAIL: expected failure for 10+5i with complex support OFF"
    subFail += 1
  else
    dim errL as String = lcase(Parser_GetLastError())
    if (instr(errL, "unknown variable") > 0 andalso instr(errL, "i") > 0) orelse instr(errL, "unexpected token") > 0 then
      print "[complex-opt] PASS: complex literal rejected when support is OFF"
      subPass += 1
    else
      print "[complex-opt] FAIL: expected parse error for 10+5i with complex support OFF, got """ & Parser_GetLastError() & """"
      subFail += 1
    end if
  end if

  g_passed += subPass
  g_failed += subFail
  print "Complex-option sub-tests: passed " & str(subPass) & ", failed " & str(subFail)
  print ""
end sub

private sub RunFloatMagnitudeLiteralTests()
  print "=== Float magnitude literals (Python reference) ==="
  dim subPass as Integer = 0
  dim subFail as Integer = 0
  dim r as Double
  dim rt as String
  dim ia as Boolean
  dim i as Integer
  for i = 1 to FLOAT_MAG_COUNT
    if floatMagIsErr(i) then
      if Parser_TryEvaluateEx(floatMagExpr(i), r, rt, ia) then
        print "[floatmag] FAIL: "; floatMagLabel(i); " expected error, got """; rt; """"
        subFail += 1
      elseif instr(lcase(Parser_GetLastError()), lcase(floatMagErr(i))) = 0 then
        print "[floatmag] FAIL: "; floatMagLabel(i); " err="; Parser_GetLastError()
        subFail += 1
      else
        subPass += 1
      end if
    else
      if Parser_TryEvaluateEx(floatMagExpr(i), r, rt, ia) = FALSE orelse ResultCloseEnough(rt, floatMagExpect(i)) = FALSE then
        print "[floatmag] FAIL: "; floatMagLabel(i); " "; floatMagExpr(i); " -> """; rt; """ want """; floatMagExpect(i); """"
        subFail += 1
      else
        subPass += 1
      end if
    end if
  next i
  g_passed += subPass
  g_failed += subFail
  print "Float-magnitude sub-tests: passed " & str(subPass) & ", failed " & str(subFail)
  print ""
end sub

private sub RunNanInfTests()
  print "=== NaN/Inf (parity with C++ buildNanInfCases, literal injection) ==="
  dim subPass as Integer = 0
  dim subFail as Integer = 0
  dim r as Double
  dim rt as String
  dim ia as Boolean
  dim i as Integer
  for i = 1 to NANINF_MIRROR_COUNT
    if naninfIsErr(i) then
      if Parser_TryEvaluateEx(naninfExpr(i), r, rt, ia) then
        print "[naninf] FAIL: "; naninfLabel(i); " expected error, got """; rt; """"
        subFail += 1
      elseif instr(lcase(Parser_GetLastError()), lcase(naninfErr(i))) = 0 then
        print "[naninf] FAIL: "; naninfLabel(i); " err="; Parser_GetLastError()
        subFail += 1
      else
        subPass += 1
      end if
    else
      if Parser_TryEvaluateEx(naninfExpr(i), r, rt, ia) = FALSE orelse ResultCloseEnough(rt, naninfExpect(i)) = FALSE then
        print "[naninf] FAIL: "; naninfLabel(i); " "; naninfExpr(i); " -> """; rt; """ want """; naninfExpect(i); """"
        subFail += 1
      else
        subPass += 1
      end if
    end if
  next i
  g_passed += subPass
  g_failed += subFail
  print "NaN/Inf sub-tests: passed " & str(subPass) & ", failed " & str(subFail)
  print ""
end sub

private sub RunBuiltinArityTableTests()
  print "=== Builtin arity table (central validation) ==="
  dim subPass as Integer = 0
  dim subFail as Integer = 0
  dim r as Double
  dim rt as String
  dim ia as Boolean

  dim errExprs(1 to 6) as String
  dim errNeed(1 to 6) as String
  errExprs(1) = "ratio(1,2)": errNeed(1) = "expects 1 argument(s)"
  errExprs(2) = "sortby()": errNeed(2) = "sortby expects a function that takes 1 parameter"
  errExprs(3) = "sortby(1)": errNeed(3) = "sortby expects a function that takes 1 parameter"
  errExprs(4) = "sortby(1,2,3)": errNeed(4) = "sortby expects"
  errExprs(5) = "lcm()": errNeed(5) = "expects 2 argument(s)"
  errExprs(6) = "npr(1)": errNeed(6) = "expects 2 argument(s)"
  dim ei as Integer
  for ei = 1 to 6
    if Parser_TryEvaluateEx(errExprs(ei), r, rt, ia) then
      print "[arity] FAIL: """ & errExprs(ei) & """ expected error containing """ & errNeed(ei) & """ but got """ & rt & """"
      subFail += 1
    elseif (instr(lcase(Parser_GetLastError()), lcase(errNeed(ei))) = 0) then
      print "[arity] FAIL: """ & errExprs(ei) & """ got """ & Parser_GetLastError()
      subFail += 1
    else
      print "[arity] PASS: """ & errExprs(ei) & """ -> """ & errNeed(ei)
      subPass += 1
    end if
  next ei

  Parser_SetSupportTimeValues(TRUE)
  if Parser_TryEvaluateEx("milliseconds()", r, rt, ia) then
    print "[arity] FAIL: milliseconds() expected arity error"
    subFail += 1
  elseif (instr(lcase(Parser_GetLastError()), "expects 1 argument(s)") = 0) then
    print "[arity] FAIL: milliseconds() got """ & Parser_GetLastError() & """"
    subFail += 1
  else
    print "[arity] PASS: milliseconds() -> expects 1 argument(s)"
    subPass += 1
  end if
  if Parser_TryEvaluateEx("seconds(1:00,2:00)", r, rt, ia) then
    print "[arity] FAIL: seconds(1:00,2:00) expected arity error"
    subFail += 1
  elseif (instr(lcase(Parser_GetLastError()), "expects 1 argument(s)") = 0) then
    print "[arity] FAIL: seconds(1:00,2:00) got """ & Parser_GetLastError() & """"
    subFail += 1
  else
    print "[arity] PASS: seconds(1:00,2:00) -> expects 1 argument(s)"
    subPass += 1
  end if
  Parser_SetSupportTimeValues(FALSE)

  g_passed += subPass
  g_failed += subFail
  print "Builtin-arity sub-tests: passed " & str(subPass) & ", failed " & str(subFail)
  print ""
end sub

private sub RunIncompleteFunctionCallHintTests()
  print "=== Incomplete function call hints (name followed by open paren) ==="
  dim subPass as Integer = 0
  dim subFail as Integer = 0
  dim r as Double
  dim rt as String
  dim ia as Boolean

  Parser_ClearVariables()
  if Parser_TryEvaluateEx("f(x)=x", r, rt, ia) = FALSE then
    print "[open-paren-hint] FAIL: setup f(x)=x -> " & Parser_GetLastError()
    subFail += 1
  end if

  dim hintExprs(1 to 8) as String
  dim hintNeed(1 to 8) as String
  hintExprs(1) = "f(": hintNeed(1) = "user-defined function: f(x)"
  hintExprs(2) = "f (": hintNeed(2) = "user-defined function: f(x)"
  hintExprs(3) = "abs(": hintNeed(3) = "function: abs(value)"
  hintExprs(4) = "abs (": hintNeed(4) = "function: abs(value)"
  hintExprs(5) = "log(": hintNeed(5) = "function: log(value, base)"
  hintExprs(6) = "log (": hintNeed(6) = "function: log(value, base)"
  hintExprs(7) = "log10(": hintNeed(7) = "function: log10(value)"
  hintExprs(8) = "log10 (": hintNeed(8) = "function: log10(value)"
  dim hi as Integer
  for hi = 1 to 8
    if Parser_TryEvaluateEx(hintExprs(hi), r, rt, ia) then
      print "[open-paren-hint] FAIL: """ & hintExprs(hi) & """ expected error containing """ & hintNeed(hi) & """ but got """ & rt & """"
      subFail += 1
    elseif instr(Parser_GetLastError(), hintNeed(hi)) = 0 then
      print "[open-paren-hint] FAIL: """ & hintExprs(hi) & """ got """ & Parser_GetLastError()
      subFail += 1
    else
      print "[open-paren-hint] PASS: """ & hintExprs(hi) & """"
      subPass += 1
    end if
  next hi

  dim endHintExprs(1 to 4) as String
  dim endHintNeed(1 to 4) as String
  endHintExprs(1) = "1+sin": endHintNeed(1) = "function: sin(angle)"
  endHintExprs(2) = "1+sin(": endHintNeed(2) = "function: sin(angle)"
  endHintExprs(3) = "1+sin(rad(": endHintNeed(3) = "function: rad(...)"
  endHintExprs(4) = "sortby((1,2,3, abs": endHintNeed(4) = "function: abs(value)"
  dim ei as Integer
  for ei = 1 to 4
    if Parser_TryEvaluateEx(endHintExprs(ei), r, rt, ia) then
      print "[expr-end-hint] FAIL: """ & endHintExprs(ei) & """ expected error containing """ & endHintNeed(ei) & """ but got """ & rt & """"
      subFail += 1
    elseif instr(Parser_GetLastError(), endHintNeed(ei)) = 0 then
      print "[expr-end-hint] FAIL: """ & endHintExprs(ei) & """ got """ & Parser_GetLastError() & """"
      subFail += 1
    else
      print "[expr-end-hint] PASS: """ & endHintExprs(ei) & """"
      subPass += 1
    end if
  next ei
  if Parser_TryEvaluateEx("1+sin(rad(45)", r, rt, ia) then
    print "[expr-end-hint] FAIL: 1+sin(rad(45) expected error"
    subFail += 1
  elseif instr(Parser_GetLastError(), "missing closing parenthesis") = 0 then
    print "[expr-end-hint] FAIL: 1+sin(rad(45) got """ & Parser_GetLastError() & """"
    subFail += 1
  else
    print "[expr-end-hint] PASS: 1+sin(rad(45)"
    subPass += 1
  end if

  dim sortbyExprs(1 to 3) as String
  dim sortbyNeed as String = "sortby expects exactly one function"
  sortbyExprs(1) = "sortby((1,2,3), f("
  sortbyExprs(2) = "sortby((1,2,3), abs("
  sortbyExprs(3) = "sortby((1,2,3), log("
  dim si as Integer
  for si = 1 to 3
    if Parser_TryEvaluateEx(sortbyExprs(si), r, rt, ia) then
      print "[open-paren-hint] FAIL: """ & sortbyExprs(si) & """ expected sortby error"
      subFail += 1
    elseif instr(lcase(Parser_GetLastError()), lcase(sortbyNeed)) = 0 then
      print "[open-paren-hint] FAIL: """ & sortbyExprs(si) & """ got """ & Parser_GetLastError() & """"
      subFail += 1
    else
      print "[open-paren-hint] PASS: """ & sortbyExprs(si) & """"
      subPass += 1
    end if
  next si

  dim semiHintExprs(1 to 10) as String
  dim semiHintMustNot(1 to 10) as String
  semiHintExprs(1) = "sin(1)/cos;": semiHintMustNot(1) = "function: cos"
  semiHintExprs(2) = "sin(1)/cos; 2": semiHintMustNot(2) = "function: cos"
  semiHintExprs(3) = "1+sin;": semiHintMustNot(3) = "function: sin"
  semiHintExprs(4) = "1+sin(;": semiHintMustNot(4) = "function: sin"
  semiHintExprs(5) = "sin(;": semiHintMustNot(5) = "function: sin"
  semiHintExprs(6) = "cos;": semiHintMustNot(6) = "function: cos"
  semiHintExprs(7) = "f(x)=x; f(;": semiHintMustNot(7) = "user-defined function:"
  semiHintExprs(8) = "f(x)=x; cos;": semiHintMustNot(8) = "function: cos"
  semiHintExprs(9) = "f(x)=x; 1+sum;": semiHintMustNot(9) = "function: sum"
  semiHintExprs(10) = "sum(;": semiHintMustNot(10) = "function: sum"
  dim shi as Integer
  for shi = 1 to 10
    if Parser_TryEvaluateEx(semiHintExprs(shi), r, rt, ia) then
      print "[semi-fn-hint] FAIL: """ & semiHintExprs(shi) & """ expected error, got """ & rt & """"
      subFail += 1
    elseif instr(Parser_GetLastError(), semiHintMustNot(shi)) > 0 then
      print "[semi-fn-hint] FAIL: """ & semiHintExprs(shi) & """ must not hint: """ & Parser_GetLastError() & """"
      subFail += 1
    else
      print "[semi-fn-hint] PASS: """ & semiHintExprs(shi) & """"
      subPass += 1
    end if
  next shi
  if Parser_TryEvaluateEx("f(x)=x", r, rt, ia) = FALSE then
    print "[semi-fn-hint] FAIL: setup f(x)=x -> " & Parser_GetLastError()
    subFail += 1
  elseif Parser_TryEvaluateEx("f(x)=x; f(", r, rt, ia) then
    print "[semi-fn-hint] FAIL: f(x)=x; f( expected error"
    subFail += 1
  elseif instr(Parser_GetLastError(), "user-defined function: f(x)") = 0 then
    print "[semi-fn-hint] FAIL: f(x)=x; f( got """ & Parser_GetLastError() & """"
    subFail += 1
  else
    print "[semi-fn-hint] PASS: f(x)=x; f("
    subPass += 1
  end if
  if Parser_TryEvaluateEx("f(x)=x; 1+f", r, rt, ia) then
    print "[semi-fn-hint] FAIL: f(x)=x; 1+f expected error"
    subFail += 1
  elseif instr(Parser_GetLastError(), "user-defined function: f(x)") = 0 then
    print "[semi-fn-hint] FAIL: f(x)=x; 1+f got """ & Parser_GetLastError() & """"
    subFail += 1
  else
    print "[semi-fn-hint] PASS: f(x)=x; 1+f"
    subPass += 1
  end if

  dim badCloserExprs(1 to 5) as String
  dim badCloserNeed(1 to 5) as String
  dim badCloserMustNot(1 to 5) as String
  badCloserExprs(1) = "sin)": badCloserNeed(1) = "mismatched closing parenthesis": badCloserMustNot(1) = "function:"
  badCloserExprs(2) = "sin]": badCloserNeed(2) = "mismatched closing bracket": badCloserMustNot(2) = "function:"
  badCloserExprs(3) = "sin}": badCloserNeed(3) = "mismatched closing brace": badCloserMustNot(3) = "function:"
  badCloserExprs(4) = "f(x)=x; f)": badCloserNeed(4) = "mismatched closing parenthesis": badCloserMustNot(4) = "user-defined function:"
  badCloserExprs(5) = "sin)(1)": badCloserNeed(5) = "mismatched closing parenthesis": badCloserMustNot(5) = "function:"
  dim bi as Integer
  for bi = 1 to 5
    if Parser_TryEvaluateEx(badCloserExprs(bi), r, rt, ia) then
      print "[bare-fn-closer] FAIL: """ & badCloserExprs(bi) & """ expected error, got """ & rt & """"
      subFail += 1
    elseif instr(Parser_GetLastError(), badCloserNeed(bi)) = 0 then
      print "[bare-fn-closer] FAIL: """ & badCloserExprs(bi) & """ got """ & Parser_GetLastError() & """"
      subFail += 1
    elseif instr(Parser_GetLastError(), badCloserMustNot(bi)) > 0 then
      print "[bare-fn-closer] FAIL: """ & badCloserExprs(bi) & """ must not hint: """ & Parser_GetLastError() & """"
      subFail += 1
    else
      print "[bare-fn-closer] PASS: """ & badCloserExprs(bi) & """"
      subPass += 1
    end if
  next bi

  g_passed += subPass
  g_failed += subFail
  print "Incomplete-open-paren-hint sub-tests: passed " & str(subPass) & ", failed " & str(subFail)
  print ""
end sub

private sub RunMultiStmtSemicolonTests()
  print "=== multi-statement semicolon stability ==="
  dim subPass as Integer = 0
  dim subFail as Integer = 0
  dim r as Double
  dim rt as String
  dim ia as Boolean

  dim okExprs(1 to 8) as String
  dim okNeed(1 to 8) as String
  okExprs(1) = "1+2; 3+4": okNeed(1) = "7"
  okExprs(2) = "1+2; 3+4;": okNeed(2) = "7"
  okExprs(3) = "f(x)=x+1; f(3)": okNeed(3) = "4"
  okExprs(4) = "f(x)=x+1; f(3);": okNeed(4) = "4"
  okExprs(5) = "0xAA; hex": okNeed(5) = "0xAA"
  okExprs(6) = "0xAA; hex;": okNeed(6) = "0xAA"
  okExprs(7) = "(10,20,30); hex()": okNeed(7) = "(0xA,0x14,0x1E)"
  okExprs(8) = "(10,20,30); hex;": okNeed(8) = "(0xA,0x14,0x1E)"
  dim mi as Integer
  for mi = 1 to 8
    if Parser_TryEvaluateEx(okExprs(mi), r, rt, ia) = FALSE then
      print "[multi-stmt-semi] FAIL: """ & okExprs(mi) & """ -> " & Parser_GetLastError()
      subFail += 1
    elseif ResultCloseEnough(rt, okNeed(mi)) = FALSE then
      print "[multi-stmt-semi] FAIL: """ & okExprs(mi) & """ expected """ & okNeed(mi) & """ got """ & rt & """"
      subFail += 1
    else
      print "[multi-stmt-semi] PASS: """ & okExprs(mi) & """"
      subPass += 1
    end if
  next mi

  if Parser_TryEvaluateEx("(10,20,30); cos;", r, rt, ia) then
    print "[multi-stmt-semi] FAIL: (10,20,30); cos; expected error"
    subFail += 1
  elseif instr(Parser_GetLastError(), "unexpected token") = 0 then
    print "[multi-stmt-semi] FAIL: (10,20,30); cos; got """ & Parser_GetLastError() & """"
    subFail += 1
  elseif instr(Parser_GetLastError(), "function: cos") > 0 then
    print "[multi-stmt-semi] FAIL: (10,20,30); cos; must not hint"
    subFail += 1
  else
    print "[multi-stmt-semi] PASS: (10,20,30); cos;"
    subPass += 1
  end if

  if Parser_TryEvaluateEx("1;;2", r, rt, ia) then
    print "[multi-stmt-semi] FAIL: 1;;2 expected error"
    subFail += 1
  elseif instr(Parser_GetLastError(), "empty statement") = 0 then
    print "[multi-stmt-semi] FAIL: 1;;2 got """ & Parser_GetLastError() & """"
    subFail += 1
  else
    print "[multi-stmt-semi] PASS: 1;;2"
    subPass += 1
  end if

  g_passed += subPass
  g_failed += subFail
  print "Multi-stmt-semicolon sub-tests: passed " & str(subPass) & ", failed " & str(subFail)
  print ""
end sub

private sub RunGcdLcmNcrNprArrayBroadcastTests()
  print "=== gcd/lcm/ncr/npr array broadcast ==="
  dim subPass as Integer = 0
  dim subFail as Integer = 0
  dim r as Double
  dim rt as String
  dim ia as Boolean

  dim okExprs(1 to 6) as String
  dim okNeed(1 to 6) as String
  okExprs(1) = "gcd((84,30),6)": okNeed(1) = "(6,6)"
  okExprs(2) = "lcm((6,8),3)": okNeed(2) = "(6,24)"
  okExprs(3) = "lcm((6,8),(3,5))": okNeed(3) = "(6,40)"
  okExprs(4) = "ncr((5,6),(2,3))": okNeed(4) = "(10,20)"
  okExprs(5) = "npr((5,6),(2,3))": okNeed(5) = "(20,120)"
  okExprs(6) = "gcd(6,(84,30))": okNeed(6) = "(6,6)"
  dim oi as Integer
  for oi = 1 to 6
    if Parser_TryEvaluateEx(okExprs(oi), r, rt, ia) = FALSE then
      print "[gcd-ncr-array] FAIL: """ & okExprs(oi) & """ -> " & Parser_GetLastError()
      subFail += 1
    elseif ResultCloseEnough(rt, okNeed(oi)) = FALSE then
      print "[gcd-ncr-array] FAIL: """ & okExprs(oi) & """ expected """ & okNeed(oi) & """ got """ & rt & """"
      subFail += 1
    else
      print "[gcd-ncr-array] PASS: """ & okExprs(oi) & """"
      subPass += 1
    end if
  next oi

  if Parser_TryEvaluateEx("gcd((1,2),(3,4,5))", r, rt, ia) then
    print "[gcd-ncr-array] FAIL: mismatched array lengths expected error"
    subFail += 1
  elseif instr(lcase(Parser_GetLastError()), "incompatible operands") = 0 then
    print "[gcd-ncr-array] FAIL: gcd length mismatch got """ & Parser_GetLastError() & """"
    subFail += 1
  else
    print "[gcd-ncr-array] PASS: gcd((1,2),(3,4,5)) -> incompatible operands"
    subPass += 1
  end if

  if Parser_TryEvaluateEx("ncr((5,6),(2,7))", r, rt, ia) then
    print "[gcd-ncr-array] FAIL: ncr domain error expected failure"
    subFail += 1
  elseif instr(lcase(Parser_GetLastError()), "numeric error in ncr()") = 0 then
    print "[gcd-ncr-array] FAIL: ncr domain got """ & Parser_GetLastError() & """"
    subFail += 1
  else
    print "[gcd-ncr-array] PASS: ncr((5,6),(2,7)) -> numeric error"
    subPass += 1
  end if

  g_passed += subPass
  g_failed += subFail
  print "gcd/lcm/ncr/npr array-broadcast sub-tests: passed " & str(subPass) & ", failed " & str(subFail)
  print ""
end sub

private sub RunRatioInExpressionTests()
  print "=== ratio() in expressions (numeric value, not numerator metadata) ==="
  dim subPass as Integer = 0
  dim subFail as Integer = 0
  dim r as Double
  dim rt as String
  dim ia as Boolean

  dim okExprs(1 to 30) as String
  dim okNeed(1 to 30) as String
  okExprs(1) = "ratio(1.3456)+1": okNeed(1) = "2.3456"
  okExprs(2) = "ratio(1.3456)*10": okNeed(2) = "13.456"
  okExprs(3) = "ratio(1.3456)-1": okNeed(3) = "0.3456"
  okExprs(4) = "int(ratio(1.3456))": okNeed(4) = "1"
  okExprs(5) = "trunc(ratio(1.3456))": okNeed(5) = "1"
  okExprs(6) = "floor(ratio(1.3456))": okNeed(6) = "1"
  okExprs(7) = "prod(ratio(1.3456),1)": okNeed(7) = "1.3456"
  okExprs(8) = "product(ratio(1.3456),1)": okNeed(8) = "1.3456"
  okExprs(9) = "sum(ratio(0.5),ratio(0.5))": okNeed(9) = "1"
  okExprs(10) = "f(x)=x-1; f(ratio(1.3456))": okNeed(10) = "0.3456"
  okExprs(11) = "g(x)=x*2; g(ratio(0.25))": okNeed(11) = "0.5"
  okExprs(12) = "ratio(1.3456)+ratio(0.5)": okNeed(12) = "1.8456"
  okExprs(13) = "abs(ratio(-0.5))": okNeed(13) = "0.5"
  okExprs(14) = "sign(ratio(-0.5))": okNeed(14) = "-1"
  okExprs(15) = "clamp(ratio(1.3456),0,2)": okNeed(15) = "841/625"
  okExprs(16) = "ratio(0.5)+1": okNeed(16) = "1.5"
  okExprs(17) = "ratio(0.5)*3": okNeed(17) = "1.5"
  okExprs(18) = "ratio(1.3456)/2": okNeed(18) = "0.6728"
  okExprs(19) = "reverse(ratio(1/2),ratio(1))": okNeed(19) = "(1, 1/2)"
  okExprs(20) = "unpack(ratio(1),ratio(1/2))": okNeed(20) = "(1, 1/2)"
  okExprs(21) = "unique(ratio(1),ratio(0.5),ratio(1))": okNeed(21) = "(1, 1/2)"
  okExprs(22) = "sort((ratio(3),ratio(1),ratio(1/2)))": okNeed(22) = "(1/2, 1, 3)"
  okExprs(23) = "sortby((ratio(3),ratio(1/2)),x:x)": okNeed(23) = "(1/2, 3)"
  okExprs(24) = "sortby((ratio(3),ratio(1/2)),abs)": okNeed(24) = "(1/2, 3)"
  okExprs(25) = "(ratio(0.5),ratio(0.25))": okNeed(25) = "(1/2, 1/4)"
  okExprs(26) = "ratio(1.3456)-ratio(1.3456)": okNeed(26) = "0"
  okExprs(27) = "round(ratio(1.3456))": okNeed(27) = "1"
  okExprs(28) = "min(ratio(1.3456), 2)": okNeed(28) = "841/625"
  okExprs(29) = "max(ratio(1.3456), 2)": okNeed(29) = "2"
  okExprs(30) = "flat(ratio(1),ratio(1/2))": okNeed(30) = "(1, 1/2)"
  dim ri as Integer
  for ri = 1 to 30
    if Parser_TryEvaluateEx(okExprs(ri), r, rt, ia) = FALSE then
      print "[ratio-expr] FAIL: """ & okExprs(ri) & """ -> " & Parser_GetLastError()
      subFail += 1
    elseif ResultCloseEnough(rt, okNeed(ri)) = FALSE then
      print "[ratio-expr] FAIL: """ & okExprs(ri) & """ expected """ & okNeed(ri) & """ got """ & rt & """"
      subFail += 1
    else
      print "[ratio-expr] PASS: """ & okExprs(ri) & """"
      subPass += 1
    end if
  next ri

  g_passed += subPass
  g_failed += subFail
  print "ratio-in-expression sub-tests: passed " & str(subPass) & ", failed " & str(subFail)
  print ""
end sub

private sub RunMinMaxPreserveWinnerTests()
  print "=== min/max preserve winning operand metadata ==="
  dim subPass as Integer = 0
  dim subFail as Integer = 0
  dim r as Double
  dim rt as String
  dim ia as Boolean

  dim okExprs(1 to 18) as String
  dim okNeed(1 to 18) as String
  okExprs(1) = "max(1,2**54,2)": okNeed(1) = "18014398509481984"
  okExprs(2) = "max(1.1,2**54,2)": okNeed(2) = "18014398509481984"
  okExprs(3) = "max(2**54,1.1,2)": okNeed(3) = "18014398509481984"
  okExprs(4) = "min(3,1,2)": okNeed(4) = "1"
  okExprs(5) = "min(3,1,2.1)": okNeed(5) = "1"
  okExprs(6) = "min(1.0,1)": okNeed(6) = "1"
  okExprs(7) = "min(1,1.0)": okNeed(7) = "1"
  okExprs(8) = "min((3,1,2))": okNeed(8) = "1"
  okExprs(9) = "max((1,2**54,2))": okNeed(9) = "18014398509481984"
  okExprs(10) = "min(ratio(0.5),1)": okNeed(10) = "1/2"
  okExprs(11) = "max(ratio(1.3456),2)": okNeed(11) = "2"
  okExprs(12) = "max(1,nan)": okNeed(12) = "1"
  okExprs(13) = "min(nan,1)": okNeed(13) = "1"
  okExprs(14) = "min(1,1,1)": okNeed(14) = "1"
  okExprs(15) = "min(nan,1,2)": okNeed(15) = "1"
  okExprs(16) = "max(nan,1,2)": okNeed(16) = "2"
  okExprs(17) = "min(nan,nan)": okNeed(17) = "nan"
  okExprs(18) = "max(nan,nan,nan)": okNeed(18) = "nan"
  dim mi as Integer
  for mi = 1 to 18
    if Parser_TryEvaluateEx(okExprs(mi), r, rt, ia) = FALSE then
      print "[minmax-winner] FAIL: """ & okExprs(mi) & """ -> " & Parser_GetLastError()
      subFail += 1
    elseif rt <> okNeed(mi) then
      print "[minmax-winner] FAIL: """ & okExprs(mi) & """ got """ & rt & """ want """ & okNeed(mi) & """"
      subFail += 1
    else
      subPass += 1
    end if
  next mi
  g_passed += subPass
  g_failed += subFail
  print "minmax-preserve-winner sub-tests: passed " & str(subPass) & ", failed " & str(subFail)
  print ""
end sub

private sub RunClampPreserveIntegerMetadataTests()
  print "=== clamp preserve integer metadata ==="
  dim subPass as Integer = 0
  dim subFail as Integer = 0
  dim r as Double
  dim rt as String
  dim ia as Boolean

  dim okExprs(1 to 5) as String
  dim okNeed(1 to 5) as String
  okExprs(1) = "clamp((0,1,2,3),1,2)": okNeed(1) = "(1, 1, 2, 2)"
  okExprs(2) = "clamp(0,1,2)": okNeed(2) = "1"
  okExprs(3) = "clamp(3,1,2)": okNeed(3) = "2"
  okExprs(4) = "clamp(1,1,2)": okNeed(4) = "1"
  okExprs(5) = "clamp(1.5,1,2)": okNeed(5) = "1.5"
  dim ci as Integer
  for ci = 1 to 5
    if Parser_TryEvaluateEx(okExprs(ci), r, rt, ia) = FALSE then
      print "[clamp-int] FAIL: """ & okExprs(ci) & """ -> " & Parser_GetLastError()
      subFail += 1
    elseif rt <> okNeed(ci) then
      print "[clamp-int] FAIL: """ & okExprs(ci) & """ got """ & rt & """ want """ & okNeed(ci) & """"
      subFail += 1
    else
      subPass += 1
    end if
  next ci
  g_passed += subPass
  g_failed += subFail
  print "clamp-preserve-integer sub-tests: passed " & str(subPass) & ", failed " & str(subFail)
  print ""
end sub

private sub RunClampInvertedBoundsTests()
  print "=== clamp inverted bounds (min > max) ==="
  dim subPass as Integer = 0
  dim subFail as Integer = 0
  dim r as Double
  dim rt as String
  dim ia as Boolean

  dim okExprs(1 to 3) as String
  dim okNeed(1 to 3) as String
  okExprs(1) = "clamp(4,4,2)": okNeed(1) = "4"
  okExprs(2) = "clamp(5,4,2)": okNeed(2) = "2"
  okExprs(3) = "clamp(4.01,4,2)": okNeed(3) = "2"
  dim ci as Integer
  for ci = 1 to 3
    if Parser_TryEvaluateEx(okExprs(ci), r, rt, ia) = FALSE then
      print "[clamp-inv] FAIL: """ & okExprs(ci) & """ -> " & Parser_GetLastError()
      subFail += 1
    elseif ResultCloseEnough(rt, okNeed(ci)) = FALSE then
      print "[clamp-inv] FAIL: """ & okExprs(ci) & """ got """ & rt & """ want """ & okNeed(ci) & """"
      subFail += 1
    else
      subPass += 1
    end if
  next ci
  g_passed += subPass
  g_failed += subFail
  print "clamp-inverted-bounds sub-tests: passed " & str(subPass) & ", failed " & str(subFail)
  print ""
end sub

private sub RunClampNanBoundTests()
  print "=== clamp NaN bounds ==="
  dim subPass as Integer = 0
  dim subFail as Integer = 0
  dim r as Double
  dim rt as String
  dim ia as Boolean

  dim okExprs(1 to 18) as String
  dim okNeed(1 to 18) as String
  okExprs(1) = "clamp(nan,1,2)": okNeed(1) = "1"
  okExprs(2) = "clamp(nan,nan,2)": okNeed(2) = "nan"
  okExprs(3) = "clamp(nan,1,nan)": okNeed(3) = "1"
  okExprs(4) = "clamp(0,1,nan)": okNeed(4) = "1"
  okExprs(5) = "clamp(2,1,nan)": okNeed(5) = "nan"
  okExprs(6) = "clamp(1,nan,2)": okNeed(6) = "1"
  okExprs(7) = "clamp(3,nan,2)": okNeed(7) = "2"
  okExprs(8) = "clamp(2,nan,2)": okNeed(8) = "2"
  okExprs(9) = "clamp(-inf,nan,2)": okNeed(9) = "-inf"
  okExprs(10) = "clamp(-inf,2,nan)": okNeed(10) = "2"
  okExprs(11) = "clamp(inf,2,nan)": okNeed(11) = "nan"
  okExprs(12) = "clamp(nan,inf,inf)": okNeed(12) = "inf"
  okExprs(13) = "clamp(nan,-inf,1)": okNeed(13) = "-inf"
  okExprs(14) = "clamp(nan,-inf,inf)": okNeed(14) = "-inf"
  okExprs(15) = "clamp(inf,nan,2)": okNeed(15) = "2"
  okExprs(16) = "clamp(-inf,-inf,nan)": okNeed(16) = "-inf"
  okExprs(17) = "clamp(inf,inf,nan)": okNeed(17) = "inf"
  okExprs(18) = "clamp(1,1,nan)": okNeed(18) = "1"
  dim ci as Integer
  for ci = 1 to 18
    if Parser_TryEvaluateEx(okExprs(ci), r, rt, ia) = FALSE then
      print "[clamp-nan] FAIL: """ & okExprs(ci) & """ -> " & Parser_GetLastError()
      subFail += 1
    elseif ResultCloseEnough(rt, okNeed(ci)) = FALSE then
      print "[clamp-nan] FAIL: """ & okExprs(ci) & """ got """ & rt & """ want """ & okNeed(ci) & """"
      subFail += 1
    else
      subPass += 1
    end if
  next ci
  g_passed += subPass
  g_failed += subFail
  print "clamp-nan-bound sub-tests: passed " & str(subPass) & ", failed " & str(subFail)
  print ""
end sub

private sub RunExactIntegerDivisionTests()
  print "=== Exact integer division (a/b) ==="
  dim subPass as Integer = 0
  dim subFail as Integer = 0
  dim r as Double
  dim rt as String
  dim ia as Boolean

  dim okExprs(1 to 19) as String
  dim okNeed(1 to 19) as String
  okExprs(1) = "10/2": okNeed(1) = "5"
  okExprs(2) = "-12/4": okNeed(2) = "-3"
  okExprs(3) = "1955685900012/338822921": okNeed(3) = "5772"
  okExprs(4) = "111*17618791892/5772": okNeed(4) = "338822921"
  okExprs(5) = "gcd(111*17618791892,222*76568758786)": okNeed(5) = "5772"
  okExprs(6) = "(2**63 - 3)/5": okNeed(6) = "1844674407370955161"
  okExprs(7) = "(-2**63 + 3)/5": okNeed(7) = "-1844674407370955161"
  okExprs(8) = "(0xFFFFFFFFFFFFFFFF - 5)/5": okNeed(8) = "3689348814741910322"
  okExprs(9) = "(2**52 - 1)/5": okNeed(9) = "900719925474099"
  okExprs(10) = "(-2**52 + 1)/5": okNeed(10) = "-900719925474099"
  okExprs(11) = "1317624576693539401/(1/11)": okNeed(11) = "14493870343628933411"
  okExprs(12) = "1317624576693539401/(1/7)": okNeed(12) = "9223372036854775807"
  okExprs(13) = "-1317624576693539401/(1/7)": okNeed(13) = "-9223372036854775807"
  okExprs(14) = "100/(5/2)": okNeed(14) = "40"
  okExprs(15) = "-50/(5/2)": okNeed(15) = "-20"
  okExprs(16) = "700000000/(5/2)": okNeed(16) = "280000000"
  okExprs(17) = "-300000000/(5/2)": okNeed(17) = "-120000000"
  okExprs(18) = "(2**51-8)/(5/2)": okNeed(18) = "900719925474096"
  okExprs(19) = "-(2**50-1)/(3/2)": okNeed(19) = "-750599937895082"
  dim di as Integer
  for di = 1 to 19
    if Parser_TryEvaluateEx(okExprs(di), r, rt, ia) = FALSE then
      print "[exact-int-div] FAIL: """ & okExprs(di) & """ -> " & Parser_GetLastError()
      subFail += 1
    elseif ResultCloseEnough(rt, okNeed(di)) = FALSE then
      print "[exact-int-div] FAIL: """ & okExprs(di) & """ expected """ & okNeed(di) & """ got """ & rt & """"
      subFail += 1
    else
      print "[exact-int-div] PASS: """ & okExprs(di) & """"
      subPass += 1
    end if
  next di

  if Parser_TryEvaluateEx("7/3", r, rt, ia) = FALSE then
    print "[exact-int-div] FAIL: ""7/3"" -> " & Parser_GetLastError()
    subFail += 1
  elseif instr(rt, ".") = 0 andalso instr(rt, "/") = 0 then
    print "[exact-int-div] FAIL: ""7/3"" must stay non-integer float, got """ & rt & """"
    subFail += 1
  else
    print "[exact-int-div] PASS: ""7/3"" non-exact quotient stays float"
    subPass += 1
  end if

  dim noExprs(1 to 10) as String
  noExprs(1) = "0x7FFFFFFFFFFFFFFF/(1/8)"
  noExprs(2) = "2**64/(1/2)"
  noExprs(3) = "1317624576693539401*(1/11)"
  noExprs(4) = "0x7FFFFFFFFFFFFFFF*(-1/8)"
  noExprs(5) = "100/(10/3+1e-15)"
  noExprs(6) = "-50/((10/3)+1e-9)"
  noExprs(7) = "700000000/(10/3+1e-15)"
  noExprs(8) = "-300000000/(10/3+1e-15)"
  noExprs(9) = "(2**51-8)/((5/2)+1e-12)"
  noExprs(10) = "-(2**50-1)/((3/2)+1e-12)"
  dim ni as Integer
  for ni = 1 to 10
    if Parser_TryEvaluateEx(noExprs(ni), r, rt, ia) = FALSE then
      print "[exact-int-div] FAIL: """ & noExprs(ni) & """ -> " & Parser_GetLastError()
      subFail += 1
    elseif instr(rt, "e+") > 0 orelse instr(rt, "e-") > 0 orelse instr(rt, ".") > 0 then
      print "[exact-int-div] PASS: """ & noExprs(ni) & """ stays float"
      subPass += 1
    else
      print "[exact-int-div] FAIL: """ & noExprs(ni) & """ must stay non-integer float, got """ & rt & """"
      subFail += 1
    end if
  next ni

  g_passed += subPass
  g_failed += subFail
  print "Exact-integer-division sub-tests: passed " & str(subPass) & ", failed " & str(subFail)
  print ""
end sub

private sub RunExactIntegerMultiplicationTests()
  print "=== Exact integer multiplication (N*f) ==="
  dim subPass as Integer = 0
  dim subFail as Integer = 0
  dim r as Double
  dim rt as String
  dim ia as Boolean

  dim okExprs(1 to 7) as String
  dim okNeed(1 to 7) as String
  okExprs(1) = "100*(5/2)": okNeed(1) = "250"
  okExprs(2) = "-50*(5/2)": okNeed(2) = "-125"
  okExprs(3) = "700000000*(5/2)": okNeed(3) = "1750000000"
  okExprs(4) = "-300000000*(5/2)": okNeed(4) = "-750000000"
  okExprs(5) = "(2**51-8)*(5/2)": okNeed(5) = "5629499534213100"
  okExprs(6) = "(2**50)*(3/2)": okNeed(6) = "1688849860263936"
  okExprs(7) = "-(2**50-1)*2": okNeed(7) = "-2251799813685246"
  dim di as Integer
  for di = 1 to 7
    if Parser_TryEvaluateEx(okExprs(di), r, rt, ia) = FALSE then
      print "[exact-int-mul] FAIL: """ & okExprs(di) & """ -> " & Parser_GetLastError()
      subFail += 1
    elseif ResultCloseEnough(rt, okNeed(di)) = FALSE then
      print "[exact-int-mul] FAIL: """ & okExprs(di) & """ expected """ & okNeed(di) & """ got """ & rt & """"
      subFail += 1
    else
      print "[exact-int-mul] PASS: """ & okExprs(di) & """"
      subPass += 1
    end if
  next di

  if Parser_TryEvaluateEx("5*(7/3)", r, rt, ia) = FALSE then
    print "[exact-int-mul] FAIL: ""5*(7/3)"" -> " & Parser_GetLastError()
    subFail += 1
  elseif instr(rt, ".") = 0 andalso instr(rt, "/") = 0 then
    print "[exact-int-mul] FAIL: ""5*(7/3)"" must stay non-integer float, got """ & rt & """"
    subFail += 1
  else
    print "[exact-int-mul] PASS: ""5*(7/3)"" non-exact product stays float"
    subPass += 1
  end if

  dim noExprs(1 to 6) as String
  noExprs(1) = "100*((5/2)+1e-12)"
  noExprs(2) = "-50*((5/2)+1e-9)"
  noExprs(3) = "700000000*((5/2)+1e-12)"
  noExprs(4) = "-300000000*((5/2)+1e-12)"
  noExprs(5) = "(2**51-8)*((5/2)+1e-12)"
  noExprs(6) = "-(2**50-1)*(2.0001)"
  dim raw as RawResult
  dim ni as Integer
  for ni = 1 to 6
    if Parser_TryEvaluateEx(noExprs(ni), r, rt, ia) = FALSE then
      print "[exact-int-mul] FAIL: """ & noExprs(ni) & """ -> " & Parser_GetLastError()
      subFail += 1
    elseif Parser_GetLastRawResult(raw) = FALSE orelse raw.kind <> RRK_SCALAR orelse raw.scalar.real.kind <> RSK_FLOATING then
      print "[exact-int-mul] FAIL: """ & noExprs(ni) & """ must stay floating scalar, got """ & rt & """"
      subFail += 1
    else
      print "[exact-int-mul] PASS: """ & noExprs(ni) & """ stays float"
      subPass += 1
    end if
  next ni

  g_passed += subPass
  g_failed += subFail
  print "Exact-integer-multiplication sub-tests: passed " & str(subPass) & ", failed " & str(subFail)
  print ""
end sub

private sub RunBinaryBuiltinArrayLengthMismatchTests()
  print "=== binary builtin array length mismatch (incompatible operands) ==="
  dim subPass as Integer = 0
  dim subFail as Integer = 0
  dim r as Double
  dim rt as String
  dim ia as Boolean

  dim badExprs(1 to 9) as String
  badExprs(1) = "atan2((1,2),(3,4,5))"
  badExprs(2) = "gcd((1,2),(3,4,5))"
  badExprs(3) = "hypot((1,2),(3,4,5))"
  badExprs(4) = "lcm((1,2),(3,4,5))"
  badExprs(5) = "log((1,2),(3,4,5))"
  badExprs(6) = "mod((1,2),(3,4,5))"
  badExprs(7) = "ncr((1,2),(3,4,5))"
  badExprs(8) = "npr((1,2),(3,4,5))"
  badExprs(9) = "pow((1,2),(3,4,5))"
  dim bi as Integer
  for bi = 1 to 9
    if Parser_TryEvaluateEx(badExprs(bi), r, rt, ia) then
      print "[bin-array-len] FAIL: """ & badExprs(bi) & """ expected error"
      subFail += 1
    elseif instr(lcase(Parser_GetLastError()), "incompatible operands") = 0 then
      print "[bin-array-len] FAIL: """ & badExprs(bi) & """ got """ & Parser_GetLastError() & """"
      subFail += 1
    else
      print "[bin-array-len] PASS: """ & badExprs(bi) & """"
      subPass += 1
    end if
  next bi

  g_passed += subPass
  g_failed += subFail
  print "binary-builtin array-length sub-tests: passed " & str(subPass) & ", failed " & str(subFail)
  print ""
end sub

private sub RunTimeValuesSupportOptionTests()
  print "=== Time value support (parser-wide option) ==="

  Parser_SetSupportTimeValues(TRUE)

  dim subPass as Integer = 0
  dim subFail as Integer = 0

  if Parser_GetSupportTimeValues() = FALSE then
    print "[time-opt] FAIL: expected support flag ON after enabling"
    subFail += 1
  else
    print "[time-opt] PASS: getter reports enabled after Parser_SetSupportTimeValues(TRUE)"
    subPass += 1
  end if

  dim r as Double
  dim rt as String
  dim ia as Boolean
  if Parser_TryEvaluateEx("1:30 + 2:45.111", r, rt, ia) = FALSE orelse rt <> "04:15.111" then
    print "[time-opt] FAIL: duration add with flag ON, got """ & rt & """ err=" & Parser_GetLastError()
    subFail += 1
  else
    print "[time-opt] PASS: 1:30 + 2:45.111 -> 04:15.111 with support flag ON"
    subPass += 1
  end if

  if Parser_TryEvaluateEx("second + 5", r, rt, ia) = FALSE orelse rt <> "00:06" then
    print "[time-opt] FAIL: second constant with flag ON, got """ & rt & """ err=" & Parser_GetLastError()
    subFail += 1
  else
    print "[time-opt] PASS: second + 5 -> 00:06 with support flag ON"
    subPass += 1
  end if

  Parser_SetSupportTimeValues(FALSE)
  if Parser_GetSupportTimeValues() <> FALSE then
    print "[time-opt] FAIL: expected support flag OFF after disable"
    subFail += 1
  else
    print "[time-opt] PASS: getter reports disabled after Parser_SetSupportTimeValues(FALSE)"
    subPass += 1
  end if

  if Parser_TryEvaluateEx("1:30 + 2:45.111", r, rt, ia) then
    print "[time-opt] FAIL: expected failure for colon time literal with support OFF but got """ & rt & """"
    subFail += 1
  else
    dim errTimeLit as String = lcase(Parser_GetLastError())
    if instr(errTimeLit, "unexpected token") > 0 orelse instr(errTimeLit, "invalid numeric") > 0 orelse instr(errTimeLit, "invalid segment") > 0 then
      print "[time-opt] PASS: colon time literal rejected when support is OFF"
      subPass += 1
    else
      print "[time-opt] FAIL: colon literal with support OFF expected parse error, got """ & Parser_GetLastError() & """"
      subFail += 1
    end if
  end if

  if Parser_TryEvaluateEx("5000ms", r, rt, ia) then
    print "[time-opt] FAIL: expected failure for compact time literal 5000ms with support OFF but got """ & rt & """"
    subFail += 1
  else
    dim errCompact as String = lcase(Parser_GetLastError())
    if instr(errCompact, "unknown variable") > 0 orelse instr(errCompact, "unexpected token") > 0 orelse instr(errCompact, "invalid numeric") > 0 then
      print "[time-opt] PASS: compact time literal rejected when support is OFF"
      subPass += 1
    else
      print "[time-opt] FAIL: compact literal with support OFF expected parse error, got """ & Parser_GetLastError() & """"
      subFail += 1
    end if
  end if

  if Parser_TryEvaluateEx("second", r, rt, ia) then
    print "[time-opt] FAIL: expected failure for second constant with support OFF but got """ & rt & """"
    subFail += 1
  else
    dim errSecond as String = lcase(Parser_GetLastError())
    if instr(errSecond, "unknown variable") > 0 then
      print "[time-opt] PASS: second constant unavailable when support is OFF"
      subPass += 1
    else
      print "[time-opt] FAIL: second with support OFF expected unknown variable, got """ & Parser_GetLastError() & """"
      subFail += 1
    end if
  end if

  if Parser_TryEvaluateEx("milliseconds(5)", r, rt, ia) then
    print "[time-opt] FAIL: expected failure for milliseconds(5) with support OFF but got """ & rt & """"
    subFail += 1
  else
    dim errMs as String = lcase(Parser_GetLastError())
    if instr(errMs, "incompatible operands") > 0 then
      print "[time-opt] PASS: milliseconds(5) with support OFF -> incompatible operands"
      subPass += 1
    else
      print "[time-opt] FAIL: milliseconds(5) with support OFF expected incompatible operands, got """ & Parser_GetLastError() & """"
      subFail += 1
    end if
  end if

  Parser_SetSupportTimeValues(TRUE)

  Parser_ClearVariables()
  Parser_SetSupportTimeValues(TRUE)
  if Parser_TryEvaluateEx("_=0:00; rd(t)=ratio(days(t)); rd(1h)", r, rt, ia) = FALSE orelse rt <> "1/24" then
    print "[time-opt] FAIL: UDF duration arg, got """ & rt & """ err=" & Parser_GetLastError()
    subFail += 1
  else
    print "[time-opt] PASS: _=0:00; rd(t)=ratio(days(t)); rd(1h) -> 1/24"
    subPass += 1
  end if

  if Parser_TryEvaluateEx("sin(1:00)", r, rt, ia) then
    print "[time-opt] FAIL: sin(1:00) should fail with incompatible operands but got """ & rt & """"
    subFail += 1
  else
    dim errSinTime as String = lcase(Parser_GetLastError())
    if instr(errSinTime, "incompatible operands") > 0 then
      print "[time-opt] PASS: sin(1:00) still rejects duration argument"
      subPass += 1
    else
      print "[time-opt] FAIL: sin(1:00) expected incompatible operands, got """ & Parser_GetLastError() & """"
      subFail += 1
    end if
  end if

  dim timeAggOk(1 to 12) as String
  dim timeAggExpect(1 to 12) as String
  timeAggOk(1) = "sum(0:30, 1:00)": timeAggExpect(1) = "01:30"
  timeAggOk(2) = "avg(0:30, 1:30)": timeAggExpect(2) = "01:00"
  timeAggOk(3) = "mean(0:20, 1:00)": timeAggExpect(3) = "00:40"
  timeAggOk(4) = "min(1:30, 0:45)": timeAggExpect(4) = "00:45"
  timeAggOk(5) = "max(1:30, 0:45)": timeAggExpect(5) = "01:30"
  timeAggOk(6) = "sum((0:15, 0:15, 0:30))": timeAggExpect(6) = "01:00"
  timeAggOk(7) = "reverse(0:30, Inf, 1:00)": timeAggExpect(7) = "(01:00, inf, 00:30)"
  timeAggOk(8) = "unpack(0:30, Inf)": timeAggExpect(8) = "(00:30, inf)"
  timeAggOk(9) = "unique(0:30, Inf, 0:30)": timeAggExpect(9) = "(00:30, inf)"
  timeAggOk(10) = "ratio((hours(1:00), 15m/1h))": timeAggExpect(10) = "(1/60, 1/4)"
  timeAggOk(11) = "repeat((1:00, 2:00), 2)": timeAggExpect(11) = "(01:00, 02:00, 01:00, 02:00)"
  timeAggOk(12) = "flat(0:30, Inf, 1:00)": timeAggExpect(12) = "(00:30, inf, 01:00)"
  dim tai as Integer
  for tai = 1 to 12
    if Parser_TryEvaluateEx(timeAggOk(tai), r, rt, ia) = FALSE orelse rt <> timeAggExpect(tai) then
      print "[time-opt] FAIL: """ & timeAggOk(tai) & """ -> """ & rt & """ err=" & Parser_GetLastError()
      subFail += 1
    else
      print "[time-opt] PASS: """ & timeAggOk(tai) & """ -> """ & rt & """"
      subPass += 1
    end if
  next tai

  dim timeAggErr(1 to 13) as String
  timeAggErr(1) = "sum(0:30, Inf, 1:00)"
  timeAggErr(2) = "avg(0:30, Inf, 1:30)"
  timeAggErr(3) = "mean(0:20, Inf, 1:00)"
  timeAggErr(4) = "min(0:45, Inf, 1:30)"
  timeAggErr(5) = "max(0:45, Inf, 1:30)"
  timeAggErr(6) = "median(0:30, Inf, 1:00)"
  timeAggErr(7) = "prod(0:30, Inf, 1:00)"
  timeAggErr(8) = "variance(0:30, Inf, 1:00)"
  timeAggErr(9) = "stddev(0:30, Inf, 1:00)"
  timeAggErr(10) = "sort(0:30, Inf, 1:00)"
  timeAggErr(11) = "sort(0:30, -Inf, 1:00)"
  timeAggErr(12) = "sort(0:30, NaN, 1:00)"
  timeAggErr(13) = "keyf(x)=x; sortby((1:30, 2), keyf)"
  dim tei as Integer
  for tei = 1 to 13
    if Parser_TryEvaluateEx(timeAggErr(tei), r, rt, ia) then
      print "[time-opt] FAIL: expected error for """ & timeAggErr(tei) & """ but got """ & rt & """"
      subFail += 1
    else
      dim errTimeAgg as String = lcase(Parser_GetLastError())
      if instr(errTimeAgg, ": time values cannot be mixed with non-time values") > 0 orelse instr(errTimeAgg, "incompatible operands") > 0 then
        print "[time-opt] PASS: """ & timeAggErr(tei) & """ -> " & Parser_GetLastError()
        subPass += 1
      else
        print "[time-opt] FAIL: """ & timeAggErr(tei) & """ expected time/incompatible error, got """ & Parser_GetLastError() & """"
        subFail += 1
      end if
    end if
  next tei

  g_passed += subPass
  g_failed += subFail
  print "Time-option sub-tests: passed " & str(subPass) & ", failed " & str(subFail)
  print ""
end sub

private sub RunLambdaFunctionsSupportOptionTests()
  print "=== Lambda function support (parser-wide option) ==="

  Parser_SetSupportLambdaFunctions(TRUE)

  dim subPass as Integer = 0
  dim subFail as Integer = 0

  if Parser_GetSupportLambdaFunctions() = FALSE then
    print "[lambda-opt] FAIL: expected support flag ON after enabling"
    subFail += 1
  else
    print "[lambda-opt] PASS: getter reports enabled after Parser_SetSupportLambdaFunctions(TRUE)"
    subPass += 1
  end if

  dim r as Double
  dim rt as String
  dim ia as Boolean
  dim isOK as Boolean = TRUE
  if Parser_TryEvaluateEx("sortby((3,1,2), x:-x)", r, rt, ia) = FALSE then
    isOK = FALSE
  elseif rt <> "(3,2,1)" andalso rt <> "(3, 2, 1)" then
    isOK = FALSE
  end if
  if not isOK then
    print "[lambda-opt] FAIL: sortby anonymous lambda with flag ON, got """ & rt & """ err=" & Parser_GetLastError()
    subFail += 1
  else
    print "[lambda-opt] PASS: sortby((3,1,2), x:-x) -> (3,2,1) with support flag ON"
    subPass += 1
  end if

  if Parser_TryEvaluateEx("f=x:x+2; f(3)", r, rt, ia) = FALSE orelse rt <> "5" then
    print "[lambda-opt] FAIL: lambda UDF with flag ON, got """ & rt & """ err=" & Parser_GetLastError()
    subFail += 1
  else
    print "[lambda-opt] PASS: f=x:x+2; f(3) -> 5 with support flag ON"
    subPass += 1
  end if

  Parser_SetSupportLambdaFunctions(FALSE)
  if Parser_GetSupportLambdaFunctions() <> FALSE then
    print "[lambda-opt] FAIL: expected support flag OFF after disable"
    subFail += 1
  else
    print "[lambda-opt] PASS: getter reports disabled after Parser_SetSupportLambdaFunctions(FALSE)"
    subPass += 1
  end if

  if Parser_TryEvaluateEx("sortby((-3,-1,2), abs)", r, rt, ia) = FALSE orelse rt <> "(-1, 2, -3)" then
    print "[lambda-opt] FAIL: sortby builtin ref with flag OFF, got """ & rt & """ err=" & Parser_GetLastError()
    subFail += 1
  else
    print "[lambda-opt] PASS: sortby((-3,-1,2), abs) still works when lambda support is OFF"
    subPass += 1
  end if

  dim rejectExprs(1 to 4) as String
  rejectExprs(1) = "f=x:x+1"
  rejectExprs(2) = "f=(x,y):(x+y)"
  rejectExprs(3) = "sortby((3,1,2), x:-x)"
  rejectExprs(4) = "sortby((1,2), (x):(1/x))"
  dim ri as Integer
  for ri = 1 to 4
    if Parser_TryEvaluateEx(rejectExprs(ri), r, rt, ia) then
      print "[lambda-opt] FAIL: expected failure for """ & rejectExprs(ri) & """ with support OFF but got """ & rt & """"
      subFail += 1
    else
      dim errLam as String = lcase(Parser_GetLastError())
      if instr(errLam, "unexpected token") > 0 then
        print "[lambda-opt] PASS: """ & rejectExprs(ri) & """ rejected when support is OFF"
        subPass += 1
      else
        print "[lambda-opt] FAIL: """ & rejectExprs(ri) & """ with support OFF expected unexpected token, got """ & Parser_GetLastError() & """"
        subFail += 1
      end if
    end if
  next ri

  Parser_SetSupportLambdaFunctions(TRUE)
  Parser_ClearVariables()

  g_passed += subPass
  g_failed += subFail
  print "Lambda-option sub-tests: passed " & str(subPass) & ", failed " & str(subFail)
  print ""
end sub

sub Main()
  dim tests(1 to 1414) as SmokeCase
  ' Inline tag legend:
  ' [spec] = intended language behavior (primary contract)
  ' [regression-lock] = current behavior intentionally locked for compatibility
  ' Default rule: unless explicitly marked [regression-lock], each test is [spec].
  ' [ok-core] [ok-func] [ok-array] [hint] [arity]
  ' [type-int-only] [shape] [shape/broadcast] [syntax] [edge] [overflow]

  ' === SPEC / intended behavior ===
  '
  ' === A) Operator precedence, core operators, and integer-only operator checks ===
  tests(1).expr = "16**-0.5":           tests(1).expected = "0.25" ' [spec][ok-core]
  tests(2).expr = "+5":                 tests(2).expected = "5" ' [spec][ok-core]
  tests(3).expr = "-5":                 tests(3).expected = "-5" ' [spec][ok-core]
  tests(4).expr = "~5":                 tests(4).expected = "-6" ' [spec][ok-core]
  tests(5).expr = "5%3":                tests(5).expected = "2" ' [spec][ok-core]
  tests(6).expr = "200 + 15%":          tests(6).expected = "230" ' [spec][ok-core]
  tests(7).expr = "200 - 15%":          tests(7).expected = "170" ' [spec][ok-core]
  tests(8).expr = "8>>1":               tests(8).expected = "4" ' [spec][ok-core]
  tests(9).expr = "3<<2":              tests(9).expected = "12" ' [spec][ok-core]
  tests(10).expr = "6&3":               tests(10).expected = "2" ' [spec][ok-core]
  tests(11).expr = "6^3":               tests(11).expected = "5" ' [spec][ok-core]
  tests(12).expr = "6|3":               tests(12).expected = "7" ' [spec][ok-core]
  tests(13).expr = "2(3+4)":            tests(13).expected = "14" ' [spec][ok-core]
  tests(14).expr = "2(3+4)**2":         tests(14).expected = "98" ' [spec][ok-core]
  tests(15).expr = "2+3<<1":            tests(15).expected = "10" ' [spec][ok-core]
  tests(16).expr = "1|2^3&6<<1":        tests(16).expected = "3" ' [spec][ok-core]
  tests(17).expr = "2(1+2)%4":          tests(17).expected = "2" ' [spec][ok-core]
  tests(18).expr = "5.5&1":             tests(18).expectedErrContains = "bitwise operands must be integer values" ' [spec][type-int-only]
  tests(19).expr = "5|1.1":             tests(19).expectedErrContains = "bitwise operands must be integer values" ' [spec][type-int-only]
  tests(20).expr = "3.2^1":             tests(20).expectedErrContains = "bitwise operands must be integer values" ' [spec][type-int-only]
  tests(21).expr = "8.1>>1":            tests(21).expectedErrContains = "bitwise operands must be integer values" ' [spec][type-int-only]
  tests(22).expr = "8<<1.2":            tests(22).expectedErrContains = "bitwise operands must be integer values" ' [spec][type-int-only]
  tests(23).expr = "~2.5":              tests(23).expectedErrContains = "bitwise operands must be integer values" ' [spec][type-int-only]
  tests(24).expr = "5.5%2":             tests(24).expectedErrContains = "modulo operands must be integer values" ' [spec][type-int-only]
  tests(25).expr = "5%2.2":             tests(25).expectedErrContains = "modulo operands must be integer values" ' [spec][type-int-only]
  tests(26).expr = "2(1+2.5)%4.2":      tests(26).expectedErrContains = "modulo operands must be integer values" ' [spec][type-int-only]

  ' === B) [spec] Function hints, comments, parser diagnostics, and literal parsing ===
  tests(27).expr = "pow(2,3)":          tests(27).expected = "8" ' [ok-func]
  tests(28).expr = "prod(2,3,4)":       tests(28).expected = "24" ' [ok-func]
  tests(29).expr = "pow":               tests(29).expectedErrContains = "function: pow(" ' [hint]
  tests(30).expr = "sin":               tests(30).expectedErrContains = "function: sin(angle)" ' [hint]
  tests(31).expr = "sum":               tests(31).expectedErrContains = "function: sum(...)" ' [hint]
  tests(32).expr = "sqr(5)":            tests(32).expected = "25" ' [ok-func]
  tests(33).expr = "sqr":               tests(33).expectedErrContains = "function: sqr(value)" ' [hint]
  tests(34).expr = "1 + 2 # calculates 1 + 2": tests(34).expected = "3" ' [syntax]
  tests(35).expr = "// this entire line is a comment": tests(35).expectNoResult = TRUE ' [syntax]
  tests(36).expr = "sin(pi/2) // calculates sin(pi/2)": tests(36).expected = "1" ' [syntax]
  tests(37).expr = "[]":                tests(37).expectedErrContains = "unexpected token" ' [syntax]
  tests(38).expr = "b=2; 2b":           tests(38).expectedErrContains = "unexpected token" ' [syntax]
  tests(39).expr = "(2;2b;3)":          tests(39).expectedErrContains = "missing closing parenthesis" ' [syntax]
  tests(40).expr = "(2,2b,3)":          tests(40).expectedErrContains = "unexpected token" ' [syntax]
  tests(41).expr = "hex(12)":           tests(41).expected = "0xC" ' [ok-func]
  tests(42).expr = "hex((12,255))":     tests(42).expected = "(0xC,0xFF)" ' [ok-array]
  tests(43).expr = "10 + hex(12) + 14": tests(43).expected = "36" ' [ok-func]
  tests(44).expr = "hex(12.5)":         tests(44).expectedErrContains = "hex() expects integer values" ' [type-int-only]
  tests(45).expr = "hex":               tests(45).expectedErrContains = "function: hex(...)" ' [hint]
  tests(46).expr = "0x":                tests(46).expectedErrContains = "invalid hex literal" ' [syntax]
  tests(47).expr = "0xG":               tests(47).expectedErrContains = "invalid hex literal" ' [syntax]
  tests(48).expr = "hex(0x7FFFFFFFFFFFFFFF)": tests(48).expected = "0x7FFFFFFFFFFFFFFF" ' [ok-func]
  tests(49).expr = "hex(0xFFFFFFFFFFFFFFFF)": tests(49).expected = "0xFFFFFFFFFFFFFFFF" ' [ok-func]
  tests(50).expr = "0b01110110011":     tests(50).expected = "947" ' [ok-core]
  tests(51).expr = "bin(13)":           tests(51).expected = "0b1101" ' [ok-func]
  tests(52).expr = "bin((1,2,5))":      tests(52).expected = "(0b1,0b10,0b101)" ' [ok-array]
  tests(53).expr = "10 + bin(12) + 14": tests(53).expected = "36" ' [ok-func]
  tests(54).expr = "0b":                tests(54).expectedErrContains = "invalid binary literal" ' [syntax]
  tests(55).expr = "bin(12.5)":         tests(55).expectedErrContains = "bin() expects integer values" ' [type-int-only]
  tests(56).expr = "bin":               tests(56).expectedErrContains = "function: bin(...)" ' [hint]

  ' === C) [spec] Integer-accuracy / overflow-path regression cases ===
  tests(57).expr = "9007199254740991+1": tests(57).expected = "9007199254740992" ' [overflow]
  tests(58).expr = "9007199254740992+1": tests(58).expected = "9007199254740993" ' [overflow]
  tests(59).expr = "3037000499*3037000499": tests(59).expected = "9223372030926249001" ' [overflow]
  tests(60).expr = "5/2":               tests(60).expected = "2.5" ' [edge]
  tests(61).expr = "2**10+1":           tests(61).expected = "1025" ' [ok-core]
  tests(62).expr = "2**-1":             tests(62).expected = "0.5" ' [edge]
  tests(63).expr = "9007199254740993&1": tests(63).expected = "1" ' [overflow]
  tests(64).expr = "9223372036854775807+1": tests(64).expected = "9223372036854775808" ' [overflow]
  tests(65).expr = "-9223372036854775808-1": tests(65).expected = "-9.223372036854778e+018" ' [overflow]
  tests(66).expr = "3037000500*3037000500": tests(66).expected = "9.223372037000249e+018" ' [overflow]
  tests(67).expr = "2**63":             tests(67).expected = "9223372036854775808" ' [overflow]
  tests(68).expr = "2**64":             tests(68).expected = "1.844674407370955e+019" ' [overflow]
  tests(69).expr = "9223372036854775807+0.5": tests(69).expected = "9.223372036854778e+018" ' [overflow]
  tests(70).expr = "hex(9223372036854775807+1)": tests(70).expected = "0x8000000000000000" ' [overflow]

  ' === D) [spec] Built-ins and ans variable baseline behavior ===
  tests(71).expr = "log(8,2)":          tests(71).expected = "3" ' [ok-func]
  tests(72).expr = "log(100,10)":       tests(72).expected = "2" ' [ok-func]
  tests(73).expr = "log(8)":            tests(73).expectedErrContains = "log() expects 2 argument(s)" ' [arity]
  tests(74).expr = "ln":                tests(74).expectedErrContains = "function: ln(value)" ' [hint]
  tests(75).expr = "a=2;sum(a,a)":      tests(75).expected = "4" ' [ok-func]
  tests(76).expr = "f(a,a)=a":          tests(76).expectedErrContains = "duplicate parameter name" ' [syntax]
  tests(77).expr = "2+3;ans":           tests(77).expected = "5" ' [ok-func]
  tests(78).expr = "(1,2,3);ans":       tests(78).expected = "(1,2,3)" ' [ok-array]
  tests(79).expr = "hex(15);ans":       tests(79).expected = "0xF" ' [ok-func]
  tests(80).expr = "7; ans*2":          tests(80).expected = "14" ' [ok-func]
  tests(81).expr = "v=(10,20);sum(ans)": tests(81).expected = "30" ' [ok-array]
  tests(82).expr = "atan2(1,1)":        tests(82).expected = "0.7853981633974483" ' [ok-func]
  tests(83).expr = "floor(2.9)":        tests(83).expected = "2" ' [ok-func]
  tests(84).expr = "ceil(2.1)":         tests(84).expected = "3" ' [ok-func]
  tests(85).expr = "trunc(-2.9)":       tests(85).expected = "-2" ' [ok-func]
  tests(86).expr = "round(2.5)":        tests(86).expected = "3" ' [ok-func]
  tests(87).expr = "sign(-123)":        tests(87).expected = "-1" ' [ok-func]
  tests(88).expr = "mod(17,5)":         tests(88).expected = "2" ' [ok-func]
  tests(89).expr = "avg(1,2,3,4)":      tests(89).expected = "2.5" ' [ok-func]
  tests(90).expr = "mean((1,2,3),9)":   tests(90).expected = "3.75" ' [ok-array]
  tests(91).expr = "clamp(15,0,10)":    tests(91).expected = "10" ' [ok-func]
  tests(92).expr = "deg(pi)":           tests(92).expected = "180" ' [ok-func]
  tests(93).expr = "rad(180)":          tests(93).expected = "3.141592653589793" ' [ok-func]
  tests(94).expr = "hypot(3,4)":        tests(94).expected = "5" ' [ok-func]
  tests(95).expr = "gcd(84,30)":        tests(95).expected = "6" ' [ok-func]
  tests(96).expr = "lcm(6,8)":          tests(96).expected = "24" ' [ok-func]
  tests(97).expr = "median(1,8,3)":     tests(97).expected = "3" ' [ok-func]
  tests(98).expr = "median((1,9),3,7)": tests(98).expected = "5" ' [ok-array]
  tests(99).expr = "variance(1,2,3)":  tests(99).expected = "0.6666666666666666" ' [ok-func]
  tests(100).expr = "stddev(1,2,3)":    tests(100).expected = "0.816496580927726" ' [ok-func]
  tests(101).expr = "fact(0)":          tests(101).expected = "1" ' [ok-func]
  tests(102).expr = "fact(5)":          tests(102).expected = "120" ' [ok-func]
  tests(103).expr = "factorial(10)":    tests(103).expected = "3628800" ' [ok-func]
  tests(104).expr = "fact(-1)":         tests(104).expectedErrContains = "fact() expects a non-negative integer" ' [edge]
  tests(105).expr = "fact(2.5)":        tests(105).expectedErrContains = "fact() expects integer values" ' [type-int-only]
  tests(106).expr = "factorial(21)":    tests(106).expected = "5.109094217170944e+019" ' [float-over-20]
  tests(107).expr = "random(5,5)":      tests(107).expected = "5" ' [edge]
  tests(108).expr = "rand(1)":          tests(108).expectedErrContains = "rand() expects 0 argument(s)" ' [arity]
  tests(109).expr = "rand":             tests(109).expectedErrContains = "function: rand()" ' [hint]
  tests(110).expr = "random":           tests(110).expectedErrContains = "function: random(min, max)" ' [hint]
  tests(111).expr = "median":           tests(111).expectedErrContains = "function: median(...)" ' [hint]

  ' === E) [spec] sort/unique baseline behavior ===
  tests(112).expr = "sort((3,1,2))":    tests(112).expected = "(1,2,3)" ' [ok-array]
  tests(113).expr = "a=(5,2,9);sort(a)": tests(113).expected = "(2,5,9)" ' [ok-array]
  tests(114).expr = "sort(5)":          tests(114).expected = "(5)" ' [ok-func]
  tests(115).expr = "sort(2,5,1)":      tests(115).expected = "(1,2,5)" ' [ok-func]
  tests(116).expr = "sort":             tests(116).expectedErrContains = "function: sort(...)" ' [hint]
  tests(117).expr = "unique((3,1,3,2,1,2))": tests(117).expected = "(3,1,2)" ' [ok-array]
  tests(118).expr = "a=(5,2,5,9,2);unique(a)": tests(118).expected = "(5,2,9)" ' [ok-array]
  tests(119).expr = "unique(5)":        tests(119).expected = "(5)" ' [ok-func]
  tests(120).expr = "unique(1,2,1,2,3)": tests(120).expected = "(1,2,3)" ' [ok-func]
  tests(121).expr = "unique":           tests(121).expectedErrContains = "function: unique(...)" ' [hint]

  ' === F) [spec] Stress matrix: argument shape, arity, syntax, and edge-case validation ===
  tests(122).expr = "(1,2)+(3)":        tests(122).expected = "(4,5)" ' [shape/broadcast]
  tests(123).expr = "(1,2)*(3,4,5)":    tests(123).expectedErrContains = "incompatible operands" ' [shape]
  tests(124).expr = "1<<64":            tests(124).expectedErrContains = "incompatible operands" ' [edge]
  tests(125).expr = "1>>-1":            tests(125).expectedErrContains = "incompatible operands" ' [edge]
  tests(126).expr = "5%0":              tests(126).expectedErrContains = "incompatible operands" ' [edge]
  tests(127).expr = "1+":               tests(127).expectedErrContains = "unexpected token" ' [syntax]
  tests(128).expr = "pow()":            tests(128).expectedErrContains = "expects 2 argument(s)" ' [arity]
  tests(129).expr = "pow(2,3,4)":       tests(129).expectedErrContains = "expects 2 argument(s)" ' [arity]
  tests(130).expr = "log()":            tests(130).expectedErrContains = "expects 2 argument(s)" ' [arity]
  tests(131).expr = "log(10,10,10)":    tests(131).expectedErrContains = "expects 2 argument(s)" ' [arity]
  tests(132).expr = "atan2()":          tests(132).expectedErrContains = "expects 2 argument(s)" ' [arity]
  tests(133).expr = "atan2(1,2,3)":     tests(133).expectedErrContains = "expects 2 argument(s)" ' [arity]
tests(134).expr = "atan2((1,2),3)":   tests(134).expected = "(0.3217505543966422,0.5880026035475675)" ' [shape]
  tests(135).expr = "hypot()":          tests(135).expectedErrContains = "expects 2 argument(s)" ' [arity]
  tests(136).expr = "hypot(3,4,5)":     tests(136).expectedErrContains = "expects 2 argument(s)" ' [arity]
  tests(137).expr = "mod()":            tests(137).expectedErrContains = "expects 2 argument(s)" ' [arity]
  tests(138).expr = "mod(10,3,1)":      tests(138).expectedErrContains = "expects 2 argument(s)" ' [arity]
  tests(139).expr = "mod(10,0)":        tests(139).expectedErrContains = "numeric error in mod()" ' [edge]
  tests(140).expr = "gcd()":            tests(140).expectedErrContains = "expects 2 argument(s)" ' [arity]
  tests(141).expr = "gcd(6,8,10)":      tests(141).expectedErrContains = "expects 2 argument(s)" ' [arity]
  tests(142).expr = "gcd(6.5,3)":       tests(142).expectedErrContains = "expects integer values" ' [type-int-only]
  tests(143).expr = "lcm(6,0)":         tests(143).expected = "0" ' [edge]
  tests(144).expr = "lcm(6.5,3)":       tests(144).expectedErrContains = "expects integer values" ' [type-int-only]
  tests(145).expr = "hex(1,2)":         tests(145).expected = "(0x1,0x2)" ' [ok-array]
  tests(146).expr = "bin()":            tests(146).expectedErrContains = "expects at least 1 argument" ' [arity]
  tests(147).expr = "bin(1,2)":         tests(147).expected = "(0b1,0b10)" ' [ok-array]
  tests(148).expr = "clamp()":          tests(148).expectedErrContains = "expects 3 argument(s)" ' [arity]
  tests(149).expr = "rand()*0":         tests(149).expected = "0" ' [edge]
  tests(150).expr = "random()":         tests(150).expectedErrContains = "expects 2 argument(s)" ' [arity]
  tests(151).expr = "random(1,2,3)":    tests(151).expectedErrContains = "expects 2 argument(s)" ' [arity]
  tests(152).expr = "sort()":           tests(152).expectedErrContains = "expects at least 1 argument" ' [arity]
  tests(153).expr = "sort((2,5,1),4,3)": tests(153).expected = "(1,2,3,4,5)" ' [ok-array]
  tests(154).expr = "unique()":         tests(154).expectedErrContains = "expects at least 1 argument" ' [arity]
  tests(155).expr = "product()":        tests(155).expectedErrContains = "expects at least 1 argument" ' [arity]
  tests(156).expr = "min()":            tests(156).expectedErrContains = "expects at least 1 argument" ' [arity]
  tests(157).expr = "max()":            tests(157).expectedErrContains = "expects at least 1 argument" ' [arity]
  tests(158).expr = "avg()":            tests(158).expectedErrContains = "expects at least 1 argument" ' [arity]
  tests(159).expr = "median()":         tests(159).expectedErrContains = "expects at least 1 argument" ' [arity]
  tests(160).expr = "variance()":       tests(160).expectedErrContains = "expects at least 1 argument" ' [arity]
  tests(161).expr = "stddev()":         tests(161).expectedErrContains = "expects at least 1 argument" ' [arity]
  tests(162).expr = "sin(1,2)":         tests(162).expectedErrContains = "expects 1 argument(s)" ' [arity]
  tests(163).expr = "cos(1,2)":         tests(163).expectedErrContains = "expects 1 argument(s)" ' [arity]
  tests(164).expr = "ln(1,2)":          tests(164).expectedErrContains = "expects 1 argument(s)" ' [arity]
  tests(165).expr = "sqrt(1,2)":        tests(165).expectedErrContains = "expects 1 argument(s)" ' [arity]
  tests(166).expr = "fact()":           tests(166).expectedErrContains = "expects 1 argument(s)" ' [arity]
  tests(167).expr = "factorial(1,2)":   tests(167).expectedErrContains = "expects 1 argument(s)" ' [arity]
  tests(168).expr = "pow(2,)":          tests(168).expectedErrContains = "unexpected comma" ' [syntax]
  tests(169).expr = "sum((1,2),)":      tests(169).expectedErrContains = "unexpected comma" ' [syntax]

  ' === G) [spec] Extended negative matrix: malformed syntax, aliases, and mismatch paths ===
  tests(170).expr = "1/":               tests(170).expectedErrContains = "unexpected token" ' [syntax]
  tests(171).expr = "1**":              tests(171).expectedErrContains = "unexpected token" ' [syntax]
  tests(172).expr = "1<<":              tests(172).expectedErrContains = "unexpected token" ' [syntax]
  tests(173).expr = "1>>":              tests(173).expectedErrContains = "unexpected token" ' [syntax]
  tests(174).expr = "1&":               tests(174).expectedErrContains = "unexpected token" ' [syntax]
  tests(175).expr = "1|":               tests(175).expectedErrContains = "unexpected token" ' [syntax]
  tests(176).expr = "1^":               tests(176).expectedErrContains = "unexpected token" ' [syntax]
  tests(177).expr = "1%":               tests(177).expected = "0.01" ' [edge]
  tests(178).expr = "pow(,2)":          tests(178).expectedErrContains = "unexpected" ' [syntax]
  tests(179).expr = "atan2(,2)":        tests(179).expectedErrContains = "unexpected" ' [syntax]
  tests(180).expr = "random(,2)":       tests(180).expectedErrContains = "unexpected" ' [syntax]
  tests(181).expr = "clamp(1,,3)":      tests(181).expectedErrContains = "unexpected" ' [syntax]
  tests(182).expr = "sum(,1)":          tests(182).expectedErrContains = "unexpected" ' [syntax]
  tests(183).expr = "sin()":            tests(183).expectedErrContains = "expects 1 argument(s)" ' [arity]
  tests(184).expr = "tan()":            tests(184).expectedErrContains = "expects 1 argument(s)" ' [arity]
  tests(185).expr = "asin()":           tests(185).expectedErrContains = "expects 1 argument(s)" ' [arity]
  tests(186).expr = "acos()":           tests(186).expectedErrContains = "expects 1 argument(s)" ' [arity]
  tests(187).expr = "atan()":           tests(187).expectedErrContains = "expects 1 argument(s)" ' [arity]
  tests(188).expr = "sinh()":           tests(188).expectedErrContains = "expects 1 argument(s)" ' [arity]
  tests(189).expr = "cosh()":           tests(189).expectedErrContains = "expects 1 argument(s)" ' [arity]
  tests(190).expr = "tanh()":           tests(190).expectedErrContains = "expects 1 argument(s)" ' [arity]
  tests(191).expr = "exp()":            tests(191).expectedErrContains = "expects 1 argument(s)" ' [arity]
  tests(192).expr = "log10()":          tests(192).expectedErrContains = "expects 1 argument(s)" ' [arity]
  tests(193).expr = "abs()":            tests(193).expectedErrContains = "expects 1 argument(s)" ' [arity]
  tests(194).expr = "floor()":          tests(194).expectedErrContains = "expects 1 argument(s)" ' [arity]
  tests(195).expr = "ceil()":           tests(195).expectedErrContains = "expects 1 argument(s)" ' [arity]
  tests(196).expr = "trunc()":          tests(196).expectedErrContains = "expects 1 argument(s)" ' [arity]
  tests(197).expr = "round()":          tests(197).expectedErrContains = "expects 1 argument(s)" ' [arity]
  tests(198).expr = "sign()":           tests(198).expectedErrContains = "expects 1 argument(s)" ' [arity]
  tests(199).expr = "ln((1,2))":        tests(199).expected = "(0,0.6931471805599453)" ' [ok-array]
  tests(200).expr = "sqrt((1,2))":      tests(200).expected = "(1,1.414213562373095)" ' [ok-array]
  tests(201).expr = "abs((1,2))":       tests(201).expected = "(1,2)" ' [ok-array]
  tests(202).expr = "arcsin(1)":        tests(202).expected = "1.570796326794897" ' [ok-func]
  tests(203).expr = "arccos(1)":        tests(203).expected = "0" ' [ok-func]
  tests(204).expr = "arctan(1)":        tests(204).expected = "0.7853981633974483" ' [ok-func]
  tests(205).expr = "prod()":           tests(205).expectedErrContains = "expects at least 1 argument" ' [arity]
  tests(206).expr = "mean()":           tests(206).expectedErrContains = "expects at least 1 argument" ' [arity]
  tests(207).expr = "variance(())":     tests(207).expectedErrContains = "expects at least 1 argument" ' [syntax]
  tests(208).expr = "stddev(())":       tests(208).expectedErrContains = "expects at least 1 argument" ' [syntax]
  tests(209).expr = "gcd(1)":           tests(209).expectedErrContains = "expects 2 argument(s)" ' [arity]
  tests(210).expr = "lcm(1)":           tests(210).expectedErrContains = "expects 2 argument(s)" ' [arity]
  tests(211).expr = "mod(1)":           tests(211).expectedErrContains = "expects 2 argument(s)" ' [arity]
  tests(212).expr = "hypot(1)":         tests(212).expectedErrContains = "expects 2 argument(s)" ' [arity]
  tests(213).expr = "atan2(1)":         tests(213).expectedErrContains = "expects 2 argument(s)" ' [arity]
  tests(214).expr = "log((1,2),10)":    tests(214).expected = "(0,0.3010299956639812)" ' [ok-array]
  tests(215).expr = "log(10,(2,3))":    tests(215).expected = "(3.321928094887362,2.095903274289384)" ' [ok-array]
  tests(216).expr = "pow((2,3),2)":     tests(216).expected = "(4,9)" ' [ok-array]
  tests(217).expr = "pow(2,(2,3))":     tests(217).expected = "(4,8)" ' [ok-array]
  tests(218).expr = "hex((1,2,3),(4))": tests(218).expected = "(0x1,0x2,0x3,0x4)" ' [ok-array]
  tests(219).expr = "bin((1,2,3),(4))": tests(219).expected = "(0b1,0b10,0b11,0b100)" ' [ok-array]
  tests(220).expr = "random(10,10)":    tests(220).expected = "10" ' [edge]
  tests(221).expr = "random(1.5,1.5)":  tests(221).expected = "1.5" ' [edge]
  tests(222).expr = "random(3.5,3.5)":  tests(222).expected = "3.5" ' [edge]
  tests(223).expr = "fact((1,2))":      tests(223).expectedErrContains = "expects a non-negative integer" ' [shape]
  tests(224).expr = "factorial((1,2))": tests(224).expectedErrContains = "expects a non-negative integer" ' [shape]
  '
  ' === REGRESSION-LOCK / compatibility behavior ===
  ' These cases intentionally lock currently observed behavior that may look odd,
  ' but should not change accidentally without an explicit decision.
  tests(225).expr = "clamp((1,2,3),(4,5),6)": tests(225).expectedErrContains = "expects scalar min/max" ' [type]
  tests(226).expr = "sum((1,2),(3,4),5)": tests(226).expected = "15" ' [ok-array]
  tests(227).expr = "sort(())":         tests(227).expectedErrContains = "expects at least 1 argument" ' [syntax]
  tests(228).expr = "unique(())":       tests(228).expectedErrContains = "expects at least 1 argument" ' [syntax]
  tests(229).expr = "RestoreAnsFromCachedRender(g_cachedRenderText(i))": tests(229).expectedErrContains = "unknown function" ' [regression-lock]
  tests(230).expr = "deg(pi/2,pi/4)":   tests(230).expected = "(90,45)" ' [ok-array]
  tests(231).expr = "rad(180,90)":      tests(231).expected = "(3.141592653589793,1.570796326794897)" ' [ok-array]
  tests(232).expr = "mean":             tests(232).expectedErrContains = "function: mean(...)" ' [hint]
  tests(233).expr = "floor":            tests(233).expectedErrContains = "function: floor(value)" ' [hint]
  tests(234).expr = "ceil":             tests(234).expectedErrContains = "function: ceil(value)" ' [hint]
  tests(235).expr = "trunc":            tests(235).expectedErrContains = "function: trunc(value)" ' [hint]
  tests(236).expr = "round":            tests(236).expectedErrContains = "function: round(value)" ' [hint]
  tests(237).expr = "sign":             tests(237).expectedErrContains = "function: sign(value)" ' [hint]
  tests(238).expr = "deg":              tests(238).expectedErrContains = "function: deg(...)" ' [hint]
  tests(239).expr = "rad":              tests(239).expectedErrContains = "function: rad(...)" ' [hint]
  tests(240).expr = "int(2.9)":         tests(240).expected = "2" ' [ok-func]
  tests(241).expr = "int(-2.9)":        tests(241).expected = "-2" ' [ok-func]
  tests(242).expr = "frac(2.9)":        tests(242).expected = "0.8999999999999999" ' [ok-func]
  tests(243).expr = "frac(-2.9)":       tests(243).expected = "-0.8999999999999999" ' [ok-func]
  tests(244).expr = "int((2.9,-2.9))":  tests(244).expected = "(2,-2)" ' [ok-array]
  tests(245).expr = "frac((2.9,-2.9))": tests(245).expected = "(0.8999999999999999,-0.8999999999999999)" ' [ok-array]
  tests(246).expr = "int":              tests(246).expectedErrContains = "function: int(value)" ' [hint]
  tests(247).expr = "frac":             tests(247).expectedErrContains = "function: frac(value)" ' [hint]
  tests(248).expr = "fract(2.9)":       tests(248).expected = "0.8999999999999999" ' [ok-func]
  tests(249).expr = "fract((2.9,-2.9))": tests(249).expected = "(0.8999999999999999,-0.8999999999999999)" ' [ok-array]
  tests(250).expr = "fract":            tests(250).expectedErrContains = "function: frac(value)" ' [hint]
  tests(251).expr = "oct(12)":          tests(251).expected = "0o14" ' [ok-func]
  tests(252).expr = "oct((12,255))":    tests(252).expected = "(0o14,0o377)" ' [ok-array]
  tests(253).expr = "10 + oct(12) + 14": tests(253).expected = "36" ' [ok-func]
  tests(254).expr = "oct(12.5)":        tests(254).expectedErrContains = "oct() expects integer values" ' [type-int-only]
  tests(255).expr = "oct":              tests(255).expectedErrContains = "function: oct(...)" ' [hint]
  tests(256).expr = "oct()":            tests(256).expectedErrContains = "expects at least 1 argument" ' [arity]
  tests(257).expr = "oct(1,2)":         tests(257).expected = "(0o1,0o2)" ' [ok-array]
  tests(258).expr = "oct((1,2,3),(4))": tests(258).expected = "(0o1,0o2,0o3,0o4)" ' [ok-array]
  tests(259).expr = "oct(15);ans":      tests(259).expected = "0o17" ' [ok-func]
  tests(260).expr = "oct(9223372036854775807+1)": tests(260).expected = "0o1000000000000000000000" ' [overflow]
  tests(261).expr = "0O77":             tests(261).expected = "63" ' [ok-core]
  tests(262).expr = "0o123 + 1":        tests(262).expected = "84" ' [ok-core]
  tests(263).expr = "0o20 & 0xF":       tests(263).expected = "0" ' [ok-core]
  tests(264).expr = "oct((0o7,0o10))":  tests(264).expected = "(0o7,0o10)" ' [ok-array]
  tests(265).expr = "0o64":             tests(265).expected = "52" ' [ok-core]
  tests(266).expr = "0o":               tests(266).expectedErrContains = "invalid octal literal" ' [syntax]
  tests(267).expr = "oct(0o17)":        tests(267).expected = "0o17" ' [ok-func]
  tests(268).expr = "0o10 + 8":         tests(268).expected = "16" ' [ok-core]
  tests(269).expr = "0b110011 & 0x37 | 0o64": tests(269).expected = "55" ' [ok-core]
  tests(270).expr = "0o8":              tests(270).expectedErrContains = "invalid octal literal" ' [syntax]
  tests(271).expr = "5=5":              tests(271).expected = "1" ' [ok-core]
  tests(272).expr = "5==4":             tests(272).expected = "0" ' [ok-core]
  tests(273).expr = "5<>4":             tests(273).expected = "1" ' [ok-core]
  tests(274).expr = "5!=5":             tests(274).expected = "0" ' [ok-core]
  tests(275).expr = "5>4":              tests(275).expected = "1" ' [ok-core]
  tests(276).expr = "5>=5":             tests(276).expected = "1" ' [ok-core]
  tests(277).expr = "4<5":              tests(277).expected = "1" ' [ok-core]
  tests(278).expr = "4<=4":             tests(278).expected = "1" ' [ok-core]
  tests(279).expr = "1|2=3":            tests(279).expected = "1" ' [ok-core]
  tests(280).expr = "1|2<2":            tests(280).expected = "0" ' [ok-core]
  tests(281).expr = "(1,2,3)=(1,2,3)":  tests(281).expected = "1" ' [ok-array]
  tests(282).expr = "(1,2,3)!=(1,2,4)": tests(282).expected = "1" ' [ok-array]
  tests(283).expr = "(1,2)<(1,2,0)":    tests(283).expected = "1" ' [ok-array]
  tests(284).expr = "(1,2,9)>(1,2,3)":  tests(284).expected = "1" ' [ok-array]
  tests(285).expr = "(1,2,3)<=(1,2,3)": tests(285).expected = "1" ' [ok-array]
  tests(286).expr = "(1)<(1,0)":        tests(286).expected = "1" ' [ok-array]
  tests(287).expr = "(1,0)>(1)":        tests(287).expected = "1" ' [ok-array]
  tests(288).expr = "5<(5,1)":          tests(288).expected = "1" ' [ok-array]
  tests(289).expr = "(5,1)>5":          tests(289).expected = "1" ' [ok-array]
  tests(290).expr = "2<3<4":            tests(290).expected = "1" ' [ok-core]
  tests(291).expr = "!0":               tests(291).expected = "1" ' [ok-core]
  tests(292).expr = "!5":               tests(292).expected = "0" ' [ok-core]
  tests(293).expr = "not 0":            tests(293).expected = "1" ' [ok-core]
  tests(294).expr = "not 2":            tests(294).expected = "0" ' [ok-core]
  tests(295).expr = "1&&1":             tests(295).expected = "1" ' [ok-core]
  tests(296).expr = "1&&0":             tests(296).expected = "0" ' [ok-core]
  tests(297).expr = "1||0":             tests(297).expected = "1" ' [ok-core]
  tests(298).expr = "0||0":             tests(298).expected = "0" ' [ok-core]
  tests(299).expr = "1 and 1":          tests(299).expected = "1" ' [ok-core]
  tests(300).expr = "1 and 0":          tests(300).expected = "0" ' [ok-core]
  tests(301).expr = "1 or 0":           tests(301).expected = "1" ' [ok-core]
  tests(302).expr = "0 or 0":           tests(302).expected = "0" ' [ok-core]
  tests(303).expr = "!1=0":             tests(303).expected = "1" ' [ok-core]
  tests(304).expr = "1=1 && 0=1":       tests(304).expected = "0" ' [ok-core]
  tests(305).expr = "1|2==3 && 5>3":    tests(305).expected = "1" ' [ok-core]
  tests(306).expr = "1 || 0 && 0":      tests(306).expected = "1" ' [ok-core]
  tests(307).expr = "(0,0) && 1":       tests(307).expected = "1" ' [ok-array]
  tests(308).expr = "(0,0) || 0":       tests(308).expected = "1" ' [ok-array]
  tests(309).expr = "not (0,0)":        tests(309).expected = "0" ' [ok-array]
  tests(310).expr = "0 or (0,0)":       tests(310).expected = "1" ' [ok-array]
  tests(311).expr = "not 1<0":          tests(311).expected = "1" ' [ok-core]
  tests(312).expr = "!1<0":             tests(312).expected = "0" ' [ok-core]
  tests(313).expr = "reverse((3,1,2))": tests(313).expected = "(2,1,3)" ' [ok-array]
  tests(314).expr = "reverse(2,5,1)":   tests(314).expected = "(1,5,2)" ' [ok-func]
  tests(315).expr = "reverse((1,2,3),(4,5,6),(7,8,9))": tests(315).expected = "(9,8,7,6,5,4,3,2,1)" ' [ok-array]
  tests(316).expr = "reverse(5)":       tests(316).expected = "(5)" ' [ok-func]
  tests(317).expr = "reverse":          tests(317).expectedErrContains = "function: reverse(...)" ' [hint]
  tests(318).expr = "reverse()":        tests(318).expectedErrContains = "expects at least 1 argument" ' [arity]
  tests(319).expr = "reverse((2,5,1),4,3)": tests(319).expected = "(3,4,1,5,2)" ' [ok-array]
  tests(320).expr = "reverse(())":      tests(320).expectedErrContains = "expects at least 1 argument" ' [syntax]
  tests(321).expr = "(10,20,30)[0]":    tests(321).expected = "10" ' [ok-array]
  tests(322).expr = "(10,20,30)[2]":    tests(322).expected = "30" ' [ok-array]
  tests(323).expr = "(10,20,30)[-1]":   tests(323).expected = "30" ' [ok-array]
  tests(324).expr = "(10,20,30)[-2]":   tests(324).expected = "20" ' [ok-array]
  tests(325).expr = "(10,20,30)[-3]":   tests(325).expected = "10" ' [ok-array]
  tests(326).expr = "(10,20,30)[-4]":   tests(326).expectedErrContains = "array index is out of range" ' [edge]
  tests(327).expr = "(10,20,30)[3]":    tests(327).expectedErrContains = "array index is out of range" ' [edge]
  tests(328).expr = "sort((3,1,2,4))[-1]": tests(328).expected = "4" ' [ok-array]
  tests(329).expr = "reverse((1,2,3,4))[-1]": tests(329).expected = "1" ' [ok-array]
  tests(330).expr = "reverse((1,2,3,4))[0]": tests(330).expected = "4" ' [ok-array]
  tests(331).expr = "sorted((3,1,2))":  tests(331).expected = "(1,2,3)" ' [ok-array]
  tests(332).expr = "sorted(2,5,1)":    tests(332).expected = "(1,2,5)" ' [ok-func]
  tests(333).expr = "sorted":           tests(333).expectedErrContains = "function: sort(...)" ' [hint]
  tests(334).expr = "reversed((1,2,3),(4,5))": tests(334).expected = "(5,4,3,2,1)" ' [ok-array]
  tests(335).expr = "reversed":         tests(335).expectedErrContains = "function: reverse(...)" ' [hint]
  tests(336).expr = "reversed((1,2,3,4))[-1]": tests(336).expected = "1" ' [ok-array]
  tests(337).expr = "3 + not 4":        tests(337).expected = "3" ' [ok-core]
  tests(338).expr = "3 + not 4 + 5":    tests(338).expected = "3" ' [ok-core]
  tests(339).expr = "oct(80)":          tests(339).expected = "0o120" ' [ok-func]
  tests(340).expr = "oct(0)":           tests(340).expected = "0o0" ' [edge]
  tests(341).expr = "oct(-1)":          tests(341).expected = "-0o1" ' [edge]
  tests(342).expr = "sin(x)=x":         tests(342).expectedErrContains = "reserved function name" ' [syntax]
  tests(343).expr = "oct(x)=x":         tests(343).expectedErrContains = "reserved function name" ' [syntax]
  tests(344).expr = "not(x)=x":         tests(344).expectedErrContains = "reserved function name" ' [syntax]
  tests(345).expr = "f(x,y)=x*y; a=(2,3); f(unpack(a))": tests(345).expected = "6" ' [ok-func]
  tests(346).expr = "f(x,y,z)=x+y+z; f(unpack((1,2,3)))": tests(346).expected = "6" ' [ok-array]
  tests(347).expr = "unpack((1,2,3))":  tests(347).expected = "(1,2,3)" ' [ok-array]
  tests(348).expr = "unpack(5)":        tests(348).expected = "5" ' [ok-func]
  tests(349).expr = "unpack":           tests(349).expectedErrContains = "function: unpack(...)" ' [hint]
  tests(350).expr = "sum(unpack((1,2,3)))": tests(350).expected = "6" ' [ok-func]
  tests(351).expr = "f(x,y)=x*y; f(unpack((2,3,4)))": tests(351).expectedErrContains = "expects 2 argument(s)" ' [arity]
  tests(352).expr = "f(x,y,z)=x+y+z; f(unpack(1,2,3))": tests(352).expected = "6" ' [ok-func]
  tests(353).expr = "f(a,b,c,d,t)=a+b+c+d+t; f(unpack((1,2),3,(4,5)))": tests(353).expected = "15" ' [ok-array]
  tests(354).expr = "unpack((1,2),3,(4,5))": tests(354).expected = "(1,2,3,4,5)" ' [ok-array]

  ' === Leading identifier: '=' assignment vs '==' equality (must not steal '=' from '==') ===
  tests(355).expr = "a=5; a==5":              tests(355).expected = "1" ' [syntax] single == after assign
  tests(356).expr = "b=3; b==4":              tests(356).expected = "0" ' [syntax]
  tests(357).expr = "7==7":                   tests(357).expected = "1" ' [ok-core] one ==, literals only
  tests(358).expr = "3==3 && 4==4 && 5==5":   tests(358).expected = "1" ' [ok-core] multiple ==
  tests(359).expr = "(1==1)==(0==0)":         tests(359).expected = "1" ' [ok-core] multiple ==
  tests(360).expr = "5==5==1":                tests(360).expected = "1" ' [ok-core] chained ==
  tests(361).expr = "u=2; v=2; u==v":         tests(361).expected = "1" ' [syntax] = and ==
  tests(362).expr = "k=9; k==9 && k==3+6":    tests(362).expected = "1" ' [syntax] = and multiple ==
  tests(363).expr = "a=1; b=1; a==b==1":      tests(363).expected = "1" ' [syntax] = and chained ==

  ' === [spec] Float-derived scalars promote to exact int64 when representable (bitwise/shift/metadata) ===
  tests(364).expr = "int(7/3) << 60":         tests(364).expected = "2305843009213693952" ' 2<<60, int from float
  tests(365).expr = "sqrt(9) << 2":          tests(365).expected = "12" ' sqrt via float, then shift
  tests(366).expr = "abs(-4) << 1":           tests(366).expected = "8"
  tests(367).expr = "sum(3,5) << 2":          tests(367).expected = "32" ' aggregate float path + shift
  tests(368).expr = "(7/3.5) << 3":           tests(368).expected = "16" ' pure float division, exact int

  ' === '=' assignment (leading identifier + single '=') vs '=' comparison in expressions ===
  tests(369).expr = "x=7; x=x":                tests(369).expected = "7" ' [syntax] leading x=x is assign, not compare
  tests(370).expr = "x=7; x==x":               tests(370).expected = "1" ' [syntax]
  tests(371).expr = "x=7; (x)=(x)":           tests(371).expected = "1" ' [syntax] paren => full expr; = is compare
  tests(372).expr = "x=5; x=x+1":             tests(372).expected = "6" ' [syntax] assign RHS uses +
  tests(373).expr = "z=0; (z)=(z)":           tests(373).expected = "1" ' [syntax] compare; z defined first
  tests(374).expr = "a=2; 1+a=3":             tests(374).expected = "1" ' [syntax] expr does not start with bare name; = compare
  tests(375).expr = "b=2; b+0=2":             tests(375).expected = "1" ' [syntax] leading term is not lone identifier assign form
  tests(376).expr = "c=1; c=2; c=c":           tests(376).expected = "2" ' [syntax] last c=c is assign
  tests(377).expr = "d=4; d = d":             tests(377).expected = "4" ' [syntax] spaces around assign =
  tests(378).expr = "t=3; (t=3)":             tests(378).expected = "1" ' [syntax] inner t=3 is compare ('e' is constant, not a var)
  tests(379).expr = "x=2;y=5;x+y=x":          tests(379).expected = "0" ' [syntax] (x+y)=x compare, not assign; 7=2 -> false
  tests(380).expr = "x=2;y=3;x*y=x*y":        tests(380).expected = "1" ' [syntax] (x*y)=(x*y) compare -> true

  ' === [spec] Built-in constants (pi, e) cannot be variable/function/param names ===
  tests(381).expr = "e=1":                    tests(381).expectedErrContains = "reserved constant name" ' [syntax]
  tests(382).expr = "PI=2":                   tests(382).expectedErrContains = "reserved constant name" ' [syntax] case-insensitive
  tests(383).expr = "f(e)=e+1":               tests(383).expectedErrContains = "reserved constant name" ' [syntax] param
  tests(384).expr = "pi(x)=x":                tests(384).expectedErrContains = "reserved constant name" ' [syntax] function name

  ' === hex() signed magnitude vs uhex/uoct/ubin (unsigned / two's complement) ===
  tests(385).expr = "hex(~0x0D)":             tests(385).expected = "-0xE" ' signed: -14
  tests(386).expr = "hex(-1)":                tests(386).expected = "-0x1"
  tests(387).expr = "uhex(~0x0D)":           tests(387).expected = "0xFFFFFFFFFFFFFFF2"
  tests(388).expr = "uhex(-1)":              tests(388).expected = "0xFFFFFFFFFFFFFFFF"
  tests(389).expr = "ubin(-1)":              tests(389).expected = "0b1111111111111111111111111111111111111111111111111111111111111111"
  tests(390).expr = "uoct(-1)":              tests(390).expected = "0o1777777777777777777777"
  tests(391).expr = "uhex()":                tests(391).expectedErrContains = "expects at least 1 argument"
  tests(392).expr = "uhex":                  tests(392).expectedErrContains = "function: uhex(...)"
  tests(393).expr = "uhex(1,2)":             tests(393).expected = "(0x1,0x2)"
  tests(394).expr = "bin(-2)":               tests(394).expected = "-0b10" ' bin/oct/hex still signed magnitude

  ' === int64 accuracy after float-path round-trip (scalar + array) ===
  tests(395).expr = "sqrt(81)&7":                                tests(395).expected = "1"
  tests(396).expr = "hex(sqrt(81))":                             tests(396).expected = "0x9"
  tests(397).expr = "mod(abs(-14),5)":                           tests(397).expected = "4"
  tests(398).expr = "pow(3,2)&7":                                tests(398).expected = "1"
  tests(399).expr = "hex(int((9007199254740992+2)/1))":          tests(399).expected = "0x20000000000002"
  tests(400).expr = "hex(int((9007199254740992+2)/2+0.0))":      tests(400).expected = "0x10000000000001"

  tests(401).expr = "a=sqrt((81,16,25)); a[0]&3":                tests(401).expected = "1"
  tests(402).expr = "a=sqrt((81,16,25)); hex(a[2])":             tests(402).expected = "0x5"
  tests(403).expr = "a=int((2.9,-2.9,7.1)); a[1]&1":             tests(403).expected = "0"
  tests(404).expr = "a=int((2.9,-2.9,7.1)); hex(a[0])":          tests(404).expected = "0x2"
  tests(405).expr = "a=int(((9007199254740992+2),(9007199254740992+6))/1); hex(a[0])": tests(405).expected = "0x20000000000002"
  tests(406).expr = "a=int(((9007199254740992+2),(9007199254740992+6))/1); mod(a[1],4)": tests(406).expected = "2"
  tests(407).expr = "a=int(((9007199254740992+2),(9007199254740992+6))/2); a[0]&1": tests(407).expected = "1"
  tests(408).expr = "a=int(((9007199254740992+2),(9007199254740992+6))/2); hex(a[1])": tests(408).expected = "0x10000000000003"
  tests(409).expr = "a=int(((5.9+0.1),(9.2+0.8))); mod(a[1],4)": tests(409).expected = "2"
  tests(410).expr = "a=int(((5.9+0.1),(9.2+0.8))); hex(a[0])":   tests(410).expected = "0x6"
  tests(411).expr = "a=(4611686018427387903,5)<<1; hex(a[0])":   tests(411).expected = "0x7FFFFFFFFFFFFFFE"
  tests(412).expr = "a=(4611686018427387903,5)<<1; mod(a[0],7)": tests(412).expected = "6"
  tests(413).expr = "a=(4611686018427387903,5)<<1; b=a>>1; hex(b[0])": tests(413).expected = "0x3FFFFFFFFFFFFFFF"
  tests(414).expr = "a=(9223372036854775806,15); b=a&7; hex(b[0])": tests(414).expected = "0x6"
  tests(415).expr = "a=(9223372036854775800,1); b=a|7; hex(b[0])": tests(415).expected = "0x7FFFFFFFFFFFFFFF"
  tests(416).expr = "a=(9223372036854775806,9223372036854775805); mod(a[1],5)": tests(416).expected = "0"
  tests(417).expr = "a=(9223372036854775806,9223372036854775805); b=a>>2; hex(b[0])": tests(417).expected = "0x1FFFFFFFFFFFFFFF"
  tests(418).expr = "-1>>1":                                     tests(418).expected = "-1"
  tests(419).expr = "uhex(1<<63)":                               tests(419).expected = "0x8000000000000000"
  tests(420).expr = "a=(-1,-2)>>1; a[0]":                        tests(420).expected = "-1"
  tests(421).expr = "a=(-1,3)<<1; uhex(a[0])":                   tests(421).expected = "0xFFFFFFFFFFFFFFFE"
  tests(422).expr = "a=(-1,-2)>>1; uhex(a[1])":                  tests(422).expected = "0xFFFFFFFFFFFFFFFF"
  tests(423).expr = "(1+2)(3+4)":                                tests(423).expected = "21"
  tests(424).expr = "2(1+pi)":                                   tests(424).expected = "8.283185307179586"
  tests(425).expr = "a=(3,4)+(5,6); hex(a[0])":                  tests(425).expected = "0x8"
  tests(426).expr = "a=(3,4)*2; hex(a[1])":                      tests(426).expected = "0x8"
  tests(427).expr = "a=2*(3,4); hex(a[1])":                      tests(427).expected = "0x8"
  tests(428).expr = "a=(1,2)&3; hex(a[1])":                      tests(428).expected = "0x2"
  tests(429).expr = "a=3&(1,2); hex(a[1])":                      tests(429).expected = "0x2"
  tests(430).expr = "a=(1,2)&(3,4); hex(a[1])":                  tests(430).expected = "0x0"
  tests(431).expr = "log((8,100),(2,10))":                       tests(431).expected = "(3,2)"
  tests(432).expr = "clamp((1,9),(0,10),(5,7))":                 tests(432).expectedErrContains = "expects scalar min/max"
  tests(433).expr = "clamp(5,(1,6),(4,7))":                      tests(433).expectedErrContains = "expects scalar min/max"
  tests(434).expr = "gcd((84,30),6)":                            tests(434).expected = "(6,6)" ' [shape/broadcast]
  tests(435).expr = "gcd(6,(84,30))":                            tests(435).expected = "(6,6)" ' [shape/broadcast]
  tests(436).expr = "lcm((6,8),3)":                              tests(436).expected = "(6,24)" ' [shape/broadcast]
  tests(437).expr = "lcm((6,8),(3,5))":                          tests(437).expected = "(6,40)" ' [shape/broadcast]
  tests(438).expr = "16>>1>>2":                                  tests(438).expected = "2"
  tests(439).expr = "7&3|8":                                     tests(439).expected = "11"
  tests(440).expr = "7^3^1":                                     tests(440).expected = "5"
  tests(441).expr = "hex(-2)":                                   tests(441).expected = "-0x2"
  tests(442).expr = "uhex(-2)":                                  tests(442).expected = "0xFFFFFFFFFFFFFFFE"
  tests(443).expr = "hex((15,-2))":                              tests(443).expected = "(0xF,-0x2)"
  tests(444).expr = "hypot(1,2,3,4)":                            tests(444).expectedErrContains = "expects 2 argument(s)"
  tests(445).expr = "a=9007199254740992+1; b=a; hex(b)":         tests(445).expected = "0x20000000000001"
  tests(446).expr = "uhex((1,-1))":                              tests(446).expected = "(0x1,0xFFFFFFFFFFFFFFFF)"
  tests(447).expr = "sum(unpack((1,2),(3,4),5))":                tests(447).expected = "15"
  tests(448).expr = "a=(1,2)+(3,4); hex(a[1])":                  tests(448).expected = "0x6"
  tests(449).expr = "NoT 0":                                     tests(449).expected = "1"
  tests(450).expr = "(5>=4)<2":                                  tests(450).expected = "1"
  tests(451).expr = "a=round((1.2,2.8)); hex(a[1])":             tests(451).expected = "0x3"
  tests(452).expr = "a=(5,6)*(7,8); hex(a[1])":                  tests(452).expected = "0x30"
  tests(453).expr = "a=(10,20)-(3,4); hex(a[1])":                tests(453).expected = "0x10"
  tests(454).expr = "a=(20,21)%(3,4); hex(a[1])":                tests(454).expected = "0x1"
  tests(455).expr = "log(8,(2,4))":                              tests(455).expected = "(3,1.5)"
  tests(456).expr = "lcm(21,6)":                                 tests(456).expected = "42"
  tests(457).expr = "a=(1,2)|(4,8); hex(a[1])":                 tests(457).expected = "0xA"
  tests(458).expr = "a=(1,2)^(4,8); hex(a[1])":                 tests(458).expected = "0xA"
  tests(459).expr = "a=(1,(2+3),7); hex(a[1])":                 tests(459).expected = "0x5"
  tests(460).expr = "a=abs((-1,-2)); hex(a[1])":                tests(460).expected = "0x2"
  tests(461).expr = "a=-(-1,-2); hex(a[1])":                    tests(461).expected = "0x2"
  tests(462).expr = "a=(8,9)>>1; hex(a[1])":                    tests(462).expected = "0x4"
  tests(463).expr = "a=(8,9)<<1; hex(a[1])":                    tests(463).expected = "0x12"
  tests(464).expr = "a=(9,10)&(3,12); hex(a[1])":               tests(464).expected = "0x8"
  tests(465).expr = "a=(9,10)|(3,12); hex(a[1])":               tests(465).expected = "0xE"
  tests(466).expr = "a=(9,10)^(3,12); hex(a[0])":               tests(466).expected = "0xA"
  tests(467).expr = "a=abs((-3,-4)); hex(a[0])":                tests(467).expected = "0x3"
  tests(468).expr = "a=(1,(2+3),4); hex(a[2])":                 tests(468).expected = "0x4"
  tests(469).expr = "a=int((7.9,8.1)); hex(a[1])":              tests(469).expected = "0x8"
  tests(470).expr = "a=(20,21)%(6,4); hex(a[0])":               tests(470).expected = "0x2"
  tests(471).expr = "a=(2,3)+(4,5); hex(a[1])":                 tests(471).expected = "0x8"
  tests(472).expr = "a=(9,7)-(4,5); hex(a[0])":                 tests(472).expected = "0x5"
  tests(473).expr = "a=(3,4)*(5,6); hex(a[0])":                 tests(473).expected = "0xF"
  tests(474).expr = "a=(8,9)/1; hex(a[1])":                     tests(474).expected = "0x9"
  tests(475).expr = "a=int((8.9,9.1)); hex(a[0])":              tests(475).expected = "0x8"
  tests(476).expr = "a=round((8.2,9.8)); hex(a[0])":            tests(476).expected = "0x8"
  tests(477).expr = "a=-(-5,-6); hex(a[0])":                    tests(477).expected = "0x5"
  tests(478).expr = "a=abs((-1,-2,-3)); hex(a[2])":             tests(478).expected = "0x3"
  tests(479).expr = "a=(8,9)|1; hex(a[0])":                     tests(479).expected = "0x9"
  tests(480).expr = "a=(8,9)&7; hex(a[0])":                     tests(480).expected = "0x0"
  tests(481).expr = "a=(8,9)<<1; hex(a[0])":                    tests(481).expected = "0x10"
  tests(482).expr = "a=(8,9)>>1; hex(a[0])":                    tests(482).expected = "0x4"
  tests(483).expr = "a=(20,21)%(6,4); hex(a[1])":               tests(483).expected = "0x1"
  tests(484).expr = "a=int((5.9,6.1,7.2)); hex(a[2])":          tests(484).expected = "0x7"
  tests(485).expr = "a=round((5.2,6.8,7.1)); hex(a[2])":        tests(485).expected = "0x7"
  tests(486).expr = "a=abs((-5,-6,-7)); hex(a[1])":             tests(486).expected = "0x6"
  tests(487).expr = "a=(8,9)|2; hex(a[1])":                     tests(487).expected = "0xB"
  tests(488).expr = "a=(8,9)&3; hex(a[1])":                     tests(488).expected = "0x1"
  tests(489).expr = "a=(8,9)^3; hex(a[1])":                     tests(489).expected = "0xA"
  tests(490).expr = "a=(10,11)+(1,2); hex(a[0])":               tests(490).expected = "0xB"
  tests(491).expr = "a=(10,11)-(1,2); hex(a[1])":               tests(491).expected = "0x9"
  tests(492).expr = "a=(10,11)*(2,3); hex(a[1])":               tests(492).expected = "0x21"
  tests(493).expr = "a=(10,11)/1; hex(a[0])":                   tests(493).expected = "0xA"
  tests(494).expr = "a=int((10.9,11.1)); hex(a[1])":            tests(494).expected = "0xB"
  tests(495).expr = "a=round((10.2,11.8)); hex(a[1])":          tests(495).expected = "0xC"
  tests(496).expr = "a=abs((-10,-11)); hex(a[0])":              tests(496).expected = "0xA"
  tests(497).expr = "a=-(-10,-11); hex(a[1])":                  tests(497).expected = "0xB"
  tests(498).expr = "a=(8,9,10)|1; hex(a[2])":                  tests(498).expected = "0xB"
  tests(499).expr = "a=(8,9,10)&3; hex(a[2])":                  tests(499).expected = "0x2"
  tests(500).expr = "a=(8,9,10)^3; hex(a[2])":                  tests(500).expected = "0x9"
  tests(501).expr = "a=(10,11)>>1; hex(a[1])":                  tests(501).expected = "0x5"
  tests(502).expr = "unique((5,3,5,2,3))":                      tests(502).expected = "(5,3,2)"
  tests(503).expr = "unique((2,1,2,1,3,2))":                    tests(503).expected = "(2,1,3)"
  tests(504).expr = "unique(7)":                                tests(504).expected = "(7)"
  tests(505).expr = "unique((0,-0,0,0))":                       tests(505).expected = "(0)"
  tests(506).expr = "unique((-0,0,1))":                         tests(506).expected = "(0,1)"
  tests(507).expr = "unique((1.5,1.5,2.5,1.5))":                tests(507).expected = "(1.5,2.5)"
  tests(508).expr = "unique((2.0,2,2.000,3))":                  tests(508).expected = "(2,3)"
  tests(509).expr = "unique(unpack((4,1),(4,2),1))":            tests(509).expected = "(4,1,2)"
  tests(510).expr = "unique((9,8,7,9,8,7,6,5,6))":              tests(510).expected = "(9,8,7,6,5)"
  tests(511).expr = "unique((3,3,3,3,2,2,1))":                  tests(511).expected = "(3,2,1)"
  tests(512).expr = "median((9,1,5,3,7))":                      tests(512).expected = "5"
  tests(513).expr = "median((10,1,7,3))":                       tests(513).expected = "5"
  tests(514).expr = "hex(fact(20))":                            tests(514).expected = "0x21C3677C82B40000"
  tests(515).expr = "factorial(30)":                            tests(515).expected = "2.65252859812191e+032"
  tests(516).expr = "sort((5,1,5,2,2,9,1))":                    tests(516).expected = "(1,1,2,2,5,5,9)"
  tests(517).expr = "sorted((4,4,3,2,2,1))":                    tests(517).expected = "(1,2,2,3,4,4)"
  tests(518).expr = "avg((2,4),6,8)":                           tests(518).expected = "5"
  tests(519).expr = "min((5,7),3,9)":                           tests(519).expected = "3"
  tests(520).expr = "sum(1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16)": tests(520).expected = "136"
  tests(521).expr = "max((5,7),3,9)":                           tests(521).expected = "9"
  tests(522).expr = "sum(unpack((1,2,3),(4,5),6,7))":          tests(522).expected = "28"
  tests(523).expr = "sorted((3,3,1,2,1))":                     tests(523).expected = "(1,1,2,3,3)"
  tests(524).expr = "reversed((1,2,3,4,5))":                   tests(524).expected = "(5,4,3,2,1)"
  tests(525).expr = "f(x,y)=x+y; f(2,3)":                      tests(525).expected = "5"
  tests(526).expr = "a=(5,2,5,9,2); unique(a)":                tests(526).expected = "(5,2,9)"
  tests(527).expr = "sum(42)":                                 tests(527).expected = "42"
  tests(528).expr = "a=(3,1,2); sorted(a)":                    tests(528).expected = "(1,2,3)"
  tests(529).expr = "unique((1,1,2,2,3,3))":                  tests(529).expected = "(1,2,3)"
  tests(530).expr = "avg(unpack((1,2),(3,4)),10)":            tests(530).expected = "4"
  tests(531).expr = "reverse((1,2),3,(4,5))":                 tests(531).expected = "(5,4,3,2,1)"
  tests(532).expr = "f(a,b)=a*b; f(unpack((3,4)))":           tests(532).expected = "12"
  tests(533).expr = "a=(9007199254740993); hex(unpack(a))":      tests(533).expected = "0x20000000000001"
  tests(534).expr = "a=-(9007199254740991,5); uhex(a[0])":    tests(534).expected = "0xFFE0000000000001"
  tests(535).expr = "hex((1,2.5))":                            tests(535).expectedErrContains = "hex() expects integer values"
  tests(536).expr = "unique((9,9,8,8,7))":                     tests(536).expected = "(9,8,7)"
  tests(537).expr = "reverse((9,8,7,6))":                      tests(537).expected = "(6,7,8,9)"
  tests(538).expr = "deg((pi/2),pi)":                          tests(538).expected = "(90,180)"
  tests(539).expr = "uhex((1,-1),2)":                          tests(539).expected = "(0x1,0xFFFFFFFFFFFFFFFF,0x2)"
  tests(540).expr = "uhex(unpack((1,-1),2))":                  tests(540).expected = "(0x1,0xFFFFFFFFFFFFFFFF,0x2)"
  tests(541).expr = "median(42)":                              tests(541).expected = "42"
  tests(542).expr = "median((42))":                            tests(542).expected = "42"
  tests(543).expr = "123":                                     tests(543).expected = "123"
  tests(544).expr = "unknownFunc(1)":                          tests(544).expectedErrContains = "unknown function"
  tests(545).expr = "2+3; ans":                                tests(545).expected = "5"
  tests(546).expr = "(e=3)":                                   tests(546).expected = "0"
  tests(547).expr = "(pi=3.141592653589793)":                  tests(547).expected = "1"
  tests(548).expr = "(1,2,3)[1.2]":                            tests(548).expectedErrContains = "array index must be an integer"
  tests(549).expr = "(1,2,3)[(1,2)]":                          tests(549).expectedErrContains = "array index must be a scalar integer"
  tests(550).expr = "1<<0":                                    tests(550).expected = "1"
  tests(551).expr = "8>>0":                                    tests(551).expected = "8"
  tests(552).expr = "1<<63":                                   tests(552).expected = "9223372036854775808"
  tests(553).expr = "-1>>63":                                  tests(553).expected = "-1"
  tests(554).expr = "a=(1,2)<<0; a":                           tests(554).expected = "(1,2)"
  tests(555).expr = "a=(8,9)>>0; a":                           tests(555).expected = "(8,9)"
  tests(556).expr = "a=(1,2)<<63; uhex(a[0])":                 tests(556).expected = "0x8000000000000000"
  tests(557).expr = "a=(-1,-2)>>63; a[1]":                     tests(557).expected = "-1"
  tests(558).expr = "0b102":                                   tests(558).expectedErrContains = "unexpected token"
  tests(559).expr = "0o89":                                    tests(559).expectedErrContains = "invalid octal literal"
  tests(560).expr = "0x1G":                                    tests(560).expectedErrContains = "unexpected token"
  tests(561).expr = "hex=1":                                   tests(561).expectedErrContains = "reserved function name"
  tests(562).expr = "HEX=1":                                   tests(562).expectedErrContains = "reserved function name"
  tests(563).expr = "random=1":                                tests(563).expectedErrContains = "reserved function name"
  tests(564).expr = "0xAA; hex":                               tests(564).expected = "0xAA"
  tests(565).expr = "0xAA; hex()":                             tests(565).expected = "0xAA"
  tests(566).expr = "(0x3C & 0x75, 0x01 | 0x30); hex":         tests(566).expected = "(0x34,0x31)"
  tests(567).expr = "(8,9); bin()":                            tests(567).expected = "(0b1000,0b1001)"
  tests(568).expr = "15; uhex":                                tests(568).expected = "0xF"
  tests(569).expr = "0xAA; foo()":                             tests(569).expectedErrContains = "unknown function"

  tests(570).expr = "x(a)=x(a); x(1)":                        tests(570).expectedErrContains = "recursive function call: x" ' [regression] direct self-call in UDF body
  tests(571).expr = "y(a)=g(a)+y(a)+4":                      tests(571).expectedErrContains = "recursive function call: y" ' [regression] self-call among other terms
  tests(572).expr = "g(a)=y(a)+1; y(a)=g(a)+2; y(5)":        tests(572).expectedErrContains = "recursive function call" ' [regression] mutual recursion y<->g
  tests(573).expr = "a(x)=b(x); b(x)=c(x); c(x)=d(x); d(x)=b(x); a(1)": tests(573).expectedErrContains = "recursive function call" ' [regression] longer cycle back to b
  tests(574).expr = "2^3":                                   tests(574).expected = "1" ' [regression] caret is bitwise XOR
  tests(575).expr = "3^2":                                   tests(575).expected = "1" ' [regression] caret is not power
  tests(576).expr = "3**2":                                  tests(576).expected = "9" ' [regression] double-star power
  tests(577).expr = "f(x)=x*p(x); f(2)":                     tests(577).expectedErrContains = "unknown function" ' [regression] late binding unresolved referenced UDF
  tests(578).expr = "f(x)=x*p(x); p(x)=x+5; f(10)":          tests(578).expected = "150" ' [regression] late binding resolved after referenced UDF definition
  tests(579).expr = "f(x)=x*p(x); p(x)=x**(1/3); f(8)":      tests(579).expected = "16" ' [regression] late binding with nonlinear referenced UDF
  tests(580).expr = "f(x)=x*p(x); p(x)=x+5; p(x)=x**(1/3); f(8)": tests(580).expected = "16" ' [regression] late binding uses latest referenced UDF definition

  ' --- Additional scalar-vs-array / precedence clarity checks (auto-loaded by C++ smoke runner) ---
  tests(581).expr = "((5))":        tests(581).expected = "5"   ' grouping only => scalar
  tests(582).expr = "!2==1":       tests(582).expected = "0"   ' (!2)==1 because ! has higher precedence
  tests(583).expr = "not 2==1":    tests(583).expected = "1"   ' not (2==1) because not is lower than ==

  ' Additional scalar-vs-array ambiguity locks:
  tests(584).expr = "!(0,0)":        tests(584).expected = "0" ' non-empty arrays are truthy, even if all elements are 0
  tests(585).expr = "0&&(0,0)":      tests(585).expected = "0" ' scalar 0 && truthy array => 0
  tests(586).expr = "((0))&((0))":   tests(586).expected = "0" ' scalar grouping for bitwise & => scalar
  tests(587).expr = "(0,0)&(0,0)":  tests(587).expected = "(0,0)" ' array & array => element-wise array

  ' Additional variants mirrored from C++ unit coverage:
  tests(588).expr = "NOT (0,0)":         tests(588).expected = "0"
  tests(589).expr = "0||(0,0)":          tests(589).expected = "1"
  tests(590).expr = "((0))|((0))":      tests(590).expected = "0"
  tests(591).expr = "(0,0)|(0,0)":     tests(591).expected = "(0,0)"

  tests(592).expr = "a=42; a+1": tests(592).expected = "43"
  tests(593).expr = "rate=1.5; rate*8": tests(593).expected = "12"
  tests(594).expr = "mul3(x)=x*3; mul3(7)": tests(594).expected = "21"
  tests(595).expr = "!(0)": tests(595).expected = "1"
  tests(596).expr = "((0))&&1": tests(596).expected = "0"
  tests(597).expr = "(0,0)&&1": tests(597).expected = "1"
  tests(598).expr = "(0,0)|1": tests(598).expected = "(1,1)"
  tests(599).expr = "k=0; k+0": tests(599).expected = "0"
  tests(600).expr = "k=0; hex(k)": tests(600).expected = "0x0"
  tests(601).expr = "k=1; k*k": tests(601).expected = "1"
  tests(602).expr = "k=1; hex(k)": tests(602).expected = "0x1"
  tests(603).expr = "k=100; k+1": tests(603).expected = "101"
  tests(604).expr = "k=100; hex(k)": tests(604).expected = "0x64"
  tests(605).expr = "k=1000000; k+1": tests(605).expected = "1000001"
  tests(606).expr = "k=1000000; hex(k)": tests(606).expected = "0xF4240"
  tests(607).expr = "k=9223372036854775807; k+0": tests(607).expected = "9223372036854775807"
  tests(608).expr = "k=-9223372036854775808; k+0": tests(608).expected = "-9223372036854775808"
  tests(609).expr = "k=9223372036854775807; (k-1)+1": tests(609).expected = "9223372036854775807"
  tests(610).expr = "k=9007199254740992; hex(k+1)": tests(610).expected = "0x20000000000001"
  tests(611).expr = "k=9007199254740991; k+1": tests(611).expected = "9007199254740992"
  tests(612).expr = "a=100; gcd(a,25)": tests(612).expected = "25"
  tests(613).expr = "a=100; mod(a,3)": tests(613).expected = "1"
  tests(614).expr = "a=100; a&7": tests(614).expected = "4"
  tests(615).expr = "x=100.7; int(x)": tests(615).expected = "100"
  tests(616).expr = "k=9223372036854775806; k+1": tests(616).expected = "9223372036854775807"
  tests(617).expr = "k=9223372036854775806; k+0": tests(617).expected = "9223372036854775806"
  tests(618).expr = "k=-9223372036854775807; k-1": tests(618).expected = "-9223372036854775808"
  tests(619).expr = "k=-9223372036854775807; k+0": tests(619).expected = "-9223372036854775807"
  tests(620).expr = "9223372036854775806+1": tests(620).expected = "9223372036854775807"
  tests(621).expr = "-9223372036854775807-1": tests(621).expected = "-9223372036854775808"
  tests(622).expr = "-9223372036854775807+0": tests(622).expected = "-9223372036854775807"
  tests(623).expr = "k=9223372036854775806; hex(k)": tests(623).expected = "0x7FFFFFFFFFFFFFFE"
  tests(624).expr = "k=-9223372036854775807; hex(k)": tests(624).expected = "-0x7FFFFFFFFFFFFFFF"
  tests(625).expr = "k=-9223372036854775807; uhex(k)": tests(625).expected = "0x8000000000000001"
  tests(626).expr = "k=9007199254740990; k+2": tests(626).expected = "9007199254740992"
  tests(627).expr = "k=9007199254740991; k+2": tests(627).expected = "9007199254740993"
  tests(628).expr = "k=9007199254740992; k+2": tests(628).expected = "9007199254740994"
  tests(629).expr = "k=9007199254740991; int(k)": tests(629).expected = "9007199254740991"
  tests(630).expr = "k=9007199254740994; hex(int(k/1))": tests(630).expected = "0x20000000000002"
  tests(631).expr = "k=9007199254740992; hex(int((k/2)+0.0))": tests(631).expected = "0x10000000000000"
  tests(632).expr = "k=9007199254740992; a=int(((k+2),(k+6))/1); hex(a[0])": tests(632).expected = "0x20000000000002"
  tests(633).expr = "k=9007199254740992; a=int(((k+2),(k+6))/1); mod(a[1],4)": tests(633).expected = "2"
  tests(634).expr = "k=9007199254740992; a=int(((k+2),(k+6))/2); a[0]&1": tests(634).expected = "1"
  tests(635).expr = "k=9007199254740992; a=int(((k+2),(k+6))/2); hex(a[1])": tests(635).expected = "0x10000000000003"
  tests(636).expr = "x=0.0/0.0; x+0": tests(636).expected = "nan"
  tests(637).expr = "x=1.0/0.0; x+0": tests(637).expected = "inf"
  tests(638).expr = "x=-1.0/0.0; x+0": tests(638).expected = "-inf"
  tests(639).expr = "x=1.0/0.0; x+1": tests(639).expected = "inf"
  tests(640).expr = "x=-1.0/0.0; x-1": tests(640).expected = "-inf"
  tests(641).expr = "x=0.0/0.0; x+1": tests(641).expected = "nan"
  tests(642).expr = "x=-1.0/0.0; abs(x)": tests(642).expected = "inf"
  tests(643).expr = "x=0.0/0.0; abs(x)": tests(643).expected = "nan"
  tests(644).expr = "x=1.0/0.0; sign(x)": tests(644).expected = "1"
  tests(645).expr = "x=-1.0/0.0; sign(x)": tests(645).expected = "-1"
  tests(646).expr = "x=0.0/0.0; sign(x)": tests(646).expected = "0"
  tests(647).expr = "x=0.0/0.0; ln(x)": tests(647).expected = "nan"
  tests(648).expr = "x=1.0/0.0; ln(x)": tests(648).expected = "inf"
  tests(649).expr = "z=0.0; ln(z)": tests(649).expected = "-inf"
  tests(650).expr = "x=0.0/0.0; sqrt(x)": tests(650).expected = "nan"
  tests(651).expr = "x=1.0/0.0; sin(x)": tests(651).expectedErrContains = "numeric error in sin()"
  tests(652).expr = "x=1.0/0.0; frac(x)": tests(652).expected = "nan"
  tests(653).expr = "x=0.0/0.0; not x": tests(653).expected = "1"
  tests(654).expr = "x=1.0/0.0; not x": tests(654).expected = "0"
  tests(655).expr = "x=0.0/0.0; x&&1": tests(655).expected = "0"
  tests(656).expr = "x=0.0/0.0; x||1": tests(656).expected = "1"
  tests(657).expr = "x=0.0/0.0; x==x": tests(657).expected = "0"
  tests(658).expr = "200 + 2*15%": tests(658).expected = "200.3"
  tests(659).expr = "unpack((1,2),3)": tests(659).expected = "(1,2,3)"
  tests(660).expr = "unpack(1,(2,3))": tests(660).expected = "(1,2,3)"
  tests(661).expr = "unpack((1,2))": tests(661).expected = "(1,2)"
  tests(662).expr = "deg((pi,pi/2))": tests(662).expected = "(180,90)"
  tests(663).expr = "rad((180,90))": tests(663).expected = "(3.141592653589793,1.570796326794896)"
  tests(664).expr = "rad((0,0))": tests(664).expected = "(0,0)"
  tests(665).expr = "clamp(9,0,7)": tests(665).expected = "7"
  tests(666).expr = "clamp((1,9),0,7)": tests(666).expected = "(1,7)"
  tests(667).expr = "fact(21)": tests(667).expected = "5.109094217170944e+019"
  tests(668).expr = "(2<3)+(2>=3)+(5==5)+(5<>4)": tests(668).expected = "3"
  tests(669).expr = "(7<>8)+(7!=8)": tests(669).expected = "2"
  tests(670).expr = "(9>8)+(9<=8)+(8>=8)": tests(670).expected = "2"
  tests(671).expr = "(3==3)+(3!=4)+(3<>3)": tests(671).expected = "2"
  tests(672).expr = "ubin(5)": tests(672).expected = "0b101"
  tests(673).expr = "log((8,64),2)": tests(673).expected = "(3,6)"
  tests(674).expr = "a=1; 2; ans": tests(674).expected = "2"
  tests(675).expr = "10+5": tests(675).expected = "15"
  tests(676).expr = "ans": tests(676).expected = "15"
  tests(677).expr = "stddev((-3,0,3,6))": tests(677).expected = "3.354101966249684"
  tests(678).expr = "stddev((1,2),3,4)": tests(678).expected = "1.118033988749894"
  tests(679).expr = "uhex(18446744073709551615)": tests(679).expected = "0xFFFFFFFFFFFFFFFF"
  tests(680).expr = "hex((1,2.5,3))": tests(680).expectedErrContains = "hex() expects integer values"
  tests(681).expr = "gcd(18446744073709551615,3)": tests(681).expected = "3"
  tests(682).expr = "lcm(18446744073709551615,3)": tests(682).expected = "18446744073709551615"
  tests(683).expr = "18446744073709551615%3": tests(683).expected = "0"
  tests(684).expr = "a=100.5; mod(a,3)": tests(684).expectedErrContains = "mod() expects integer values"
  tests(685).expr = "mod(18446744073709551615,3)": tests(685).expected = "0"
  tests(686).expr = "a=100.5; a&7": tests(686).expectedErrContains = "bitwise operands must be integer values"
  tests(687).expr = "mod(7,(2,3))": tests(687).expected = "(1,1)"
  tests(688).expr = "mod((7,8),3)": tests(688).expected = "(1,2)"
  tests(689).expr = "a=(1,2)<<64": tests(689).expectedErrContains = "incompatible operands"
  tests(690).expr = "x=0.0/0.0; hex(x)": tests(690).expected = "nan"
  tests(691).expr = "x=-1.0/0.0; oct(x)": tests(691).expected = "-inf"
  tests(692).expr = "x=0.0/0.0; mod(x,3)": tests(692).expectedErrContains = "mod() expects integer values"
  tests(693).expr = "x=0.0/0.0; gcd(x,1)": tests(693).expectedErrContains = "gcd() expects integer values"
  tests(694).expr = "x=0.0/0.0; x&1": tests(694).expectedErrContains = "bitwise operands must be integer values"
  tests(695).expr = "clamp(5,1,(4,7))": tests(695).expectedErrContains = "clamp() expects scalar min/max"
  tests(696).expr = "hypot((3,4),5)": tests(696).expected = "(5.830951894845301,6.403124237432849)"
  tests(697).expr = "fact((3,4))": tests(697).expectedErrContains = "fact() expects a non-negative integer"
  tests(698).expr = "mod(5)": tests(698).expectedErrContains = "mod() expects 2 argument(s)"
  tests(699).expr = "pow(5)": tests(699).expectedErrContains = "pow() expects 2 argument(s)"
  tests(700).expr = "log(5)": tests(700).expectedErrContains = "log() expects 2 argument(s)"
  tests(701).expr = "(1,2)+(3,4,5)": tests(701).expectedErrContains = "incompatible operands"
  tests(702).expr = "200 + (2*15)%": tests(702).expected = "260"
  tests(703).expr = "x=1.0/0.0; uhex(x)": tests(703).expected = "inf"
  tests(704).expr = "x=1.0/0.0; sqrt(x)": tests(704).expected = "inf"
  tests(705).expr = "a=(0.0/0.0,1.0/0.0,-1.0/0.0); a": tests(705).expected = "(nan,inf,-inf)"
  tests(706).expr = "x=0.0/0.0; sum(1,x)": tests(706).expected = "nan"

  tests(707).expr = "log(e,e)": tests(707).expected = "1"
  tests(708).expr = "unique((1,2),2,1,3)": tests(708).expected = "(1,2,3)"
  tests(709).expr = "variance((-3,0,3,6))": tests(709).expected = "11.25"
  tests(710).expr = "variance((1,2),3,4)": tests(710).expected = "1.25"
  tests(711).expr = "2**3": tests(711).expected = "8"
  tests(712).expr = "unpack()": tests(712).expectedErrContains = "unpack() expects at least 1 argument"
  tests(713).expr = "deg()": tests(713).expectedErrContains = "deg() expects at least 1 argument"
  tests(714).expr = "rad()": tests(714).expectedErrContains = "rad() expects at least 1 argument"
  tests(715).expr = "clamp(1,2)": tests(715).expectedErrContains = "clamp() expects 3 argument(s)"
  tests(716).expr = "clamp(1,2,3,4)": tests(716).expectedErrContains = "clamp() expects 3 argument(s)"
  tests(717).expr = "clamp((1,2),(3,4),4)": tests(717).expectedErrContains = "clamp() expects scalar min/max"
  tests(718).expr = "random((1,2),3)": tests(718).expectedErrContains = "random() expects scalar values"
  tests(719).expr = "random(1)": tests(719).expectedErrContains = "random() expects 2 argument(s)"
  tests(720).expr = "sum()": tests(720).expectedErrContains = "sum() expects at least 1 argument"
  tests(721).expr = "hex()": tests(721).expectedErrContains = "hex() expects at least 1 argument"
  tests(722).expr = "k=9007199254740992; v=(k+2)-2; hex(int(v))": tests(722).expected = "0x20000000000000"
  tests(723).expr = "k=18446744073709551615; v=k+1; uhex(v)": tests(723).expectedErrContains = "uhex() expects integer values"
  tests(724).expr = "a=(18446744073709551615,1); b=(1,2); c=a+b; uhex(c[0])": tests(724).expectedErrContains = "uhex() expects integer values"
  tests(725).expr = "a=(18446744073709551615,3); b=(3,4); c=a*b; uhex(c[0])": tests(725).expectedErrContains = "uhex() expects integer values"
  tests(726).expr = "a=floor((9007199254740993,2.9)); hex(a[0])": tests(726).expected = "0x20000000000001"
  tests(727).expr = "a=(9007199254740993,2); hex(sum(a[0]))": tests(727).expected = "0x20000000000001"
  tests(728).expr = "a=abs((18446744073709551615,-7)); uhex(a[0])": tests(728).expected = "0xFFFFFFFFFFFFFFFF"
  tests(729).expr = "a=sort((18446744073709551615,2)); uhex(a[1])": tests(729).expected = "0xFFFFFFFFFFFFFFFF"
  tests(730).expr = "a=reverse((18446744073709551615,2)); uhex(a[1])": tests(730).expected = "0xFFFFFFFFFFFFFFFF"
  tests(731).expr = "a=unique((18446744073709551615,18446744073709551615,2)); uhex(a[0])": tests(731).expected = "0xFFFFFFFFFFFFFFFF"
  tests(732).expr = "a=(1.0/0.0,2,3); sum(a)": tests(732).expected = "inf"
  tests(733).expr = "a=(1.0/0.0,2); mod(a,3)": tests(733).expectedErrContains = "mod() expects integer values"
  tests(734).expr = "x=1.0/0.0; hypot(x,3)": tests(734).expected = "inf"
  tests(735).expr = "x=1.0/0.0; clamp(x,0,7)": tests(735).expected = "7"
  tests(736).expr = "x=1.0/0.0; f(a)=a+1; f(x)": tests(736).expected = "inf"
  tests(737).expr = "x=1.0/0.0; f(a,b)=a*b; f(x,2)": tests(737).expected = "inf"
  tests(738).expr = "x=1.0/0.0; x*2": tests(738).expected = "inf"
  tests(739).expr = "x=1.0/0.0; x|1": tests(739).expectedErrContains = "bitwise operands must be integer values"
  tests(740).expr = "x=1.0/0.0; x&&0": tests(740).expected = "0"
  tests(741).expr = "x=1.0/0.0; x>1": tests(741).expected = "1"
  tests(742).expr = "x=1.0/0.0; sum(1,2,x)": tests(742).expected = "inf"
  tests(743).expr = "x=1.0/0.0; max(1,2,x)": tests(743).expected = "inf"
  tests(744).expr = "x=1.0/0.0; avg(1,x,3)": tests(744).expected = "inf"
  tests(745).expr = "x=1.0/0.0; x==(1.0/0.0)": tests(745).expected = "1"
  tests(746).expr = "a=(-1.0/0.0,2,3); sum(a)": tests(746).expected = "-inf"
  tests(747).expr = "a=(-1.0/0.0,2); mod(a,3)": tests(747).expectedErrContains = "mod() expects integer values"
  tests(748).expr = "x=-1.0/0.0; hypot(x,3)": tests(748).expected = "inf"
  tests(749).expr = "x=-1.0/0.0; clamp(x,0,7)": tests(749).expected = "0"
  tests(750).expr = "x=-1.0/0.0; f(a)=a+1; f(x)": tests(750).expected = "-inf"
  tests(751).expr = "x=-1.0/0.0; f(a,b)=a*b; f(x,2)": tests(751).expected = "-inf"
  tests(752).expr = "x=-1.0/0.0; x*2": tests(752).expected = "-inf"
  tests(753).expr = "x=-1.0/0.0; x|1": tests(753).expectedErrContains = "bitwise operands must be integer values"
  tests(754).expr = "x=-1.0/0.0; x&&0": tests(754).expected = "0"
  tests(755).expr = "x=-1.0/0.0; x>1": tests(755).expected = "0"
  tests(756).expr = "x=-1.0/0.0; sum(1,2,x)": tests(756).expected = "-inf"
  tests(757).expr = "x=-1.0/0.0; max(1,2,x)": tests(757).expected = "2"
  tests(758).expr = "x=-1.0/0.0; avg(1,x,3)": tests(758).expected = "-inf"
  tests(759).expr = "x=-1.0/0.0; y=-1.0/0.0; x==y": tests(759).expected = "1"
  tests(760).expr = "x=1.0/0.0; x**2": tests(760).expected = "inf"
  tests(761).expr = "x=1.0/0.0; 2**x": tests(761).expected = "inf"
  tests(762).expr = "x=1.0/0.0; x**x": tests(762).expected = "inf"
  tests(763).expr = "x=1.0/0.0; pow(x,2)": tests(763).expected = "inf"
  tests(764).expr = "x=1.0/0.0; pow(2,x)": tests(764).expected = "inf"
  tests(765).expr = "x=1.0/0.0; pow(x,x)": tests(765).expected = "inf"
  tests(766).expr = "x=1.0/0.0; atan2(x,1)": tests(766).expected = "1.570796326794897"
  tests(767).expr = "x=-1.0/0.0; atan2(x,1)": tests(767).expected = "-1.570796326794897"
  tests(768).expr = "x=-1.0/0.0; atan2(1,x)": tests(768).expected = "3.141592653589793"
  tests(769).expr = "x=-1.0/0.0; sin(x)": tests(769).expectedErrContains = "numeric error in sin()"
  tests(770).expr = "inf+1": tests(770).expected = "inf"
  tests(771).expr = "INF+1": tests(771).expected = "inf"
  tests(772).expr = "inf=1": tests(772).expectedErrContains = "reserved constant name"
  tests(773).expr = "Inf=1": tests(773).expectedErrContains = "reserved constant name"
  tests(774).expr = "f(inf)=inf+1": tests(774).expectedErrContains = "reserved constant name"
  tests(775).expr = "inf(x)=x": tests(775).expectedErrContains = "reserved constant name"
  tests(776).expr = "-inf": tests(776).expected = "-inf"
  tests(777).expr = "-Inf+1": tests(777).expected = "-inf"
  tests(778).expr = "sum(-inf,2,3)": tests(778).expected = "-inf"
  tests(779).expr = "max(-inf,2,3)": tests(779).expected = "3"
  tests(780).expr = "(-inf)&1": tests(780).expectedErrContains = "bitwise operands must be integer values"
  tests(781).expr = "(-inf)==(-inf)": tests(781).expected = "1"
  tests(782).expr = "-inf + inf": tests(782).expected = "nan"
  tests(783).expr = "inf/-inf": tests(783).expected = "nan"
  tests(784).expr = "pow(inf,-inf)": tests(784).expected = "0"
  tests(785).expr = "pow(-inf,3)": tests(785).expected = "-inf"
  tests(786).expr = "pow(-inf,2)": tests(786).expected = "inf"
  tests(787).expr = "inf-inf": tests(787).expected = "nan"
  tests(788).expr = "(-inf)/(-inf)": tests(788).expected = "nan"
  tests(789).expr = "inf*0": tests(789).expected = "nan"
  tests(790).expr = "0*inf": tests(790).expected = "nan"
  tests(791).expr = "pow(inf,-2)": tests(791).expected = "0"
  tests(792).expr = "ncr(5,2)": tests(792).expected = "10"
  tests(793).expr = "npr(5,2)": tests(793).expected = "20"
  tests(794).expr = "ncr(10,0)": tests(794).expected = "1"
  tests(795).expr = "npr(10,0)": tests(795).expected = "1"
  tests(796).expr = "ncr(5,7)": tests(796).expectedErrContains = "numeric error in ncr()"
  tests(797).expr = "npr(5,7)": tests(797).expectedErrContains = "numeric error in npr()"
  tests(798).expr = "ncr(-1,0)": tests(798).expectedErrContains = "numeric error in ncr()"
  tests(799).expr = "npr(5,-1)": tests(799).expectedErrContains = "numeric error in npr()"
  tests(800).expr = "ncr(5.5,2)": tests(800).expectedErrContains = "ncr() expects integer values"
  tests(801).expr = "npr(inf,2)": tests(801).expectedErrContains = "npr() expects integer values"
  tests(802).expr = "ncr": tests(802).expectedErrContains = "function: ncr(n, r)"
  tests(803).expr = "npr": tests(803).expectedErrContains = "function: npr(n, r)"

  tests(804).expr = "nan+1": tests(804).expected = "nan"
  tests(805).expr = "NAN+1": tests(805).expected = "nan"
  tests(806).expr = "nan=1": tests(806).expectedErrContains = "reserved constant name"
  tests(807).expr = "NaN=1": tests(807).expectedErrContains = "reserved constant name"
  tests(808).expr = "f(nan)=nan+1": tests(808).expectedErrContains = "reserved constant name"
  tests(809).expr = "nan(x)=x": tests(809).expectedErrContains = "reserved constant name"
  tests(810).expr = "(nan, inf)": tests(810).expected = "(nan, inf)"

  tests(811).expr = "acosh(1)": tests(811).expected = "0"
  tests(812).expr = "acosh(2)": tests(812).expected = "1.3169578969248166"
  tests(813).expr = "acosh(0)": tests(813).expected = "nan"
  tests(814).expr = "acosh(inf)": tests(814).expected = "inf"
  tests(815).expr = "asinh(0)": tests(815).expected = "0"
  tests(816).expr = "asinh(1)": tests(816).expected = "0.8813735870195431"
  tests(817).expr = "asinh(-1)": tests(817).expected = "-0.8813735870195431"
  tests(818).expr = "asinh((0,1))": tests(818).expected = "(0,0.8813735870195431)"
  tests(819).expr = "atanh(0)": tests(819).expected = "0"
  tests(820).expr = "atanh(0.5)": tests(820).expected = "0.5493061443340549"
  tests(821).expr = "atanh(1)": tests(821).expected = "inf"
  tests(822).expr = "atanh(-1)": tests(822).expected = "-inf"
  tests(823).expr = "atanh(2)": tests(823).expected = "nan"
  tests(824).expr = "acosh()": tests(824).expectedErrContains = "expects 1 argument(s)"
  tests(825).expr = "asinh()": tests(825).expectedErrContains = "expects 1 argument(s)"
  tests(826).expr = "atanh()": tests(826).expectedErrContains = "expects 1 argument(s)"

  tests(827).expr = "sin(pi/2)": tests(827).expected = "1"
  tests(828).expr = "sin(-pi/2)": tests(828).expected = "-1"
  tests(829).expr = "sin(77777*pi/2)": tests(829).expected = "1"
  tests(830).expr = "sin(-77777*pi/2)": tests(830).expected = "-1"
  tests(831).expr = "sin(0)": tests(831).expected = "0"
  tests(832).expr = "sin(pi)": tests(832).expected = "0"
  tests(833).expr = "sin(-pi)": tests(833).expected = "0"
  tests(834).expr = "sin(2*pi)": tests(834).expected = "0"
  tests(835).expr = "sin(-2*pi)": tests(835).expected = "0"
  tests(836).expr = "sin(77777*pi)": tests(836).expected = "0"
  tests(837).expr = "sin(-77777*pi)": tests(837).expected = "0"
  tests(838).expr = "sin(77778*pi)": tests(838).expected = "0"
  tests(839).expr = "sin(-77778*pi)": tests(839).expected = "0"
  tests(840).expr = "cos(pi/2)": tests(840).expected = "0"
  tests(841).expr = "cos(-pi/2)": tests(841).expected = "0"
  tests(842).expr = "cos(77777*pi/2)": tests(842).expected = "0"
  tests(843).expr = "cos(-77777*pi/2)": tests(843).expected = "0"
  tests(844).expr = "cos(0)": tests(844).expected = "1"
  tests(845).expr = "cos(pi)": tests(845).expected = "-1"
  tests(846).expr = "cos(-pi)": tests(846).expected = "-1"
  tests(847).expr = "cos(2*pi)": tests(847).expected = "1"
  tests(848).expr = "cos(-2*pi)": tests(848).expected = "1"
  tests(849).expr = "cos(77777*pi)": tests(849).expected = "-1"
  tests(850).expr = "cos(-77777*pi)": tests(850).expected = "-1"
  tests(851).expr = "cos(77778*pi)": tests(851).expected = "1"
  tests(852).expr = "cos(-77778*pi)": tests(852).expected = "1"
  tests(853).expr = "tan(pi/2)": tests(853).expected = "inf"
  tests(854).expr = "tan(-pi/2)": tests(854).expected = "-inf"
  tests(855).expr = "tan(77777*pi/2)": tests(855).expected = "inf"
  tests(856).expr = "tan(-77777*pi/2)": tests(856).expected = "-inf"
  tests(857).expr = "tan(0)": tests(857).expected = "0"
  tests(858).expr = "tan(pi)": tests(858).expected = "0"
  tests(859).expr = "tan(-pi)": tests(859).expected = "0"
  tests(860).expr = "tan(2*pi)": tests(860).expected = "0"
  tests(861).expr = "tan(-2*pi)": tests(861).expected = "0"
  tests(862).expr = "tan(77777*pi)": tests(862).expected = "0"
  tests(863).expr = "tan(-77777*pi)": tests(863).expected = "0"
  tests(864).expr = "tan(77778*pi)": tests(864).expected = "0"
  tests(865).expr = "tan(-77778*pi)": tests(865).expected = "0"

  tests(866).expr = "bad1(x)=x)": tests(866).expectedErrContains = "unexpected characters" ' [regression] UDF body must be one expression; no extra )
  tests(867).expr = "bad2(x)=x$$5": tests(867).expectedErrContains = "unexpected characters" ' [regression] garbage after a valid expression prefix in UDF body

  tests(868).expr = "!nan": tests(868).expected = "1" ' [spec] NaN is falsy in logical ops (IEEE: NaN <> 0 is true, so do not use <>0 alone for truthiness)
  tests(869).expr = "nan && 1": tests(869).expected = "0" ' [spec] NaN is falsy
  tests(870).expr = "nan || 0": tests(870).expected = "0" ' [spec] NaN is falsy
  tests(871).expr = "nan==nan": tests(871).expected = "0" ' [spec] IEEE: NaN never equals NaN
  tests(872).expr = "nan!=nan": tests(872).expected = "1" ' [spec] inequality is true for unordered
  tests(873).expr = "nan<>nan": tests(873).expected = "1"
  tests(874).expr = "nan<nan": tests(874).expected = "0"
  tests(875).expr = "nan>nan": tests(875).expected = "0"
  tests(876).expr = "nan<=nan": tests(876).expected = "0"
  tests(877).expr = "nan>=nan": tests(877).expected = "0"

  tests(878).expr = "x=-nan; int(x)": tests(878).expected = "nan"
  tests(879).expr = "x=nan; int(x)": tests(879).expected = "nan"
  tests(880).expr = "x=-inf; int(x)": tests(880).expected = "-inf"
  tests(881).expr = "x=inf; int(x)": tests(881).expected = "inf"
  tests(882).expr = "x=-nan; ceil(x)": tests(882).expected = "nan"
  tests(883).expr = "x=nan; floor(x)": tests(883).expected = "nan"
  tests(884).expr = "x=-inf; round(x)": tests(884).expected = "-inf"
  tests(885).expr = "x=inf; trunc(x)": tests(885).expected = "inf"
  tests(886).expr = "ceil((1e14+0.5, 1e30, -1e30))": tests(886).expected = "(100000000000001, 1e+30, -1e+30)"

  tests(887).expr = "(45,60,90); rad": tests(887).expected = "(0.7853981633974483, 1.047197551196598, 1.570796326794897)"
  tests(888).expr = "(pi/4,pi/3,pi/2); deg": tests(888).expected = "(45, 60, 90)"

  tests(889).expr = "(2**58, 2**58+123)": tests(889).expected = "(288230376151711744, 288230376151711867)"
  tests(890).expr = "(2**58, 2**58+123);hex": tests(890).expected = "(0x400000000000000, 0x40000000000007B)"
  tests(891).expr = "(2**61, 2**61+123)": tests(891).expected = "(2305843009213693952, 2305843009213694075)"
  tests(892).expr = "(2**61, 2**61+123);hex": tests(892).expected = "(0x2000000000000000, 0x200000000000007B)"
  tests(893).expr = "2**58+123": tests(893).expected = "288230376151711867"
  tests(894).expr = "2**61+123": tests(894).expected = "2305843009213694075"
  tests(895).expr = "2**58+123;hex": tests(895).expected = "0x40000000000007B"
  tests(896).expr = "2**61+123;hex": tests(896).expected = "0x200000000000007B"
  tests(897).expr = "9.0*10**18": tests(897).expected = "9000000000000000000"
  tests(898).expr = "9.2e17": tests(898).expected = "920000000000000000"
  tests(899).expr = "9e18": tests(899).expected = "9000000000000000000"
  tests(900).expr = "90.123e15": tests(900).expected = "90123000000000000"
  tests(901).expr = "1.23456789e18": tests(901).expected = "1234567890000000000"
  tests(902).expr = "(9.0*10**18,9.2e17,9e18,90.123e15,1.23456789e18)": tests(902).expected = "(9000000000000000000, 920000000000000000, 9000000000000000000, 90123000000000000, 1234567890000000000)"
  tests(903).expr = "(9.2233e18,9.2234e18,0.123,0.123e3,0.12345e4,0.123e5,0.012345678901234e18,1.234567890123456e18,222.0,0)": tests(903).expected = "(9223300000000000000, 9.2234e+18, 0.123, 123, 1234.5, 12300, 12345678901234000, 1234567890123456000, 222, 0)"
  tests(904).expr = "0x10000000000000000": tests(904).expectedErrContains = "invalid hex literal"
  tests(905).expr = "0b10000000000000000000000000000000000000000000000000000000000000000": tests(905).expectedErrContains = "invalid binary literal"
  tests(906).expr = "0o2000000000000000000000": tests(906).expectedErrContains = "invalid octal literal"
  tests(907).expr = "1e": tests(907).expectedErrContains = "unexpected token"
  tests(908).expr = "1e+": tests(908).expectedErrContains = "unexpected token"
  tests(909).expr = "1e-": tests(909).expectedErrContains = "unexpected token"
  tests(910).expr = ".": tests(910).expectedErrContains = "unexpected token"
  tests(911).expr = ".0": tests(911).expected = "0"
  tests(912).expr = "0.": tests(912).expected = "0"
  tests(913).expr = "0.0": tests(913).expected = "0"
  tests(914).expr = "0xFFFFFFFFFFFFFFFF+2": tests(914).expected = "1.844674407370955e+019"
  tests(915).expr = "0xFFFFFFFFFFFFFFFF-2": tests(915).expected = "18446744073709551613"
  tests(916).expr = "0xFFFFFFFFFFFFFFFF*2": tests(916).expected = "3.689348814741911e+019"
  tests(917).expr = "0xFFFFFFFFFFFFFFFF/2": tests(917).expected = "9.223372036854776e+018"
  tests(918).expr = "0xFFFFFFFFFFFFFFFF%2": tests(918).expected = "1"
  tests(919).expr = "0xFFFFFFFFFFFFFFFF>>1": tests(919).expected = "9223372036854775807"
  tests(920).expr = "0xFFFFFFFFFFFFFFFF<<1": tests(920).expected = "3.689348814741911e+019"
  tests(921).expr = "0xFFFFFFFFFFFFFFFF**2": tests(921).expected = "3.402823669209385e+038"
  tests(922).expr = "0x7FFFFFFFFFFFFFFF**2": tests(922).expected = "8.507059173023462e+037"
  tests(923).expr = "-0xFFFFFFFFFFFFFFFF+2": tests(923).expected = "-1.844674407370955e+019"
  tests(924).expr = "-0xFFFFFFFFFFFFFFFF-2": tests(924).expected = "-1.844674407370955e+019"
  tests(925).expr = "(-0xFFFFFFFFFFFFFFFF)%2": tests(925).expectedErrContains = "modulo operands must be integer values"
  tests(926).expr = "(-0xFFFFFFFFFFFFFFFF)>>1": tests(926).expectedErrContains = "bitwise operands must be integer values"
  tests(927).expr = "-0x7FFFFFFFFFFFFFFF+2": tests(927).expected = "-9223372036854775805"
  tests(928).expr = "0x7FFFFFFFFFFFFFFF+2": tests(928).expected = "9223372036854775809"
  tests(929).expr = "0x7FFFFFFFFFFFFFFF-2": tests(929).expected = "9223372036854775805"
  tests(930).expr = "0x7FFFFFFFFFFFFFFF<<1": tests(930).expected = "18446744073709551614"
  tests(931).expr = "a=(-0xFFFFFFFFFFFFFFFF,-0x7FFFFFFFFFFFFFFF,0x7FFFFFFFFFFFFFFF); mod(a,2)": tests(931).expectedErrContains = "mod() expects integer values"
  tests(932).expr = "a=(-0x7FFFFFFFFFFFFFFF,0x7FFFFFFFFFFFFFFF); a>>1": tests(932).expected = "(-4611686018427387904, 4611686018427387903)"
  tests(933).expr = "a=(0x7FFFFFFFFFFFFFFF,0x7FFFFFFFFFFFFFFF); b=(2,2); a+b": tests(933).expected = "(9223372036854775809, 9223372036854775809)"
  tests(934).expr = "-0x7FFFFFFFFFFFFFFF-2": tests(934).expected = "-9.223372036854776e+018"
  tests(935).expr = "0x7FFFFFFFFFFFFFFF*2": tests(935).expected = "18446744073709551614"
  tests(936).expr = "a=(-0xFFFFFFFFFFFFFFFF,-0x7FFFFFFFFFFFFFFF,0x7FFFFFFFFFFFFFFF); b=(2,2,2); a+b": tests(936).expected = "(-1.844674407370955e+019, -9223372036854775805, 9223372036854775809)"
  tests(937).expr = "a=(-0xFFFFFFFFFFFFFFFF,-0x7FFFFFFFFFFFFFFF,0x7FFFFFFFFFFFFFFF); b=(2,2,2); a*b": tests(937).expected = "(-3.689348814741911e+019, -1.844674407370955e+019, 18446744073709551614)"
  tests(938).expr = "a=(-0xFFFFFFFFFFFFFFFF,-0x7FFFFFFFFFFFFFFF,0x7FFFFFFFFFFFFFFF); a>>1": tests(938).expectedErrContains = "bitwise operands must be integer values"
  tests(939).expr = "0x7FFFFFFFFFFFFFFF<<2": tests(939).expected = "3.68934881474191e+019"
  tests(940).expr = "0x7FFFFFFFFFFFFFFF>>1": tests(940).expected = "4611686018427387903"
  tests(941).expr = "a=(0xFFFFFFFFFFFFFFFF,0x7FFFFFFFFFFFFFFF,0x7FFFFFFFFFFFFFF); a>>1": tests(941).expected = "(9223372036854775807, 4611686018427387903, 288230376151711743)"
  tests(942).expr = "a=(0xFFFFFFFFFFFFFFFF,0x7FFFFFFFFFFFFFFF,0x7FFFFFFFFFFFFFF); a<<1": tests(942).expected = "(3.68934881474191e+19, 18446744073709551614, 1152921504606846974)"
  tests(943).expr = "3.68934881474191e+19>>1": tests(943).expectedErrContains = "bitwise operands must be integer values"
  tests(944).expr = "3.68934881474191e+19/2": tests(944).expected = "1.844674407370955e+019"
  tests(945).expr = "3.68934881474191e+19**0.5": tests(945).expected = "6074000999.952098"
  tests(946).expr = "3.68934881474191e+19%3": tests(946).expectedErrContains = "modulo operands must be integer values"
  tests(947).expr = "sum(18446744073709551614,1)": tests(947).expected = "18446744073709551615"
  tests(948).expr = "max(3,18446744073709551615)": tests(948).expected = "18446744073709551615"
  tests(949).expr = "min(18446744073709551615,5)": tests(949).expected = "5"
  tests(950).expr = "sum(-9223372036854775807,-1)": tests(950).expected = "-9223372036854775808"
  tests(951).expr = "max(-9,-1)": tests(951).expected = "-1"
  tests(952).expr = "int(0xFFFFFFFFFFFFFFFF+1)": tests(952).expected = "1.844674407370955e+019"
  tests(953).expr = "trunc(0xFFFFFFFFFFFFFFFF+1)": tests(953).expected = "1.844674407370955e+019"
  tests(954).expr = "floor(0xFFFFFFFFFFFFFFFF+1)": tests(954).expected = "1.844674407370955e+019"
  tests(955).expr = "ceil(0xFFFFFFFFFFFFFFFF+1)": tests(955).expected = "1.844674407370955e+019"
  tests(956).expr = "uhex(product(18446744073709551615,1))": tests(956).expected = "0xFFFFFFFFFFFFFFFF"
  tests(957).expr = "prod(-9223372036854775807,1)&1": tests(957).expected = "1"
  tests(958).expr = "uhex(product(-9223372036854775807-1,-1))": tests(958).expected = "0x8000000000000000"
  tests(959).expr = "uhex(2**63)": tests(959).expected = "0x8000000000000000"
  tests(960).expr = "uhex(pow(2,63))": tests(960).expected = "0x8000000000000000"
  tests(961).expr = "uhex(pow((2,3),(63,2)))": tests(961).expected = "(0x8000000000000000,0x9)"
  tests(962).expr = "uhex((-2)**63)": tests(962).expected = "0x8000000000000000"
  tests(963).expr = "sort(0xFFFFFFFFFFFFFFFF,nan,0x7FFFFFFFFFFFFFFF);hex": tests(963).expected = "(nan,0x7FFFFFFFFFFFFFFF,0xFFFFFFFFFFFFFFFF)"
  tests(964).expr = "sort(-5,nan,3,-inf,inf,0)": tests(964).expected = "(nan,-inf,-5,0,3,inf)"
  tests(965).expr = "sort(nan,inf,2,-inf,nan,-2)": tests(965).expected = "(nan,nan,-inf,-2,2,inf)"
  tests(966).expr = "sorted((nan,-3,4,inf,-inf))": tests(966).expected = "(nan,-inf,-3,4,inf)"
  tests(967).expr = "x=0.5; f(x)=x<<2; f(3)": tests(967).expected = "12"
  tests(968).expr = "x=0.5; f(x)=x<<2; x": tests(968).expected = "0.5"
  tests(969).expr = "x=0; f(x)=1/x; f(2)": tests(969).expected = "0.5"
  tests(970).expr = "f(x)=1%x; f(7)": tests(970).expected = "1"
  tests(971).expr = "_": tests(971).expected = "1"
  tests(972).expr = "_+5": tests(972).expected = "6"
  tests(973).expr = "_=10; g(x)=1%(x-1); g(5)": tests(973).expected = "1"
  tests(974).expr = "f(x)=x*newfn(x)": tests(974).expected = "defined f(x)"
  tests(975).expr = "f(2)": tests(975).expectedErrContains = "unknown function: newfn"
  tests(976).expr = "newfn(x)=x**(1/3); f(8)": tests(976).expected = "16"
  tests(977).expr = "sin(0)==0": tests(977).expected = "1" ' [syntax] builtin call + == not UDF
  tests(978).expr = "x=0; sin(x)==x": tests(978).expected = "1" ' [syntax] same as user report
  tests(979).expr = "f(t)=t*3; f(2)==6": tests(979).expected = "1" ' [syntax] UDF then ==
  tests(980).expr = "0:60": tests(980).expected = "01:00" ' [time] MM:SS carry
  tests(981).expr = "1:30 + 2:45.111": tests(981).expected = "04:15.111" ' [time] add with fractional seconds
  tests(982).expr = "second + 5": tests(982).expected = "00:06" ' [time] constant + seconds
  tests(983).expr = "minute - second": tests(983).expected = "00:59" ' [time] constants
  tests(984).expr = "1:00 == 0:60": tests(984).expected = "1" ' [time] compare equal durations
  tests(985).expr = "0:30 > 20": tests(985).expected = "1" ' [time] compare scalar seconds
  tests(986).expr = "milliseconds(minute + 30*second)": tests(986).expected = "90000" ' [time] converter
  tests(987).expr = "seconds(2:00)": tests(987).expected = "120" ' [time] HH:MM as two-seg is invalid... 2:00 is MM:SS = 2 min = 120000 ms -> seconds 120. Good
  tests(988).expr = "minutes(0:45)": tests(988).expected = "0.75" ' [time]
  tests(989).expr = "hours(12:00:00)": tests(989).expected = "12" ' [time] 12 hours literal
  tests(990).expr = "days(12:00:00)": tests(990).expected = "0.5" ' [time]
  tests(991).expr = "0:25 * 6": tests(991).expected = "02:30" ' [time] multiply unitless
  tests(992).expr = "1:30 / 0:30": tests(992).expected = "3" ' [time] ratio
  tests(993).expr = "1::0": tests(993).expectedErrContains = "empty segment" ' [time] error
  tests(994).expr = "second=1": tests(994).expectedErrContains = "reserved constant name" ' [time] reserved
  tests(995).expr = "minute=1": tests(995).expectedErrContains = "reserved constant name"
  tests(996).expr = "hour=1": tests(996).expectedErrContains = "reserved constant name"
  tests(997).expr = "day=1": tests(997).expectedErrContains = "reserved constant name"
  tests(998).expr = "sin(minute)": tests(998).expectedErrContains = "incompatible operands" ' [time] reject
  tests(999).expr = "milliseconds(5)": tests(999).expectedErrContains = "milliseconds() expects a time value" ' [time] converter non-time
  tests(1000).expr = "sum(0:30,1:00)": tests(1000).expected = "01:30" ' [time] aggregate
  tests(1001).expr = "1:00*1:00": tests(1001).expectedErrContains = "incompatible operands" ' [time] no duration*duration
  tests(1002).expr = "product(1:00,0:30)": tests(1002).expectedErrContains = "incompatible operands" ' [time] product needs unitless factors
  tests(1003).expr = "prod(1:00)": tests(1003).expectedErrContains = "incompatible operands" ' [time] prod disallowed with duration
  tests(1004).expr = "product(1:00)": tests(1004).expectedErrContains = "incompatible operands" ' [time] product disallowed with duration
  tests(1005).expr = "ans(x)=x*x": tests(1005).expectedErrContains = "reserved built-in variable name" ' [syntax] UDF name collides with ans
  tests(1006).expr = "_(x)=x*x": tests(1006).expectedErrContains = "reserved built-in variable name" ' [syntax] UDF name collides with formal probe _
  tests(1007).expr = "ANS(x)=x": tests(1007).expectedErrContains = "reserved built-in variable name" ' [syntax] ans reserved case-insensitively for UDF
  tests(1008).expr = "milliseconds((0:01:00,0:02:00))": tests(1008).expected = "(60000, 120000)" ' [time] converter element-wise int array
  tests(1009).expr = "seconds((0:30,1:00))": tests(1009).expected = "(30, 60)" ' [time] converter element-wise float tuple
  tests(1010).expr = "hours((1:00:00,0:30:00))": tests(1010).expected = "(1, 0.5)" ' [time] hours array
  tests(1011).expr = "minutes((0:30,1:00))": tests(1011).expected = "(0.5, 1)" ' [time] minutes array
  tests(1012).expr = "milliseconds((5,6))": tests(1012).expectedErrContains = "milliseconds() expects a time value" ' [time] array element must be duration
  tests(1013).expr = "1d2h3m4s5ms": tests(1013).expected = "1:02:03:04.005" ' [time] compact suffix literal
  tests(1014).expr = "1d 2h 3m 4s 5ms": tests(1014).expected = "1:02:03:04.005" ' [time] compact with spaces between fields
  tests(1015).expr = "1d3m + 2h5ms + 4s == 1:02:03:04.005": tests(1015).expected = "1" ' [time] compact reorder with +
  tests(1016).expr = "5000ms": tests(1016).expected = "00:05" ' [time] compact ms carry
  tests(1017).expr = "2h90m": tests(1017).expected = "03:30:00" ' [time] compact field overflow
  tests(1018).expr = "23h3600s": tests(1018).expected = "1:00:00:00" ' [time] compact to day
  tests(1019).expr = "-1m 1s": tests(1019).expected = "-01:01" ' [time] unary minus on compact literal
  tests(1020).expr = "1d2d": tests(1020).expectedErrContains = "compact time literal: unit order or duplicate unit"
  tests(1021).expr = "1h9d": tests(1021).expectedErrContains = "compact time literal: unit order or duplicate unit"
  tests(1022).expr = "45s37m": tests(1022).expectedErrContains = "compact time literal: unit order or duplicate unit"
  tests(1023).expr = "1day": tests(1023).expectedErrContains = "compact time literal: invalid suffix"
  tests(1024).expr = "1e2": tests(1024).expected = "100" ' [time] scientific e not compact ms
  tests(1025).expr = "milliseconds(1.5*1ms)": tests(1025).expected = "2" ' [time] fractional scale matches colon literal
  tests(1026).expr = "milliseconds(1.5*00:00.001)": tests(1026).expected = "2"
  tests(1027).expr = "seconds(-1s-1m)": tests(1027).expected = "-61" ' [time] unary vs binary minus
  tests(1028).expr = "1d == 1:00:00:00": tests(1028).expected = "1" ' [time] compact equals colon form
  tests(1029).expr = "1d2": tests(1029).expectedErrContains = "compact time literal: expected unit suffix"
  tests(1030).expr = "1ms1m": tests(1030).expectedErrContains = "compact time literal: unit order or duplicate unit" ' [time] ms must follow s
  tests(1031).expr = "(2+3x)": tests(1031).expectedErrContains = "unexpected token" ' [syntax] junk after inner expr, not ``)`` missing
  tests(1032).expr = "(2+3t)": tests(1032).expectedErrContains = "unexpected token" ' [syntax] same for letter after number in parens
  tests(1033).expr = "20m or 10h": tests(1033).expected = "1" ' [time] logical or after compact suffix (not ``invalid suffix``)
  tests(1034).expr = "20m10s and 10h5m": tests(1034).expected = "1" ' [time] logical and after compact suffix
  tests(1035).expr = "1h and not 0": tests(1035).expected = "1" ' [time] ``not`` keyword after compact + ``and``
  tests(1036).expr = "real(5)":           tests(1036).expected = "5" ' [ok-func] complex component unary on real
  tests(1037).expr = "imag(-3.5)":        tests(1037).expected = "0" ' [ok-func]
  tests(1038).expr = "phase(42)":         tests(1038).expected = "0" ' [ok-func]
  tests(1039).expr = "polar(10)":         tests(1039).expected = "(10, 0)" ' [ok-func]
  tests(1040).expr = "conj(-7)":          tests(1040).expected = "-7" ' [ok-func]
  tests(1041).expr = "cart(3.5)":         tests(1041).expected = "3.5" ' [ok-func]
  tests(1042).expr = "cart(polar(2))":    tests(1042).expected = "2" ' [ok-func]
  tests(1043).expr = "imag((5,10))":      tests(1043).expected = "(0, 0)" ' [ok-array]
  tests(1044).expr = "cart((5,0))":       tests(1044).expected = "5" ' [ok-func] polar cart with zero angle
  tests(1045).expr = "cart((5,1))":       tests(1045).expectedErrContains = "incompatible operands" ' [ok-func]
  tests(1046).expr = "sortby((3,-1,2), abs)": tests(1046).expected = "(-1,2,3)" ' [ok-func]
  tests(1047).expr = "sortby((5,1,5,2,2,9,1), abs)": tests(1047).expected = "(1,1,2,2,5,5,9)" ' [ok-func]
  tests(1048).expr = "sortby((), abs)": tests(1048).expectedErrContains = "expects at least 1 argument" ' [arity]
  tests(1049).expr = "f(x)=x*x; sortby((3,1,2), f)": tests(1049).expected = "(1,2,3)" ' [ok-func]
  tests(1050).expr = "sortby((1,2), abs())": tests(1050).expectedErrContains = "sortby expects exactly one function" ' [err]
  tests(1051).expr = "sortby((1,2), pow)": tests(1051).expectedErrContains = "sortby expects a function that takes 1 parameter" ' [err]
  tests(1052).expr = "sortby((2,1), polar)": tests(1052).expected = "(1,2)"                 ' tuple keys (r, angle), lexicographic order
  tests(1053).expr = "sortby((1:30,0:30,1:00), milliseconds)": tests(1053).expected = "(00:30,01:00,01:30)" ' [time]
  tests(1054).expr = "ratio(5)": tests(1054).expected = "5" ' [ok-func]
  tests(1055).expr = "ratio(0)": tests(1055).expected = "0" ' [ok-func]
  tests(1056).expr = "ratio(0.5)": tests(1056).expected = "1/2" ' [ok-func]
  tests(1057).expr = "ratio(0.3333333333333333)": tests(1057).expected = "1/3" ' [ok-func]
  tests(1058).expr = "ratio(nan)": tests(1058).expected = "nan" ' [ok-func]
  tests(1059).expr = "ratio(inf)": tests(1059).expected = "inf" ' [ok-func]
  tests(1060).expr = "ratio((1,2,3))": tests(1060).expected = "(1,2,3)" ' [ok-array]
  tests(1061).expr = "ratio(1:00)": tests(1061).expectedErrContains = "incompatible operands" ' [time]
  tests(1062).expr = "ratio(hours(1:00))": tests(1062).expected = "1/60" ' [time]
  tests(1063).expr = "ratio(1e-7)": tests(1063).expected = "1/10000000" ' [ok-func]
  tests(1064).expr = "ratio(1e-8)": tests(1064).expected = "1/100000000" ' [ok-func]
  tests(1065).expr = "ratio(sqrt(2))": tests(1065).expected = "13250218/9369319" ' [ok-func]
  tests(1066).expr = "ratio((sqrt(2), e, pi))": tests(1066).expected = "(13250218/9369319, 14665106/5394991, 5419351/1725033)" ' [ok-array]
  tests(1067).expr = "f(x)=1/x; sortby((0:30, 1:00, 0:45), f)": tests(1067).expectedErrContains = ", |f)" ' [error-snippet] key fn failure -> sortby key arg
  tests(1068).expr = "x=10; sortby((0:30, 1:00, 0:45), x)": tests(1068).expectedErrContains = "sortby expects a function that takes 1 parameter" ' [err] variable name, not function ref value
  tests(1069).expr = "f(x)=x*2; f=10; f(3)": tests(1069).expectedErrContains = "unknown function: f" ' [udf-var] assignment drops same-named user function
  tests(1070).expr = "f(x)=x*2; f=10; f": tests(1070).expected = "10" ' [udf-var]
  tests(1071).expr = "f=10; f(x)=x*2; f": tests(1071).expectedErrContains = "user-defined function: f(x)" ' [udf-var] bare name hints signature
  tests(1072).expr = "f(x)=x+1; f": tests(1072).expectedErrContains = "user-defined function: f(x)" ' [udf-hint]
  tests(1073).expr = "g(a)=1/a; sortby((0:30, 1:00, 0:45), g)": tests(1073).expectedErrContains = ", |g)" ' [error-snippet] key fn failure -> sortby key arg
  tests(1074).expr = "a(x)=1/a; sortby((0:30, 1:00, 0:45), a)": tests(1074).expectedErrContains = "unknown variable: a" ' [sortby] body uses a but param is x; not UDF hint
  tests(1075).expr = "a(x)=1/x; sortby((0:30, 1:00, 0:45), a)": tests(1075).expectedErrContains = "incompatible operands" ' [sortby] duration key fn
  tests(1076).expr = "ratio(())": tests(1076).expectedErrContains = "expects at least 1 argument" ' [arity]
  tests(1077).expr = "ratio(0.14285714285714285)": tests(1077).expected = "1/7" ' [ok-func] repeating decimal
  tests(1078).expr = "ratio(-0.3333333333333333)": tests(1078).expected = "-1/3" ' [ok-func] negative
  tests(1079).expr = "ratio(1234567)": tests(1079).expected = "1234567" ' [ok-func] exact integer
  tests(1080).expr = "ratio(0.0000001)": tests(1080).expected = "1/10000000" ' [ok-func] power-of-10 small
  tests(1081).expr = "ratio(0.4142135623730951)": tests(1081).expected = "3880899/9369319" ' [ok-func] semiconvergent
  tests(1082).expr = "ratio((0.14285714285714285, sqrt(2)))": tests(1082).expected = "(1/7, 13250218/9369319)" ' [ok-array]
  tests(1083).expr = "ratio(0.9999999)": tests(1083).expected = "9999999/10000000" ' [ok-func] near-one large denom
  tests(1084).expr = "ratio(0.123456789012345)": tests(1084).expected = "10/81" ' [ok-func] modest denominator
  tests(1085).expr = "ratio(1/7)": tests(1085).expected = "1/7" ' [ok-func] exact rational input
  tests(1086).expr = "0xFFFFFFFFFFFFFFFF+0": tests(1086).expected = "18446744073709551615" ' [hex-exact] uint64 metadata preserved
  tests(1087).expr = "0x7FFFFFFFFFFFFFFF+1": tests(1087).expected = "9223372036854775808" ' [hex-exact] int64-range hex add
  tests(1088).expr = "uhex(0xFFFFFFFFFFFFFFFF+0)": tests(1088).expected = "0xFFFFFFFFFFFFFFFF" ' [hex-exact] uhex on exact uint result
  tests(1089).expr = "100000000000000000000": tests(1089).expected = "1.0000000000000000e+20" ' [numeric] long decimal integer, not compact time
  tests(1090).expr = "d(x)=(ratio(x), x-ratio(x)); d(seconds(20ms))": tests(1090).expected = "(1/50, 0)"
  tests(1091).expr = "f(x)=x*(2,3,4); f(10)": tests(1091).expected = "(20, 30, 40)"
  tests(1092).expr = "1+2;": tests(1092).expected = "3"                   ' trailing semicolon (top-level)
  tests(1093).expr = "   1+2 ; ": tests(1093).expected = "3"
  tests(1094).expr = ";": tests(1094).expectedErrContains = "empty statement"
  tests(1095).expr = " ; ": tests(1095).expectedErrContains = "empty statement"
  tests(1096).expr = "1;;2": tests(1096).expectedErrContains = "empty statement"  ' interior empty stmt
  tests(1097).expr = "f(x)=x*(10,20); sortby((3,1,2), f)": tests(1097).expected = "(1,2,3)" ' UDF returns array key
  tests(1098).expr = "f=x:x+2; f(3)": tests(1098).expected = "5"                               ' lambda UDF
  tests(1099).expr = "f=(x):x+2; f(4)": tests(1099).expected = "6"                           ' parens around parameter
  tests(1100).expr = "f=():42; f()": tests(1100).expected = "42"                            ' zero-parameter lambda / UDF
  tests(1101).expr = "sortby((3,1,2), x:-x)": tests(1101).expected = "(3,2,1)"
  tests(1102).expr = "sortby((5,1), ():1)": tests(1102).expectedErrContains = "sortby expects a function that takes 1 parameter" ' [err] zero-param lambda in sortby
  tests(1103).expr = "sortby((1,2), x,y:x+y)": tests(1103).expectedErrContains = "sortby expects a function that takes 1 parameter" ' [err] multi-param lambda in sortby
  tests(1104).expr = "g=((x):(1/x)); g(2)": tests(1104).expected = "0.5"                     ' nested wrapping on rhs
  tests(1105).expr = "sortby((2,1), (x):(1/x))": tests(1105).expected = "(2,1)"
  tests(1106).expr = "sortby((1:30,1:00),x:)": tests(1106).expectedErrContains = "function body is empty" ' [err] empty sortby lambda body
  tests(1107).expr = "qqq(x)=polar": tests(1107).expectedErrContains = "function:" ' [udf] bare builtin at body tail
  tests(1108).expr = "qqq=x:polar": tests(1108).expectedErrContains = "function:" ' [udf] lambda bare builtin at tail
  tests(1109).expr = "qqq(x)=polar+x": tests(1109).expectedErrContains = "unknown variable: polar" ' [udf] bare builtin mid-body
  tests(1110).expr = "qqq(x)=sin+cos": tests(1110).expectedErrContains = "unknown variable: sin" ' [udf] bare builtin mid-body
  tests(1111).expr = "polar+123": tests(1111).expectedErrContains = "unknown variable: polar" ' [err] bare builtin mid-expr
  tests(1112).expr = "sortby((1:30,1:00),x:polar)": tests(1112).expectedErrContains = "unknown variable: polar" ' [sortby] lambda body tail eval
  tests(1113).expr = "sortby((1:30,1:00),x:polar+2)": tests(1113).expectedErrContains = "unknown variable: polar" ' [sortby] lambda body mid
  tests(1114).expr = "1+2; polar+2": tests(1114).expectedErrContains = "unknown variable: polar" ' [err] statement mid-expr
  tests(1115).expr = "h(x)=x+1; h+2": tests(1115).expectedErrContains = "unknown variable: h" ' [udf] bare name mid-expr
  tests(1116).expr = "h(x)=h+1": tests(1116).expectedErrContains = "unknown variable: h" ' [udf] bare name mid-body
  tests(1117).expr = "g(x)=x; sortby((3,1,2),x:g)": tests(1117).expectedErrContains = "unknown variable: g" ' [sortby] lambda body tail eval
  tests(1118).expr = "g(x)=x; sortby((3,1,2),x:g+1)": tests(1118).expectedErrContains = "unknown variable: g" ' [sortby] lambda body mid
  tests(1119).expr = "f=x,y:x+y; f(1,2)": tests(1119).expected = "3" ' [udf] lambda two unwrapped params
  tests(1120).expr = "f=x,y,z:x+y+z; f(1,2,3)": tests(1120).expected = "6" ' [udf] lambda three unwrapped params
  tests(1121).expr = "f=(x,y,z):(x+y+z); f(1,2,3)": tests(1121).expected = "6" ' [udf] lambda paren param list
  tests(1122).expr = "f=():100; f()": tests(1122).expected = "100" ' [udf] lambda zero params
  tests(1123).expr = "_=(1,2,3,4); ff(a)=(a[1],a[-1]); ff(_)": tests(1123).expected = "(2,4)" ' [udf] formal array indexing
  tests(1124).expr = "t=(1,2,3,4,5,6); fff(a)=(a[0],a[-3]); fff(t)": tests(1124).expected = "(1,4)" ' [udf] formal index bounds deferred to call
  tests(1125).expr = "f=():()": tests(1125).expectedErrContains = "function body is empty" ' [udf] reject empty tuple body
  tests(1126).expr = "f()=( )": tests(1126).expectedErrContains = "function body is empty" ' [udf] reject whitespace-only tuple body
  tests(1127).expr = "f=():1+99; f()": tests(1127).expected = "100" ' [udf] zero-param lambda body still allowed (distinct from 1122)
  tests(1128).expr = "atan2;": tests(1128).expectedErrContains = "unexpected token" ' [err] bare builtin before semicolon
  tests(1129).expr = "atan2;(2,3)": tests(1129).expectedErrContains = "unexpected token" ' [err] first statement bare builtin;
  tests(1130).setup = "nnnn=5": tests(1130).expr = "atan2;": tests(1130).expectedErrContains = "unexpected token" ' [err] error location not from prior expr
  tests(1131).setup = "nnnn=5": tests(1131).expr = "atan2;(2,3)": tests(1131).expectedErrContains = "unexpected token" ' [err] error location not from prior expr
  tests(1132).expr = "ratio(1,2)": tests(1132).expectedErrContains = "expects 1 argument(s)" ' [arity-table]
  tests(1133).expr = "sortby()": tests(1133).expectedErrContains = "sortby expects a function that takes 1 parameter" ' [arity-table] parse-time
  tests(1134).expr = "sortby(1)": tests(1134).expectedErrContains = "sortby expects a function that takes 1 parameter" ' [arity-table] parse-time
  tests(1135).expr = "sortby(1,2,3)": tests(1135).expectedErrContains = "sortby expects" ' [arity-table] parse-time (exactly one function or unary function)
  tests(1136).expr = "lcm()": tests(1136).expectedErrContains = "expects 2 argument(s)" ' [arity-table]
  tests(1137).expr = "npr()": tests(1137).expectedErrContains = "expects 2 argument(s)" ' [arity-table]
  tests(1138).expr = "random(1)": tests(1138).expectedErrContains = "expects 2 argument(s)" ' [arity-table]
  tests(1139).expr = "clamp(1)": tests(1139).expectedErrContains = "expects 3 argument(s)" ' [arity-table]
  tests(1140).expr = "fact(5,6)": tests(1140).expectedErrContains = "expects 1 argument(s)" ' [arity-table]
  tests(1141).expr = "deg(1,2)": tests(1141).expected = "(57.29577951308232, 114.5915590261647)" ' [arity-table] variadic deg
  tests(1142).expr = "hex(1,2,3)": tests(1142).expected = "(0x1,0x2,0x3)" ' [arity-table] variadic hex
  tests(1143).expr = "sum(1,2,3)": tests(1143).expected = "6" ' [arity-table] variadic sum
  tests(1144).expr = "unpack(1,2)": tests(1144).expected = "(1,2)" ' [arity-table] variadic unpack
  tests(1145).expr = "sin(1,2,3)": tests(1145).expectedErrContains = "expects 1 argument(s)" ' [arity-table]
  tests(1146).expr = "abs(0x7FFFFFFFFFFFFFFF+20)": tests(1146).expected = "9223372036854775827" ' [abs-exact-int]
  tests(1147).expr = "hex(abs(0x7FFFFFFFFFFFFFFF+20))": tests(1147).expected = "0x8000000000000013" ' [abs-exact-int]
  tests(1148).expr = "abs(-9223372036854775808)": tests(1148).expected = "9223372036854775808" ' [abs-exact-int] |INT64_MIN|
  tests(1149).expr = "uhex(abs((18446744073709551615,-7)))": tests(1149).expected = "(0xFFFFFFFFFFFFFFFF,0x7)" ' [abs-exact-int]
  tests(1150).expr = "hex(sqr(9))": tests(1150).expected = "0x51" ' [sqr-exact-int]
  tests(1151).expr = "sqr(3037000499)": tests(1151).expected = "9223372030926249001" ' [sqr-exact-int] large exact int square (differs from 3037000499*3037000499 float path)
  tests(1152).expr = "hypot(3037000499,0)": tests(1152).expected = "3037000499" ' [hypot-exact-int]
  tests(1153).expr = "sqr(1.5)": tests(1153).expected = "2.25" ' [sqr-float] non-exact input uses float path
  tests(1154).expr = "sqrt(4611686014132420611)*sqrt(4611686014132420611)": tests(1154).expected = "4611686014132420609" ' [sqrt-no-fake-exact] float sqrt; product != radicand
  tests(1155).expr = "sqr(sqrt(81))": tests(1155).expected = "81" ' [sqrt-exact-int] verified perfect square keeps int through sqr
  tests(1156).expr = "sqr(sqrt(4611686014132420611))": tests(1156).expected = "4611686014132420609" ' [sqrt-no-fake-exact] fake exact int would yield 2147483647^2
  tests(1157).expr = "125**(1/3)": tests(1157).expected = "5" ' [pow-exact-int] verified fractional root
  tests(1158).expr = "823543**(1/7)": tests(1158).expected = "7" ' [pow-exact-int] verified 7th root
  tests(1159).expr = "pow(27,1/3)": tests(1159).expected = "3" ' [pow-exact-int] builtin same as **
  tests(1160).expr = "pow(1,99)": tests(1160).expected = "1" ' [pow-exact-int] base 1 fast path
  tests(1161).expr = "2**10": tests(1161).expected = "1024" ' [pow-exact-int] integer exponent verify
  tests(1162).expr = "126**(1/3)": tests(1162).expected = "5.013297934964584" ' [pow-float] non-perfect cube stays float
  tests(1163).expr = "(-5)**3": tests(1163).expected = "-125" ' [pow-exact-int] negative base, odd integer exponent
  tests(1164).expr = "(-125)**(1/3)": tests(1164).expected = "-5" ' [pow-exact-int] negative base, odd root
  tests(1165).expr = "pow(-32,1/5)": tests(1165).expected = "-2" ' [pow-exact-int] negative base, odd root via pow()
  tests(1166).expr = "pow(pow(-3,39), 1/39)": tests(1166).expected = "-3" ' [pow-real] odd unit root, not principal complex
  tests(1167).expr = "pow(pow(-3,45), 1/45)": tests(1167).expected = "-3" ' [pow-real] odd unit root (regression: was complex)
  tests(1168).expr = "cos((pi/2, pi, -pi/2, -pi, 2*pi, 0, 77777*pi/2, -77777*pi/2))": tests(1168).expected = "(0, -1, 0, -1, 1, 1, 0, 0)" ' [trig-array] half-pi multiples incl. large N
  tests(1169).expr = "tan((pi, pi/2, -pi, 2*pi, -pi/2, 0, 77777*pi/2, -77777*pi/2))": tests(1169).expected = "(0, inf, 0, 0, -inf, 0, inf, -inf)" ' [trig-array] poles at odd half-pi
  tests(1170).expr = "fact(171)": tests(1170).expected = "inf" ' [fact-overflow] double factorial overflows
  tests(1171).expr = "fact(170000000)": tests(1171).expected = "inf" ' [fact-overflow] stop loop once non-finite
  tests(1172).expr = "avg(0:30, 1:30)": tests(1172).expected = "01:00" ' [time] aggregate avg
  tests(1173).expr = "mean(0:20, 1:00)": tests(1173).expected = "00:40" ' [time] aggregate mean
  tests(1174).expr = "min(0:45, 1:30)": tests(1174).expected = "00:45" ' [time] aggregate min
  tests(1175).expr = "max(0:45, 1:30)": tests(1175).expected = "01:30" ' [time] aggregate max
  tests(1176).expr = "sum((0:15, 0:15, 0:30))": tests(1176).expected = "01:00" ' [time] aggregate array
  tests(1177).expr = "sum(1, Inf, 2)": tests(1177).expected = "inf" ' [agg-inf] real sum
  tests(1178).expr = "prod(1, Inf, 2)": tests(1178).expected = "inf" ' [agg-inf] real product
  tests(1179).expr = "product(1, Inf, 2)": tests(1179).expected = "inf" ' [agg-inf] real product alias
  tests(1180).expr = "avg(1, Inf, 2)": tests(1180).expected = "inf" ' [agg-inf] real average
  tests(1181).expr = "mean(1, Inf, 2)": tests(1181).expected = "inf" ' [agg-inf] real mean
  tests(1182).expr = "min(1, Inf, 2)": tests(1182).expected = "1" ' [agg-inf] real min
  tests(1183).expr = "max(1, Inf, 2)": tests(1183).expected = "inf" ' [agg-inf] real max
  tests(1184).expr = "median(1, Inf, 2)": tests(1184).expected = "2" ' [agg-inf] real median
  tests(1185).expr = "variance(1, Inf, 2)": tests(1185).expected = "nan" ' [agg-inf] real variance
  tests(1186).expr = "stddev(1, Inf, 2)": tests(1186).expected = "nan" ' [agg-inf] real stddev
  tests(1187).expr = "sort(1, Inf, 2)": tests(1187).expected = "(1, 2, inf)" ' [agg-inf] real sort
  tests(1188).expr = "sorted(1, Inf, 2)": tests(1188).expected = "(1, 2, inf)" ' [agg-inf] real sorted alias
  tests(1189).expr = "reverse(1, Inf, 2)": tests(1189).expected = "(2, inf, 1)" ' [agg-inf] real reverse
  tests(1190).expr = "reversed(1, Inf, 2)": tests(1190).expected = "(2, inf, 1)" ' [agg-inf] real reversed alias
  tests(1191).expr = "unique(1, Inf, 2, Inf)": tests(1191).expected = "(1, inf, 2)" ' [agg-inf] real unique
  tests(1192).expr = "unpack(1, Inf, 2)": tests(1192).expected = "(1, inf, 2)" ' [agg-inf] real unpack
  tests(1193).expr = "sum(unpack(1, Inf, 2))": tests(1193).expected = "inf" ' [agg-inf] unpack then sum
  tests(1194).expr = "sortby((1, Inf, 2), abs)": tests(1194).expected = "(1, 2, inf)" ' [agg-inf] sortby with Inf

  dim smokeAi as Integer
  for smokeAi = 1 to SMOKE_APPEND_COUNT
    dim ti as Integer = 1194 + smokeAi
    tests(ti).expr = smokeAppendExpr(smokeAi)
    if smokeAppendIsErr(smokeAi) then
      tests(ti).expectedErrContains = smokeAppendErr(smokeAi)
    else
      tests(ti).expected = smokeAppendExpected(smokeAi)
    end if
  next smokeAi

  tests(1209).expr = "factorint(33)": tests(1209).expected = "(3, 11)" ' [factorint]
  tests(1210).expr = "factorint(12)": tests(1210).expected = "(2**2, 3)" ' [factorint]
  tests(1211).expr = "factorint(13)": tests(1211).expected = "(13)" ' [factorint]
  tests(1212).expr = "factorint(-33)": tests(1212).expected = "(-3, 11)" ' [factorint]
  tests(1213).expr = "factorint(-13)": tests(1213).expected = "(-13)" ' [factorint]
  tests(1214).expr = "factorint(-12)": tests(1214).expected = "(-2**2, 3)" ' [factorint]
  tests(1215).expr = "factorint(0)": tests(1215).expected = "(0)" ' [factorint]
  tests(1216).expr = "factorint(1)": tests(1216).expected = "(1)" ' [factorint]
  tests(1217).expr = "factorint(-1)": tests(1217).expected = "(-1)" ' [factorint]
  tests(1218).expr = "factorint(2**52)": tests(1218).expected = "(2**52)" ' [factorint]
  tests(1219).expr = "factorint(2**63-1)": tests(1219).expected = "(7**2, 73, 127, 337, 92737, 649657)" ' [factorint]
  tests(1220).expr = "factorint((33,12))": tests(1220).expectedErrContains = "expects scalar values" ' [factorint]
  tests(1221).expr = "factorint(2**64)": tests(1221).expectedErrContains = "expects integer values" ' [factorint]
  tests(1222).expr = "factorint(18446744073709551615)": tests(1222).expected = "(3, 5, 17, 257, 641, 65537, 6700417)" ' [factorint]
  tests(1223).expr = "prod(factorint(12))": tests(1223).expected = "12" ' [factorint] prod inverse
  tests(1224).expr = "prod(factorint(-33))": tests(1224).expected = "-33" ' [factorint] prod inverse
  tests(1225).expr = "prod(factorint(2**52))": tests(1225).expected = "4503599627370496" ' [factorint] prod inverse
  tests(1226).expr = "factorint(33.0)": tests(1226).expected = "(3, 11)" ' [factorint]
  tests(1227).expr = "factorint(33.5)": tests(1227).expectedErrContains = "expects integer values" ' [factorint]
  tests(1228).expr = "factorint(90)": tests(1228).expected = "(2, 3**2, 5)" ' [factorint]
  tests(1229).expr = "factorint(9007)": tests(1229).expected = "(9007)" ' [factorint]
  tests(1230).expr = "factorint(900719)": tests(1230).expected = "(900719)" ' [factorint]
  tests(1231).expr = "factorint(90071992)": tests(1231).expected = "(2**3, 11258999)" ' [factorint]
  tests(1232).expr = "factorint(9007199254)": tests(1232).expected = "(2, 89, 50602243)" ' [factorint]
  tests(1233).expr = "factorint(900719925474)": tests(1233).expected = "(2, 3, 12907, 11630897)" ' [factorint]
  tests(1234).expr = "factorint(76568758722)": tests(1234).expected = "(2, 3**2, 47, 101, 896107)" ' [factorint]
  tests(1235).expr = "factorint(-3333*9)": tests(1235).expected = "(-3**3, 11, 101)" ' [factorint] signed prime power
  tests(1236).expr = "factorint(-9999)": tests(1236).expected = "(-3**2, 11, 101)" ' [factorint] signed prime power
  tests(1237).expr = "6*(1/2)": tests(1237).expected = "3" ' [exact-int] integer multiplied by fractional float
  tests(1238).expr = "6*(1/3)": tests(1238).expected = "2" ' [exact-int] integer multiplied by fractional float
  tests(1239).expr = "5*(1/3)": tests(1239).expected = "1.666666666666667" ' [exact-int] non-integer product stays float
  tests(1240).expr = "0xFFFFFFFFFFFFFFFF*(1/5)": tests(1240).expected = "3689348814741910323" ' [exact-int] uint64 * fractional float
  tests(1241).expr = "0x7FFFFFFFFFFFFFFF*(1/7)": tests(1241).expected = "1317624576693539401" ' [exact-int] int64 * fractional float
  tests(1242).expr = "2**64*(1/2)": tests(1242).expected = "9.223372036854776e+018" ' [exact-int] non-exact input stays float
  tests(1243).expr = "0x7FFFFFFFFFFFFFFF*(1/8)": tests(1243).expected = "1.152921504606847e+018" ' [exact-int] non-integer product stays float
  tests(1244).expr = "2**52*(1/3)": tests(1244).expected = "1.501199875790165e+015" ' [exact-int] non-integer product stays float
  tests(1245).expr = "-0x7FFFFFFFFFFFFFFF*(1/7)": tests(1245).expected = "-1317624576693539401" ' [exact-int] signed integer with positive reciprocal
  tests(1246).expr = "0x7FFFFFFFFFFFFFFF*(-1/7)": tests(1246).expected = "-1317624576693539401" ' [exact-int] signed integer with negative reciprocal
  tests(1247).expr = "-2**52*(1/2)": tests(1247).expected = "-2251799813685248" ' [exact-int] signed exact power-of-two halving
  tests(1248).expr = "2**52*(-1/2)": tests(1248).expected = "-2251799813685248" ' [exact-int] signed exact power-of-two halving
  tests(1249).expr = "1317624576693539401/(1/11)": tests(1249).expected = "14493870343628933411" ' [exact-int] integer divided by fractional float
  tests(1250).expr = "1317624576693539401/(1/7)": tests(1250).expected = "9223372036854775807" ' [exact-int] integer divided by fractional float
  tests(1251).expr = "-1317624576693539401/(1/7)": tests(1251).expected = "-9223372036854775807" ' [exact-int] signed integer divided by fractional float
  tests(1252).expr = "1317624576693539401/(-1/7)": tests(1252).expected = "-9223372036854775807" ' [exact-int] integer divided by negative fractional float
  tests(1253).expr = "0x7FFFFFFFFFFFFFFF/(1/8)": tests(1253).expected = "7.37869762948382e+019" ' [exact-int] overflow quotient stays float
  tests(1254).expr = "2**64/(1/2)": tests(1254).expected = "3.68934881474191e+019" ' [exact-int] overflow quotient stays float
  tests(1255).expr = "1317624576693539401*(1/11)": tests(1255).expected = "1.197840524266854e+017" ' [exact-int] non-integer product stays float
  tests(1256).expr = "0x7FFFFFFFFFFFFFFF*(-1/8)": tests(1256).expected = "-1.152921504606847e+018" ' [exact-int] non-integer product stays float
  tests(1257).expr = "1317624576693539401/(1/8)": tests(1257).expected = "10540996613548315208" ' [exact-int] large exact quotient via uint64
  tests(1258).expr = "5/(1/3)": tests(1258).expected = "15" ' [exact-int] small exact quotient
  tests(1259).expr = "-5/(1/3)": tests(1259).expected = "-15" ' [exact-int] signed small exact quotient
  tests(1260).expr = "5/(-1/3)": tests(1260).expected = "-15" ' [exact-int] signed divisor reciprocal
  tests(1261).expr = "100/(5/2)": tests(1261).expected = "40" ' [exact-int] abs(f)>=1 float divisor
  tests(1262).expr = "100/(10/3+1e-15)": tests(1262).expected = "29.99999999999999" ' [exact-int] near-integer quotient stays float
  tests(1263).expr = "-50/(5/2)": tests(1263).expected = "-20" ' [exact-int] abs(f)>=1 negative N
  tests(1264).expr = "-50/((10/3)+1e-9)": tests(1264).expected = "-14.9999999955" ' [exact-int] near-integer quotient stays float
  tests(1265).expr = "700000000/(5/2)": tests(1265).expected = "280000000" ' [exact-int] abs(f)>=1 large N
  tests(1266).expr = "700000000/(10/3+1e-15)": tests(1266).expected = "209999999.9999999" ' [exact-int] near-integer quotient stays float
  tests(1267).expr = "-300000000/(5/2)": tests(1267).expected = "-120000000" ' [exact-int] abs(f)>=1 large negative N
  tests(1268).expr = "-300000000/(10/3+1e-15)": tests(1268).expected = "-89999999.99999997" ' [exact-int] near-integer quotient stays float
  tests(1269).expr = "(2**51-8)/(5/2)": tests(1269).expected = "900719925474096" ' [exact-int] abs(f)>=1 near 2^53
  tests(1270).expr = "(2**51-8)/((5/2)+1e-12)": tests(1270).expected = "900719925473735.6" ' [exact-int] near-integer quotient stays float
  tests(1271).expr = "-(2**50-1)/(3/2)": tests(1271).expected = "-750599937895082" ' [exact-int] abs(f)>=1 negative N
  tests(1272).expr = "-(2**50-1)/((3/2)+1e-12)": tests(1272).expected = "-750599937894581.5" ' [exact-int] near-integer quotient stays float
  tests(1273).expr = "100*(5/2)": tests(1273).expected = "250" ' [exact-int] abs(f)>1 float factor
  tests(1274).expr = "100*((5/2)+1e-12)": tests(1274).expected = "250.0000000001" ' [exact-int] near-integer product stays float
  tests(1275).expr = "-50*(5/2)": tests(1275).expected = "-125" ' [exact-int] abs(f)>1 negative N
  tests(1276).expr = "-50*((5/2)+1e-9)": tests(1276).expected = "-125.00000005" ' [exact-int] near-integer product stays float
  tests(1277).expr = "700000000*(5/2)": tests(1277).expected = "1750000000" ' [exact-int] abs(f)>1 large N
  tests(1278).expr = "700000000*((5/2)+1e-12)": tests(1278).expected = "1750000000.0007" ' [exact-int] near-integer product stays float
  tests(1279).expr = "-300000000*(5/2)": tests(1279).expected = "-750000000" ' [exact-int] abs(f)>1 large negative N
  tests(1280).expr = "-300000000*((5/2)+1e-12)": tests(1280).expected = "-750000000.0003" ' [exact-int] near-integer product stays float
  tests(1281).expr = "(2**51-8)*(5/2)": tests(1281).expected = "5629499534213100" ' [exact-int] abs(f)>1 near 2^53
  tests(1282).expr = "(2**51-8)*((5/2)+1e-12)": tests(1282).expected = "5629499534215352" ' [exact-int] near-integer product stays float
  tests(1283).expr = "-(2**50-1)*2": tests(1283).expected = "-2251799813685246" ' [exact-int] abs(f)>1 negative N
  tests(1284).expr = "-(2**50-1)*(2.0001)": tests(1284).expected = "-2251912403675931" ' [exact-int] near-integer product stays float
  tests(1285).expr = "6/(10,20,30)": tests(1285).expected = "(0.6, 0.3, 0.2)" ' array division
  tests(1286).expr = "ratio(6/(10,20,30))": tests(1286).expected = "(3/5, 3/10, 1/5)" ' array ratio
  tests(1287).expr = "sort(ratio(6/(10,20,30)))": tests(1287).expected = "(1/5, 3/10, 3/5)" ' sorted array ratio
  tests(1288).expr = "unique(abs((2,-1,2,3,-3,1)))": tests(1288).expected = "(2, 1, 3)" ' unique elements
  tests(1289).expr = "log10": tests(1289).expectedErrContains = "function: log10(value)" ' [hint] digit in builtin name
  tests(1290).expr = "log10+1": tests(1290).expectedErrContains = "unknown variable: log10" ' [hint] not bare at expression tail
  tests(1291).expr = "1+sin": tests(1291).expectedErrContains = "function: sin(angle)" ' [hint] bare builtin at expression end
  tests(1292).expr = "1+sum": tests(1292).expectedErrContains = "function: sum(...)" ' [hint] bare builtin at expression end
  tests(1293).expr = "(1+2)/sum": tests(1293).expectedErrContains = "function: sum(...)" ' [hint] bare builtin at expression end
  tests(1294).expr = "cos(pi/4)*sin(pi/4)-exp": tests(1294).expectedErrContains = "function: exp(value)" ' [hint] bare builtin at expression end
  tests(1295).expr = "1+sin(": tests(1295).expectedErrContains = "function: sin(angle)" ' [hint] incomplete call at expression end
  tests(1296).expr = "(1+2)/sum(": tests(1296).expectedErrContains = "function: sum(...)" ' [hint] incomplete call at expression end
  tests(1297).expr = "cos(pi/4)*sin(pi/4)-exp(": tests(1297).expectedErrContains = "function: exp(value)" ' [hint] incomplete call at expression end
  tests(1298).expr = "1+sin(rad(": tests(1298).expectedErrContains = "function: rad(...)" ' [hint] innermost incomplete call at end
  tests(1299).expr = "1+sin(rad(45)": tests(1299).expectedErrContains = "missing closing parenthesis" ' [syntax] outer sin call unclosed
  tests(1300).expr = "sortby((1,2,3, abs": tests(1300).expectedErrContains = "function: abs(value)" ' [hint] bare builtin in incomplete array at end
  tests(1301).expr = "f(x)=x; 1+f": tests(1301).expectedErrContains = "user-defined function: f(x)" ' [hint] bare UDF at expression end
  tests(1302).expr = "sortby((1,2,3), abs": tests(1302).expectedErrContains = "missing closing parenthesis" ' [syntax] sortby call unclosed after key ref
  tests(1303).expr = "sortby((1,2,3),x:x": tests(1303).expectedErrContains = "missing closing parenthesis" ' [sortby] inline lambda key at EOF
  tests(1304).expr = "sortby((1,2,3),x:x+sin": tests(1304).expectedErrContains = "missing closing parenthesis" ' [sortby] lambda body at EOF
  tests(1305).expr = "sortby((1,2,3),x:x+sin(": tests(1305).expectedErrContains = "missing closing parenthesis" ' [sortby] lambda body with open paren at EOF
  tests(1306).expr = "sortby((1,2,3),x:": tests(1306).expectedErrContains = "function body is empty" ' [sortby] empty lambda body at EOF
  tests(1307).expr = "sortby((1,2,3),x:x+sin)": tests(1307).expectedErrContains = "unknown variable: sin" ' [sortby] lambda body tail eval
  tests(1308).expr = "sortby((1,2,3),x:x:x)": tests(1308).expectedErrContains = "at col 19" ' [sortby] lambda body error column
  tests(1309).expr = "sortby((1:30,1:00),x:polar+2)": tests(1309).expectedErrContains = "at col 22" ' [sortby] lambda body unknown ident column
  tests(1310).expr = "a=1; b=2; c=3; a+b+c;": tests(1310).expected = "6" ' [syntax] multi-stmt trailing semicolon
  tests(1311).expr = "r=10; r;": tests(1311).expected = "10" ' [syntax] multi-stmt trailing semicolon
  tests(1312).expr = "r()=5; r();": tests(1312).expected = "5" ' [syntax] multi-stmt trailing semicolon
  tests(1313).expr = "1+2; 3+4": tests(1313).expected = "7" ' [syntax] multi-stmt no trailing semicolon
  tests(1314).expr = "1+2; 3+4;": tests(1314).expected = "7" ' [syntax] multi-stmt trailing semicolon
  tests(1315).expr = "f(x)=x+1; f(3)": tests(1315).expected = "4" ' [syntax] multi-stmt UDF then call
  tests(1316).expr = "f(x)=x+1; f(3);": tests(1316).expected = "4" ' [syntax] multi-stmt UDF call trailing semicolon
  tests(1317).expr = "0xAA; hex": tests(1317).expected = "0xAA" ' [syntax] trailing formatter sugar
  tests(1318).expr = "0xAA; hex;": tests(1318).expected = "0xAA" ' [syntax] trailing formatter sugar + semicolon
  tests(1319).expr = "(10,20,30); hex()": tests(1319).expected = "(0xA,0x14,0x1E)" ' [syntax] formatter sugar on array ans
  tests(1320).expr = "(10,20,30); hex;": tests(1320).expected = "(0xA,0x14,0x1E)" ' [syntax] formatter sugar bare name + semicolon
  tests(1321).expr = "repeat(0, 3)": tests(1321).expected = "(0, 0, 0)" ' [repeat] scalar
  tests(1322).expr = "repeat((10), 5)": tests(1322).expected = "(10, 10, 10, 10, 10)" ' [repeat] single-element array
  tests(1323).expr = "repeat((1,2), 1)": tests(1323).expected = "(1, 2)" ' [repeat] array once
  tests(1324).expr = "repeat(5, 2)": tests(1324).expected = "(5, 5)" ' [repeat] scalar twice
  tests(1325).expr = "repeat(1, 0)": tests(1325).expectedErrContains = "repeat() expects a positive integer" ' [repeat] zero count
  tests(1326).expr = "repeat(10, -3)": tests(1326).expectedErrContains = "repeat() expects a positive integer" ' [repeat] negative count
  tests(1327).expr = "repeat((1,2), -1)": tests(1327).expectedErrContains = "repeat() expects a positive integer" ' [repeat] negative count array
  tests(1328).expr = "repeat((1,2), 1.23)": tests(1328).expectedErrContains = "repeat() expects integer values" ' [repeat] non-integer count
  tests(1329).expr = "repeat": tests(1329).expectedErrContains = "function: repeat(x, n)" ' [repeat] hint
  tests(1330).expr = "repeat()": tests(1330).expectedErrContains = "expects 2 argument(s)" ' [repeat] arity
  tests(1331).expr = "repeat(1)": tests(1331).expectedErrContains = "expects 2 argument(s)" ' [repeat] arity
  tests(1332).expr = "repeat(1,2,3)": tests(1332).expectedErrContains = "expects 2 argument(s)" ' [repeat] arity
  tests(1333).expr = "repeat((1,2), 1:00)": tests(1333).expectedErrContains = "repeat() expects integer values" ' [repeat] time count
  tests(1334).expr = "a=repeat((18446744073709551615,2), 2); uhex(a[2])": tests(1334).expected = "0xFFFFFFFFFFFFFFFF" ' [repeat] int64 metadata on repeated copy
  tests(1335).expr = "range(1, 4)": tests(1335).expected = "(1, 2, 3)" ' [range] two-arg exclusive stop
  tests(1336).expr = "range(1, 10, 3)": tests(1336).expected = "(1, 4, 7)" ' [range] three-arg step
  tests(1337).expr = "range(10, 0, -2)": tests(1337).expected = "(10, 8, 6, 4, 2)" ' [range] negative step
  tests(1338).expr = "range(1, 2)": tests(1338).expected = "(1)" ' [range] single value
  tests(1339).expr = "range(-1, 2)": tests(1339).expected = "(-1, 0, 1)" ' [range] negative start
  tests(1340).expr = "range(1, 10, 0)": tests(1340).expectedErrContains = "range() step must not be zero" ' [range] zero step
  tests(1341).expr = "range(1, 10, -1)": tests(1341).expectedErrContains = "range() produces no values" ' [range] empty forward
  tests(1342).expr = "range(10, 0, 2)": tests(1342).expectedErrContains = "range() produces no values" ' [range] empty reverse
  tests(1343).expr = "range(0, 0, 2)": tests(1343).expectedErrContains = "range() produces no values" ' [range] empty equal bounds
  tests(1344).expr = "range(1, 1)": tests(1344).expectedErrContains = "range() produces no values" ' [range] empty default step
  tests(1345).expr = "range(5, 1)": tests(1345).expectedErrContains = "range() produces no values" ' [range] empty start>=stop
  tests(1346).expr = "range(1, 2, 1.5)": tests(1346).expectedErrContains = "range() expects integer values" ' [range] float step
  tests(1347).expr = "range((1,2), 5)": tests(1347).expectedErrContains = "range() expects integer values" ' [range] array arg
  tests(1348).expr = "range(1, 1:00)": tests(1348).expectedErrContains = "range() expects integer values" ' [range] time arg
  tests(1349).expr = "range": tests(1349).expectedErrContains = "function: range(start, stop [, step])" ' [range] hint
  tests(1350).expr = "range()": tests(1350).expectedErrContains = "expects at least 2 arguments" ' [range] arity
  tests(1351).expr = "range(1)": tests(1351).expectedErrContains = "expects at least 2 arguments" ' [range] arity
  tests(1352).expr = "range(1,2,3,4)": tests(1352).expectedErrContains = "expects 3 argument(s)" ' [range] arity
  tests(1353).expr = "hex(range(1,4)[0])": tests(1353).expected = "0x1" ' [range] int64 metadata
  tests(1354).expr = "sum(range(1,4))": tests(1354).expected = "6" ' [range] aggregate use
  tests(1355).expr = "range(1, 10, 2)": tests(1355).expected = "(1, 3, 5, 7, 9)" ' [range] even step
  tests(1356).expr = "range(1, 1001)[999]": tests(1356).expected = "1000" ' [range] max length boundary
  tests(1357).expr = "range(1, 1002)": tests(1357).expectedErrContains = "range() produces too many values" ' [range] over max length
  tests(1358).expr = "range(1, 2**53)": tests(1358).expectedErrContains = "range() produces too many values" ' [range] huge length guard
  tests(1359).expr = "repeat(0, 1000)[999]": tests(1359).expected = "0" ' [repeat] max length boundary
  tests(1360).expr = "repeat(0, 1001)": tests(1360).expectedErrContains = "repeat() produces too many values" ' [repeat] over max length
  tests(1361).expr = "repeat((1,2), 500)[999]": tests(1361).expected = "2" ' [repeat] array at max length
  tests(1362).expr = "repeat((1,2), 501)": tests(1362).expectedErrContains = "repeat() produces too many values" ' [repeat] array over max
  tests(1363).expr = "repeat(1, 2**53)": tests(1363).expectedErrContains = "repeat() produces too many values" ' [repeat] huge length guard
  tests(1364).expr = "len((10,20,30))": tests(1364).expected = "3" ' [len] array
  tests(1365).expr = "len(1:30)": tests(1365).expected = "1" ' [len] time scalar
  tests(1366).expr = "len(range(0,10))": tests(1366).expected = "10" ' [len] range result
  tests(1367).expr = "len(1,2,3)": tests(1367).expected = "3" ' [len] variadic
  tests(1368).expr = "len()": tests(1368).expected = "0" ' [len] empty call
  tests(1369).expr = "len(())": tests(1369).expected = "0" ' [len] empty array
  tests(1370).expr = "len(1, (2, 3))": tests(1370).expected = "3" ' [len] flatten mixed args
  tests(1371).expr = "len(42)": tests(1371).expected = "1" ' [len] numeric scalar
  tests(1372).expr = "len(nan)": tests(1372).expected = "1" ' [len] non-finite scalar
  tests(1373).expr = "length((1,2))": tests(1373).expected = "2" ' [len] alias
  tests(1374).expr = "len": tests(1374).expectedErrContains = "function: len(x)" ' [len] hint
  tests(1375).expr = "length": tests(1375).expectedErrContains = "function: len(x)" ' [len] alias hint

  ' === Array slicing (Python-compatible start:stop:step; empty -> ()) ===
  tests(1376).expr = "a=(1,2,3,4,5,6,7,8,9); a[2:7]": tests(1376).expected = "(3, 4, 5, 6, 7)" ' [slice]
  tests(1377).expr = "a=(1,2,3,4,5,6,7,8,9); a[3:-3]": tests(1377).expected = "(4, 5, 6)" ' [slice]
  tests(1378).expr = "a=(1,2,3,4,5,6,7,8,9); a[0:1]": tests(1378).expected = "(1)" ' [slice]
  tests(1379).expr = "a=(1,2,3,4,5,6,7,8,9); a[-3:-2]": tests(1379).expected = "(7)" ' [slice]
  tests(1380).expr = "a=(1,2,3,4,5,6,7,8,9); a[-2:-3]": tests(1380).expected = "()" ' [slice] empty
  tests(1381).expr = "a=(1,2,3,4,5,6,7,8,9); a[3:1]": tests(1381).expected = "()" ' [slice] empty
  tests(1382).expr = "a=(1,2,3,4,5,6,7,8,9); a[1:1]": tests(1382).expected = "()" ' [slice] empty
  tests(1383).expr = "a=(1,2,3,4,5,6,7,8,9); a[11:11]": tests(1383).expected = "()" ' [slice] empty clamp
  tests(1384).expr = "a=(1,2,3,4,5,6,7,8,9); a[:3]": tests(1384).expected = "(1, 2, 3)" ' [slice] omit start
  tests(1385).expr = "a=(1,2,3,4,5,6,7,8,9); a[3:]": tests(1385).expected = "(4, 5, 6, 7, 8, 9)" ' [slice] omit stop
  tests(1386).expr = "a=(1,2,3,4,5,6,7,8,9); a[:]": tests(1386).expected = "(1, 2, 3, 4, 5, 6, 7, 8, 9)" ' [slice] copy
  tests(1387).expr = "a=(1,2,3,4,5,6,7,8,9); a[::2]": tests(1387).expected = "(1, 3, 5, 7, 9)" ' [slice] step
  tests(1388).expr = "a=(1,2,3,4,5,6,7,8,9); a[1:8:2]": tests(1388).expected = "(2, 4, 6, 8)" ' [slice] step
  tests(1389).expr = "a=(1,2,3,4,5,6,7,8,9); a[::-1]": tests(1389).expected = "(9, 8, 7, 6, 5, 4, 3, 2, 1)" ' [slice] reverse
  tests(1390).expr = "a=(1,2,3,4,5,6,7,8,9); a[8:1:-2]": tests(1390).expected = "(9, 7, 5, 3)" ' [slice] neg step
  tests(1391).expr = "a=(1,2,3,4,5,6,7,8,9); a[1:30]": tests(1391).expected = "(2, 3, 4, 5, 6, 7, 8, 9)" ' [slice] not time lit
  tests(1392).expr = "a=(1,2,3,4,5,6,7,8,9); a[::0]": tests(1392).expectedErrContains = "slice step must not be zero" ' [slice]
  tests(1393).expr = "a=(1,2,3,4,5,6,7,8,9); a[1.5:3]": tests(1393).expectedErrContains = "array index must be an integer" ' [slice]
  tests(1394).expr = "(10,20,30)[1:2]": tests(1394).expected = "(20)" ' [slice] literal base
  tests(1395).expr = "a=(1,2,3,4,5); a[1:4][0]": tests(1395).expected = "2" ' [slice] chain to index
  tests(1396).expr = "a=(1,2,3,4,5,6,7,8,9); a[-100:100]": tests(1396).expected = "(1, 2, 3, 4, 5, 6, 7, 8, 9)" ' [slice] clamp
  tests(1397).expr = "a=(1,2,3,4,5,6,7,8,9); a[100:-100:-1]": tests(1397).expected = "(9, 8, 7, 6, 5, 4, 3, 2, 1)" ' [slice] clamp neg
  tests(1398).expr = "_=(1,2,3,4,5,6); ff(a)=a[1:4]; ff(_)": tests(1398).expected = "(2, 3, 4)" ' [udf] formal slice
  tests(1399).expr = "a=(1,2,3,4,5,6,7,8,9); i=2; j=7; a[i:j]": tests(1399).expected = "(3, 4, 5, 6, 7)" ' [slice] expr bounds
  tests(1400).expr = "5[0:1]": tests(1400).expectedErrContains = "indexing requires an array value" ' [slice]
  tests(1401).expr = "a=range(1,101); a[1m3s]": tests(1401).expectedErrContains = "array index must be an integer" ' [slice] time index
  tests(1402).expr = "a=range(1,101); a[1m:1m3s]": tests(1402).expectedErrContains = "array index must be an integer" ' [slice] time slice bounds

  ' === flat(...) - concatenate to one array (no unpack expand-args) ===
  tests(1403).expr = "a=(1,2); b=(4,5); flat(a,3,b,6)": tests(1403).expected = "(1,2,3,4,5,6)" ' [flat] mixed arrays/scalars
  tests(1404).expr = "flat(5)": tests(1404).expected = "(5)" ' [flat] single scalar -> array
  tests(1405).expr = "flat((1,2,3))": tests(1405).expected = "(1,2,3)" ' [flat] single array passthrough
  tests(1406).expr = "flat((1,2),3,(4,5))": tests(1406).expected = "(1,2,3,4,5)" ' [flat] multi-arg
  tests(1407).expr = "flat()": tests(1407).expectedErrContains = "expects at least 1 argument" ' [flat] arity
  tests(1408).expr = "flat(())": tests(1408).expectedErrContains = "expects at least 1 argument" ' [flat] empty array
  tests(1409).expr = "flat": tests(1409).expectedErrContains = "function: flat(...)" ' [flat] hint
  tests(1410).expr = "f(x,y)=x*y; f(flat((2,3)))": tests(1410).expectedErrContains = "expects 2 argument(s)" ' [flat] no call spreading
  tests(1411).expr = "flat(1,Inf,2)": tests(1411).expected = "(1,inf,2)" ' [flat] keep non-finite
  tests(1412).expr = "a=flat((18446744073709551615,2)); uhex(a[0])": tests(1412).expected = "0xFFFFFFFFFFFFFFFF" ' [flat] int64 metadata
  tests(1413).expr = "flat(1,2,1,2,3)": tests(1413).expected = "(1,2,1,2,3)" ' [flat] keep duplicates
  tests(1414).expr = "flat((9,8,7))": tests(1414).expected = "(9,8,7)" ' [flat] single-arg array edge

  dim uniqueTotal as Integer
  dim duplicateTotal as Integer
  dim sigI as String
  dim isDup as Boolean
  uniqueTotal = 0
  duplicateTotal = 0
  for i as Integer = lbound(tests) to ubound(tests)
    sigI = SmokeCaseSignature(tests(i))
    isDup = FALSE
    for j as Integer = lbound(tests) to i - 1
      if SmokeCaseSignature(tests(j)) = sigI then
        isDup = TRUE
      end if
    next j
    if isDup = FALSE then
      uniqueTotal += 1
    else
      duplicateTotal += 1
    end if
  next i
  g_total = uniqueTotal

  print "=== SmartMath parser smoke tests ==="
  print "Total cases: " & g_total
  print "Detected duplicate cases: " & duplicateTotal
  print ""

  Parser_ClearVariables()
  for i as Integer = lbound(tests) to ubound(tests)
    sigI = SmokeCaseSignature(tests(i))
    isDup = FALSE
    for j as Integer = lbound(tests) to i - 1
      if SmokeCaseSignature(tests(j)) = sigI then
        isDup = TRUE
      end if
    next j
    if isDup = FALSE then
      RunCase(tests(i))
    else
      print "[DUPLICATE] SKIP : " & tests(i).expr
    end if
  next i

  RunRawResultApiTests()

  RunOperatorPrecedenceDocTests()

  RunTrigAngleReductionTests()

  RunComplexNumberSupportOptionTests()

  RunNegativeArgumentMagnitudeBandTests()

  RunPositiveArgumentMagnitudeBandTests()

  RunFloatMagnitudeLiteralTests()

  RunNanInfTests()

  RunBuiltinArityTableTests()

  RunIncompleteFunctionCallHintTests()

  RunMultiStmtSemicolonTests()

  RunGcdLcmNcrNprArrayBroadcastTests()
  RunRatioInExpressionTests()
  RunMinMaxPreserveWinnerTests()
  RunClampPreserveIntegerMetadataTests()
  RunClampInvertedBoundsTests()
  RunClampNanBoundTests()

  RunExactIntegerDivisionTests()
  RunExactIntegerMultiplicationTests()
  RunBinaryBuiltinArrayLengthMismatchTests()

  RunTimeValuesSupportOptionTests()

  RunLambdaFunctionsSupportOptionTests()

  print "=== Result ==="
  print "Passed: " & g_passed
  print "Failed: " & g_failed

  if g_failed > 0 then
    end 1
  else
    end 0
  end if
end sub

Main()
