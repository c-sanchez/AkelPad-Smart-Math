''
'' Formatter regression tests for SmartMath_Format.bas (link with FormatterTest_Globals.bas + SmartMath_Format.bas).
''
#include once "SmartMath_Globals.bi"

#ifndef LOCALE_SDECIMAL
const LOCALE_SDECIMAL = &h0000000E
#endif

'' Primary Win32 LCIDs for Format()/locale decimal smoke.
private const LCID_EN_US as DWORD = &h0409
private const LCID_DE_DE as DWORD = &h0407
private const LCID_FR_FR as DWORD = &h040C

'' Mirror SmartMath_Format.bas UI labels (not exported from formatter module).
private const FMT_NAN as String = "NaN"
private const FMT_INF as String = "Inf"
private const FMT_NEGINF as String = "-Inf"

private sub ApplySeparatorDefaults()
  g_sDecimalSeparator = SMARTMATH_DECIMAL_SEPARATOR_DEFAULT
  g_sThousandsSeparator = SMARTMATH_THOUSANDS_SEPARATOR_DEFAULT
  g_sArrayOutputSeparator = SMARTMATH_ARRAY_OUTPUT_SEPARATOR_DEFAULT
end sub

private sub FormatterTestSetup()
  g_nDecimals = -1
  g_bUseThousandsSeparator = FALSE
  ApplySeparatorDefaults()
  ResetSmartMathFormatLocaleCache()
end sub

private function ThreadLocaleDecimalA() as String
  dim buf as zstring * 8 = any
  buf = ""
  if GetLocaleInfoA(GetThreadLocale(), LOCALE_SDECIMAL, @buf, 8) = 0 then return "."
  return Trim(buf)
end function

private sub AssertEq(byref caseName as String, byref got as String, byref wantVal as String, byref failCount as Integer)
  if got = wantVal then
    print !"[PASS] "; caseName
  else
    print !"[FAIL] "; caseName
    print !"  expected: ["; wantVal; !"]"
    print !"  got:      ["; got; !"]"
    failCount += 1
  end if
end sub

private sub AssertTrue(byref caseName as String, byval cond as Boolean, byref failCount as Integer)
  if cond then
    print !"[PASS] "; caseName
  else
    print !"[FAIL] "; caseName
    failCount += 1
  end if
end sub

private sub AssertNoSubstr(byref caseName as String, byref haystack as String, byref bad as String, byref failCount as Integer)
  if Len(bad) = 0 then
    print !"[PASS] "; caseName; !" (skip empty needle)"
    exit sub
  end if
  if InStr(1, haystack, bad) = 0 then
    print !"[PASS] "; caseName
  else
    print !"[FAIL] "; caseName; !" — unexpected substring ["; bad; !"] in ["; haystack; !"]"
    failCount += 1
  end if
end sub

private function CountCharAsc(byref s as String, byval chAsc as Integer) as Integer
  dim c as Integer = 0
  dim i as Integer
  for i = 1 to Len(s)
    if Asc(Mid(s, i, 1)) = chAsc then c += 1
  next i
  return c
end function

private sub FormatterTestInitRawScalarFloat(byref s as RawScalar, byval v as Double)
  s.kind = RSK_FLOATING
  s.real.kind = RSK_FLOATING
  s.real.floatValue = v
  RawCartesianScalarClear(s.imag)
  s.renderBase = 0
  s.renderUnsigned = FALSE
end sub

private sub FormatterTestInitRawScalarInt64(byref s as RawScalar, byval v as LongInt)
  s.kind = RSK_INT64
  s.real.kind = RSK_INT64
  s.real.intValue = v
  RawCartesianScalarClear(s.imag)
  s.renderBase = 0
  s.renderUnsigned = FALSE
end sub

private sub FormatterTestSetRawFloatArray(byref r as RawResult, vals() as Double)
  RawResultClear(r)
  r.kind = RRK_ARRAY
  dim lb as Integer = lbound(vals)
  dim ub as Integer = ubound(vals)
  redim r.arr(0 to ub - lb)
  dim i as Integer
  for i = lb to ub
    FormatterTestInitRawScalarFloat r.arr(i - lb), vals(i)
  next i
end sub

private sub FormatterTestSetRawInt64Array(byref r as RawResult, vals() as LongInt)
  RawResultClear(r)
  r.kind = RRK_ARRAY
  dim lb as Integer = lbound(vals)
  dim ub as Integer = ubound(vals)
  redim r.arr(0 to ub - lb)
  dim i as Integer
  for i = lb to ub
    FormatterTestInitRawScalarInt64 r.arr(i - lb), vals(i)
  next i
end sub

private function FormatTestRawArray(byref r as RawResult) as String
  return FormatRawResultForDisplay(r)
end function

'' Array display via RawResult (same output as former FormatArrayResultText parser-text path).
private sub RunArrayDisplayTests(byref failCount as Integer)
  dim savedLc as DWORD = GetThreadLocale()
  if SetThreadLocale(LCID_EN_US) = 0 then
    print !"[SKIP] RunArrayDisplayTests: SetThreadLocale en-US failed"
    exit sub
  end if
  ResetSmartMathFormatLocaleCache()
  print !"-- Array display (en-US for stable Format)"

  '' Auto decimals, no thousands — integer elements, custom array separator
  FormatterTestSetup()
  g_sArrayOutputSeparator = ";"
  dim rawFast as RawResult
  dim intsFast(0 to 1) as LongInt = {10, 20}
  FormatterTestSetRawInt64Array rawFast, intsFast()
  dim rFast as String = FormatTestRawArray(rawFast)
  AssertEq("arr/fast/semicolon elem sep", rFast, SMARTMATH_RESULT_PREFIX & "(10; 20)", failCount)

  '' Fixed decimals, dot decimal, comma as g_sArrayOutputSeparator (e.g. (12.345, 1.444, 7.890))
  g_nDecimals = 3
  g_bUseThousandsSeparator = FALSE
  g_sDecimalSeparator = "."
  g_sArrayOutputSeparator = ","
  '' Use values that keep a non-zero fractional part after rounding/trim (avoids all-integer tuple).
  dim rawDot as RawResult
  dim floatsDot(0 to 2) as Double = {12.345, 1.444, 7.890}
  FormatterTestSetRawFloatArray rawDot, floatsDot()
  dim rDot as String = FormatTestRawArray(rawDot)
  AssertTrue("arr/fixed/dot-dec/has prefix", Left(rDot, Len(SMARTMATH_RESULT_PREFIX)) = SMARTMATH_RESULT_PREFIX, failCount)
  AssertTrue("arr/fixed/dot-dec/parens", InStr(1, rDot, "(") > 0 andalso InStr(1, rDot, ")") > 0, failCount)
  AssertTrue("arr/fixed/dot-dec/elem gap uses array sep", InStr(1, rDot, ", ") > 0, failCount)
  AssertTrue("arr/fixed/dot-dec/first elt uses ascii dot", InStr(1, rDot, "12.345") > 0, failCount)

  '' Thousands + comma decimal + comma array sep: (1'024,333, 2'048,666, 4'096,999) style
  g_nDecimals = 3
  g_bUseThousandsSeparator = TRUE
  g_sThousandsSeparator = "'"
  g_sDecimalSeparator = ","
  g_sArrayOutputSeparator = ","
  dim rawThou as RawResult
  dim floatsThou(0 to 2) as Double = {1024.333, 2048.666, 4096.999}
  FormatterTestSetRawFloatArray rawThou, floatsThou()
  dim rThou as String = FormatTestRawArray(rawThou)
  AssertTrue("arr/thou-comma/thou marks", CountCharAsc(rThou, 39) >= 3, failCount)
  AssertTrue("arr/thou-comma/has 1'024", InStr(1, rThou, "1'024") > 0, failCount)
  AssertTrue("arr/thou-comma/has 2'048", InStr(1, rThou, "2'048") > 0, failCount)
  AssertTrue("arr/thou-comma/has 4'096", InStr(1, rThou, "4'096") > 0, failCount)
  AssertTrue("arr/thou-comma/dec comma in first elt", InStr(1, rThou, "1'024,") > 0, failCount)
  '' Element boundary: fraction ends then array sep ", " before next thousand group
  AssertTrue("arr/thou-comma/elt boundary", InStr(1, rThou, ",333, 2'048") > 0, failCount)

  '' Alternate array separator with same numeric styling
  g_sArrayOutputSeparator = "|"
  dim rawPipe as RawResult
  dim floatsPipe(0 to 1) as Double = {1024.333, 2048.666}
  FormatterTestSetRawFloatArray rawPipe, floatsPipe()
  dim rPipe as String = FormatTestRawArray(rawPipe)
  AssertTrue("arr/pipe-elem/has pipe-space", InStr(1, rPipe, "| ") > 0, failCount)
  AssertTrue("arr/pipe-elem/still thousands", InStr(1, rPipe, "1'024") > 0, failCount)

  SetThreadLocale(savedLc)
  ResetSmartMathFormatLocaleCache()
  FormatterTestSetup()
end sub

private sub RunLocaleBlock(byval lcid as DWORD, byref label as String, byref failCount as Integer)
  dim prev as DWORD = GetThreadLocale()
  if SetThreadLocale(lcid) = 0 then
    print !"[SKIP] SetThreadLocale failed for "; label
    exit sub
  end if
  ResetSmartMathFormatLocaleCache()
  print !"-- Locale block: "; label; !" (thread dec sep=["; ThreadLocaleDecimalA(); !"])"

  g_nDecimals = 2
  g_bUseThousandsSeparator = TRUE
  g_sDecimalSeparator = ","
  g_sThousandsSeparator = "'"

  dim r73 as String = FormatResult(73.5)
  AssertNoSubstr(label & !"/73.5 no mangled thou", r73, "7'3", failCount)

  dim r105 as String = FormatResult(10.5 * 7.0)
  AssertNoSubstr(label & !"/10.5*7 no mangled thou", r105, "7'3", failCount)

  dim r1k as String = FormatResult(12345.67)
  AssertTrue(label & !"/12345 has thou sep", InStr(1, r1k, "'") > 0, failCount)

  SetThreadLocale(prev)
  ResetSmartMathFormatLocaleCache()
end sub

private sub RunSeparatorMatrix(byref failCount as Integer)
  print !"-- Separator matrix (1-char ini-style, thousands on, 2 decimals)"

  dim decChars(0 to 2) as String
  decChars(0) = "."
  decChars(1) = ","
  decChars(2) = ";"
  dim thouChars(0 to 2) as String
  thouChars(0) = "'"
  thouChars(1) = "."
  thouChars(2) = ","
  dim arrayChars(0 to 2) as String
  arrayChars(0) = ";"
  arrayChars(1) = "|"
  arrayChars(2) = ":"

  dim di as Integer, ti as Integer, ai as Integer
  for di = 0 to 2
    for ti = 0 to 2
      for ai = 0 to 2
        dim dch as String = decChars(di)
        dim tch as String = thouChars(ti)
        dim ach as String = arrayChars(ai)
        if dch <> tch then
          FormatterTestSetup()
          g_nDecimals = 2
          g_bUseThousandsSeparator = TRUE
          g_sDecimalSeparator = dch
          g_sThousandsSeparator = tch
          g_sArrayOutputSeparator = ach

          dim tag as String = !"dec=" & dch & !"/thou=" & tch & !"/arr=" & ach
          dim r as String = FormatResult(8888.25)
          AssertTrue(tag & !"/len", Len(r) > Len(SMARTMATH_RESULT_PREFIX), failCount)
          AssertNoSubstr(tag & !"/no double-dec glitch", r, dch & dch, failCount)

          dim rawArr as RawResult
          dim intsArr(0 to 2) as LongInt = {1, 2, 3}
          FormatterTestSetRawInt64Array rawArr, intsArr()
          dim a as String = FormatTestRawArray(rawArr)
          AssertTrue(tag & !"/array has sep", InStr(1, a, ach) > 0, failCount)
        end if
      next ai
    next ti
  next di

  FormatterTestSetup()
end sub

'' --------- main ----------
dim failures as Integer = 0

print !"=== SmartMath formatter regression tests ==="
print !""

FormatterTestSetup()

'' --- Non-finite ---
dim nanv as Double
nanv = 0.0 / 0.0
dim pInf as Double = 1.0e200 * 1.0e200
dim nInf as Double = -pInf

AssertEq("NaN display", FormatResult(nanv), SMARTMATH_RESULT_PREFIX & FMT_NAN, failures)
AssertEq("+Inf display", FormatResult(pInf), SMARTMATH_RESULT_PREFIX & FMT_INF, failures)
AssertEq("-Inf display", FormatResult(nInf), SMARTMATH_RESULT_PREFIX & FMT_NEGINF, failures)

AssertEq("parser nan token", FormatNonFiniteDisplayFromParserScalar("nan"), FMT_NAN, failures)
AssertEq("parser inf token", FormatNonFiniteDisplayFromParserScalar("inf"), FMT_INF, failures)
AssertEq("parser -inf token", FormatNonFiniteDisplayFromParserScalar("-inf"), FMT_NEGINF, failures)
AssertEq("parser finite", FormatNonFiniteDisplayFromParserScalar("3.14"), "", failures)

'' --- Auto decimals (Str path, ASCII dot) ---
AssertTrue("auto decimals has prefix", Left(FormatResult(73.5), Len(SMARTMATH_RESULT_PREFIX)) = SMARTMATH_RESULT_PREFIX, failures)

'' --- Fixed decimals, no thousands ---
g_nDecimals = 2
dim rNorm as String = FormatResult(73.5)
AssertTrue("fixed73 no empty", Len(rNorm) > Len(SMARTMATH_RESULT_PREFIX), failures)
AssertNoSubstr("fixed73 no apostrophe glitch", rNorm, "7'3", failures)

'' --- Thousands grouping shape ---
FormatterTestSetup()
g_nDecimals = 1
g_bUseThousandsSeparator = TRUE
g_sThousandsSeparator = "'"
dim rBig as String = FormatResult(12345.6)
AssertTrue("thousands near 12k", InStr(1, rBig, "'") > 0, failures)

'' --- Scientific normalization lowercase e ---
FormatterTestSetup()
g_nDecimals = 3
g_bUseThousandsSeparator = FALSE
dim rSci as String = FormatResult(0.0000012)
AssertTrue("sci uses e", InStr(1, LCase(rSci), "e") > 0, failures)

RunArrayDisplayTests(failures)

private sub RunRawResultFormatTests(byref failCount as Integer)
  FormatterTestSetup()
  dim nanv as Double
  nanv = 0.0 / 0.0
  dim pInf as Double = 1.0e200 * 1.0e200
  dim r as RawResult
  RawResultClear(r)
  r.kind = RRK_SCALAR
  r.scalar.kind = RSK_RATIONAL
  r.scalar.real.kind = RSK_RATIONAL
  r.scalar.real.ratNum = 1
  r.scalar.real.ratDen = 2
  RawCartesianScalarClear(r.scalar.imag)
  AssertEq("raw format rational scalar", FormatRawResultForDisplay(r), SMARTMATH_RESULT_PREFIX & "1/2", failCount)

  RawResultClear(r)
  r.kind = RRK_SCALAR
  r.scalar.kind = RSK_COMPLEX
  r.scalar.real.kind = RSK_INT64
  r.scalar.real.intValue = 10
  r.scalar.imag.kind = RSK_INT64
  r.scalar.imag.intValue = 5
  AssertEq("raw format complex int", FormatRawResultForDisplay(r), SMARTMATH_RESULT_PREFIX & "10 + 5i", failCount)

  RawResultClear(r)
  r.kind = RRK_SCALAR
  r.scalar.kind = RSK_COMPLEX
  r.scalar.real.kind = RSK_FLOATING
  r.scalar.real.floatValue = -0.0
  r.scalar.imag.kind = RSK_FLOATING
  r.scalar.imag.floatValue = -0.22
  AssertEq("raw format pure imag neg zero real", FormatRawResultForDisplay(r), SMARTMATH_RESULT_PREFIX & "-0.22i", failCount)

  RawResultClear(r)
  r.kind = RRK_SCALAR
  r.scalar.kind = RSK_COMPLEX
  r.scalar.real.kind = RSK_FLOATING
  r.scalar.real.floatValue = pInf
  r.scalar.imag.kind = RSK_FLOATING
  r.scalar.imag.floatValue = pInf
  AssertEq("raw format complex inf inf*i", FormatRawResultForDisplay(r), SMARTMATH_RESULT_PREFIX & "Inf + Inf*i", failCount)

  RawResultClear(r)
  r.kind = RRK_SCALAR
  r.scalar.kind = RSK_COMPLEX
  r.scalar.real.kind = RSK_FLOATING
  r.scalar.real.floatValue = nanv
  r.scalar.imag.kind = RSK_FLOATING
  r.scalar.imag.floatValue = nanv
  AssertEq("raw format complex nan nan*i", FormatRawResultForDisplay(r), SMARTMATH_RESULT_PREFIX & "NaN", failCount)

  RawResultClear(r)
  r.kind = RRK_ARRAY
  redim r.arr(0 to 1)
  r.arr(0).kind = RSK_RATIONAL
  r.arr(0).real.kind = RSK_RATIONAL
  r.arr(0).real.ratNum = 1
  r.arr(0).real.ratDen = 2
  RawCartesianScalarClear(r.arr(0).imag)
  r.arr(1).kind = RSK_INT64
  r.arr(1).real.kind = RSK_INT64
  r.arr(1).real.intValue = 3
  RawCartesianScalarClear(r.arr(1).imag)
  AssertEq("raw format array mix", FormatRawResultForDisplay(r), SMARTMATH_RESULT_PREFIX & "(1/2, 3)", failCount)

  FormatterTestSetup()
  g_bUseThousandsSeparator = TRUE
  g_sThousandsSeparator = "'"
  RawResultClear(r)
  r.kind = RRK_SCALAR
  FormatterTestInitRawScalarFloat r.scalar, 2147483647.0
  AssertEq("raw float integer-like shows .0", FormatRawResultForDisplay(r), SMARTMATH_RESULT_PREFIX & "2'147'483'647.0", failCount)

  FormatterTestSetup()
  RawResultClear(r)
  r.kind = RRK_SCALAR
  FormatterTestInitRawScalarInt64 r.scalar, 2147483647
  AssertEq("raw int64 integer-like no .0", FormatRawResultForDisplay(r), SMARTMATH_RESULT_PREFIX & "2147483647", failCount)

  FormatterTestSetup()
  g_nDecimals = -1
  g_bUseThousandsSeparator = FALSE
  ApplySeparatorDefaults()
  '' Python format(4285714285714285.5, '.16g') is 4285714285714286; auto mode appends .0.
  AssertEq("raw float hidden fraction shows .0", FormatResult(4285714285714285.5), SMARTMATH_RESULT_PREFIX & "4285714285714286.0", failCount)
  '' Python format((2**53)*1.1, '.15e') is 9.907919180215092e+15; auto mode strips the padded exponent.
  AssertEq("raw float above 2^53 uses scientific", FormatResult(9007199254740992.0 * 1.1), SMARTMATH_RESULT_PREFIX & "9.907919180215092e+15", failCount)

  RawResultClear(r)
  r.kind = RRK_SCALAR
  r.scalar.kind = RSK_COMPLEX
  r.scalar.real.kind = RSK_FLOATING
  r.scalar.real.floatValue = 2147483647.0
  r.scalar.imag.kind = RSK_FLOATING
  r.scalar.imag.floatValue = 2.0
  AssertEq("raw complex float integer-like shows .0", FormatRawResultForDisplay(r), SMARTMATH_RESULT_PREFIX & "2147483647.0 + 2.0i", failCount)
end sub

RunRawResultFormatTests(failures)

'' --- Locale emulation via SetThreadLocale ---
dim savedLc as DWORD = GetThreadLocale()
RunLocaleBlock(LCID_EN_US, "en-US", failures)
RunLocaleBlock(LCID_DE_DE, "de-DE", failures)
RunLocaleBlock(LCID_FR_FR, "fr-FR", failures)
if savedLc <> GetThreadLocale() then
  SetThreadLocale(savedLc)
  ResetSmartMathFormatLocaleCache()
end if

'' --- 1-char separator matrix ---
RunSeparatorMatrix(failures)
FormatterTestSetup()

#ifdef PARSER_FORMATTER_INTEGRATION

private function FormatterTryEvaluateAndFormat(byref expr as String, byref fmtOut as String, byref rawOut as RawResult) as Boolean
  fmtOut = ""
  RawResultClear(rawOut)
  if Parser_TryEvaluateExRaw(expr, rawOut) = FALSE then return FALSE
  fmtOut = FormatRawEvaluationResult(rawOut)
  return TRUE
end function

private sub AssertRawComplexInt64Parts(byref caseName as String, byref raw as RawResult, byref failCount as Integer)
  if raw.scalar.kind <> RSK_COMPLEX then
    print !"[FAIL] "; caseName; !" — expected RSK_COMPLEX, kind="; raw.scalar.kind
    failCount += 1
    exit sub
  end if
  if raw.scalar.real.kind <> RSK_INT64 orelse raw.scalar.imag.kind <> RSK_INT64 then
    print !"[FAIL] "; caseName; !" — expected int64 real/imag, got re="; raw.scalar.real.kind; !" im="; raw.scalar.imag.kind
    failCount += 1
  else
    print !"[PASS] "; caseName; !" raw int64 complex"
  end if
end sub

private sub AssertRawScalarInt64(byref caseName as String, byref raw as RawResult, byref failCount as Integer)
  if raw.scalar.kind <> RSK_INT64 then
    print !"[FAIL] "; caseName; !" — expected RSK_INT64 scalar, kind="; raw.scalar.kind
    failCount += 1
  else
    print !"[PASS] "; caseName; !" raw int64 scalar"
  end if
end sub

private sub AssertRawFactorintIntPowerTerm( _
  byref caseName as String, _
  byref raw as RawResult, _
  byval termIdx as Integer, _
  byval baseV as LongInt, _
  byval expV as ULongInt, _
  byref failCount as Integer _
)
  if raw.kind <> RRK_ARRAY then
    print !"[FAIL] "; caseName; !"/raw — expected RRK_ARRAY, kind="; raw.kind
    failCount += 1
    exit sub
  end if
  if termIdx < lbound(raw.arr) orelse termIdx > ubound(raw.arr) then
    print !"[FAIL] "; caseName; !"/raw — term index out of range"
    failCount += 1
    exit sub
  end if
  if raw.arr(termIdx).real.kind <> RSK_INT_POWER orelse raw.arr(termIdx).real.ratNum <> baseV orelse raw.arr(termIdx).real.ratDen <> expV then
    print !"[FAIL] "; caseName; !"/raw — expected INT_POWER "; baseV; !"**"; expV; !", got kind="; raw.arr(termIdx).real.kind; !" num="; raw.arr(termIdx).real.ratNum; !" den="; raw.arr(termIdx).real.ratDen
    failCount += 1
  else
    print !"[PASS] "; caseName; !"/raw int-power term"
  end if
end sub

private sub RunFactorintFormatterTests(byref failCount as Integer)
  '' Mirror SmokeTest_MathParser.bas factorint cases (tests 1209..1234).
  print !""
  print !"-- Parser + formatter integration (factorint, decimals=-1)"

  Parser_ClearVariables()
  Parser_SetSupportComplexNumbers(FALSE)
  FormatterTestSetup()
  g_nDecimals = -1
  g_bUseThousandsSeparator = FALSE
  ApplySeparatorDefaults()

  const FACTORINT_FMT_CASE_COUNT as Integer = 28
  dim expr(1 to FACTORINT_FMT_CASE_COUNT) as String
  dim want(1 to FACTORINT_FMT_CASE_COUNT) as String
  dim expectFail(1 to FACTORINT_FMT_CASE_COUNT) as Boolean
  dim expectArray(1 to FACTORINT_FMT_CASE_COUNT) as Boolean

  expr(1) = "factorint(33)": want(1) = "(3, 11)"
  expr(2) = "factorint(12)": want(2) = "(2**2, 3)"
  expr(3) = "factorint(13)": want(3) = "(13)"
  expr(4) = "factorint(-33)": want(4) = "(-3, 11)"
  expr(5) = "factorint(-13)": want(5) = "(-13)"
  expr(6) = "factorint(-12)": want(6) = "(-2**2, 3)"
  expr(7) = "factorint(0)": want(7) = "(0)"
  expr(8) = "factorint(1)": want(8) = "(1)"
  expr(9) = "factorint(-1)": want(9) = "(-1)"
  expr(10) = "factorint(2**52)": want(10) = "(2**52)"
  expr(11) = "factorint(2**63-1)": want(11) = "(7**2, 73, 127, 337, 92737, 649657)"
  expr(12) = "factorint((33,12))": expectFail(12) = TRUE
  expr(13) = "factorint(2**64)": expectFail(13) = TRUE
  expr(14) = "factorint(18446744073709551615)": want(14) = "(3, 5, 17, 257, 641, 65537, 6700417)"
  expr(15) = "prod(factorint(12))": want(15) = "12": expectArray(15) = FALSE
  expr(16) = "prod(factorint(-33))": want(16) = "-33": expectArray(16) = FALSE
  expr(17) = "prod(factorint(2**52))": want(17) = "4503599627370496": expectArray(17) = FALSE
  expr(18) = "factorint(33.0)": want(18) = "(3, 11)"
  expr(19) = "factorint(33.5)": expectFail(19) = TRUE
  expr(20) = "factorint(90)": want(20) = "(2, 3**2, 5)"
  expr(21) = "factorint(9007)": want(21) = "(9007)"
  expr(22) = "factorint(900719)": want(22) = "(900719)"
  expr(23) = "factorint(90071992)": want(23) = "(2**3, 11258999)"
  expr(24) = "factorint(9007199254)": want(24) = "(2, 89, 50602243)"
  expr(25) = "factorint(900719925474)": want(25) = "(2, 3, 12907, 11630897)"
  expr(26) = "factorint(76568758722)": want(26) = "(2, 3**2, 47, 101, 896107)"
  expr(27) = "factorint(-3333*9)": want(27) = "(-3**3, 11, 101)"
  expr(28) = "factorint(-9999)": want(28) = "(-3**2, 11, 101)"

  dim ci as Integer
  for ci = 1 to FACTORINT_FMT_CASE_COUNT
    if ci <> 12 andalso ci <> 13 andalso ci <> 15 andalso ci <> 16 andalso ci <> 17 andalso ci <> 19 then
      expectArray(ci) = TRUE
    end if
  next ci

  dim raw as RawResult
  dim fmt as String
  for ci = 1 to FACTORINT_FMT_CASE_COUNT
    dim tag as String = "parser-fmt/factorint/" & expr(ci)
    dim okEval as Boolean = FormatterTryEvaluateAndFormat(expr(ci), fmt, raw)
    if expectFail(ci) then
      if okEval then
        print !"[FAIL] "; tag; !" — expected evaluation error"
        failCount += 1
      else
        print !"[PASS] "; tag; !" (expected error)"
      end if
      continue for
    end if
    if okEval = FALSE then
      print !"[FAIL] "; tag; !" — evaluate failed"
      failCount += 1
      continue for
    end if
    AssertEq(tag, fmt, SMARTMATH_RESULT_PREFIX & want(ci), failCount)
    if expectArray(ci) then
      if raw.kind <> RRK_ARRAY then
        print !"[FAIL] "; tag; !"/raw — expected RRK_ARRAY, kind="; raw.kind
        failCount += 1
      else
        print !"[PASS] "; tag; !"/raw array"
      end if
    else
      if raw.kind <> RRK_SCALAR then
        print !"[FAIL] "; tag; !"/raw — expected RRK_SCALAR, kind="; raw.kind
        failCount += 1
      else
        print !"[PASS] "; tag; !"/raw scalar"
      end if
    end if
  next ci

  '' Spot-check INT_POWER raw metadata on cases that use ** in formatted output.
  if FormatterTryEvaluateAndFormat("factorint(12)", fmt, raw) then
    AssertRawFactorintIntPowerTerm("parser-fmt/factorint/12/raw", raw, 0, 2, 2, failCount)
  end if
  if FormatterTryEvaluateAndFormat("factorint(90)", fmt, raw) then
    AssertRawFactorintIntPowerTerm("parser-fmt/factorint/90/raw", raw, 1, 3, 2, failCount)
  end if
  if FormatterTryEvaluateAndFormat("factorint(76568758722)", fmt, raw) then
    AssertRawFactorintIntPowerTerm("parser-fmt/factorint/76568758722/raw", raw, 1, 3, 2, failCount)
  end if
  if FormatterTryEvaluateAndFormat("factorint(90071992)", fmt, raw) then
    AssertRawFactorintIntPowerTerm("parser-fmt/factorint/90071992/raw", raw, 0, 2, 3, failCount)
  end if
  if FormatterTryEvaluateAndFormat("factorint(-3333*9)", fmt, raw) then
    AssertRawFactorintIntPowerTerm("parser-fmt/factorint/-3333*9/raw", raw, 0, -3, 3, failCount)
  end if
  if FormatterTryEvaluateAndFormat("factorint(-9999)", fmt, raw) then
    AssertRawFactorintIntPowerTerm("parser-fmt/factorint/-9999/raw", raw, 0, -3, 2, failCount)
  end if
  if FormatterTryEvaluateAndFormat("factorint(-12)", fmt, raw) then
    AssertRawFactorintIntPowerTerm("parser-fmt/factorint/-12/raw", raw, 0, -2, 2, failCount)
  end if
end sub

private sub RunParserFormatterIntegrationTests(byref failCount as Integer)
  print !""
  print !"-- Parser + formatter integration (ComplexNumbers ON, 2 decimal places)"

  Parser_ClearVariables()
  Parser_SetSupportComplexNumbers(TRUE)
  FormatterTestSetup()
  g_nDecimals = 2
  g_bUseThousandsSeparator = FALSE
  ApplySeparatorDefaults()

  dim cxLit as String = "0x7FFFFFFFFFFFFFFF-0x7FFFFFFFFFFFFFFFi"
  dim cxLit2 as String = "-0x7FFFFFFFFFFFFFFF-0x7FFFFFFFFFFFFFFFi"
  dim cxMaxLit as String = "0x7FFFFFFFFFFFFFFF+0x7FFFFFFFFFFFFFFFi"
  dim raw as RawResult
  dim fmt as String

  if FormatterTryEvaluateAndFormat("hex(10+15i)", fmt, raw) = FALSE then
    print !"[FAIL] parser-fmt/hex(10+15i) — evaluate failed"
    failCount += 1
  else
    AssertEq("parser-fmt/hex(10+15i)", fmt, SMARTMATH_RESULT_PREFIX & "0xA + 0xF*i", failCount)
  end if
  if FormatterTryEvaluateAndFormat("hex(cart(" & cxMaxLit & "))", fmt, raw) = FALSE then
    print !"[FAIL] parser-fmt/hex(cart max i64) — evaluate failed"
    failCount += 1
  else
    AssertEq("parser-fmt/hex(cart max i64)", fmt, SMARTMATH_RESULT_PREFIX & "0x7FFFFFFFFFFFFFFF + 0x7FFFFFFFFFFFFFFF*i", failCount)
  end if
  if FormatterTryEvaluateAndFormat("hex(conj(" & cxMaxLit & "))", fmt, raw) = FALSE then
    print !"[FAIL] parser-fmt/hex(conj max i64) — evaluate failed"
    failCount += 1
  else
    AssertEq("parser-fmt/hex(conj max i64)", fmt, SMARTMATH_RESULT_PREFIX & "0x7FFFFFFFFFFFFFFF - 0x7FFFFFFFFFFFFFFF*i", failCount)
  end if
  if FormatterTryEvaluateAndFormat("(0xFFFFFFFFFFFFFF + 0x7FFFFFFFFFFFFFi);uhex", fmt, raw) = FALSE then
    print !"[FAIL] parser-fmt/uhex mixed hex complex — evaluate failed"
    failCount += 1
  else
    AssertEq("parser-fmt/uhex mixed hex complex", fmt, SMARTMATH_RESULT_PREFIX & "0xFFFFFFFFFFFFFF + 0x7FFFFFFFFFFFFF*i", failCount)
  end if

  dim bareExpr(1 to 2) as String
  dim bareWant(1 to 2) as String
  bareExpr(1) = cxLit
  bareWant(1) = SMARTMATH_RESULT_PREFIX & "9223372036854775807 - 9223372036854775807i"
  bareExpr(2) = cxLit2
  bareWant(2) = SMARTMATH_RESULT_PREFIX & "-9223372036854775807 - 9223372036854775807i"

  dim bi as Integer
  for bi = 1 to 2
    dim tagBare as String = "parser-fmt/bare/" & bareExpr(bi)
    if FormatterTryEvaluateAndFormat(bareExpr(bi), fmt, raw) = FALSE then
      print !"[FAIL] "; tagBare; !" — evaluate failed"
      failCount += 1
    else
      AssertEq(tagBare, fmt, bareWant(bi), failCount)
      AssertRawComplexInt64Parts(tagBare & "/raw", raw, failCount)
      AssertNoSubstr(tagBare & "/no-sci", fmt, "e+", failCount)
      AssertNoSubstr(tagBare & "/no-sci", fmt, "e-", failCount)
      AssertNoSubstr(tagBare & "/no-sci", fmt, "E+", failCount)
      AssertNoSubstr(tagBare & "/no-sci", fmt, "E-", failCount)
    end if
  next bi

  dim fnExpr(1 to 8) as String
  dim fnWant(1 to 8) as String
  fnExpr(1) = "real(" & cxLit & ")": fnWant(1) = SMARTMATH_RESULT_PREFIX & "9223372036854775807"
  fnExpr(2) = "imag(" & cxLit & ")": fnWant(2) = SMARTMATH_RESULT_PREFIX & "-9223372036854775807"
  fnExpr(3) = "cart(" & cxLit & ")": fnWant(3) = bareWant(1)
  fnExpr(4) = "conj(" & cxLit & ")": fnWant(4) = SMARTMATH_RESULT_PREFIX & "9223372036854775807 + 9223372036854775807i"
  fnExpr(5) = "real(" & cxLit2 & ")": fnWant(5) = SMARTMATH_RESULT_PREFIX & "-9223372036854775807"
  fnExpr(6) = "imag(" & cxLit2 & ")": fnWant(6) = SMARTMATH_RESULT_PREFIX & "-9223372036854775807"
  fnExpr(7) = "cart(" & cxLit2 & ")": fnWant(7) = bareWant(2)
  fnExpr(8) = "conj(" & cxLit2 & ")": fnWant(8) = SMARTMATH_RESULT_PREFIX & "-9223372036854775807 + 9223372036854775807i"

  dim fi as Integer
  for fi = 1 to 8
    dim tagFn as String = "parser-fmt/" & fnExpr(fi)
    if FormatterTryEvaluateAndFormat(fnExpr(fi), fmt, raw) = FALSE then
      print !"[FAIL] "; tagFn; !" — evaluate failed"
      failCount += 1
    else
      AssertEq(tagFn, fmt, fnWant(fi), failCount)
      AssertNoSubstr(tagFn & "/no-sci", fmt, "e+", failCount)
      AssertNoSubstr(tagFn & "/no-sci", fmt, "e-", failCount)
      if fi = 1 orelse fi = 2 orelse fi = 5 orelse fi = 6 then
        AssertRawScalarInt64(tagFn & "/raw", raw, failCount)
      else
        AssertRawComplexInt64Parts(tagFn & "/raw", raw, failCount)
      end if
    end if
  next fi

  if FormatterTryEvaluateAndFormat("6*(1/2)", fmt, raw) then
    AssertEq("parser-fmt/exact-mul/6*(1/2)", fmt, SMARTMATH_RESULT_PREFIX & "3", failCount)
    AssertRawScalarInt64("parser-fmt/exact-mul/6*(1/2)/raw", raw, failCount)
  else
    print !"[FAIL] parser-fmt/exact-mul/6*(1/2) — evaluate failed"
    failCount += 1
  end if

  if FormatterTryEvaluateAndFormat("6*(1/3)", fmt, raw) then
    AssertEq("parser-fmt/exact-mul/6*(1/3)", fmt, SMARTMATH_RESULT_PREFIX & "2", failCount)
    AssertRawScalarInt64("parser-fmt/exact-mul/6*(1/3)/raw", raw, failCount)
  else
    print !"[FAIL] parser-fmt/exact-mul/6*(1/3) — evaluate failed"
    failCount += 1
  end if

  dim ratioCxExpr(1 to 4) as String
  dim ratioCxWant(1 to 4) as String
  ratioCxExpr(1) = "ratio(0.5+0.25i)": ratioCxWant(1) = SMARTMATH_RESULT_PREFIX & "1/2 + 1/4*i"
  ratioCxExpr(2) = "ratio(0.25i)": ratioCxWant(2) = SMARTMATH_RESULT_PREFIX & "1/4*i"
  ratioCxExpr(3) = "ratio(2+3i)": ratioCxWant(3) = SMARTMATH_RESULT_PREFIX & "2 + 3i"
  ratioCxExpr(4) = "ratio(0.5+0.25i)+1": ratioCxWant(4) = SMARTMATH_RESULT_PREFIX & "1.5 + 0.25i"
  dim rcx as Integer
  for rcx = 1 to 4
    dim tagRatioCx as String = "parser-fmt/" & ratioCxExpr(rcx)
    if FormatterTryEvaluateAndFormat(ratioCxExpr(rcx), fmt, raw) = FALSE then
      print !"[FAIL] "; tagRatioCx; !" — evaluate failed"
      failCount += 1
    else
      AssertEq(tagRatioCx, fmt, ratioCxWant(rcx), failCount)
      if rcx = 4 then
        AssertNoSubstr(tagRatioCx & "/no-wrong-imag-unit", fmt, "1*i", failCount)
      end if
    end if
  next rcx

  if FormatterTryEvaluateAndFormat("0xFFFFFFFFFFFFFFFF*(1/5)", fmt, raw) then
    AssertEq("parser-fmt/exact-mul/0xFFFFFFFFFFFFFFFF*(1/5)", fmt, SMARTMATH_RESULT_PREFIX & "3689348814741910323", failCount)
    if raw.kind <> RRK_SCALAR orelse (raw.scalar.kind <> RSK_UINT64 andalso raw.scalar.kind <> RSK_INT64) then
      print !"[FAIL] parser-fmt/exact-mul/0xFFFFFFFFFFFFFFFF*(1/5)/raw — expected integer scalar"
      failCount += 1
    else
      print !"[PASS] parser-fmt/exact-mul/0xFFFFFFFFFFFFFFFF*(1/5)/raw raw integer scalar"
    end if
  else
    print !"[FAIL] parser-fmt/exact-mul/0xFFFFFFFFFFFFFFFF*(1/5) — evaluate failed"
    failCount += 1
  end if

  if FormatterTryEvaluateAndFormat("0x7FFFFFFFFFFFFFFF*(1/7)", fmt, raw) then
    AssertEq("parser-fmt/exact-mul/0x7FFFFFFFFFFFFFFF*(1/7)", fmt, SMARTMATH_RESULT_PREFIX & "1317624576693539401", failCount)
    if raw.kind <> RRK_SCALAR orelse (raw.scalar.kind <> RSK_INT64 andalso raw.scalar.kind <> RSK_UINT64) then
      print !"[FAIL] parser-fmt/exact-mul/0x7FFFFFFFFFFFFFFF*(1/7)/raw — expected integer scalar"
      failCount += 1
    else
      print !"[PASS] parser-fmt/exact-mul/0x7FFFFFFFFFFFFFFF*(1/7)/raw raw integer scalar"
    end if
  else
    print !"[FAIL] parser-fmt/exact-mul/0x7FFFFFFFFFFFFFFF*(1/7) — evaluate failed"
    failCount += 1
  end if

  if FormatterTryEvaluateAndFormat("2**64*(1/2)", fmt, raw) then
    AssertTrue("parser-fmt/exact-mul/2**64*(1/2)/sci", InStr(1, fmt, "e+") > 0 orelse InStr(1, fmt, "e-") > 0, failCount)
    if raw.kind <> RRK_SCALAR orelse raw.scalar.kind <> RSK_FLOATING then
      print !"[FAIL] parser-fmt/exact-mul/2**64*(1/2)/raw — expected RSK_FLOATING scalar"
      failCount += 1
    else
      print !"[PASS] parser-fmt/exact-mul/2**64*(1/2)/raw raw float scalar"
    end if
  else
    print !"[FAIL] parser-fmt/exact-mul/2**64*(1/2) — evaluate failed"
    failCount += 1
  end if

  if FormatterTryEvaluateAndFormat("0x7FFFFFFFFFFFFFFF*(1/8)", fmt, raw) then
    AssertTrue("parser-fmt/exact-mul/0x7FFFFFFFFFFFFFFF*(1/8)/sci", InStr(1, fmt, "e+") > 0 orelse InStr(1, fmt, "e-") > 0, failCount)
    if raw.kind <> RRK_SCALAR orelse raw.scalar.kind <> RSK_FLOATING then
      print !"[FAIL] parser-fmt/exact-mul/0x7FFFFFFFFFFFFFFF*(1/8)/raw — expected RSK_FLOATING scalar"
      failCount += 1
    else
      print !"[PASS] parser-fmt/exact-mul/0x7FFFFFFFFFFFFFFF*(1/8)/raw raw float scalar"
    end if
  else
    print !"[FAIL] parser-fmt/exact-mul/0x7FFFFFFFFFFFFFFF*(1/8) — evaluate failed"
    failCount += 1
  end if

  if FormatterTryEvaluateAndFormat("2**52*(1/3)", fmt, raw) then
    AssertTrue("parser-fmt/exact-mul/2**52*(1/3)/sci", InStr(1, fmt, "e+") > 0 orelse InStr(1, fmt, "e-") > 0, failCount)
    if raw.kind <> RRK_SCALAR orelse raw.scalar.kind <> RSK_FLOATING then
      print !"[FAIL] parser-fmt/exact-mul/2**52*(1/3)/raw — expected RSK_FLOATING scalar"
      failCount += 1
    else
      print !"[PASS] parser-fmt/exact-mul/2**52*(1/3)/raw raw float scalar"
    end if
  else
    print !"[FAIL] parser-fmt/exact-mul/2**52*(1/3) — evaluate failed"
    failCount += 1
  end if

  if FormatterTryEvaluateAndFormat("-0x7FFFFFFFFFFFFFFF*(1/7)", fmt, raw) then
    AssertEq("parser-fmt/exact-mul/-0x7FFFFFFFFFFFFFFF*(1/7)", fmt, SMARTMATH_RESULT_PREFIX & "-1317624576693539401", failCount)
    if raw.kind <> RRK_SCALAR orelse (raw.scalar.kind <> RSK_INT64 andalso raw.scalar.kind <> RSK_UINT64) then
      print !"[FAIL] parser-fmt/exact-mul/-0x7FFFFFFFFFFFFFFF*(1/7)/raw — expected integer scalar"
      failCount += 1
    else
      print !"[PASS] parser-fmt/exact-mul/-0x7FFFFFFFFFFFFFFF*(1/7)/raw raw integer scalar"
    end if
  else
    print !"[FAIL] parser-fmt/exact-mul/-0x7FFFFFFFFFFFFFFF*(1/7) — evaluate failed"
    failCount += 1
  end if

  if FormatterTryEvaluateAndFormat("0x7FFFFFFFFFFFFFFF*(-1/7)", fmt, raw) then
    AssertEq("parser-fmt/exact-mul/0x7FFFFFFFFFFFFFFF*(-1/7)", fmt, SMARTMATH_RESULT_PREFIX & "-1317624576693539401", failCount)
    if raw.kind <> RRK_SCALAR orelse (raw.scalar.kind <> RSK_INT64 andalso raw.scalar.kind <> RSK_UINT64) then
      print !"[FAIL] parser-fmt/exact-mul/0x7FFFFFFFFFFFFFFF*(-1/7)/raw — expected integer scalar"
      failCount += 1
    else
      print !"[PASS] parser-fmt/exact-mul/0x7FFFFFFFFFFFFFFF*(-1/7)/raw raw integer scalar"
    end if
  else
    print !"[FAIL] parser-fmt/exact-mul/0x7FFFFFFFFFFFFFFF*(-1/7) — evaluate failed"
    failCount += 1
  end if

  if FormatterTryEvaluateAndFormat("-2**52*(1/2)", fmt, raw) then
    AssertEq("parser-fmt/exact-mul/-2**52*(1/2)", fmt, SMARTMATH_RESULT_PREFIX & "-2251799813685248", failCount)
    AssertRawScalarInt64("parser-fmt/exact-mul/-2**52*(1/2)/raw", raw, failCount)
  else
    print !"[FAIL] parser-fmt/exact-mul/-2**52*(1/2) — evaluate failed"
    failCount += 1
  end if

  if FormatterTryEvaluateAndFormat("2**52*(-1/2)", fmt, raw) then
    AssertEq("parser-fmt/exact-mul/2**52*(-1/2)", fmt, SMARTMATH_RESULT_PREFIX & "-2251799813685248", failCount)
    AssertRawScalarInt64("parser-fmt/exact-mul/2**52*(-1/2)/raw", raw, failCount)
  else
    print !"[FAIL] parser-fmt/exact-mul/2**52*(-1/2) — evaluate failed"
    failCount += 1
  end if

  if FormatterTryEvaluateAndFormat("1317624576693539401*(1/11)", fmt, raw) then
    AssertTrue("parser-fmt/exact-mul/1317624576693539401*(1/11)/sci", InStr(1, fmt, "e+") > 0 orelse InStr(1, fmt, "e-") > 0, failCount)
    if raw.kind <> RRK_SCALAR orelse raw.scalar.kind <> RSK_FLOATING then
      print !"[FAIL] parser-fmt/exact-mul/1317624576693539401*(1/11)/raw — expected RSK_FLOATING scalar"
      failCount += 1
    else
      print !"[PASS] parser-fmt/exact-mul/1317624576693539401*(1/11)/raw raw float scalar"
    end if
  else
    print !"[FAIL] parser-fmt/exact-mul/1317624576693539401*(1/11) — evaluate failed"
    failCount += 1
  end if

  if FormatterTryEvaluateAndFormat("0x7FFFFFFFFFFFFFFF*(-1/8)", fmt, raw) then
    AssertTrue("parser-fmt/exact-mul/0x7FFFFFFFFFFFFFFF*(-1/8)/sci", InStr(1, fmt, "e+") > 0 orelse InStr(1, fmt, "e-") > 0, failCount)
    if raw.kind <> RRK_SCALAR orelse raw.scalar.kind <> RSK_FLOATING then
      print !"[FAIL] parser-fmt/exact-mul/0x7FFFFFFFFFFFFFFF*(-1/8)/raw — expected RSK_FLOATING scalar"
      failCount += 1
    else
      print !"[PASS] parser-fmt/exact-mul/0x7FFFFFFFFFFFFFFF*(-1/8)/raw raw float scalar"
    end if
  else
    print !"[FAIL] parser-fmt/exact-mul/0x7FFFFFFFFFFFFFFF*(-1/8) — evaluate failed"
    failCount += 1
  end if

  if FormatterTryEvaluateAndFormat("1317624576693539401/(1/11)", fmt, raw) then
    AssertEq("parser-fmt/exact-div/1317624576693539401/(1/11)", fmt, SMARTMATH_RESULT_PREFIX & "14493870343628933411", failCount)
    if raw.kind <> RRK_SCALAR orelse (raw.scalar.kind <> RSK_INT64 andalso raw.scalar.kind <> RSK_UINT64) then
      print !"[FAIL] parser-fmt/exact-div/1317624576693539401/(1/11)/raw — expected integer scalar"
      failCount += 1
    else
      print !"[PASS] parser-fmt/exact-div/1317624576693539401/(1/11)/raw raw integer scalar"
    end if
  else
    print !"[FAIL] parser-fmt/exact-div/1317624576693539401/(1/11) — evaluate failed"
    failCount += 1
  end if

  if FormatterTryEvaluateAndFormat("1317624576693539401/(1/7)", fmt, raw) then
    AssertEq("parser-fmt/exact-div/1317624576693539401/(1/7)", fmt, SMARTMATH_RESULT_PREFIX & "9223372036854775807", failCount)
    if raw.kind <> RRK_SCALAR orelse (raw.scalar.kind <> RSK_INT64 andalso raw.scalar.kind <> RSK_UINT64) then
      print !"[FAIL] parser-fmt/exact-div/1317624576693539401/(1/7)/raw — expected integer scalar"
      failCount += 1
    else
      print !"[PASS] parser-fmt/exact-div/1317624576693539401/(1/7)/raw raw integer scalar"
    end if
  else
    print !"[FAIL] parser-fmt/exact-div/1317624576693539401/(1/7) — evaluate failed"
    failCount += 1
  end if

  if FormatterTryEvaluateAndFormat("-1317624576693539401/(1/7)", fmt, raw) then
    AssertEq("parser-fmt/exact-div/-1317624576693539401/(1/7)", fmt, SMARTMATH_RESULT_PREFIX & "-9223372036854775807", failCount)
    AssertRawScalarInt64("parser-fmt/exact-div/-1317624576693539401/(1/7)/raw", raw, failCount)
  else
    print !"[FAIL] parser-fmt/exact-div/-1317624576693539401/(1/7) — evaluate failed"
    failCount += 1
  end if

  if FormatterTryEvaluateAndFormat("0x7FFFFFFFFFFFFFFF/(1/8)", fmt, raw) then
    AssertTrue("parser-fmt/exact-div/0x7FFFFFFFFFFFFFFF/(1/8)/sci", InStr(1, fmt, "e+") > 0 orelse InStr(1, fmt, "e-") > 0, failCount)
    if raw.kind <> RRK_SCALAR orelse raw.scalar.kind <> RSK_FLOATING then
      print !"[FAIL] parser-fmt/exact-div/0x7FFFFFFFFFFFFFFF/(1/8)/raw — expected RSK_FLOATING scalar"
      failCount += 1
    else
      print !"[PASS] parser-fmt/exact-div/0x7FFFFFFFFFFFFFFF/(1/8)/raw raw float scalar"
    end if
  else
    print !"[FAIL] parser-fmt/exact-div/0x7FFFFFFFFFFFFFFF/(1/8) — evaluate failed"
    failCount += 1
  end if

  if FormatterTryEvaluateAndFormat("2**64/(1/2)", fmt, raw) then
    AssertTrue("parser-fmt/exact-div/2**64/(1/2)/sci", InStr(1, fmt, "e+") > 0 orelse InStr(1, fmt, "e-") > 0, failCount)
    if raw.kind <> RRK_SCALAR orelse raw.scalar.kind <> RSK_FLOATING then
      print !"[FAIL] parser-fmt/exact-div/2**64/(1/2)/raw — expected RSK_FLOATING scalar"
      failCount += 1
    else
      print !"[PASS] parser-fmt/exact-div/2**64/(1/2)/raw raw float scalar"
    end if
  else
    print !"[FAIL] parser-fmt/exact-div/2**64/(1/2) — evaluate failed"
    failCount += 1
  end if

  '' Regression guard: float raw export must not be used for these magnitudes (shows sci at 2 decimals).
  FormatterTestSetup()
  g_nDecimals = 2
  RawResultClear(raw)
  raw.kind = RRK_SCALAR
  raw.scalar.kind = RSK_COMPLEX
  raw.scalar.real.kind = RSK_FLOATING
  raw.scalar.real.floatValue = 9223372036854775808.0
  raw.scalar.imag.kind = RSK_FLOATING
  raw.scalar.imag.floatValue = -9223372036854775808.0
  dim floatFmt as String = FormatRawEvaluationResult(raw)
  AssertTrue("parser-fmt/float-raw/shows-sci", InStr(1, floatFmt, "e+") > 0 orelse InStr(1, floatFmt, "e-") > 0, failCount)

  '' User SmartMath.ini profile: Decimals=-1, ThousandsSeparator=1, ComplexNumbers=1.
  print !""
  print !"-- Parser + formatter integration (user ini: decimals=-1, thousands ON)"
  Parser_ClearVariables()
  Parser_SetSupportComplexNumbers(TRUE)
  FormatterTestSetup()
  g_nDecimals = -1
  g_bUseThousandsSeparator = TRUE
  g_sThousandsSeparator = "'"
  ApplySeparatorDefaults()

  dim cxLitIni as String = "0x7FFFFFFFFFFFFFFF-0x7FFFFFFFFFFFFFFFi"
  dim cxLitIni2 as String = "-0x7FFFFFFFFFFFFFFF-0x7FFFFFFFFFFFFFFFi"
  dim iniBareWant(1 to 2) as String
  iniBareWant(1) = SMARTMATH_RESULT_PREFIX & "9'223'372'036'854'775'807 - 9'223'372'036'854'775'807i"
  iniBareWant(2) = SMARTMATH_RESULT_PREFIX & "-9'223'372'036'854'775'807 - 9'223'372'036'854'775'807i"
  dim iniBareExpr(1 to 2) as String
  iniBareExpr(1) = cxLitIni
  iniBareExpr(2) = cxLitIni2

  for bi = 1 to 2
    dim tagIniBare as String = "parser-fmt/ini-bare/" & iniBareExpr(bi)
    if FormatterTryEvaluateAndFormat(iniBareExpr(bi), fmt, raw) = FALSE then
      print !"[FAIL] "; tagIniBare; !" — evaluate failed"
      failCount += 1
    else
      AssertEq(tagIniBare, fmt, iniBareWant(bi), failCount)
      AssertRawComplexInt64Parts(tagIniBare & "/raw", raw, failCount)
      AssertNoSubstr(tagIniBare & "/no-sci", fmt, "e+", failCount)
      AssertNoSubstr(tagIniBare & "/no-sci", fmt, "e-", failCount)
    end if
  next bi

  dim iniFnExpr(1 to 4) as String
  dim iniFnWant(1 to 4) as String
  iniFnExpr(1) = "real(" & cxLitIni & ")": iniFnWant(1) = SMARTMATH_RESULT_PREFIX & "9'223'372'036'854'775'807"
  iniFnExpr(2) = "imag(" & cxLitIni & ")": iniFnWant(2) = SMARTMATH_RESULT_PREFIX & "-9'223'372'036'854'775'807"
  iniFnExpr(3) = "real(" & cxLitIni2 & ")": iniFnWant(3) = SMARTMATH_RESULT_PREFIX & "-9'223'372'036'854'775'807"
  iniFnExpr(4) = "imag(" & cxLitIni2 & ")": iniFnWant(4) = SMARTMATH_RESULT_PREFIX & "-9'223'372'036'854'775'807"
  for fi = 1 to 4
    dim tagIniFn as String = "parser-fmt/ini-fn/" & iniFnExpr(fi)
    if FormatterTryEvaluateAndFormat(iniFnExpr(fi), fmt, raw) = FALSE then
      print !"[FAIL] "; tagIniFn; !" — evaluate failed"
      failCount += 1
    else
      AssertEq(tagIniFn, fmt, iniFnWant(fi), failCount)
      AssertRawScalarInt64(tagIniFn & "/raw", raw, failCount)
      AssertNoSubstr(tagIniFn & "/no-sci", fmt, "e+", failCount)
      AssertNoSubstr(tagIniFn & "/no-sci", fmt, "e-", failCount)
    end if
  next fi

  '' sqrt(2**70) and sqrt(-2**70): operand is float (2^70 > i64); sqrt uses float unless verified exact int.
  FormatterTestSetup()
  g_nDecimals = -1
  g_bUseThousandsSeparator = FALSE
  ApplySeparatorDefaults()
  dim sqrtPosFmt as String, sqrtNegFmt as String
  if FormatterTryEvaluateAndFormat("sqrt(2**70)", sqrtPosFmt, raw) = FALSE then
    print !"[FAIL] parser-fmt/sqrt(2**70) — evaluate failed"
    failCount += 1
  else
    AssertEq("parser-fmt/sqrt(2**70)", sqrtPosFmt, SMARTMATH_RESULT_PREFIX & "34359738368.0", failCount)
    if raw.kind <> RRK_SCALAR orelse raw.scalar.real.kind <> RSK_FLOATING then
      print !"[FAIL] parser-fmt/sqrt(2**70)/raw — expected RSK_FLOATING scalar"
      failCount += 1
    else
      print !"[PASS] parser-fmt/sqrt(2**70)/raw float"
    end if
  end if
  if FormatterTryEvaluateAndFormat("sqrt(-2**70)", sqrtNegFmt, raw) = FALSE then
    print !"[FAIL] parser-fmt/sqrt(-2**70) — evaluate failed"
    failCount += 1
  else
    AssertEq("parser-fmt/sqrt(-2**70)", sqrtNegFmt, SMARTMATH_RESULT_PREFIX & "34359738368.0i", failCount)
    if raw.scalar.kind <> RSK_COMPLEX orelse raw.scalar.imag.kind <> RSK_FLOATING then
      print !"[FAIL] parser-fmt/sqrt(-2**70)/raw — expected complex float imag"
      failCount += 1
    else
      print !"[PASS] parser-fmt/sqrt(-2**70)/raw float imag"
    end if
  end if

  RunFactorintFormatterTests(failCount)

  FormatterTestSetup()
end sub

#endif

#ifdef PARSER_FORMATTER_INTEGRATION
RunParserFormatterIntegrationTests(failures)
#endif

print !""
print !"=== Result ==="
if failures = 0 then
  print !"All formatter tests passed."
  end 0
else
  print !"Failures: "; failures
  end 1
end if
