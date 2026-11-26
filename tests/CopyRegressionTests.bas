''
'' Regression tests for double-click clipboard normalization (SmartMath_CopyNormalize.bas).
''
#include once "..\SmartMath_Globals.bi"

private sub ApplySepDefaults()
  g_sDecimalSeparator = SMARTMATH_DECIMAL_SEPARATOR_DEFAULT
  g_sThousandsSeparator = SMARTMATH_THOUSANDS_SEPARATOR_DEFAULT
  g_sArrayOutputSeparator = SMARTMATH_ARRAY_OUTPUT_SEPARATOR_DEFAULT
  g_bUseThousandsSeparator = FALSE
end sub

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

dim fails as Integer = 0
print !"=== SmartMath copy normalization tests ==="
print !""

ApplySepDefaults()

'' User example: " as thousands, comma decimal, | array sep (rendered text after prefix)
g_bUseThousandsSeparator = TRUE
g_sThousandsSeparator = Chr(34)
g_sDecimalSeparator = ","
g_sArrayOutputSeparator = "|"
dim sQuote as String = SMARTMATH_RESULT_PREFIX & "(1" & Chr(34) & "024,222| 2" & Chr(34) & "048,444)"
AssertEq("tuple/quote-thou/comma-dec/pipe-array", NormalizeCopiedResult(sQuote), "(1024.222, 2048.444)", fails)

ApplySepDefaults()
g_bUseThousandsSeparator = TRUE
g_sThousandsSeparator = "'"
g_sDecimalSeparator = ","
g_sArrayOutputSeparator = "|"
dim sApo as String = SMARTMATH_RESULT_PREFIX & "(1'024,222| 2'048,444)"
AssertEq("tuple/apostrophe-thou/comma-dec/pipe-array", NormalizeCopiedResult(sApo), "(1024.222, 2048.444)", fails)

ApplySepDefaults()
g_sDecimalSeparator = "."
g_sArrayOutputSeparator = "|"
'' Simulated locale leak: comma mantissa while ini says dot decimal
dim sLeak as String = SMARTMATH_RESULT_PREFIX & "(1024,222| 2048,444)"
AssertEq("tuple/comma-mantissa-with-dot-dec-config", NormalizeCopiedResult(sLeak), "(1024.222, 2048.444)", fails)

ApplySepDefaults()
g_bUseThousandsSeparator = TRUE
g_sThousandsSeparator = "'"
g_sDecimalSeparator = ","
g_sArrayOutputSeparator = ","
dim sCommaElem as String = SMARTMATH_RESULT_PREFIX & "(1'024,333, 2'048,666)"
AssertEq("tuple/comma-elem-sep European", NormalizeCopiedResult(sCommaElem), "(1024.333, 2048.666)", fails)

'' Decimal and array sep both comma: split on ", " only; do not strip decimal commas when thousands flag is off.
ApplySepDefaults()
g_sDecimalSeparator = ","
g_sArrayOutputSeparator = ","
g_bUseThousandsSeparator = FALSE
dim sBothComma as String = SMARTMATH_RESULT_PREFIX & "(123,456, 789,012)"
AssertEq("tuple/comma-dec and comma-array sep", NormalizeCopiedResult(sBothComma), "(123.456, 789.012)", fails)

ApplySepDefaults()
g_bUseThousandsSeparator = TRUE
g_sDecimalSeparator = ";"
g_sThousandsSeparator = "'"
dim sSemi as String = SMARTMATH_RESULT_PREFIX & "1'024;222"
AssertEq("scalar/semi-decimal", NormalizeCopiedResult(sSemi), "1024.222", fails)

ApplySepDefaults()
dim sHex as String = SMARTMATH_RESULT_PREFIX & "0xFF"
AssertEq("scalar/hex passthrough", NormalizeCopiedResult(sHex), "0xFF", fails)

print !""
print !"=== Result ==="
if fails = 0 then
  print !"All copy normalization tests passed."
  end 0
else
  print !"Failures: "; fails
  end 1
end if
