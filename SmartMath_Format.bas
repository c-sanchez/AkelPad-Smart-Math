#include once "SmartMath_Globals.bi"

' -----------------------------------------------------------------------------
'  Result Formatting
' -----------------------------------------------------------------------------
function FormatResult(byval d as Double) as String
  dim sRes as String

  if g_nDecimals < 0 then
    sRes = LTrim(Str(d))
  else
    dim sFmt as String
    if g_nDecimals = 0 then
      sFmt = "0"
    else
      sFmt = "0." & String(g_nDecimals, "0")
    end if

    sRes = Format(d, sFmt)

    if InStr(sRes, ".") > 0 orelse InStr(sRes, ",") > 0 then
      while Right(sRes, 1) = "0"
        sRes = Left(sRes, Len(sRes) - 1)
      wend
      if Right(sRes, 1) = "." orelse Right(sRes, 1) = "," then
        sRes = Left(sRes, Len(sRes) - 1)
      end if
    end if
  end if

  ' Force the decimal point to be a comma
  dim dotPos as Integer = InStr(sRes, ".")
  if dotPos > 0 then
    Mid(sRes, dotPos, 1) = ","
  end if

  if g_bUseThousandsSeparator then
    dim ePos as Integer = InStr(UCase(sRes), "E")
    dim expPart as String = ""
    if ePos > 0 then
      expPart = Mid(sRes, ePos)
      sRes = Left(sRes, ePos - 1)
    end if

    ' We now know that if there are decimals, the separator is always a comma ','
    dim decPos as Integer = InStr(sRes, ",")
    ' And we force the thousands separator to be a dot '.'
    dim localThouSep as String = "."

    dim intPart as String
    dim decPart as String

    if decPos > 0 then
      intPart = Left(sRes, decPos - 1)
      decPart = Mid(sRes, decPos)
    else
      intPart = sRes
      decPart = ""
    end if

    dim isNeg as Boolean = False
    if Left(intPart, 1) = "-" then
      isNeg = True
      intPart = Mid(intPart, 2)
    end if

    dim withCommas as String = ""
    dim count as Integer = 0
    dim i as Integer
    for i = Len(intPart) to 1 step -1
      count += 1
      withCommas = Mid(intPart, i, 1) & withCommas
      if count = 3 andalso i > 1 then
        withCommas = localThouSep & withCommas
        count = 0
      end if
    next i

    if isNeg then withCommas = "-" & withCommas
    sRes = withCommas & decPart & expPart
  end if

  return " = " & sRes
end function