#include once "crt.bi"
#include once "Inc\MathParser.bi"

type VarEntry
  name as String
  value as Double
end type

dim shared variables() as VarEntry
dim shared pStream as ZString ptr
dim shared parseError as Integer
dim shared wasPercentage as Boolean

private sub SkipSpaces()
  while (pStream[0] = 32) orelse (pStream[0] = 9) orelse (pStream[0] = 10) orelse (pStream[0] = 13)
    pStream += 1
  wend
end sub

private function GetVariable(byref n as String, byref v as Double) as Boolean
  dim i as Integer
  for i = lbound(variables) to ubound(variables)
    if variables(i).name = n then
      v = variables(i).value
      return TRUE
    end if
  next i
  return FALSE
end function

private sub SetVariable(byref n as String, v as Double)
  dim i as Integer
  for i = lbound(variables) to ubound(variables)
    if variables(i).name = n then
      variables(i).value = v
      exit sub
    end if
  next i
  if ubound(variables) = -1 then
    redim variables(0)
  else
    redim preserve variables(ubound(variables) + 1)
  end if
  variables(ubound(variables)).name = n
  variables(ubound(variables)).value = v
end sub

declare function ParseExpression() as Double

private function ParseFactor() as Double
  dim n as Double = 0
  wasPercentage = FALSE

  SkipSpaces()
  if (pStream[0] >= 48 andalso pStream[0] <= 57) orelse (pStream[0] = 46) then
    dim dVal as Double = 0, fract as Double = 1
    while pStream[0] >= 48 andalso pStream[0] <= 57
      dVal = dVal * 10 + (pStream[0] - 48)
      pStream += 1
    wend
    if pStream[0] = 46 then ' "."
      pStream += 1
      while pStream[0] >= 48 andalso pStream[0] <= 57
        fract /= 10
        dVal += (pStream[0] - 48) * fract
        pStream += 1
      wend
    end if
    n = dVal
  elseif (pStream[0] >= 65 andalso pStream[0] <= 90) orelse (pStream[0] >= 97 andalso pStream[0] <= 122) orelse (pStream[0] = 95) then
    dim pStart as ZString ptr = pStream
    while (pStream[0] >= 65 andalso pStream[0] <= 90) orelse (pStream[0] >= 97 andalso pStream[0] <= 122) orelse (pStream[0] >= 48 andalso pStream[0] <= 57) orelse (pStream[0] = 95)
      pStream += 1
    wend
    dim varName as String = Left(*pStart, pStream - pStart)
    if GetVariable(varName, n) = FALSE then parseError = 1
  elseif pStream[0] = 40 then ' (
    pStream += 1
    n = ParseExpression()
    SkipSpaces()
    if pStream[0] = 41 then pStream += 1 else parseError = 1
  else
    parseError = 1
  end if

  SkipSpaces()
  if pStream[0] = 37 then ' %
    pStream += 1
    n = n / 100.0
    wasPercentage = TRUE
  end if
  return n
end function

private function ParseTerm() as Double
  dim n as Double = ParseFactor()
  dim termWasPercentage as Boolean = wasPercentage
  SkipSpaces()
  while (pStream[0] = 42) orelse (pStream[0] = 47)
    if parseError then exit while
    dim op as UByte = pStream[0]
    pStream += 1
    dim n2 as Double = ParseFactor()
    if op = 42 then
      n *= n2
    else
      if n2 = 0 then parseError = 1 else n /= n2
    end if
    termWasPercentage = FALSE
    SkipSpaces()
  wend
  wasPercentage = termWasPercentage
  return n
end function

private function ParseExpression() as Double
  SkipSpaces()
  dim sign as Integer = 1
  if pStream[0] = 45 then ' "-"
    sign = -1
    pStream += 1
  elseif pStream[0] = 43 then ' "+"
    pStream += 1
  end if

  dim n as Double = ParseTerm() * sign
  SkipSpaces()
  while (pStream[0] = 43) orelse (pStream[0] = 45)
    if parseError then exit while
    dim op as UByte = pStream[0]
    pStream += 1
    dim n2 as Double = ParseTerm()

    if wasPercentage then
      SkipSpaces()
      if (pStream[0] = 0) orelse (pStream[0] = 41) orelse (pStream[0] = 43) orelse (pStream[0] = 45) then
        n2 = n * n2
      end if
    end if

    if op = 43 then n += n2 else n -= n2
    SkipSpaces()
  wend
  return n
end function

sub Parser_ClearVariables()
  erase variables
end sub

function Parser_TryEvaluate(byref sExpr as String, byref result as Double) as Boolean
  if Len(sExpr) = 0 then return FALSE

  dim i as Integer, hasDigitOrVar as Integer = 0
  for i = 1 to Len(sExpr)
    dim c as Integer = Asc(Mid(sExpr, i, 1))
    if (c >= 48 andalso c <= 57) orelse (c >= 65 andalso c <= 90) orelse (c >= 97 andalso c <= 122) then
      hasDigitOrVar = 1
      exit for
    end if
  next i
  if hasDigitOrVar = 0 then return FALSE

  pStream = StrPtr(sExpr)
  parseError = 0

  SkipSpaces()
  dim pStart as ZString ptr = pStream
  if (pStream[0] >= 65 andalso pStream[0] <= 90) orelse (pStream[0] >= 97 andalso pStream[0] <= 122) orelse (pStream[0] = 95) then
    while (pStream[0] >= 65 andalso pStream[0] <= 90) orelse (pStream[0] >= 97 andalso pStream[0] <= 122) orelse (pStream[0] >= 48 andalso pStream[0] <= 57) orelse (pStream[0] = 95)
      pStream += 1
    wend
    dim varName as String = Left(*pStart, pStream - pStart)
    SkipSpaces()
    if pStream[0] = 61 then ' =
      pStream += 1
      result = ParseExpression()
      SkipSpaces()
      if pStream[0] = 0 andalso parseError = 0 then
        SetVariable(varName, result)
        return TRUE
      end if
      return FALSE
    end if
  end if

  pStream = StrPtr(sExpr)
  parseError = 0
  result = ParseExpression()
  SkipSpaces()
  
  if pStream[0] <> 0 then parseError = 1
  if parseError = 1 then return FALSE
  return TRUE
end function