#ifndef __WIDTHGENERICDECIMAL_BI__
#define __WIDTHGENERICDECIMAL_BI__

'' Independent reference arithmetic for the width-generic tests (Basic mirror of the widthgeneric helpers in
'' cpp/tests/Tests_MathParser.cpp). Expected values are derived from SM_INT_BITS / SM_FLOAT_* by schoolbook
'' decimal arithmetic on digit strings, never from the parser's own limit tables or text helpers.
#include once "..\Inc\MathParserTypes.bi"

'' Decimal digits of 2^n.
private function WgDecPow2(byval n as Integer) as String
  dim d as String = "1"
  dim carry as Integer
  dim v as Integer
  for i as Integer = 1 to n
    carry = 0
    for k as Integer = len(d) to 1 step -1
      v = (asc(mid(d, k, 1)) - 48) * 2 + carry
      mid(d, k, 1) = chr(48 + v mod 10)
      carry = v \ 10
    next k
    if carry <> 0 then d = chr(48 + carry) & d
  next i
  return d
end function

'' 2^n - 1 (2^n never ends in 0, so only the last digit changes).
private function WgDecPow2Minus1(byval n as Integer) as String
  dim d as String = WgDecPow2(n)
  mid(d, len(d), 1) = chr(asc(right(d, 1)) - 1)
  return d
end function

private function WgDecPow10(byval e as Integer) as String
  return "1" & string(e, "0")
end function

private function WgDecMul(byref a as String, byval m as Integer) as String
  dim r as String = a
  dim carry as Integer = 0
  dim v as Integer
  for k as Integer = len(r) to 1 step -1
    v = (asc(mid(r, k, 1)) - 48) * m + carry
    mid(r, k, 1) = chr(48 + v mod 10)
    carry = v \ 10
  next k
  while carry <> 0
    r = chr(48 + carry mod 10) & r
    carry = carry \ 10
  wend
  return r
end function

private function WgDecFactorial(byval n as Integer) as String
  dim f as String = "1"
  for i as Integer = 2 to n
    f = WgDecMul(f, i)
  next i
  return f
end function

private function WgDecLessOrEqual(byref a as String, byref b as String) as Boolean
  if len(a) <> len(b) then return (len(a) < len(b))
  return (a <= b)
end function

private function WgIntMaxText() as String
  return WgDecPow2Minus1(SM_INT_BITS - 1)
end function

private function WgIntMinMagText() as String
  return WgDecPow2(SM_INT_BITS - 1)
end function

private function WgUIntMaxText() as String
  return WgDecPow2Minus1(SM_INT_BITS)
end function

'' Largest n with n! <= SmInt max.
private function WgMaxFactorialArg() as Integer
  dim limit as String = WgIntMaxText()
  dim n as Integer = 0
  while WgDecLessOrEqual(WgDecFactorial(n + 1), limit)
    n += 1
  wend
  return n
end function

'' Largest e with 10^e <= bound.
private function WgMaxPow10Exponent(byref bound as String) as Integer
  dim e as Integer = 0
  while WgDecLessOrEqual(WgDecPow10(e + 1), bound)
    e += 1
  wend
  return e
end function

'' 2^n by repeated doubling (no libm / exponent operator).
private function WgPow2Float(byval n as Integer) as Double
  dim r as Double = 1.0
  for i as Integer = 1 to n
    r *= 2.0
  next i
  return r
end function

'' 10^e by repeated multiplication / division.
private function WgPow10Float(byval e as Integer) as Double
  dim r as Double = 1.0
  for i as Integer = 1 to abs(e)
    r *= 10.0
  next i
  if e < 0 then return 1.0 / r
  return r
end function

#endif
