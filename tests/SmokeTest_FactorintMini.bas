#include once "..\Inc\MathParser.bi"

sub Main()
  Parser_ClearVariables()
  dim result as Double
  dim resultText as String
  dim isArray as Boolean

  print "evaluating 1+1..."
  if Parser_TryEvaluateEx("1+1", result, resultText, isArray) then
    print "ok: " & resultText
  else
    print "err: " & Parser_GetLastError()
  end if

  print "evaluating factorint(33)..."
  dim t0 as Double = timer
  if Parser_TryEvaluateEx("factorint(33)", result, resultText, isArray) then
    print "ok (" & (timer - t0) & "s): " & resultText
  else
    print "err (" & (timer - t0) & "s): " & Parser_GetLastError()
  end if
  end 0
end sub

Main()
