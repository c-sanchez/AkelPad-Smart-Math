#include once "..\Inc\MathParser.bi"

function main() as Integer
  dim expr as String = ""
  dim useComplex as Boolean = FALSE
  if lcase(trim(command(1))) = "--complex" then
    useComplex = TRUE
    expr = command(2)
  else
    expr = command(1)
  end if
  if len(expr) = 0 then return 1
  if useComplex then Parser_SetSupportComplexNumbers(TRUE)
  dim r as Double
  dim rt as String
  dim ia as Boolean
  if Parser_TryEvaluateEx(expr, r, rt, ia) then
    if len(rt) > 0 then
      print rt;
    else
      print str(r);
    end if
    return 0
  end if
  print "ERROR:"; Parser_GetLastError();
  return 1
end function

end main()
