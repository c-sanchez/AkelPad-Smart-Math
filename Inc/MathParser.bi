#ifndef __MATHPARSER_BI__
#define __MATHPARSER_BI__

declare sub Parser_ClearVariables()
declare function Parser_TryEvaluate(byref sExpr as String, byref result as Double) as Boolean

#endif