'' Minimal shared state for SmartMath_Format.bas when linked outside SmartMath.dll.
#include once "SmartMath_Globals.bi"

dim shared lpEditProcData as WNDPROCDATA ptr = 0
dim shared g_bSmartMathActive as BOOL = FALSE
dim shared g_bOldRichEdit as BOOL = FALSE

dim shared rcOldMargin as RECT
dim shared nOldFirstLine as Integer = 0
dim shared nOldCaretLine as Integer = 0
dim shared nOldMargin as Integer = 0

dim shared g_nDecimals as Integer = -1
dim shared g_crResultColor as COLORREF = 0
dim shared g_bUseThousandsSeparator as BOOL = FALSE
dim shared g_bSupportComplexNumbers as BOOL = FALSE
dim shared g_bLogParsedLines as BOOL = FALSE
dim shared g_sDecimalSeparator as String
dim shared g_sThousandsSeparator as String
dim shared g_sArrayOutputSeparator as String
dim shared hSmartMathMenu as HMENU = 0
dim shared hSubMenuDecimals as HMENU = 0
dim shared hSubMenuColor as HMENU = 0
dim shared g_hMainWnd as HWND = 0
dim shared g_hMainMenu as HMENU = 0
dim shared g_hWndEdit as HWND = 0
dim shared g_bShuttingDown as BOOL = FALSE
