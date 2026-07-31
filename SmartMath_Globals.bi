#pragma once
#include once "windows.bi"
#include once "win/shellapi.bi"
#include once "vbcompat.bi"
#include once "win/richedit.bi"
#include once "Inc\AkelEdit.bi"
#include once "Inc\AkelDLL.bi"
#include once "Inc\MathParser.bi"
#include once "Inc\MathParserRawResult.bi"

#ifndef EM_GETFIRSTVISIBLELINE
  const EM_GETFIRSTVISIBLELINE = &h00CE
#endif
#ifndef EM_REPLACESEL
  const EM_REPLACESEL = &hC2
#endif

' Max fixed decimal places in UI and in stored Decimals= option (values above clamp).
const SMARTMATH_DECIMALS_MAX = 14

' -----------------------------------------------------------------------------
'  Menu IDs for SmartMath decimal-place selector.
'  Range 12000-12015 avoids AkelPad built-ins (<=10019) and common plugins.
'    IDM_DECIMAL_BASE + 0  = "Auto (full precision)"
'    IDM_DECIMAL_BASE + 1  = "0 decimal places"
'    ...
'    IDM_DECIMAL_BASE + 15 = "14 decimal places"
' -----------------------------------------------------------------------------
const IDM_DECIMAL_BASE = 12000

' -----------------------------------------------------------------------------
'  Menu IDs for SmartMath color selector.
' -----------------------------------------------------------------------------
const IDM_COLOR_BASE   = 12100
const IDM_COLOR_GREEN  = IDM_COLOR_BASE + 0
const IDM_COLOR_BLUE   = IDM_COLOR_BASE + 1
const IDM_COLOR_RED    = IDM_COLOR_BASE + 2
const IDM_COLOR_YELLOW = IDM_COLOR_BASE + 3
const IDM_COLOR_WHITE  = IDM_COLOR_BASE + 4
const IDM_COLOR_BLACK  = IDM_COLOR_BASE + 5
const MAX_COLORS       = 6

' -----------------------------------------------------------------------------
'  Menu ID for Options & About
' -----------------------------------------------------------------------------
const IDM_ABOUT               = 12200
const IDM_THOUSANDS_SEPARATOR = 12201
const IDM_COMPLEX_NUMBERS     = 12202
const IDM_ACTIVE_ON_CURR_TAB  = 12203

' -----------------------------------------------------------------------------
'  Shared formatting/render defaults
' -----------------------------------------------------------------------------
const SMARTMATH_DECIMAL_SEPARATOR_DEFAULT      = "."
const SMARTMATH_THOUSANDS_SEPARATOR_DEFAULT    = "'"
const SMARTMATH_ARRAY_OUTPUT_SEPARATOR_DEFAULT = ","
const SMARTMATH_RESULT_PREFIX       = " = "
const SMARTMATH_ERROR_PREFIX        = " ! "

' Pixel gap between the end of a line's source text and the result gutter, and
' between the result and the client right edge.
const SMARTMATH_RESULT_GUTTER_GAP   = 6
const SMARTMATH_RESULT_RIGHT_MARGIN = 10

' -----------------------------------------------------------------------------
'  Global Variables (Extern)
' -----------------------------------------------------------------------------
extern lpEditProcData as WNDPROCDATA ptr
extern g_bSmartMathActive as BOOL
extern g_bOldRichEdit as BOOL

extern rcOldMargin as RECT
extern nOldFirstLine as Integer
extern nOldCaretLine as Integer
extern nOldMargin as Integer

extern g_nDecimals as Integer
extern g_crResultColor as COLORREF
extern g_bUseThousandsSeparator as BOOL
extern g_bSupportComplexNumbers as BOOL
extern g_bLogParsedLines as BOOL
extern g_sDecimalSeparator as String
extern g_sThousandsSeparator as String
extern g_sArrayOutputSeparator as String
extern hSmartMathMenu as HMENU
extern hSubMenuDecimals as HMENU
extern hSubMenuColor as HMENU
extern g_hMainWnd as HWND
extern g_hMainMenu as HMENU
extern g_hWndEdit as HWND
extern g_bShuttingDown as BOOL

' -----------------------------------------------------------------------------
'  Global Function Prototypes
' -----------------------------------------------------------------------------
declare sub LoadSettings()
declare sub SaveSettings()
declare function LoadSettings_ActiveFrames(wszActiveFrameFiles() as UShort) as Integer
declare sub SaveSettings_ActiveFrames(wszActiveFrameFiles() as UShort, byval nLen as Integer)
declare function FormatResult(byval d as Double) as String
declare function FormatNonFiniteDisplayFromParserScalar(byref s as String) as String
declare function FormatRawResultForDisplay(byref r as RawResult) as String
declare function FormatRawEvaluationResult(byref raw as RawResult) as String
declare function AddThousandsSeparator(byref sRes as String) as String
declare function NormalizeCopiedResult(byref sRes as String) as String
declare sub ResetSmartMathFormatLocaleCache()
declare sub InitSmartMathMenu()
declare sub UpdateMenuChecks()
declare sub UpdateMenuActiveOnCurrTab(byval bActive as BOOL)
declare sub UninitSmartMathMenu(byval bAppClosing as BOOL = FALSE)
declare sub ShowAboutDialog(byval hWnd as HWND)
declare sub UpdateInternalState(byval hWnd as HWND, byref bVisible as BOOL)
declare sub LogInfo(byref sMsg as String)