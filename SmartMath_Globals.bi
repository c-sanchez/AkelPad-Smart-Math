#pragma once
#include once "windows.bi"
#include once "win/shellapi.bi"
#include once "vbcompat.bi"
#include once "win/richedit.bi"
#include once "Inc\AkelEdit.bi"
#include once "Inc\AkelDLL.bi"
#include once "Inc\MathParser.bi"

#ifndef EM_GETFIRSTVISIBLELINE
  const EM_GETFIRSTVISIBLELINE = &h00CE
#endif

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

' -----------------------------------------------------------------------------
'  Global Variables (Extern)
' -----------------------------------------------------------------------------
extern lpEditProcData as WNDPROCDATA ptr
extern bSmartMathActive as BOOL
extern g_bOldRichEdit as BOOL

extern rcOldMargin as RECT
extern nOldFirstLine as Integer
extern nOldCaretLine as Integer
extern nOldMargin as Integer
extern dwOldAkelOptions as DWORD

extern g_nDecimals as Integer
extern g_crResultColor as COLORREF
extern g_bUseThousandsSeparator as BOOL
extern hSmartMathMenu as HMENU
extern hSubMenuDecimals as HMENU
extern hSubMenuColor as HMENU
extern g_hMainWnd as HWND
extern g_hMainMenu as HMENU
extern g_hWndEdit as HWND
extern g_bShuttingDown as BOOL
extern g_pfnOldMainProc as WNDPROC
extern g_wszIniPath as WString * 260

' -----------------------------------------------------------------------------
'  Global Function Prototypes
' -----------------------------------------------------------------------------
declare sub LoadSettings()
declare sub SaveSettings()
declare function FormatResult(byval d as Double) as String
declare sub InitSmartMathMenu()
declare sub UpdateMenuChecks()
declare sub UninitSmartMathMenu(byval bAppClosing as BOOL = FALSE)
declare sub ShowAboutDialog(byval hWnd as HWND)
declare sub UpdateMarginAndState(byval hWnd as HWND, byref bVisible as BOOL)
declare sub LogInfo(byref sMsg as String)
declare function SmartMathMainProc stdcall(byval hWnd as HWND, byval uMsg as UINT, byval wParam as WPARAM, byval lParam as LPARAM) as LRESULT