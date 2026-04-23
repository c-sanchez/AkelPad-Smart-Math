#ifndef __AKELDLL_BI__
#define __AKELDLL_BI__

#include once "windows.bi"
#include once "AkelEdit.bi"

#ifndef MAKE_IDENTIFIER
  #define MAKE_IDENTIFIER(a, b, c, d) cast(DWORD, MAKELONG(MAKEWORD(a, b), MAKEWORD(c, d)))
#endif

const AKELDLL = MAKE_IDENTIFIER(2, 2, 0, 4)

#define PDS_NOAUTOLOAD &h00000001
#define PDS_GETSUPPORT &h10000000
#define UD_UNLOAD &h00000000
#define UD_NONUNLOAD_ACTIVE &h00000001

#ifndef WM_USER
  const WM_USER = &h0400
#endif

' RichEdit Messages
const EM_EXGETSEL64       = (WM_USER + 1952)
const EM_EXLINEFROMCHAR   = (WM_USER + 54)
const EM_LINEINDEX        = &h00BB
const EM_LINELENGTH       = &h00C1

#ifndef EM_POSFROMCHAR
  const EM_POSFROMCHAR    = (WM_USER + 38)
#endif

' Native text margin messages
#ifndef EM_SETMARGINS
  const EM_SETMARGINS     = &h00D3
#endif

#ifndef EC_RIGHTMARGIN
  const EC_RIGHTMARGIN    = 2
#endif

' AkelPad Constants
const AKD_FRAMEFIND             = (WM_USER + 50)
' AKD_SETMAINPROC: corrected to match AkelPad 4.9.x+ SDK (was WM_USER + 52)
const AKD_SETMAINPROC           = (WM_USER + 1217)
const AKD_SETEDITPROC           = (WM_USER + 106)
const AKD_FRAMEFINDW            = (WM_USER + 55)
const AKDN_FRAME_ACTIVATE       = (WM_USER + 256)
const AKDN_OPENDOCUMENT_FINISH  = (WM_USER + 263)

' Main-window notification sent when AkelPad is closing (added)
const AKDN_MAIN_ONFINISH        = (WM_USER + 6)

' Plugin Manager Constants
const AKD_DLLFINDW = (WM_USER + 14)
const AKD_DLLSAVE  = (WM_USER + 25)
const DLLSF_NOW    = 1
const DLLSF_ONEXIT = 2

#ifndef MAX_PATH
  const MAX_PATH = 260
#endif

' AkelPad menu bar position constants (added)
#define MENU_FILE_POSITION    0
#define MENU_EDIT_POSITION    1
#define MENU_VIEW_POSITION    2
#define MENU_OPTIONS_POSITION 3
#define MENU_MDI_POSITION     4
#define MENU_ABOUT_POSITION   5

type CHARRANGE64
  cpMin as Integer
  cpMax as Integer
end type

type PLUGINVERSION
  cb as DWORD
  hMainWnd as HWND
  dwAkelDllVersion as DWORD
  dwExeMinVersion3x as DWORD
  dwExeMinVersion4x as DWORD
  pPluginName as ZString ptr
end type

' Callback type used for all WNDPROCDATA proc fields
type AKEL_WNDPROC as function stdcall(byval as HWND, byval as UINT, byval as WPARAM, byval as LPARAM) as LRESULT

' Matches AkelPad SDK WNDPROCDATA exactly (5 fields)
type WNDPROCDATA
  pNext as WNDPROCDATA ptr
  pPrev as WNDPROCDATA ptr
  CurProc as AKEL_WNDPROC
  NextProc as AKEL_WNDPROC
  PrevProc as AKEL_WNDPROC
end type

type PLUGINDATA
  cb as DWORD
  pcs as any ptr
  dwSupport as DWORD
  pFunction as UBYTE ptr
  szFunction as ZString ptr
  wszFunction as WString ptr
  lParam as LPARAM
  hInstanceDLL as HINSTANCE
  lpPluginFunction as any ptr
  nUnload as Integer
  bInMemory as WINBOOL
  bOnStart as WINBOOL
  pAkelDir as UBYTE ptr
  szAkelDir as ZString ptr
  wszAkelDir as WString ptr
  hInstanceEXE as HINSTANCE
  hPluginsStack as any ptr
  nSaveSettings as Integer
  hMainWnd as HWND
  lpFrameData as any ptr
  hWndEdit as HWND
  hDocEdit as any ptr
  hStatus as HWND
  hMdiClient as HWND
  hTab as HWND
  hMainMenu as HMENU
  hMenuRecentFiles as HMENU
  hMenuLanguage as HMENU
  hPopupMenu as HMENU
  hMainIcon as HICON
  hGlobalAccel as HACCEL
  hMainAccel as HACCEL
  bOldWindows as WINBOOL
  bOldRichEdit as WINBOOL
  dwVerComctl32 as DWORD
  bAkelEdit as WINBOOL
  nMDI as Integer
  pLangModule as UBYTE ptr
  szLangModule as ZString ptr
  wszLangModule as WString ptr
  hLangModule as HMODULE
  wLangSystem as LANGID
  wLangModule as LANGID
end type

type EDITINFO
  hWndEdit as HWND
  hDocEdit as HWND
end type

type FRAMEDATA
  pNextFrame as FRAMEDATA ptr
  pPrevFrame as FRAMEDATA ptr
  cb as DWORD
  nFrameID as UInteger
  hWndEditParent as HWND
  ei as EDITINFO
end type

' Plugin Stack Structure
type PLUGINFUNCTION
  pNext as PLUGINFUNCTION ptr
  pPrev as PLUGINFUNCTION ptr
  pFunction as UBYTE ptr
  szFunction as ZString * MAX_PATH
  wszFunction as WString * MAX_PATH
  nFunctionLen as Integer
  wHotkey as WORD
  bAutoLoad as WINBOOL
  bRunning as WINBOOL
  PluginProc as any ptr
  lpParameter as any ptr
  nRefCount as Integer
end type

#endif