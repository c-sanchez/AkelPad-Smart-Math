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

#define WMD_SDI   0  ' Single document interface (SDI)
#define WMD_MDI   1  ' Multiple document interface (MDI)
#define WMD_PMDI  2  ' Pseudo-Multiple document interface (PMDI)

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

' AkelPad Control IDs
const ID_EDIT             = 10001

' AkelPad Notifications
const AKDN_MAIN_ONSTART_SHOW      = (WM_USER + 3)
const AKDN_MAIN_ONSTART_FINISH    = (WM_USER + 4)
const AKDN_MAIN_ONFINISH          = (WM_USER + 6)
const AKDN_MAIN_ONCLOSE_PREFINISH = (WM_USER + 8) 
const AKDN_FRAME_ACTIVATE         = (WM_USER + 22)
const AKDN_FRAME_DESTROY          = (WM_USER + 24)
const AKDN_OPENDOCUMENT_FINISH    = (WM_USER + 54)

' AkelPad Messages
const AKD_SETMAINPROC           = (WM_USER + 102)
const AKD_SETEDITPROC           = (WM_USER + 106)
const AKD_SETFRAMEPROC          = (WM_USER + 110)
const AKD_GETFRAMEINFO          = (WM_USER + 199)
const AKD_GETEDITINFO           = (WM_USER + 200)
const AKD_FRAMEFIND             = (WM_USER + 264)
const AKD_FRAMEFINDW            = (WM_USER + 266)
const AKD_DLLCALLW              = (WM_USER + 303)
const AKD_DLLFINDW              = (WM_USER + 307)
const AKD_DLLSAVE               = (WM_USER + 312)
const AKD_BEGINOPTIONSA         = (WM_USER + 332)
const AKD_BEGINOPTIONSW         = (WM_USER + 333)
const AKD_OPTIONA               = (WM_USER + 335)
const AKD_OPTIONW               = (WM_USER + 336)
const AKD_ENDOPTIONS            = (WM_USER + 341)

const DLLSF_NOW    = &h01
const DLLSF_ONEXIT = &h02
const POB_READ     = &h01
const POB_SAVE     = &h02
const PO_DWORD     = 1
const PO_BINARY    = 2
const PO_STRING    = 3
const FWF_CURRENT    = 1
const FWF_NEXT       = 2
const FWF_PREV       = 3
const FWF_BYFILENAME = 5
const FI_WNDEDIT     = 2

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

type PLUGINVERSION
  cb as DWORD
  hMainWnd as HWND
  dwAkelDllVersion as DWORD
  dwExeMinVersion3x as DWORD
  dwExeMinVersion4x as DWORD
  pPluginName as ZString ptr
end type

type PLUGINOPTIONW
  pOptionName as WString ptr
  dwType as DWORD
  lpData as UByte ptr
  dwData as DWORD
end type

type PLUGINCALLSENDW
  pFunction as WString ptr
  lParam as LPARAM
  dwSupport as DWORD
  nResult as LPARAM
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
  nUnload as Long ' corresponds to 32-bit `int` in C!
  bInMemory as WINBOOL
  bOnStart as WINBOOL
  pAkelDir as UBYTE ptr
  szAkelDir as ZString ptr
  wszAkelDir as WString ptr
  hInstanceEXE as HINSTANCE
  hPluginsStack as any ptr
  nSaveSettings as Long ' corresponds to 32-bit `int` in C!
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
  nMDI as Long ' corresponds to 32-bit `int` in C!
  pLangModule as UBYTE ptr
  szLangModule as ZString ptr
  wszLangModule as WString ptr
  hLangModule as HMODULE
  wLangSystem as LANGID
  wLangModule as LANGID
  nSaveHistory as Long ' corresponds to 32-bit `int` in C!
end type

type EDITINFO
  hWndEdit as HWND
  hDocEdit as any ptr
  pFile as UBYTE ptr
  szFile as ZString ptr
  wszFile as WString ptr
  nCodePage as Long ' corresponds to 32-bit `int` in C!
  bBOM as WINBOOL
  nNewLine as Long ' corresponds to 32-bit `int` in C!
  bModified as WINBOOL
  bReadOnly as WINBOOL
  bWordWrap as WINBOOL
  bOvertypeMode as WINBOOL
  hWndMaster as HWND
  hDocMaster as any ptr
  hWndClone1 as HWND
  hDocClone1 as any ptr
  hWndClone2 as HWND
  hDocClone2 as any ptr
  hWndClone3 as HWND
  hDocClone3 as any ptr
end type

type FRAMEDATA
  pNextFrame as FRAMEDATA ptr
  pPrevFrame as FRAMEDATA ptr
  cb as DWORD
  nFrameID as INT_PTR
  hWndEditParent as HWND
  ei as EDITINFO
  ' note: this structure is incomplete, be sure to complete it if needed
end type

' Plugin Stack Structure
type PLUGINFUNCTION
  pNext as PLUGINFUNCTION ptr
  pPrev as PLUGINFUNCTION ptr
  pFunction as UBYTE ptr
  szFunction as ZString * MAX_PATH
  wszFunction as WString * MAX_PATH
  nFunctionLen as Long ' corresponds to 32-bit `int` in C!
  wHotkey as WORD
  bAutoLoad as WINBOOL
  bRunning as WINBOOL
  PluginProc as any ptr
  lpParameter as any ptr
  nRefCount as Long ' corresponds to 32-bit `int` in C!
end type

#endif