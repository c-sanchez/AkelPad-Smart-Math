#ifndef __AKELEDIT_BI__
#define __AKELEDIT_BI__
#include once "windows.bi"

' AkelEdit Messages
const AEN_TEXTINSERTEND = (WM_USER + 1057)
const AEN_TEXTCHANGED   = (WM_USER + 1060)
const AEM_GETRECT       = (WM_USER + 2153)
const AEM_GETSCROLLPOS  = (WM_USER + 2155)
const AEM_GETCHARSIZE   = (WM_USER + 2164)
const AEM_GETOPTIONS    = (WM_USER + 2203)
const AEM_SETOPTIONS    = (WM_USER + 2204)
const AEM_GETCOLORS     = (WM_USER + 2207)
const AEM_SETCOLORS     = (WM_USER + 2208)
const AEM_GETMARKER     = (WM_USER + 2233)

' AEM_GETCHARSIZE Flags
const AECS_HEIGHT   = 0
const AECS_AVEWIDTH = 1

' AEM_GETMARKER / AEM_SETMARKER Types
const AEMT_PIXEL  = 0
const AEMT_SYMBOL = 1

' AEM_SETOPTIONS Flags
const AECOOP_SET = 1
const AECOOP_OR  = 2
const AECOOP_AND = 3
const AECOOP_XOR = 4

' AECO Options
const AECO_ACTIVELINE       = &h00000800
const AECO_ACTIVELINEBORDER = &h00001000

' AECOLORS dwFlags
const AECLR_BASICBK          = &h00000008
const AECLR_ACTIVELINETEXT   = &h00000040
const AECLR_ACTIVELINEBK     = &h00000080
const AECLR_COLUMNMARKER     = &h00000400
const AECLR_ACTIVELINEBORDER = &h00002000

' AkelEdit Colors Structure
type AECOLORS
  dwFlags as DWORD
  crCaret as COLORREF
  crBasicText as COLORREF
  crBasicBk as COLORREF
  crSelText as COLORREF
  crSelBk as COLORREF
  crActiveLineText as COLORREF
  crActiveLineBk as COLORREF
  crUrlText as COLORREF
  crActiveColumn as COLORREF
  crColumnMarker as COLORREF
  crUrlCursorText as COLORREF
  crUrlVisitText as COLORREF
  crActiveLineBorder as COLORREF
  crAltLineText as COLORREF
  crAltLineBk as COLORREF
  crAltLineBorder as COLORREF
end type

' x64-aware point in the virtual text space of a document
type POINT64
  x as INT_PTR
  y as INT_PTR
end type

type AENMHDR
  hwndFrom as HWND
  idFrom as UINT_PTR
  code as UINT
  docFrom as any ptr
end type

type AECHARINDEX
  nLine as Long ' corresponds to 32-bit `int` in C!
  lpLine as any ptr
  nCharInLine as Long ' corresponds to 32-bit `int` in C!
end type

type AECHARRANGE
  ciMin as AECHARINDEX
  ciMax as AECHARINDEX
end type

type CHARRANGE64
  cpMin as INT_PTR
  cpMax as INT_PTR
end type

type AENTEXTCHANGE
  hdr as AENMHDR
  crSel as AECHARRANGE
  ciCaret as AECHARINDEX
  dwType as DWORD
  bColumnSel as BOOL
  crRichSel as CHARRANGE64
end type

type AENTEXTINSERT
  hdr as AENMHDR
  crSel as AECHARRANGE
  ciCaret as AECHARINDEX
  dwType as DWORD
  wpText as WString ptr
  dwTextLen as UINT_PTR
  nNewLine as Long ' corresponds to 32-bit `int` in C!
  bColumnSel as BOOL
  dwInsertFlags as DWORD
  crAkelRange as AECHARRANGE
  crRichRange as CHARRANGE64
end type

#endif