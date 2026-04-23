#ifndef __AKELEDIT_BI__
#define __AKELEDIT_BI__
#include once "windows.bi"

' AkelEdit Messages
const AEM_GETCHARSIZE = (WM_USER + 2164)
const AEM_GETOPTIONS  = (WM_USER + 2203)
const AEM_SETOPTIONS  = (WM_USER + 2204)
const AEM_GETCOLORS   = (WM_USER + 2207)
const AEM_SETCOLORS   = (WM_USER + 2208)

' AEM_SETOPTIONS Flags
const AECOOP_SET = 1
const AECOOP_OR  = 2
const AECOOP_XOR = 4
const AECOOP_AND = 8

' AECO Options
const AECO_ACTIVELINE       = &h00000800
const AECO_ACTIVELINEBORDER = &h00001000

' AECOLORS dwFlags
const AECLR_ACTIVELINETEXT   = &h00000040
const AECLR_ACTIVELINEBK     = &h00000080
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

#endif