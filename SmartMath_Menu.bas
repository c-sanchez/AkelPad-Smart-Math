#include once "SmartMath_Globals.bi"

' -----------------------------------------------------------------------------
'  Update checkmarks on submenus to reflect loaded settings
' -----------------------------------------------------------------------------
sub UpdateMenuChecks()
  if hSubMenuDecimals <> 0 then
    dim stAuto as UINT
    if g_nDecimals = -1 then stAuto = MF_CHECKED else stAuto = MF_UNCHECKED
    CheckMenuItem(hSubMenuDecimals, IDM_DECIMAL_BASE, MF_BYCOMMAND or stAuto)

    dim i as Integer
    for i = 0 to SMARTMATH_DECIMALS_MAX
      dim st as UINT
      if g_nDecimals = i then st = MF_CHECKED else st = MF_UNCHECKED
      CheckMenuItem(hSubMenuDecimals, IDM_DECIMAL_BASE + 1 + i, MF_BYCOMMAND or st)
    next i
  end if

  if hSubMenuColor <> 0 then
    dim idCheck as Integer = -1
    select case g_crResultColor
      case &H008000: idCheck = IDM_COLOR_GREEN
      case &HFF0000: idCheck = IDM_COLOR_BLUE
      case &H0000FF: idCheck = IDM_COLOR_RED
      case &H00FFFF: idCheck = IDM_COLOR_YELLOW
      case &HFFFFFF: idCheck = IDM_COLOR_WHITE
      case &H000000: idCheck = IDM_COLOR_BLACK
    end select

    dim j as Integer
    for j = 0 to MAX_COLORS - 1
      dim st as UINT
      if (IDM_COLOR_BASE + j) = idCheck then st = MF_CHECKED else st = MF_UNCHECKED
      CheckMenuItem(hSubMenuColor, IDM_COLOR_BASE + j, MF_BYCOMMAND or st)
    next j
  end if

  if hSmartMathMenu <> 0 then
    dim stThou as UINT
    if g_bUseThousandsSeparator then stThou = MF_CHECKED else stThou = MF_UNCHECKED
    CheckMenuItem(hSmartMathMenu, IDM_THOUSANDS_SEPARATOR, MF_BYCOMMAND or stThou)

    dim stComplex as UINT
    if g_bSupportComplexNumbers then stComplex = MF_CHECKED else stComplex = MF_UNCHECKED
    Parser_SetSupportComplexNumbers(g_bSupportComplexNumbers)
    CheckMenuItem(hSmartMathMenu, IDM_COMPLEX_NUMBERS, MF_BYCOMMAND or stComplex)
  end if
end sub

sub UpdateMenuActiveOnCurrTab(byval bActive as BOOL)
  if hSmartMathMenu = 0 then exit sub

  dim stActive as UINT
  if bActive then stActive = MF_CHECKED else stActive = MF_UNCHECKED
  CheckMenuItem(hSmartMathMenu, IDM_ACTIVE_ON_CURR_TAB, MF_BYCOMMAND or stActive)
end sub

' -----------------------------------------------------------------------------
'  Menu setup
' -----------------------------------------------------------------------------
sub InitSmartMathMenu()
  if hSmartMathMenu <> 0 then exit sub

  hSubMenuDecimals = CreatePopupMenu()
  if hSubMenuDecimals <> 0 then
    AppendMenuW(hSubMenuDecimals, MF_STRING, IDM_DECIMAL_BASE, wstr("Auto (full precision)"))
    AppendMenuW(hSubMenuDecimals, MF_SEPARATOR, 0, NULL)

    dim i as Integer
    dim sLabel as WString * 64
    for i = 0 to SMARTMATH_DECIMALS_MAX
      if i = 1 then
        sLabel = wstr("1 decimal place")
      else
        sLabel = WStr(i) & wstr(" decimal places")
      end if
      AppendMenuW(hSubMenuDecimals, MF_STRING, IDM_DECIMAL_BASE + 1 + i, @sLabel)
    next i
  end if

  hSubMenuColor = CreatePopupMenu()
  if hSubMenuColor <> 0 then
    AppendMenuW(hSubMenuColor, MF_STRING, IDM_COLOR_GREEN, wstr("Green"))
    AppendMenuW(hSubMenuColor, MF_STRING, IDM_COLOR_BLUE, wstr("Blue"))
    AppendMenuW(hSubMenuColor, MF_STRING, IDM_COLOR_RED, wstr("Red"))
    AppendMenuW(hSubMenuColor, MF_STRING, IDM_COLOR_YELLOW, wstr("Yellow"))
    AppendMenuW(hSubMenuColor, MF_STRING, IDM_COLOR_WHITE, wstr("White"))
    AppendMenuW(hSubMenuColor, MF_STRING, IDM_COLOR_BLACK, wstr("Black"))
  end if

  hSmartMathMenu = CreatePopupMenu()
  if hSmartMathMenu = 0 then
    if hSubMenuDecimals then DestroyMenu(hSubMenuDecimals)
    hSubMenuDecimals = 0
    if hSubMenuColor then DestroyMenu(hSubMenuColor)
    hSubMenuColor = 0
    exit sub
  end if

  AppendMenuW(hSmartMathMenu, MF_STRING, IDM_ACTIVE_ON_CURR_TAB, wstr("Active on Current Tab"))
  AppendMenuW(hSmartMathMenu, MF_SEPARATOR, 0, NULL)

  if hSubMenuDecimals <> 0 then
    AppendMenuW(hSmartMathMenu, MF_POPUP, cast(UINT_PTR, hSubMenuDecimals), wstr("Decimal Places"))
  end if

  if hSubMenuColor <> 0 then
    AppendMenuW(hSmartMathMenu, MF_POPUP, cast(UINT_PTR, hSubMenuColor), wstr("Text Color"))
  end if

  AppendMenuW(hSmartMathMenu, MF_SEPARATOR, 0, NULL)
  AppendMenuW(hSmartMathMenu, MF_STRING, IDM_THOUSANDS_SEPARATOR, wstr("Use Thousands Separator"))
  AppendMenuW(hSmartMathMenu, MF_STRING, IDM_COMPLEX_NUMBERS, wstr("Complex Numbers"))
  AppendMenuW(hSmartMathMenu, MF_SEPARATOR, 0, NULL)
  AppendMenuW(hSmartMathMenu, MF_STRING, IDM_ABOUT, wstr("About..."))

  dim nPos as Integer = GetMenuItemCount(g_hMainMenu)
  InsertMenuW(g_hMainMenu, nPos, MF_BYPOSITION or MF_POPUP, cast(UINT_PTR, hSmartMathMenu), wstr("SmartMath"))
  DrawMenuBar(g_hMainWnd)
  UpdateMenuChecks()
end sub

' -----------------------------------------------------------------------------
'  Menu teardown
' -----------------------------------------------------------------------------
sub UninitSmartMathMenu(byval bAppClosing as BOOL = FALSE)
  if hSmartMathMenu = 0 then exit sub

  if (g_hMainMenu <> 0) andalso IsMenu(g_hMainMenu) then
    dim nIndex as Integer = -1
    dim j as Integer
    for j = 0 to GetMenuItemCount(g_hMainMenu) - 1
      if GetSubMenu(g_hMainMenu, j) = hSmartMathMenu then
        nIndex = j
        exit for
      end if
    next j
    if nIndex <> -1 then
      RemoveMenu(g_hMainMenu, nIndex, MF_BYPOSITION)
      if not bAppClosing then
        if (g_hMainWnd <> 0) andalso IsWindow(g_hMainWnd) then
          DrawMenuBar(g_hMainWnd)
        end if
      end if
    end if
  end if

  if IsMenu(hSmartMathMenu) then
    DestroyMenu(hSmartMathMenu)
  end if
  hSmartMathMenu = 0
  hSubMenuDecimals = 0
  hSubMenuColor = 0
end sub