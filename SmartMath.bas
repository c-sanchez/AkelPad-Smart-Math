#include once "SmartMath_Globals.bi"
#include once "Inc\PtrArray.bi"

' --- Shared global variables definitions ---
dim shared lpEditProcData as WNDPROCDATA ptr = 0
dim shared lpMainProcData as WNDPROCDATA ptr = 0
dim shared lpFrameProcData as WNDPROCDATA ptr = 0
dim shared g_bSmartMathActive as BOOL = FALSE
dim shared g_bAkelPadReady as BOOL = FALSE
dim shared g_bOldRichEdit as BOOL = FALSE

dim shared rcOldMargin as RECT
dim shared nOldFirstLine as Integer = -1
dim shared nOldCaretLine as Integer = -1
dim shared nLastNavSelStart as Integer = -1
dim shared nLastNavSelEnd as Integer = -1
dim shared dwOldAkelOptions as DWORD = 0

dim shared g_nDecimals as Integer = -1
dim shared g_crResultColor as COLORREF = &H008000
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
dim shared g_bActiveFramesSaved as BOOL = FALSE
dim shared g_cacheLineCount as Integer = -1
dim shared g_cacheReady as BOOL = FALSE
dim shared g_cacheDocGeneration as UInteger = 0
dim shared g_cacheSyncedGeneration as UInteger = 0
dim shared g_cacheDirtyFromLine as Integer = -1
redim shared g_cachedLineText(0 to 0) as String
redim shared g_cachedRenderText(0 to 0) as String

const ACTIVE_FRAME_FILE_SEPARATOR as UShort = 10 '' LF
type WStringPtr as WString ptr

type FrameItem
  declare constructor(byval pFrameInit as FRAMEDATA ptr)
  declare destructor()

  pFrame as FRAMEDATA ptr
  wszFile as WString ptr
  isWorking as BOOL
end type

constructor FrameItem(byval pFrameInit as FRAMEDATA ptr)
  pFrame = pFrameInit
  wszFile = 0
  isWorking = FALSE

  if pFrameInit <> 0 andalso pFrameInit->ei.wszFile <> 0 then
    dim pSrc as WString ptr = pFrameInit->ei.wszFile
    dim nFileChars as Integer = Len(*pSrc)
    wszFile = cast(WString ptr, Allocate((nFileChars + 1) * SizeOf(WString)))
    if wszFile <> 0 then
      *wszFile = *pSrc
    end if
  end if
end constructor

destructor FrameItem()
  if wszFile <> 0 then
    Deallocate(wszFile)
    wszFile = 0
  end if
  pFrame = 0
end destructor

private sub FrameItemDeleter(byval p as Any Ptr)
  dim pFrameItem as FrameItem ptr = cast(FrameItem ptr, p)
  if pFrameItem <> 0 then
    Delete pFrameItem
  end if
end sub

private function FrameItemMatchesFrame(byval p as Any Ptr, byval key as Any Ptr) as Boolean
  dim pFrameItem as FrameItem ptr = cast(FrameItem ptr, p)
  if pFrameItem = 0 then return FALSE
  return pFrameItem->pFrame = cast(FRAMEDATA ptr, key)
end function

dim shared g_framesWithSmartMathEnabled as PtrArray = PtrArray(@FrameItemDeleter, @FrameItemMatchesFrame)

' -----------------------------------------------------------------------------
'  Debug logging
' -----------------------------------------------------------------------------
sub LogInfo(byref sMsg as String)
  dim sOut as String = "[SmartMath] " & sMsg
  OutputDebugString(strptr(sOut))
end sub

private sub MarkRenderCacheDocumentDirty(byval dirtyFromLine as Integer = -1)
  g_cacheDocGeneration += 1u
  if dirtyFromLine >= 0 then
    if g_cacheDirtyFromLine < 0 orelse dirtyFromLine < g_cacheDirtyFromLine then
      g_cacheDirtyFromLine = dirtyFromLine
    end if
  else
    g_cacheDirtyFromLine = 0
  end if
end sub

private function IsRenderCacheSyncedWithDocument(byval hWnd as HWND) as BOOL
  if g_cacheReady = FALSE then return FALSE
  if g_cacheSyncedGeneration <> g_cacheDocGeneration then return FALSE
  if SendMessage(hWnd, EM_GETLINECOUNT, 0, 0) <> g_cacheLineCount then return FALSE
  return TRUE
end function

private sub InvalidateRenderCache()
  MarkRenderCacheDocumentDirty(0)
  g_cacheReady = FALSE
  g_cacheLineCount = -1
end sub

private function IsVolatilePureExpressionLine(byref sLine as String) as BOOL
  dim t as String = lcase(trim(sLine))
  if len(t) = 0 then return FALSE
  if instr(t, "=") > 0 then return FALSE
  if instr(t, ";") > 0 then return FALSE
  if instr(t, "rand(") > 0 then return TRUE
  if instr(t, "random(") > 0 then return TRUE
  return FALSE
end function

private function IsLikelyCodeLine(byref sLine as String) as BOOL
  dim t as String = lcase(trim(sLine))
  if len(t) = 0 then return FALSE
  if left(t, 1) = "'" orelse left(t, 1) = "#" orelse left(t, 2) = "//" then return TRUE
  if left(t, 8) = "#include" then return TRUE
  if left(t, 4) = "dim " orelse left(t, 6) = "const " orelse left(t, 5) = "type " _
     orelse left(t, 5) = "enum " orelse left(t, 4) = "sub " orelse left(t, 9) = "function " _
     orelse left(t, 8) = "private " orelse left(t, 7) = "public " orelse left(t, 8) = "declare " _
     orelse left(t, 7) = "extern " orelse left(t, 4) = "end " orelse left(t, 7) = "return " then
    return TRUE
  end if

  ' Defensive skip for source-like call chains that are not typical math input.
  if instr(t, "_") > 0 andalso instr(t, "(") > 0 andalso instr(t, ")") > 0 _
     andalso instr(t, "=") = 0 andalso instr(t, ";") = 0 andalso instr(t, ",") = 0 then
    dim hasDigit as BOOL = FALSE
    for i as Integer = 1 to len(t)
      dim ch as String = mid(t, i, 1)
      if ch >= "0" andalso ch <= "9" then
        hasDigit = TRUE
        exit for
      end if
    next i
    if hasDigit = FALSE then return TRUE
  end if
  return FALSE
end function

private sub RestoreAnsFromCachedRender(byref sRender as String)
  if left(sRender, Len(SMARTMATH_RESULT_PREFIX)) <> SMARTMATH_RESULT_PREFIX then exit sub
  dim sExpr as String = mid(sRender, Len(SMARTMATH_RESULT_PREFIX) + 1)
  if len(trim(sExpr)) = 0 then exit sub

  dim dResult as Double
  dim sResult as String
  dim bIsArray as Boolean
  Parser_TryEvaluateEx(sExpr, dResult, sResult, bIsArray)
end sub

dim shared g_bSmartMathDocActive as BOOL = FALSE
dim shared g_lastDocModeFrame as FRAMEDATA ptr = 0
dim shared g_lastDocModeEdit as HWND = 0
dim shared SMARTMATH_CODER_HIGHLIGHT_FUNC as WString * 32 = WStr("Coder::HighLight")
dim shared SMARTMATH_CODER_SETTINGS_FUNC as WString * 32 = WStr("Coder::Settings")
dim shared SMARTMATH_CODER_ALIAS as WString * 32 = WStr(".smartmath")

const DLLA_CODER_SETALIAS = 6

type DLLECCODERSETTINGS_SETALIAS
  dwStructSize as UINT_PTR
  nAction as INT_PTR
  pszAlias as const UByte ptr
end type

#if sizeof(DLLECCODERSETTINGS_SETALIAS) <> (sizeof(UINT_PTR) + sizeof(INT_PTR) + sizeof(any ptr))
  #error "DLLECCODERSETTINGS_SETALIAS layout mismatch"
#endif

private sub CallPluginFuncW(byval pFunctionName as WString ptr, byval params as any ptr)
  dim pcsW as PLUGINCALLSENDW
  pcsW.pFunction = pFunctionName
  pcsW.lParam = cast(LPARAM, params)
  pcsW.dwSupport = 0
  pcsW.nResult = 0
  SendMessageW(g_hMainWnd, AKD_DLLCALLW, 0, cast(LPARAM, @pcsW))
end sub

private sub SetCoderAliasW(byval pAliasName as WString ptr)
  dim stParams as DLLECCODERSETTINGS_SETALIAS
  stParams.dwStructSize = sizeof(DLLECCODERSETTINGS_SETALIAS)
  stParams.nAction = DLLA_CODER_SETALIAS
  stParams.pszAlias = cast(UByte ptr, pAliasName)
  CallPluginFuncW(strptr(SMARTMATH_CODER_SETTINGS_FUNC), @stParams)
end sub

private sub TryApplySmartMathCoderTheme()
  ' "Coder::HighLight" function must be running to apply the SmartMath coder theme.
  dim pf as PLUGINFUNCTION ptr = cast(PLUGINFUNCTION ptr, SendMessageW(g_hMainWnd, AKD_DLLFINDW, cast(WPARAM, strptr(SMARTMATH_CODER_HIGHLIGHT_FUNC)), 0))
  ' LogInfo("TryApplySmartMathCoderTheme: pf=" & pf & ", bRunning=" & pf->bRunning)
  if pf = 0 orelse pf->bRunning = FALSE then exit sub

  ' LogInfo("TryApplySmartMathCoderTheme: calling SetCoderAliasW")

  ' Set the SmartMath coder alias.
  SetCoderAliasW(strptr(SMARTMATH_CODER_ALIAS))
end sub

declare function GetLineText(byval hWnd as HWND, byval lineIdx as Integer, byval lineLen as Integer) as String

private sub SetSmartMathDocActiveState(byval hWndEdit as HWND, byval bActive as BOOL, byval bApplyTheme as BOOL = TRUE)
  if (hWndEdit = 0) orelse (IsWindow(hWndEdit) = FALSE) then
    g_hWndEdit = 0
    g_bSmartMathDocActive = FALSE
    exit sub
  end if

  g_hWndEdit = hWndEdit
  g_bSmartMathDocActive = bActive
  InvalidateRenderCache()

  if bActive then
    if bApplyTheme then
      TryApplySmartMathCoderTheme()
    end if
    dim bVisible as BOOL
    UpdateInternalState(hWndEdit, bVisible)
    InvalidateRect(hWndEdit, 0, TRUE)
  else
    rcOldMargin.left = 0 : rcOldMargin.right = 0
    rcOldMargin.top = 0  : rcOldMargin.bottom = 0
    nOldFirstLine = -1
    nOldCaretLine = -1
    nLastNavSelStart = -1
    nLastNavSelEnd = -1
    InvalidateRect(hWndEdit, 0, TRUE)
  end if
end sub

private function GetWndEdit(byval hMainWnd as HWND) as HWND
  dim ei as EDITINFO
  ei.hWndEdit = 0
  SendMessage(hMainWnd, AKD_GETEDITINFO, 0, cast(LPARAM, @ei))
  return ei.hWndEdit
end function

private function IsSpaceOrTabAfterHash(byval ch as String) as BOOL
  if Len(ch) <> 1 then return FALSE
  select case Asc(ch)
    case 9, 32  '' HT, space
      return TRUE
    case else
      return FALSE
  end select
end function

private sub RefreshSmartMathDocMode(byval hWndEdit as HWND)
  ' LogInfo("RefreshSmartMathDocMode: entering")

  dim pFrameCurrent as FRAMEDATA ptr = cast(FRAMEDATA ptr, SendMessage(g_hMainWnd, AKD_FRAMEFIND, FWF_CURRENT, 0))
  dim hWndEditCurrent as HWND = hWndEdit

  if (pFrameCurrent = g_lastDocModeFrame) andalso (hWndEditCurrent = g_lastDocModeEdit) then
    exit sub
  end if

  g_lastDocModeFrame = pFrameCurrent
  g_lastDocModeEdit = hWndEditCurrent

  if hWndEditCurrent = 0 then
    g_bSmartMathDocActive = FALSE
    g_hWndEdit = 0
    exit sub
  end if
  if IsWindow(hWndEditCurrent) = FALSE then
    g_bSmartMathDocActive = FALSE
    g_hWndEdit = 0
    exit sub
  end if

  dim idx as Integer = g_framesWithSmartMathEnabled.Find(pFrameCurrent)
  if idx >= 0 then
    dim pFrameItem as FrameItem ptr = cast(FrameItem ptr, g_framesWithSmartMathEnabled[idx])
    if pFrameItem <> 0 then
      pFrameItem->isWorking = TRUE
    end if
  end if
  SetSmartMathDocActiveState(hWndEditCurrent, idx >= 0, TRUE)
  UpdateMenuActiveOnCurrTab(idx >= 0)
end sub

private sub SaveActiveSmartMathFrames()
  dim nTotalChars as Integer = 0
  dim nFiles as Integer = 0

  for i as Integer = 0 to g_framesWithSmartMathEnabled.Count() - 1
    dim pFrameItem as FrameItem ptr = cast(FrameItem ptr, g_framesWithSmartMathEnabled[i])
    if pFrameItem <> 0 andalso pFrameItem->wszFile <> 0 then
      dim nFileChars as Integer = Len(*(pFrameItem->wszFile))
      if nFileChars > 0 then
        if nFiles > 0 then nTotalChars += 1
        nTotalChars += nFileChars
        nFiles += 1
      end if
    end if
  next i

  ' LogInfo("SaveActiveSmartMathFrames: nTotalChars=" & nTotalChars & ", nFiles=" & nFiles)

  redim wszActiveFrameFiles(0 to nTotalChars) as UShort
  dim nPos as Integer = 0
  dim nWrittenFiles as Integer = 0

  for i as Integer = 0 to g_framesWithSmartMathEnabled.Count() - 1
    dim pFrameItem as FrameItem ptr = cast(FrameItem ptr, g_framesWithSmartMathEnabled[i])
    if pFrameItem <> 0 andalso pFrameItem->wszFile <> 0 then
      dim nFileChars as Integer = Len(*(pFrameItem->wszFile))
      if nFileChars > 0 then
        if nWrittenFiles > 0 then
          wszActiveFrameFiles(nPos) = ACTIVE_FRAME_FILE_SEPARATOR
          nPos += 1
        end if

        dim pSrc as UShort ptr = cast(UShort ptr, pFrameItem->wszFile)
        for j as Integer = 0 to nFileChars - 1
          wszActiveFrameFiles(nPos) = pSrc[j]
          nPos += 1
        next j

        nWrittenFiles += 1
      end if
    end if
  next i
  wszActiveFrameFiles(nPos) = 0

  SaveSettings_ActiveFrames(wszActiveFrameFiles(), nPos)
end sub

private sub FreeActiveSmartMathFrameFiles(frameFiles() as WStringPtr, byval nFiles as Integer)
  for i as Integer = 0 to nFiles - 1
    if frameFiles(i) <> 0 then
      Deallocate(frameFiles(i))
      frameFiles(i) = 0
    end if
  next i

  erase frameFiles
end sub

private function LoadActiveSmartMathFrameFiles(frameFiles() as WStringPtr) as Integer
  erase frameFiles

  dim wszActiveFrameFiles() as UShort
  dim nCharsRead as Integer = LoadSettings_ActiveFrames(wszActiveFrameFiles())
  if nCharsRead <= 0 then return 0

  dim nFiles as Integer = 0
  dim nSegmentChars as Integer = 0

  for i as Integer = 0 to nCharsRead - 1
    if wszActiveFrameFiles(i) = ACTIVE_FRAME_FILE_SEPARATOR then
      if nSegmentChars > 0 then nFiles += 1
      nSegmentChars = 0
    else
      nSegmentChars += 1
    end if
  next i
  if nSegmentChars > 0 then nFiles += 1

  if nFiles = 0 then return 0

  redim frameFiles(0 to nFiles - 1)

  dim nOut as Integer = 0
  dim nSegmentStart as Integer = 0
  nSegmentChars = 0

  for i as Integer = 0 to nCharsRead
    if (i = nCharsRead) orelse (wszActiveFrameFiles(i) = ACTIVE_FRAME_FILE_SEPARATOR) then
      if nSegmentChars > 0 then
        dim pFile as WString ptr = cast(WString ptr, Allocate((nSegmentChars + 1) * SizeOf(WString)))
        if pFile <> 0 then
          dim pDst as UShort ptr = cast(UShort ptr, pFile)
          for j as Integer = 0 to nSegmentChars - 1
            pDst[j] = wszActiveFrameFiles(nSegmentStart + j)
          next j
          pDst[nSegmentChars] = 0
        end if

        frameFiles(nOut) = pFile
        nOut += 1
      end if

      nSegmentStart = i + 1
      nSegmentChars = 0
    else
      nSegmentChars += 1
    end if
  next i

  return nOut
end function

private sub ActivateSmartMathFiles()
  dim frameFiles() as WStringPtr
  dim nFiles as Integer = LoadActiveSmartMathFrameFiles(frameFiles())
  if nFiles > 0 then
    dim pStartFrame as FRAMEDATA ptr = cast(FRAMEDATA ptr, SendMessage(g_hMainWnd, AKD_FRAMEFINDW, FWF_CURRENT, 0))
    dim pFrame as FRAMEDATA ptr = pStartFrame
    while pFrame <> 0
      for i as Integer = 0 to nFiles - 1
        if frameFiles(i) <> 0 andalso pFrame->ei.wszFile <> 0 then
          if *(frameFiles(i)) = *(pFrame->ei.wszFile) andalso g_framesWithSmartMathEnabled.Find(pFrame) < 0 then
            dim pFrameItem as FrameItem ptr = New FrameItem(pFrame)
            if pFrameItem <> 0 then
              g_framesWithSmartMathEnabled.Append(pFrameItem)
            end if
          end if
        end if
      next i

      pFrame = cast(FRAMEDATA ptr, SendMessage(g_hMainWnd, AKD_FRAMEFINDW, FWF_PREV, cast(LPARAM, pFrame)))
      if pFrame = pStartFrame then exit while
    wend

    dim hWndEditCurrent as HWND = GetWndEdit(g_hMainWnd)
    if hWndEditCurrent <> 0 then
      g_lastDocModeFrame = 0
      RefreshSmartMathDocMode(hWndEditCurrent)
    end if
  end if

  FreeActiveSmartMathFrameFiles(frameFiles(), nFiles)
end sub

private sub CheckEditNotifications(byval hWnd as HWND, byval uMsg as UINT, byval wParam as WPARAM, byval lParam as LPARAM)
  if (uMsg <> WM_NOTIFY) orelse (wParam <> ID_EDIT) orelse (lParam = 0) orelse (g_bShuttingDown) then exit sub

  dim pHdr as AENMHDR ptr = cast(AENMHDR ptr, lParam)

  if pHdr->code = AEN_TEXTCHANGED then
    dim pChange as AENTEXTCHANGE ptr = cast(AENTEXTCHANGE ptr, lParam)
    dim nChangedLine as Integer = pChange->ciCaret.nLine
    if nChangedLine >= 0 then
      nOldCaretLine = nChangedLine
    end if
    dim nDirtyFromLine as Integer = nChangedLine
    if pChange->crSel.ciMin.nLine >= 0 then
      if nDirtyFromLine < 0 orelse pChange->crSel.ciMin.nLine < nDirtyFromLine then
        nDirtyFromLine = pChange->crSel.ciMin.nLine
      end if
    end if
    MarkRenderCacheDocumentDirty(nDirtyFromLine)
    ' Dependent lines can change results (e.g. rand / UDF) without AkelEdit
    ' invalidating them. Erase+repaint the full client so overlays do not stack.
    dim hWndEdit as HWND = pHdr->hwndFrom
    if (hWndEdit = 0) orelse (IsWindow(hWndEdit) = FALSE) then hWndEdit = g_hWndEdit
    if (hWndEdit <> 0) andalso (IsWindow(hWndEdit) <> 0) andalso g_bSmartMathDocActive then
      InvalidateRect(hWndEdit, 0, TRUE)
    end if
    exit sub
  end if

  'if pHdr->code = AEN_TEXTINSERTEND then
  '  dim pIns as AENTEXTINSERT ptr = cast(AENTEXTINSERT ptr, lParam)
  '  ' LogInfo("AEN_TEXTINSERTEND: ciMin.nLine=" & pIns->crAkelRange.ciMin.nLine)
  '  exit sub
  'end if
end sub

' After hooks are live: sync doc mode, margins, AkelEdit options.
private sub SyncSmartMathActiveEditorState(byval hWndEdit as HWND)
  ' LogInfo("SyncSmartMathActiveEditorState: entering, hWndEdit=" & hWndEdit)
  if (hWndEdit = 0) orelse (IsWindow(hWndEdit) = FALSE) then exit sub
  g_hWndEdit = hWndEdit
  RefreshSmartMathDocMode(hWndEdit)
  if not g_bOldRichEdit then
    dwOldAkelOptions = SendMessage(hWndEdit, AEM_GETOPTIONS, 0, 0)
    SendMessage(hWndEdit, AEM_SETOPTIONS, AECOOP_OR, AECO_ACTIVELINE)
  end if
  dim bInitialVis as BOOL
  UpdateInternalState(hWndEdit, bInitialVis)
  InvalidateRect(hWndEdit, 0, TRUE)
end sub

' -----------------------------------------------------------------------------
'  Text safety (universal ANSI/Unicode support for AkelPad)
' -----------------------------------------------------------------------------
function GetLineText(byval hWnd as HWND, byval lineIdx as Integer, byval lineLen as Integer) as String
  if lineLen <= 0 then return ""
  const MAX_EM_GETLINE_CHARS as Integer = 32760
  dim safeLen as Integer = lineLen
  if safeLen > MAX_EM_GETLINE_CHARS then safeLen = MAX_EM_GETLINE_CHARS
  dim as String sRet = ""

  if IsWindowUnicode(hWnd) then
    dim wBuf as WString ptr = cast(WString ptr, callocate((safeLen + 2) * 2))
    cast(Short ptr, wBuf)[0] = safeLen
    dim lenRead as Integer = SendMessageW(hWnd, EM_GETLINE, lineIdx, cast(LPARAM, wBuf))
    if lenRead > 0 then sRet = Left(*wBuf, lenRead)
    deallocate(wBuf)
  else
    dim zBuf as ZString ptr = cast(ZString ptr, callocate(safeLen + 2))
    cast(Short ptr, zBuf)[0] = safeLen
    dim lenRead as Integer = SendMessageA(hWnd, EM_GETLINE, lineIdx, cast(LPARAM, zBuf))
    if lenRead > 0 then sRet = Left(*zBuf, lenRead)
    deallocate(zBuf)
  end if

  return sRet
end function

private function RawCartesianScalarToString(byref scalar as RawCartesianScalar) as String
  dim s as String = "kind=" & ltrim(str(scalar.kind))
  select case scalar.kind
    case RSK_INT64: s &= " int64=" & ltrim(str(scalar.intValue))
    case RSK_UINT64: s &= " uint64=" & ltrim(str(scalar.uintValue))
    case RSK_FLOATING: s &= " float=" & ltrim(str(scalar.floatValue))
    case RSK_RATIONAL: s &= " ratio=" & ltrim(str(scalar.ratNum)) & "/" & ltrim(str(scalar.ratDen))
    case else: s &= " ..."
  end select
  return s
end function

private sub LogInfo_InputAndResult(byref sLine as String, byref sRes as String, byref raw as RawResult)
  dim rawDbg as String = "; "
  if RawScalarIsComplex(raw.scalar) then
    rawDbg &= "real: " & RawCartesianScalarToString(raw.scalar.real)
    rawDbg &= ", imag: " & RawCartesianScalarToString(raw.scalar.imag)
  else
    rawDbg &= "real: " & RawCartesianScalarToString(raw.scalar.real)
  end if
  LogInfo("`" & sLine & "` -> `" & sRes & "`" & rawDbg)
end sub

private sub BuildRenderedResultText(byref sLine as String, byref sRes as String, byref bIsError as Boolean, byval lineIdx as Integer = -1)
  sRes = ""
  bIsError = FALSE
  if IsLikelyCodeLine(sLine) then exit sub

  if g_bLogParsedLines then
    const MAX_LOG_LINE_LEN as Integer = 220
    dim logLine as String = sLine
    if len(logLine) > MAX_LOG_LINE_LEN then
      logLine = left(logLine, MAX_LOG_LINE_LEN) & " ...[truncated]"
    end if
    if lineIdx >= 0 then
      LogInfo("parse-line begin [" & ltrim(str(lineIdx)) & "]: " & logLine)
    else
      LogInfo("parse-line begin: " & logLine)
    end if
  end if

  dim raw as RawResult

  if Parser_TryEvaluateExRaw(sLine, raw) then
    sRes = FormatRawEvaluationResult(raw)
    if g_bLogParsedLines then
      LogInfo_InputAndResult(sLine, sRes, raw)
    end if
  else
    dim sErr as String = Parser_GetLastError()
    if len(sErr) > 0 then
      sRes = SMARTMATH_ERROR_PREFIX & sErr
      bIsError = TRUE
    end if
  end if

  if g_bLogParsedLines then
    if lineIdx >= 0 then
      LogInfo("parse-line end [" & ltrim(str(lineIdx)) & "]")
    else
      LogInfo("parse-line end")
    end if
  end if
end sub

private sub EnsureRenderCache(byval hWnd as HWND)
  dim nLineCount as Integer = SendMessage(hWnd, EM_GETLINECOUNT, 0, 0)
  if nLineCount <= 0 then
    g_cacheLineCount = 0
    erase g_cachedLineText
    erase g_cachedRenderText
    g_cacheReady = TRUE
    g_cacheSyncedGeneration = g_cacheDocGeneration
    g_cacheDirtyFromLine = -1
    exit sub
  end if

  if IsRenderCacheSyncedWithDocument(hWnd) then exit sub

  dim needsFullRebuild as BOOL = (g_cacheReady = FALSE) orelse (nLineCount <> g_cacheLineCount)
  dim firstChanged as Integer = -1
  dim minCount as Integer = IIf(nLineCount < g_cacheLineCount, nLineCount, g_cacheLineCount)

  if needsFullRebuild = FALSE then
    dim scanFrom as Integer = 0
    if g_cacheDirtyFromLine >= 0 and g_cacheDirtyFromLine < nLineCount then
      scanFrom = g_cacheDirtyFromLine
    end if

    dim iScan as Integer
    for iScan = scanFrom to nLineCount - 1
      dim nLineIndexScan as Integer = SendMessage(hWnd, EM_LINEINDEX, iScan, 0)
      dim nLineLenScan as Integer = SendMessage(hWnd, EM_LINELENGTH, nLineIndexScan, 0)
      dim sLineScan as String = GetLineText(hWnd, iScan, nLineLenScan)
      if sLineScan <> g_cachedLineText(iScan) then
        firstChanged = iScan
        exit for
      end if
    next iScan

    if firstChanged < 0 then
      g_cacheSyncedGeneration = g_cacheDocGeneration
      g_cacheDirtyFromLine = -1
      exit sub
    end if
  end if

  if firstChanged < 0 then
    if g_cacheReady = FALSE then
      firstChanged = 0
    else
      firstChanged = minCount
    end if
  end if

  redim currentLineText(0 to nLineCount - 1) as String
  dim i as Integer
  for i = 0 to nLineCount - 1
    dim nLineIndex as Integer = SendMessage(hWnd, EM_LINEINDEX, i, 0)
    dim nLineLen as Integer = SendMessage(hWnd, EM_LINELENGTH, nLineIndex, 0)
    currentLineText(i) = GetLineText(hWnd, i, nLineLen)
  next i

  redim oldLineText(0 to 0) as String
  redim oldRenderText(0 to 0) as String
  dim oldCount as Integer = g_cacheLineCount
  if g_cacheReady andalso oldCount > 0 then
    redim oldLineText(0 to oldCount - 1)
    redim oldRenderText(0 to oldCount - 1)
    for i = 0 to oldCount - 1
      oldLineText(i) = g_cachedLineText(i)
      oldRenderText(i) = g_cachedRenderText(i)
    next i
  end if

  redim g_cachedLineText(0 to nLineCount - 1)
  redim g_cachedRenderText(0 to nLineCount - 1)
  for i = 0 to nLineCount - 1
    g_cachedLineText(i) = currentLineText(i)
  next i

  g_cacheLineCount = nLineCount

  Parser_ClearVariables()
  Parser_SetSupportComplexNumbers(g_bSupportComplexNumbers)
  for i = 0 to nLineCount - 1
    dim sRes as String = ""
    dim bIsError as Boolean = FALSE
    BuildRenderedResultText(g_cachedLineText(i), sRes, bIsError, i)

    if (i < firstChanged) andalso g_cacheReady andalso (i < oldCount) andalso (g_cachedLineText(i) = oldLineText(i)) then
      g_cachedRenderText(i) = oldRenderText(i)
      if IsVolatilePureExpressionLine(g_cachedLineText(i)) then
        ' Keep ans chain consistent with frozen rendered output for volatile pure expressions.
        RestoreAnsFromCachedRender(g_cachedRenderText(i))
      end if
    else
      g_cachedRenderText(i) = sRes
    end if
  next i

  g_cacheReady = TRUE
  g_cacheSyncedGeneration = g_cacheDocGeneration
  g_cacheDirtyFromLine = -1
end sub

function BuildLineRenderText(byval hWnd as HWND, byval lineIdx as Integer) as String
  EnsureRenderCache(hWnd)
  if lineIdx < 0 orelse lineIdx >= g_cacheLineCount then return ""
  return g_cachedRenderText(lineIdx)
end function

function BuildResultTextForLine(byval hWnd as HWND, byval targetLine as Integer) as String
  dim nLineCount as Integer = SendMessage(hWnd, EM_GETLINECOUNT, 0, 0)
  if targetLine < 0 orelse targetLine >= nLineCount then return ""
  EnsureRenderCache(hWnd)
  return BuildLineRenderText(hWnd, targetLine)
end function

function CopyTextToClipboard(byval hWndOwner as HWND, byref sText as String) as BOOL
  if len(sText) = 0 then return FALSE
  if OpenClipboard(hWndOwner) = FALSE then return FALSE
  EmptyClipboard()

  dim cbBytes as Integer = len(sText) + 1
  dim hMem as HGLOBAL = GlobalAlloc(GMEM_MOVEABLE, cbBytes)
  if hMem = 0 then
    CloseClipboard()
    return FALSE
  end if
  dim pMem as any ptr = GlobalLock(hMem)
  if pMem then
    memcpy(pMem, strptr(sText), cbBytes)
    GlobalUnlock(hMem)
    SetClipboardData(CF_TEXT, hMem)
    CloseClipboard()
    return TRUE
  end if
  GlobalFree(hMem)
  CloseClipboard()
  return FALSE
end function

' -----------------------------------------------------------------------------
'  Drawing Logic & Positioning
' -----------------------------------------------------------------------------
sub UpdateInternalState(byval hWnd as HWND, byref bVisible as BOOL)
  if g_bShuttingDown orelse (g_bSmartMathDocActive = FALSE) then
    bVisible = FALSE
    exit sub
  end if

  bVisible = FALSE

  if IsRenderCacheSyncedWithDocument(hWnd) = FALSE then
    EnsureRenderCache(hWnd)
  end if

  if g_cacheLineCount > 0 then
    bVisible = TRUE
  end if
end sub

' Client-area strip for one edit line (full width). Returns FALSE if unknown.
private function GetEditCharHeight(byval hWnd as HWND) as Integer
  dim nCharHeight as Integer = 0
  if not g_bOldRichEdit then
    nCharHeight = SendMessage(hWnd, AEM_GETCHARSIZE, 0, 0)
  end if
  if nCharHeight <= 0 then
    dim hDC as HDC = GetDC(hWnd)
    if hDC <> 0 then
      dim tm as TEXTMETRIC
      GetTextMetrics(hDC, @tm)
      nCharHeight = tm.tmHeight
      ReleaseDC(hWnd, hDC)
    end if
  end if
  if nCharHeight <= 0 then nCharHeight = 16
  return nCharHeight
end function

private function GetEditLineClientRect(byval hWnd as HWND, byval lineIdx as Integer, byref rcOut as RECT) as BOOL
  if lineIdx < 0 then return FALSE

  dim rcClient as RECT
  GetClientRect(hWnd, @rcClient)

  dim nLineIndex as Integer = SendMessage(hWnd, EM_LINEINDEX, lineIdx, 0)
  if nLineIndex < 0 then return FALSE

  dim ptY as Integer = -10001
  if g_bOldRichEdit then
    dim res as LRESULT = SendMessage(hWnd, EM_POSFROMCHAR, nLineIndex, 0)
    ptY = cast(short, HiWord(res))
  else
    dim pt as POINT
    SendMessage(hWnd, EM_POSFROMCHAR, cast(WPARAM, @pt), nLineIndex)
    ptY = pt.y
  end if
  if ptY < -10000 then return FALSE

  dim nCharHeight as Integer = GetEditCharHeight(hWnd)

  rcOut.left = rcClient.left
  rcOut.right = rcClient.right
  rcOut.top = ptY
  rcOut.bottom = ptY + nCharHeight
  return TRUE
end function

private sub InvalidateEditLine(byval hWnd as HWND, byval lineIdx as Integer)
  dim rcLine as RECT
  if GetEditLineClientRect(hWnd, lineIdx, rcLine) then
    InvalidateRect(hWnd, @rcLine, TRUE)
  end if
end sub

' Newly exposed client strip after a vertical scroll of (nNewFirst - nOldFirst) lines.
private function BuildScrollExposedRect(byval hWnd as HWND, byval nOldFirst as Integer, byval nNewFirst as Integer, byref rcOut as RECT) as BOOL
  if nOldFirst = nNewFirst then return FALSE

  dim rcClient as RECT
  GetClientRect(hWnd, @rcClient)
  dim nCharHeight as Integer = GetEditCharHeight(hWnd)
  dim delta as Integer = nNewFirst - nOldFirst
  dim nLines as Integer = abs(delta)
  dim nStrip as Integer = nLines * nCharHeight
  if nStrip > (rcClient.bottom - rcClient.top) then nStrip = rcClient.bottom - rcClient.top
  if nStrip <= 0 then return FALSE

  rcOut.left = rcClient.left
  rcOut.right = rcClient.right
  if delta > 0 then
    ' Content moved up: new lines appear at the bottom.
    rcOut.bottom = rcClient.bottom
    rcOut.top = rcClient.bottom - nStrip
  else
    ' Content moved down: new lines appear at the top.
    rcOut.top = rcClient.top
    rcOut.bottom = rcClient.top + nStrip
  end if
  return TRUE
end function

declare sub DrawDynamicMathResults(byval hWnd as HWND, byval prcClip as RECT ptr = 0, byval bContentUnchanged as BOOL = FALSE)

' Union line strips and/or the scroll-exposed edge, then draw overlays immediately
' (key-repeat starves WM_PAINT, so InvalidateRect alone leaves blank result areas).
private sub DrawMathResultsForScroll(byval hWnd as HWND, byval nOldFirst as Integer, byval nNewFirst as Integer, byval nOldCaret as Integer, byval nNewCaret as Integer)
  dim rcClip as RECT
  dim bHaveClip as BOOL = FALSE
  dim rcPart as RECT

  if BuildScrollExposedRect(hWnd, nOldFirst, nNewFirst, rcPart) then
    rcClip = rcPart
    bHaveClip = TRUE
  end if
  if (nOldCaret >= 0) andalso GetEditLineClientRect(hWnd, nOldCaret, rcPart) then
    if bHaveClip then
      UnionRect(@rcClip, @rcClip, @rcPart)
    else
      rcClip = rcPart
      bHaveClip = TRUE
    end if
  end if
  if (nNewCaret >= 0) andalso GetEditLineClientRect(hWnd, nNewCaret, rcPart) then
    if bHaveClip then
      UnionRect(@rcClip, @rcClip, @rcPart)
    else
      rcClip = rcPart
      bHaveClip = TRUE
    end if
  end if

  if bHaveClip then
    DrawDynamicMathResults(hWnd, @rcClip, TRUE)
  else
    DrawDynamicMathResults(hWnd, 0, TRUE)
  end if
end sub

' prcClip: when non-null, only paint result glyphs that intersect this rect
' (typically the WM_PAINT update region). Avoids transparent overdraw on
' lines the edit control did not erase.
' Before text, fill drawClip with the line background: word-select / dblclick
' paints often report a tall update region while only erasing the clicked line,
' which would otherwise stack transparent glyphs on lines below.
sub DrawDynamicMathResults(byval hWnd as HWND, byval prcClip as RECT ptr = 0, byval bContentUnchanged as BOOL = FALSE)
  if g_bShuttingDown then exit sub
  if g_bSmartMathDocActive = FALSE then exit sub
  if (prcClip <> 0) andalso (IsRectEmpty(prcClip) <> 0) then exit sub

  if bContentUnchanged = FALSE orelse IsRenderCacheSyncedWithDocument(hWnd) = FALSE then
    EnsureRenderCache(hWnd)
  end if

  dim hDC as HDC = GetDC(hWnd)
  if hDC = 0 then exit sub

  dim rcClient as RECT
  GetClientRect(hWnd, @rcClient)

  dim nFirstVisible as Integer = SendMessage(hWnd, EM_GETFIRSTVISIBLELINE, 0, 0)
  dim nLineCount as Integer = SendMessage(hWnd, EM_GETLINECOUNT, 0, 0)
  dim nCaretLine as Integer = SendMessage(hWnd, EM_EXLINEFROMCHAR, 0, -1)

  dim hFont as HFONT = cast(HFONT, SendMessage(hWnd, WM_GETFONT, 0, 0))
  dim hOldFont as HFONT
  if hFont then hOldFont = cast(HFONT, SelectObject(hDC, hFont))

  dim nCharHeight as Integer = 0
  if not g_bOldRichEdit then
    nCharHeight = SendMessage(hWnd, AEM_GETCHARSIZE, 0, 0)
  end if

  dim crBk as COLORREF = GetSysColor(COLOR_WINDOW)
  dim crActiveBk as COLORREF = crBk
  dim bUseActiveLineBk as BOOL = FALSE
  dim hBrushBk as HBRUSH = 0
  dim hBrushActive as HBRUSH = 0
  if not g_bOldRichEdit then
    dim aec as AECOLORS
    aec.dwFlags = AECLR_BASICBK or AECLR_ACTIVELINEBK
    SendMessage(hWnd, AEM_GETCOLORS, 0, cast(LPARAM, @aec))
    crBk = aec.crBasicBk
    crActiveBk = aec.crActiveLineBk
    dim dwOptions as DWORD = SendMessage(hWnd, AEM_GETOPTIONS, 0, 0)
    if (dwOptions and AECO_ACTIVELINE) then bUseActiveLineBk = TRUE
  end if
  hBrushBk = CreateSolidBrush(crBk)
  if bUseActiveLineBk then hBrushActive = CreateSolidBrush(crActiveBk)

  SetBkMode(hDC, TRANSPARENT)

  dim i as Integer
  for i = 0 to nLineCount - 1
    dim nLineIndex as Integer = SendMessage(hWnd, EM_LINEINDEX, i, 0)
    dim nLineLen as Integer = SendMessage(hWnd, EM_LINELENGTH, nLineIndex, 0)

    dim ptClient_x as Integer = -10001
    dim ptClient_y as Integer = -10001

    if i >= nFirstVisible then
      dim ptClient as POINT
      dim res as LRESULT
      if g_bOldRichEdit then
        res = SendMessage(hWnd, EM_POSFROMCHAR, nLineIndex + nLineLen, 0)
        ptClient_x = cast(short, LoWord(res))
        ptClient_y = cast(short, HiWord(res))
      else
        SendMessage(hWnd, EM_POSFROMCHAR, cast(WPARAM, @ptClient), nLineIndex + nLineLen)
        ptClient_x = ptClient.x
        ptClient_y = ptClient.y
      end if

      if ptClient_y > rcClient.bottom then exit for
      if prcClip <> 0 then
        ' Below the clip: no further visible lines.
        if ptClient_y >= prcClip->bottom then exit for
        ' Above the clip: skip until line height is known and line intersects.
        if (nCharHeight > 0) andalso (ptClient_y + nCharHeight <= prcClip->top) then
          continue for
        end if
      end if
    end if

    dim sRes as String = ""
    if i >= 0 andalso i < g_cacheLineCount then sRes = g_cachedRenderText(i)
    dim bIsError as Boolean = (left(sRes, Len(SMARTMATH_ERROR_PREFIX)) = SMARTMATH_ERROR_PREFIX)

    if len(sRes) > 0 then
      if i >= nFirstVisible andalso ptClient_y > -10000 then
        dim sz as SIZE
        GetTextExtentPoint32(hDC, strptr(sRes), Len(sRes), @sz)

        if nCharHeight <= 0 then nCharHeight = sz.cy

        dim lineRect as RECT
        lineRect.left = rcClient.left
        lineRect.right = rcClient.right
        lineRect.top = ptClient_y
        lineRect.bottom = ptClient_y + nCharHeight

        if bIsError then
          SetTextColor(hDC, &H0000FF)
        else
          SetTextColor(hDC, g_crResultColor)
        end if
        dim drawX as Integer = rcClient.right - sz.cx - 10
        dim drawY as Integer = lineRect.top + ((lineRect.bottom - lineRect.top) - sz.cy) \ 2
        dim minDrawX as Integer = ptClient_x + 6
        if minDrawX < rcClient.left then minDrawX = rcClient.left
        if minDrawX < rcClient.right then
          ' Clear the full result gutter (after source text to client right), not
          ' only the new string width. Otherwise a shorter/recalculated result is
          ' drawn over leftover glyphs from the previous overlay (e.g. after paste
          ' that reshuffles rand/UDF-dependent lines).
          dim clearRect as RECT
          clearRect.left = minDrawX
          clearRect.top = lineRect.top
          clearRect.right = rcClient.right
          clearRect.bottom = lineRect.bottom
          dim drawClip as RECT = clearRect
          if prcClip <> 0 then
            if IntersectRect(@drawClip, @clearRect, prcClip) = 0 then continue for
          end if
          dim hFill as HBRUSH = hBrushBk
          if bUseActiveLineBk andalso (i = nCaretLine) andalso (hBrushActive <> 0) then
            hFill = hBrushActive
          end if
          if hFill <> 0 then FillRect(hDC, @drawClip, hFill)
          dim clipRect as RECT
          clipRect.left = IIf(drawX > minDrawX, drawX, minDrawX)
          clipRect.top = lineRect.top
          clipRect.right = rcClient.right
          clipRect.bottom = lineRect.bottom
          dim textClip as RECT = clipRect
          if prcClip <> 0 then
            if IntersectRect(@textClip, @clipRect, prcClip) = 0 then continue for
          end if
          ExtTextOut(hDC, drawX, drawY, ETO_CLIPPED, @textClip, strptr(sRes), Len(sRes), 0)
        end if
      end if
    elseif i >= nFirstVisible andalso ptClient_y > -10000 andalso (bContentUnchanged = FALSE) then
      ' Content changed but this line has no result now: still erase any old overlay.
      if nCharHeight <= 0 then nCharHeight = GetEditCharHeight(hWnd)
      dim emptyLineRect as RECT
      emptyLineRect.left = rcClient.left
      emptyLineRect.right = rcClient.right
      emptyLineRect.top = ptClient_y
      emptyLineRect.bottom = ptClient_y + nCharHeight
      dim minClearX as Integer = ptClient_x + 6
      if minClearX < rcClient.left then minClearX = rcClient.left
      if minClearX < rcClient.right andalso ptClient_x > -10000 then
        dim clearEmpty as RECT
        clearEmpty.left = minClearX
        clearEmpty.top = emptyLineRect.top
        clearEmpty.right = rcClient.right
        clearEmpty.bottom = emptyLineRect.bottom
        dim eraseClip as RECT = clearEmpty
        if prcClip <> 0 then
          if IntersectRect(@eraseClip, @clearEmpty, prcClip) = 0 then continue for
        end if
        dim hFillEmpty as HBRUSH = hBrushBk
        if bUseActiveLineBk andalso (i = nCaretLine) andalso (hBrushActive <> 0) then
          hFillEmpty = hBrushActive
        end if
        if hFillEmpty <> 0 then FillRect(hDC, @eraseClip, hFillEmpty)
      end if
    end if
  next i

  if hBrushActive then DeleteObject(hBrushActive)
  if hBrushBk then DeleteObject(hBrushBk)
  if hFont then SelectObject(hDC, hOldFont)
  ReleaseDC(hWnd, hDC)
end sub

declare function EditGlobalProc stdcall(byval hWnd as HWND, byval uMsg as UINT, byval wParam as WPARAM, byval lParam as LPARAM) as LRESULT
declare function MainGlobalProc stdcall(byval hWnd as HWND, byval uMsg as UINT, byval wParam as WPARAM, byval lParam as LPARAM) as LRESULT
declare function FrameGlobalProc stdcall(byval hWnd as HWND, byval uMsg as UINT, byval wParam as WPARAM, byval lParam as LPARAM) as LRESULT

private sub SetSmartMathProcData(byval pd as PLUGINDATA ptr)
  SendMessage(pd->hMainWnd, AKD_SETEDITPROC, cast(WPARAM, @EditGlobalProc), cast(LPARAM, @lpEditProcData))
  SendMessage(pd->hMainWnd, AKD_SETMAINPROC, cast(WPARAM, @MainGlobalProc), cast(LPARAM, @lpMainProcData))
  if pd->nMDI = WMD_MDI then
    SendMessage(pd->hMainWnd, AKD_SETFRAMEPROC, cast(WPARAM, @FrameGlobalProc), cast(LPARAM, @lpFrameProcData))
  end if
end sub

private sub SetOriginalProcData()
  if lpEditProcData <> 0 then
    SendMessage(g_hMainWnd, AKD_SETEDITPROC, 0, cast(LPARAM, @lpEditProcData))
    lpEditProcData = 0
  end if
  if lpMainProcData <> 0 then
    SendMessage(g_hMainWnd, AKD_SETMAINPROC, 0, cast(LPARAM, @lpMainProcData))
    lpMainProcData = 0
  end if
  if lpFrameProcData <> 0 then
    SendMessage(g_hMainWnd, AKD_SETFRAMEPROC, 0, cast(LPARAM, @lpFrameProcData))
    lpFrameProcData = 0
  end if
end sub

' -----------------------------------------------------------------------------
'  Main-window subclass procedure
' -----------------------------------------------------------------------------
function MainGlobalProc stdcall(byval hWnd as HWND, byval uMsg as UINT, byval wParam as WPARAM, byval lParam as LPARAM) as LRESULT
  if uMsg = WM_COMMAND then
    dim nCmd as Integer = LOWORD(wParam)

    if nCmd = IDM_ACTIVE_ON_CURR_TAB then
      dim pFrameCurrent as FRAMEDATA ptr = cast(FRAMEDATA ptr, SendMessage(g_hMainWnd, AKD_FRAMEFIND, FWF_CURRENT, 0))
      dim idx as Integer = g_framesWithSmartMathEnabled.Find(pFrameCurrent)
      if idx >= 0 then
        g_framesWithSmartMathEnabled.RemoveAt(idx)
        SetSmartMathDocActiveState(pFrameCurrent->ei.hWndEdit, FALSE, FALSE)
        UpdateMenuActiveOnCurrTab(FALSE)
      else
        dim pFrameItem as FrameItem ptr = New FrameItem(pFrameCurrent)
        if pFrameItem <> 0 then
          pFrameItem->isWorking = TRUE
          g_framesWithSmartMathEnabled.Append(pFrameItem)
          SetSmartMathDocActiveState(pFrameCurrent->ei.hWndEdit, TRUE, TRUE)
          UpdateMenuActiveOnCurrTab(TRUE)
        end if
      end if
      return 0

    elseif nCmd = IDM_DECIMAL_BASE then
      g_nDecimals = -1
      InvalidateRenderCache()
      SaveSettings()
      UpdateMenuChecks()
      if g_hWndEdit then
        dim bVis as BOOL
        UpdateInternalState(g_hWndEdit, bVis)
        InvalidateRect(g_hWndEdit, 0, TRUE)
      end if
      return 0

    elseif nCmd > IDM_DECIMAL_BASE andalso nCmd <= IDM_DECIMAL_BASE + 1 + SMARTMATH_DECIMALS_MAX then
      g_nDecimals = nCmd - IDM_DECIMAL_BASE - 1
      InvalidateRenderCache()
      SaveSettings()
      UpdateMenuChecks()
      if g_hWndEdit then
        dim bVis as BOOL
        UpdateInternalState(g_hWndEdit, bVis)
        InvalidateRect(g_hWndEdit, 0, TRUE)
      end if
      return 0

    elseif nCmd >= IDM_COLOR_BASE andalso nCmd < IDM_COLOR_BASE + MAX_COLORS then
      select case nCmd
        case IDM_COLOR_GREEN:  g_crResultColor = &H008000
        case IDM_COLOR_BLUE:   g_crResultColor = &HFF0000
        case IDM_COLOR_RED:    g_crResultColor = &H0000FF
        case IDM_COLOR_YELLOW: g_crResultColor = &H00FFFF
        case IDM_COLOR_WHITE:  g_crResultColor = &HFFFFFF
        case IDM_COLOR_BLACK:  g_crResultColor = &H000000
      end select
      SaveSettings()
      UpdateMenuChecks()
      if g_hWndEdit then
        InvalidateRect(g_hWndEdit, 0, TRUE)
      end if
      return 0

    elseif nCmd = IDM_THOUSANDS_SEPARATOR then
      if g_bUseThousandsSeparator then
        g_bUseThousandsSeparator = FALSE
      else
        g_bUseThousandsSeparator = TRUE
      end if
      InvalidateRenderCache()
      SaveSettings()
      UpdateMenuChecks()
      if g_hWndEdit then
        dim bVis as BOOL
        UpdateInternalState(g_hWndEdit, bVis)
        InvalidateRect(g_hWndEdit, 0, TRUE)
      end if
      return 0

    elseif nCmd = IDM_COMPLEX_NUMBERS then
      if g_bSupportComplexNumbers then
        g_bSupportComplexNumbers = FALSE
      else
        g_bSupportComplexNumbers = TRUE
      end if
      InvalidateRenderCache()
      SaveSettings()
      UpdateMenuChecks()
      if g_hWndEdit then
        dim bVis as BOOL
        UpdateInternalState(g_hWndEdit, bVis)
        InvalidateRect(g_hWndEdit, 0, TRUE)
      end if
      return 0

    elseif nCmd = IDM_ABOUT then
      ShowAboutDialog(hWnd)
      return 0
    end if

  elseif (uMsg = AKDN_FRAME_ACTIVATE) orelse (uMsg = AKDN_OPENDOCUMENT_FINISH) then
    dim hWndEditCurrent as HWND = GetWndEdit(g_hMainWnd)
    if hWndEditCurrent <> 0 then
      if uMsg = AKDN_OPENDOCUMENT_FINISH then
        g_lastDocModeFrame = 0
      end if
      RefreshSmartMathDocMode(hWndEditCurrent)
    end if

  elseif uMsg = AKDN_FRAME_DESTROY then
    dim pFrame as FRAMEDATA ptr = cast(FRAMEDATA ptr, lParam)
    if pFrame <> 0 then
      g_framesWithSmartMathEnabled.RemoveAt(g_framesWithSmartMathEnabled.Find(pFrame))
    end if

  elseif uMsg = AKDN_MAIN_ONSTART_FINISH then
    if g_bSmartMathActive then
      dim result as LRESULT = 0
      if lpMainProcData andalso lpMainProcData->NextProc then
        result = lpMainProcData->NextProc(hWnd, uMsg, wParam, lParam)
      end if

      ' LogInfo("MainGlobalProc: AKDN_MAIN_ONSTART_FINISH, g_bAkelPadReady=TRUE")
      g_bAkelPadReady = TRUE
      dim hEditStart as HWND = GetWndEdit(g_hMainWnd)
      if hEditStart <> 0 then
        SyncSmartMathActiveEditorState(hEditStart)
      end if

      return result
    end if

  elseif uMsg = AKDN_MAIN_ONSTART_SHOW then
    ' On AKDN_MAIN_ONSTART_SHOW, the Sessions plugin loads all the files
    dim result as LRESULT = 0
    if lpMainProcData andalso lpMainProcData->NextProc then
      result = lpMainProcData->NextProc(hWnd, uMsg, wParam, lParam)
    end if
    ' The Sessions plugin has loaded all the files at this point

    ActivateSmartMathFiles()
    return result

  elseif uMsg = WM_CLOSE orelse uMsg = WM_QUERYENDSESSION orelse (uMsg = WM_SYSCOMMAND andalso wParam = SC_CLOSE) then
    if g_bActiveFramesSaved = FALSE then
      SaveActiveSmartMathFrames()
      g_bActiveFramesSaved = TRUE
    end if

  elseif uMsg = AKDN_MAIN_ONFINISH then
    g_bShuttingDown = TRUE
    dim bWasSmartMathActive as BOOL = g_bSmartMathActive

    g_bSmartMathActive = FALSE

    dim result as LRESULT = 0
    if lpMainProcData andalso lpMainProcData->NextProc then
      result = lpMainProcData->NextProc(hWnd, uMsg, wParam, lParam)
    end if

    if bWasSmartMathActive then
      SetOriginalProcData()
    end if

    UninitSmartMathMenu(TRUE)

    return result
  end if

  CheckEditNotifications(hWnd, uMsg, wParam, lParam)

  if lpMainProcData andalso lpMainProcData->NextProc then
    return lpMainProcData->NextProc(hWnd, uMsg, wParam, lParam)
  end if
  return 0
end function

' -----------------------------------------------------------------------------
'  Global Edit Window Subclass
' -----------------------------------------------------------------------------
function EditGlobalProc stdcall(byval hWnd as HWND, byval uMsg as UINT, byval wParam as WPARAM, byval lParam as LPARAM) as LRESULT
  if g_bShuttingDown then
    if lpEditProcData andalso lpEditProcData->NextProc then
      return lpEditProcData->NextProc(hWnd, uMsg, wParam, lParam)
    else
      return 0
    end if
  end if

  g_hWndEdit = hWnd

  dim lRes as LRESULT
  dim rcUpdate as RECT
  dim bNeedRedraw as BOOL = FALSE

  select case uMsg
    case WM_PAINT
      if g_bSmartMathDocActive then
        if GetUpdateRect(hWnd, @rcUpdate, FALSE) then
          dim rcIntersect as RECT
          if (rcOldMargin.right = 0) orelse IntersectRect(@rcIntersect, @rcUpdate, @rcOldMargin) then
            bNeedRedraw = TRUE
          end if
        end if
      end if

    case WM_SETFOCUS
      if not g_bOldRichEdit then
        SendMessage(hWnd, AEM_SETOPTIONS, AECOOP_OR, AECO_ACTIVELINE)
      end if

    case WM_SIZE, WM_MOUSEWHEEL, WM_VSCROLL, WM_HSCROLL
      if rcOldMargin.right > 0 then InvalidateRect(hWnd, @rcOldMargin, TRUE)

    case WM_LBUTTONDBLCLK
      dim xPos as Integer = cast(short, loword(lParam))
      dim yPos as Integer = cast(short, hiword(lParam))
      dim rcClient as RECT
      GetClientRect(hWnd, @rcClient)

      dim charPos as Integer = -1
      if g_bOldRichEdit then
        dim posRes as LRESULT = SendMessage(hWnd, EM_CHARFROMPOS, 0, cast(LPARAM, MAKELPARAM(xPos, yPos)))
        charPos = cast(short, loword(posRes))
      else
        dim pt as POINT
        pt.x = xPos
        pt.y = yPos
        charPos = SendMessage(hWnd, EM_CHARFROMPOS, 0, cast(LPARAM, @pt))
      end if

      if charPos >= 0 then
        dim lineIdx as Integer = SendMessage(hWnd, EM_EXLINEFROMCHAR, 0, charPos)
        dim sRes as String = BuildResultTextForLine(hWnd, lineIdx)
        if len(sRes) > 0 then
          dim hitResultArea as BOOL = FALSE
          dim nLineIndex as Integer = SendMessage(hWnd, EM_LINEINDEX, lineIdx, 0)
          dim nLineLen as Integer = SendMessage(hWnd, EM_LINELENGTH, nLineIndex, 0)
          dim ptLineEndX as Integer = rcClient.left
          if nLineIndex >= 0 then
            if g_bOldRichEdit then
              dim res as LRESULT = SendMessage(hWnd, EM_POSFROMCHAR, nLineIndex + nLineLen, 0)
              ptLineEndX = cast(short, LoWord(res))
            else
              dim ptLine as POINT
              SendMessage(hWnd, EM_POSFROMCHAR, cast(WPARAM, @ptLine), nLineIndex + nLineLen)
              ptLineEndX = ptLine.x
            end if
          end if

          dim hDC as HDC = GetDC(hWnd)
          if hDC <> 0 then
            dim hFont as HFONT = cast(HFONT, SendMessage(hWnd, WM_GETFONT, 0, 0))
            dim hOldFont as HFONT
            if hFont then hOldFont = cast(HFONT, SelectObject(hDC, hFont))
            dim sz as SIZE
            GetTextExtentPoint32(hDC, strptr(sRes), Len(sRes), @sz)
            if hFont then SelectObject(hDC, hOldFont)
            ReleaseDC(hWnd, hDC)

            dim drawX as Integer = rcClient.right - sz.cx - 10
            dim minDrawX as Integer = ptLineEndX + 6
            if minDrawX < rcClient.left then minDrawX = rcClient.left
            dim clipLeft as Integer = IIf(drawX > minDrawX, drawX, minDrawX)
            if xPos >= clipLeft andalso xPos <= rcClient.right then hitResultArea = TRUE
          end if

          if hitResultArea then
            dim sCopy as String = NormalizeCopiedResult(sRes)
            if CopyTextToClipboard(hWnd, sCopy) then
              return 0
            end if
          end if
        end if
      end if
  end select

  if lpEditProcData andalso lpEditProcData->NextProc then
    lRes = lpEditProcData->NextProc(hWnd, uMsg, wParam, lParam)
  else
    lRes = 0
  end if

  if uMsg = WM_PAINT then
    if bNeedRedraw then
      ' Clip overlays to the pre-NextProc update region so partial paints
      ' (e.g. same-line click) do not transparently overdraw other lines.
      DrawDynamicMathResults(hWnd, @rcUpdate)
    end if
  else
    select case uMsg
      case WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, WM_SYSKEYUP, _
           WM_LBUTTONDOWN, WM_LBUTTONUP, WM_MOUSEMOVE, _
           WM_MOUSEWHEEL, WM_HSCROLL, WM_VSCROLL, WM_CHAR, WM_SIZE
        if g_bSmartMathDocActive = FALSE then
          return lRes
        end if

        if (uMsg <> WM_MOUSEMOVE) orelse (wParam and MK_LBUTTON) then
          dim bVisible as BOOL
          UpdateInternalState(hWnd, bVisible)

          dim rcClient as RECT
          GetClientRect(hWnd, @rcClient)

          dim rcNewMargin as RECT = rcClient

          dim nFirstVisible as Integer = SendMessage(hWnd, EM_GETFIRSTVISIBLELINE, 0, 0)
          dim nCaretLine as Integer = SendMessage(hWnd, EM_EXLINEFROMCHAR, 0, -1)

          ' Arrow/nav keys must not force a full-client erase (that flickers every
          ' visible result). Same-line moves need no SmartMath invalidate; caret
          ' line changes only refresh the old and new line strips.
          dim bIsNavKey as BOOL = FALSE
          dim bForceKeyRedraw as BOOL = FALSE
          if (uMsg = WM_KEYDOWN) orelse (uMsg = WM_SYSKEYDOWN) then
            if (wParam = VK_UP) orelse (wParam = VK_DOWN) orelse (wParam = VK_LEFT) orelse (wParam = VK_RIGHT) orelse _
               (wParam = VK_HOME) orelse (wParam = VK_END) orelse (wParam = VK_PRIOR) orelse (wParam = VK_NEXT) then
              bIsNavKey = TRUE
              dim selStart as Integer = 0, selEnd as Integer = 0
              SendMessage(hWnd, EM_GETSEL, cast(WPARAM, @selStart), cast(LPARAM, @selEnd))
              nLastNavSelStart = selStart
              nLastNavSelEnd = selEnd
            else
              ' Backspace/Delete/etc.: content may change without WM_CHAR.
              bForceKeyRedraw = TRUE
              nLastNavSelStart = -1
              nLastNavSelEnd = -1
            end if
          end if

          if bVisible then
            dim bSizeChanged as BOOL = (rcNewMargin.left <> rcOldMargin.left) orelse _
                                       (rcNewMargin.right <> rcOldMargin.right)
            dim bScrollChanged as BOOL = (nFirstVisible <> nOldFirstLine)
            dim bCaretLineChanged as BOOL = (nCaretLine <> nOldCaretLine)

            if bSizeChanged orelse (uMsg = WM_CHAR) orelse bForceKeyRedraw then
              if rcOldMargin.right > 0 then InvalidateRect(hWnd, @rcOldMargin, TRUE)
              InvalidateRect(hWnd, @rcNewMargin, TRUE)

              rcOldMargin = rcNewMargin
              nOldFirstLine = nFirstVisible
              nOldCaretLine = nCaretLine
            elseif bScrollChanged then
              ' Do not full-erase invalidate: key-repeat starves WM_PAINT and leaves
              ' blank result strips on newly exposed lines. Draw the edge (and caret
              ' lines) synchronously after the edit control has scrolled.
              DrawMathResultsForScroll(hWnd, nOldFirstLine, nFirstVisible, nOldCaretLine, nCaretLine)
              rcOldMargin = rcNewMargin
              nOldFirstLine = nFirstVisible
              nOldCaretLine = nCaretLine
            elseif bCaretLineChanged then
              ' Active-line background moved: refresh only those two lines.
              if nOldCaretLine >= 0 then InvalidateEditLine(hWnd, nOldCaretLine)
              InvalidateEditLine(hWnd, nCaretLine)
              rcOldMargin = rcNewMargin
              nOldFirstLine = nFirstVisible
              nOldCaretLine = nCaretLine
            elseif bIsNavKey then
              ' Left/Right/Home/End on the same line: leave overlays alone.
              rcOldMargin = rcNewMargin
              nOldFirstLine = nFirstVisible
              nOldCaretLine = nCaretLine
            end if
          else
            if rcOldMargin.right > 0 then
              InvalidateRect(hWnd, @rcOldMargin, TRUE)
              rcOldMargin.left = 0 : rcOldMargin.right = 0
              rcOldMargin.top = 0  : rcOldMargin.bottom = 0
              nOldFirstLine = -1
              nOldCaretLine = -1
              nLastNavSelStart = -1
              nLastNavSelEnd = -1
            end if
          end if
        end if
    end select
  end if

  return lRes
end function

' -----------------------------------------------------------------------------
'  Global Frame Window Subclass
' -----------------------------------------------------------------------------
function FrameGlobalProc stdcall(byval hWnd as HWND, byval uMsg as UINT, byval wParam as WPARAM, byval lParam as LPARAM) as LRESULT
  CheckEditNotifications(hWnd, uMsg, wParam, lParam)
  if lpFrameProcData andalso lpFrameProcData->NextProc then
    return lpFrameProcData->NextProc(hWnd, uMsg, wParam, lParam)
  else
    return 0
  end if
end function

' -----------------------------------------------------------------------------
'  Exported Functions
' -----------------------------------------------------------------------------
extern "C"

sub DllAkelPadID alias "DllAkelPadID" (byval pv as PLUGINVERSION ptr) export
  pv->dwAkelDllVersion = AKELDLL
  pv->dwExeMinVersion3x = MAKE_IDENTIFIER(-1, -1, -1, -1)
  pv->dwExeMinVersion4x = MAKE_IDENTIFIER(4, 9, 7, 0)
  pv->pPluginName = @"SmartMath"
end sub

sub ToggleSmartMath alias "ToggleSmartMath" (byval pd as PLUGINDATA ptr) export
  if (pd->dwSupport and PDS_GETSUPPORT) then exit sub

  LogInfo("--- ToggleSmartMath Called ---")

  g_hMainWnd  = pd->hMainWnd
  g_hMainMenu = pd->hMainMenu
  g_hWndEdit  = pd->hWndEdit
  g_bOldRichEdit = pd->bOldRichEdit

  ' dim pf as PLUGINFUNCTION ptr

  if g_bSmartMathActive then
    LogInfo("Deactivating Global Edit Hook...")
    pd->nUnload = UD_UNLOAD

    SetOriginalProcData()
    UninitSmartMathMenu(FALSE)

    SaveActiveSmartMathFrames()
    g_framesWithSmartMathEnabled.Clear()
    g_bActiveFramesSaved = FALSE
    g_bSmartMathActive = FALSE
    InvalidateRenderCache()

    if pd->hWndEdit then
      if not g_bOldRichEdit then
        SendMessage(pd->hWndEdit, AEM_SETOPTIONS, AECOOP_SET, dwOldAkelOptions)
      end if
      InvalidateRect(pd->hWndEdit, 0, TRUE)
    end if

    ' The code below leads to undesired effects;
    ' instead, use AkelPad's native way of auto-load
    ' by checking the plugin's function checkbox.
    'pf = cast(PLUGINFUNCTION ptr, SendMessageW(pd->hMainWnd, AKD_DLLFINDW, cast(WPARAM, @wstr("SmartMath::ToggleSmartMath")), 0))
    'if pf <> 0 then
    '  pf->bAutoLoad = FALSE
    '  SendMessage(pd->hMainWnd, AKD_DLLSAVE, DLLSF_NOW, 0)
    'end if

  else
    LogInfo("Activating Global Edit Hook...")
    pd->nUnload = UD_NONUNLOAD_ACTIVE

    LoadSettings()
    ActivateSmartMathFiles()

    g_bShuttingDown = FALSE
    InvalidateRenderCache()

    rcOldMargin.left = 0 : rcOldMargin.right = 0
    rcOldMargin.top = 0  : rcOldMargin.bottom = 0
    nOldFirstLine = -1
    nOldCaretLine = -1
    nLastNavSelStart = -1
    nLastNavSelEnd = -1

    SetSmartMathProcData(pd)

    g_bSmartMathActive = TRUE

    InitSmartMathMenu()

    dim hEditAct as HWND = pd->hWndEdit
    if hEditAct = 0 then hEditAct = GetWndEdit(pd->hMainWnd)

    ' LogInfo("ToggleSmartMath: hEditAct=" & hEditAct)
    if hEditAct <> 0 then
      SyncSmartMathActiveEditorState(hEditAct)
    end if

    ' The code below leads to undesired effects;
    ' instead, use AkelPad's native way of auto-load
    ' by checking the plugin's function checkbox .
    'pf = cast(PLUGINFUNCTION ptr, SendMessageW(pd->hMainWnd, AKD_DLLFINDW, cast(WPARAM, @wstr("SmartMath::ToggleSmartMath")), 0))
    'if pf <> 0 then
    '  pf->bAutoLoad = TRUE
    '  if pd->bOnStart = FALSE then
    '    SendMessage(pd->hMainWnd, AKD_DLLSAVE, DLLSF_NOW, 0)
    '  end if
    'end if

  end if

  LogInfo("--- ToggleSmartMath Finished ---")
end sub

end extern

function DllMain(byval hinstDLL as HINSTANCE, byval fdwReason as DWORD, byval lpvReserved as LPVOID) as WINBOOL
  return TRUE
end function