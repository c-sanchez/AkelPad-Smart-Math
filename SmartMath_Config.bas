#include once "SmartMath_Globals.bi"

' -----------------------------------------------------------------------------
'  Settings Loading & Saving
' -----------------------------------------------------------------------------
sub LoadSettings()
  if g_wszIniPath <> "" then
    dim sVal as WString * 32
    GetPrivateProfileStringW(wstr("Settings"), wstr("Decimals"), wstr("-1"), @sVal, 32, g_wszIniPath)
    g_nDecimals = Val(sVal)

    ' 32768 corresponds to &H008000 (Green)
    GetPrivateProfileStringW(wstr("Settings"), wstr("Color"), wstr("32768"), @sVal, 32, g_wszIniPath)
    g_crResultColor = Val(sVal)
    
    GetPrivateProfileStringW(wstr("Settings"), wstr("ThousandsSeparator"), wstr("0"), @sVal, 32, g_wszIniPath)
    if Val(sVal) <> 0 then
      g_bUseThousandsSeparator = TRUE
    else
      g_bUseThousandsSeparator = FALSE
    end if
  end if
end sub

sub SaveSettings()
  if g_wszIniPath <> "" then
    dim sVal as WString * 32 = WStr(g_nDecimals)
    WritePrivateProfileStringW(wstr("Settings"), wstr("Decimals"), sVal, g_wszIniPath)

    sVal = WStr(g_crResultColor)
    WritePrivateProfileStringW(wstr("Settings"), wstr("Color"), sVal, g_wszIniPath)
    
    if g_bUseThousandsSeparator then sVal = wstr("1") else sVal = wstr("0")
    WritePrivateProfileStringW(wstr("Settings"), wstr("ThousandsSeparator"), sVal, g_wszIniPath)
  end if
end sub