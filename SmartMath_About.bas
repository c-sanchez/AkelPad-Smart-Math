#include once "SmartMath_Globals.bi"

sub ShowAboutDialog(byval hWnd as HWND)
  dim sAbout as WString * 1024
  sAbout = "SmartMath Plugin" & wchr(13, 10) & _
           "Real-time Mathematics for AkelPad" & wchr(13, 10, 13, 10) & _
           "Transform AkelPad into a high-performance scratchpad." & wchr(13, 10, 13, 10) & _
           wchr(&h2714) & " REAL-TIME: No need to press 'Equal'." & wchr(13, 10) & _
           wchr(&h2714) & " SMART: Logic-aware percentage handling." & wchr(13, 10) & _
           wchr(&h2714) & " VISUAL: Customizable colors." & wchr(13, 10) & _
           wchr(&h2714) & " PERSISTENT: Auto-loads your preferred settings." & wchr(13, 10, 13, 10) & _
           "Designed and Coded by Carlos Sánchez (2026)." & wchr(13, 10) & _
           "Special thanks to the FreeBASIC community." & wchr(13, 10, 13, 10) & _
           "Open official GitHub repository?"
  
  if MessageBoxW(hWnd, @sAbout, wstr("SmartMath - Information"), MB_YESNOCANCEL or MB_ICONINFORMATION) = IDYES then
    ShellExecuteW(0, wstr("open"), wstr("https://github.com/c-sanchez/AkelPad-Smart-Math"), 0, 0, SW_SHOWNORMAL)
  end if
end sub