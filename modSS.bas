Attribute VB_Name = "modSS"
Option Explicit

'Gets & Sets info from system-- in this case the active state of the SS
Private Declare Function SystemParametersInfo _
    Lib "user32" Alias "SystemParametersInfoA" _
    (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Long, _
    ByVal fuWinIni As Long) As Long
'Constants for SystemParametersInfo
Private Const SPI_SETSCREENSAVEACTIVE = 17
Private Const SPI_GETSCREENSAVEACTIVE = 16
Private Const SPI_GETSCREENSAVETIMEOUT = 14
Private Const SPI_SETSCREENSAVETIMEOUT = 15

'Gets handle of desktop window
Public Declare Function GetDesktopWindow Lib "user32" () As Long

'Send message to a window
Private Declare Function SendMessage _
    Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
'Windows Messages for SendMessage
Private Const WM_SYSCOMMAND = &H112&
Private Const SC_SCREENSAVE = &HF140&

'Is Screensaver Enabled?
Public Function IsSSEnabled() As Boolean
Dim lTemp As Long
    lTemp = 0&
    SystemParametersInfo SPI_GETSCREENSAVEACTIVE, 0&, lTemp, 0&
    IsSSEnabled = (lTemp = 1&)
End Function

'Set Screensaver Enabled
Public Sub SetSSEnabled(blnEnable As Boolean)
Dim lFlag As Long
    lFlag = IIf(blnEnable, 1&, 0&) '1 to enable, 0 to disable
    SystemParametersInfo SPI_SETSCREENSAVEACTIVE, lFlag, 0&, 0&
End Sub

'Run the current SS
Public Sub StartSS()
Dim lDesktop As Long
    lDesktop = GetDesktopWindow()
    'tell screensaver to run in front of desktop window
    SendMessage lDesktop, WM_SYSCOMMAND, SC_SCREENSAVE, 0&
End Sub

'Get/Set path of current SS
Public Function GetCurrSS()
    GetCurrSS = GetStringRegValue(HKEY_CURRENT_USER, "Control Panel\Desktop", "SCRNSAVE.EXE", "")
End Function

Public Sub SetCurrSS(strPath As String)
    SetStringRegValue HKEY_CURRENT_USER, "Control Panel\Desktop", "SCRNSAVE.EXE", strPath
End Sub

'Get/Set password protected setting
Public Function GetSSSecure() As Boolean
    GetSSSecure = CBool(GetLongRegValue(HKEY_CURRENT_USER, "Control Panel\Desktop", "ScreenSaverIsSecure", CLng(False)))
End Function

Public Function SetSSSecure(blnSecure As Boolean)
    SetLongRegValue HKEY_CURRENT_USER, "Control Panel\Desktop", "ScreenSaverIsSecure", Abs(CLng(blnSecure))
End Function

'Get/Set timeout minutes
Public Function GetSSTimeout() As Long
Dim lTemp As Long
    SystemParametersInfo SPI_GETSCREENSAVETIMEOUT, 0&, lTemp, 0&
    GetSSTimeout = lTemp \ 60
End Function

Public Function SetSSTimeout(lMin As Long)
    SystemParametersInfo SPI_SETSCREENSAVETIMEOUT, lMin * 60&, 0&, 0&
End Function

