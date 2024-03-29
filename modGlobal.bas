Attribute VB_Name = "modGlobal"
Option Explicit

'Tooltip constants
Public Const DMESS As String = "Screensaver Disabled"
Public Const EMESS As String = "Screensaver Enabled"
Public Const NMESS As String = "No Screensaver Selected"

'Find a window and get the handle
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'check that the handle returned (above) really IS a window
Public Declare Function IsWindow Lib "user32" _
    (ByVal hwnd As Long) As Long
'Pretty self-explanatory...
Public Declare Function GetDoubleClickTime Lib "user32" () As Long
'Returns time in seconds since midnight
Public Declare Function GetTickCount Lib "kernel32" () As Long

'Similar to 'Shell' - Launches program or file
Public Declare Function ShellExecute _
    Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowbut As Long) As Long
    
Private Const OPTIONSKEY As String = "SOFTWARE\Blue Knot\SaverSwitch\Options"

Public blnStartup As Boolean

Public Sub LoadSettings()
    blnStartup = CBool(GetStringRegValue(HKEY_CURRENT_USER, OPTIONSKEY, "LoadAtStartup", CStr(False)))
End Sub

Public Sub SaveSettings()
    CreateRegKey HKEY_CURRENT_USER, "SOFTWARE\Blue Knot"
    CreateRegKey HKEY_CURRENT_USER, "SOFTWARE\Blue Knot\SaverSwitch"
    CreateRegKey HKEY_CURRENT_USER, OPTIONSKEY

    SetStringRegValue HKEY_CURRENT_USER, OPTIONSKEY, "LoadAtStartup", CStr(blnStartup)

End Sub

Public Sub RunAtStartup(blnRun As Boolean)
    If blnRun Then
        SetStringRegValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", App.Title, getFullPath(App.Path, App.EXEName & ".exe")
    Else
        DeleteRegValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", App.Title
    End If
End Sub

Public Function getFullPath(strPath As String, strFile As String) As String
    If Right$(strPath, 1) = "\" Then
        getFullPath = strPath & strFile
    ElseIf strPath = "" Then
        getFullPath = strFile
    Else
        getFullPath = strPath & "\" & strFile
    End If
End Function

Public Function getFileName(strFile As String) As String
Dim intDot As Integer
    intDot = InStrRev(strFile, ".")
    If intDot > 0 Then
        getFileName = Left$(strFile, intDot - 1)
    Else
        getFileName = strFile
    End If
End Function

