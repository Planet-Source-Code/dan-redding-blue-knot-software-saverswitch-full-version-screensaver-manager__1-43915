Attribute VB_Name = "modSpecialFolder"
Option Explicit

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
    (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" _
    (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
    
Public Function SpecialFolder(ByVal CSIDL As Long) As String
'Another Mr. BoBo help
Dim r As Long
Dim sPath As String
Dim IDL As ITEMIDLIST
Const NOERROR = 0
Const MAX_LENGTH = 260
r = SHGetSpecialFolderLocation(frmTray.hwnd, CSIDL, IDL)
If r = NOERROR Then
    sPath = Space$(MAX_LENGTH)
    r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
    If r Then
        SpecialFolder = Left$(sPath, InStr(sPath, Chr$(0)) - 1)
    End If
End If
 'Here's the list of special folders
 
 '0=Desktop
 '2=StartMenu\Programs
 '5=My Documents
 '6=Favorites
 '7=Startup
 '8=Recent
 '9=SendTo
 '11=StartMenu
 '13=My Music
 '14=My Videos
 '16=Desktop
 '19=Nethood
 '20=Fonts
 '21=ShellNew
 '22=All Users\Start Menu
 '23=All Users\Start Menu\Programs
 '24=All Users\Start Menu\Programs\Startup
 '25=All users\desktop
 '26=Application Data
 '27=PrintHood
 '28=Local Settings\Application Data
 '31=All Users\Favorites
 '32=Temporary Internet Files
 '33=Cookies
 '34=History
 '35 All Users\Application Data
 '36=Windows
 '37=Windows\System
 '38=Program Files
 '39=My Pictures
 '40=Current User Root
 '41=Windows\System
 '43=Program Files\Common Files
 '45=All Users\Templates
 '46=All Users\Documents
 '47=All Users\Start Menu\Programs\Adminstrative Tools
End Function


