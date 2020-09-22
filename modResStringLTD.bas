Attribute VB_Name = "modResStringLTD"
Option Explicit

'This mod is what's left of a great project "Resource Viewer/Extractor" AT
'http://www.planet-source-code.com/vb/scripts/showcode.asp?txtCodeId=25890&lngWId=1
'by 'Ark'
'I'm not adding many comments, I've hacked it to the point where it's almost single-purpose.
'(opens a file to access resources and retrieve a certain string)
'You should really check out the original.

Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
Private Declare Function SizeofResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function FreeResource Lib "kernel32" (ByVal hResData As Long) As Long
Private Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As String) As Long
Private Declare Function FindResourceByNum Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As Long) As Long

Public hModule As Long

Public Function GetString(ByVal ResName As String) As String
   Dim arr() As Byte
   Dim nPos As Long, wID As Long, uLength As Long
   Dim s As String, sText As String
   On Error Resume Next
   arr = GetDataArray("6", ResName)
   If Err.Number = 0 Then
       For wID = (CLng(Mid(ResName, 2)) - 1) * 16 To CLng(Mid(ResName, 2)) * 16 - 1
           Call CopyMemory(uLength, arr(nPos), 2)
           If uLength Then
              s = String(uLength, 0)
              CopyMemory ByVal StrPtr(s), arr(nPos + 2), uLength * 2
              If wID = 1 Then 'special for this project
                s = Replace$(s, vbLf, vbCrLf)
                s = Replace$(s, vbCr & vbCrLf, vbCrLf)
                sText = TrimNULL(s)
                Exit For
              End If
              nPos = nPos + uLength * 2 + 2
           Else
              nPos = nPos + 2
           End If
       Next wID
       GetString = sText
    End If
End Function

Public Function GetDataArray(ByVal ResType As String, ByVal ResName As String) As Variant
   Dim hRsrc As Long
   Dim hGlobal As Long
   Dim arrData() As Byte
   Dim lpData As Long
   Dim arrSize As Long
   If IsNumeric(ResType) Then hRsrc = FindResourceByNum(hModule, ResName, CLng(ResType))
   If hRsrc = 0 Then hRsrc = FindResource(hModule, ResName, ResType)
   If hRsrc = 0 Then Exit Function
   hGlobal = LoadResource(hModule, hRsrc)
   lpData = LockResource(hGlobal)
   arrSize = SizeofResource(hModule, hRsrc)
   If arrSize = 0 Then Exit Function
   ReDim arrData(arrSize - 1)
   Call CopyMemory(arrData(0), ByVal lpData, arrSize)
   Call FreeResource(hGlobal)
   GetDataArray = arrData
End Function


Public Function InitResource(ByVal sLibName As String) As Boolean
  On Error Resume Next
  hModule = LoadLibraryEx(sLibName, 0, 1)
  InitResource = (hModule <> 0)
End Function

Public Function TrimNULL(ByVal str As String) As String
    If InStr(str, Chr$(0)) > 0& Then
        TrimNULL = Left$(str, InStr(str, Chr$(0)) - 1&)
    Else
        TrimNULL = str
    End If
End Function

Public Sub ClearResource()
   If hModule Then FreeLibrary (hModule)
   hModule = 0
End Sub

