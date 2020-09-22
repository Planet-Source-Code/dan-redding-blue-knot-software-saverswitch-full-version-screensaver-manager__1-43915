Attribute VB_Name = "modSSInfo"
Option Explicit

'just a convenient structure for sorting
Public Type SSInfo
    File As String
    Path As String
    Name As String
End Type

'Array of the same
Public SSList() As SSInfo

'Add an item to the list
Public Sub AddSSFile(strPath As String, strFile As String, strName As String)
Dim iUBound As Integer
    On Error Resume Next 'UBound on an empty array throws an error
    iUBound = 0
    iUBound = UBound(SSList)
    If Err.Number <> 0 Then
        ReDim SSList(0) As SSInfo
    End If
    On Error GoTo 0
    iUBound = iUBound + 1 'index of new
    ReDim Preserve SSList(0 To iUBound)
    SSList(iUBound).File = strFile
    SSList(iUBound).Path = strPath
    SSList(iUBound).Name = strName
End Sub

'check if there's a matching .File on the list
Public Function CheckSSFile(strFile As String) As Boolean
Dim iLoop As Integer, iUBound As Integer
    On Error Resume Next
    iUBound = UBound(SSList)
    If Err.Number <> 0 Then
        CheckSSFile = False
        Exit Function
    End If
    On Error GoTo 0
    For iLoop = 1 To iUBound
        If SSList(iLoop).File = strFile Then
            CheckSSFile = True
            Exit Function
        End If
    Next iLoop
    CheckSSFile = False
End Function

Public Function GetSSName(strPath As String, strFile As String) As String
Dim strName As String
    'Display name is in a string resource with ID#1
    'load & retrieve (see modResStringLTD.bas)
    If InitResource(strPath) Then
        strName = GetString("#1")
        ClearResource
    End If
    'if that turned up a blank, just use the filename - .scr
    If strName = "" Then
        strName = getFileName(strFile)
    End If
    GetSSName = strName
End Function

'Simple bubble sort.
Public Sub SortSS()
Dim iLoop As Integer, iUBound As Integer, blnBubble As Boolean, SSX As SSInfo
    On Error Resume Next
    iUBound = UBound(SSList)
    If Err.Number <> 0 Or iUBound = 1 Then
        Exit Sub
    End If
    Do
        blnBubble = False
        For iLoop = 1 To iUBound - 1
            If UCase$(SSList(iLoop).Name) > UCase$(SSList(iLoop + 1).Name) Then
                SSX = SSList(iLoop)
                SSList(iLoop) = SSList(iLoop + 1)
                SSList(iLoop + 1) = SSX
                blnBubble = True
            End If
        Next iLoop
    Loop While blnBubble
End Sub

