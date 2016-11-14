Attribute VB_Name = "MdlNode"
Option Explicit

Dim IoI, File2Num, FileNum, text, X
Const FILE_BEGIN As Long = 0
Const FILE_SHARE_READ As Long = &H1
Const FILE_SHARE_WRITE As Long = &H2
Const OPEN_EXISTING As Long = 3
Const OPEN_ALWAYS As Long = 4
Const GENERIC_READ As Long = &H80000000
Const GENERIC_WRITE As Long = &H40000000

Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Public N, M, Y, FileSize, MM

' -----------
' Functions
' -----------


Public Function OpenFile(FileName As String) As String
On Error Resume Next
Dim hFile As Long, l As Long, result As Long
Dim sText As String
''///////////////////////////////////////////////////'
hFile = CreateFile(FileName, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
l = GetFileSize(hFile, 0)
SetFilePointer hFile, 0, 0, FILE_BEGIN
sText = Space$(l)
ReadFile hFile, ByVal sText, l, result, ByVal 0&
If result <> l Then Exit Function
CloseHandle hFile
'///////////////////////////////////////////////////'
'FileNum = FreeFile
'Open FileName For Binary As #FileNum
'sText = Input(LOF(FileNum), #FileNum)
'Close #FileNum
OpenFile = sText
End Function

Public Sub SaveFile(FileName As String, sText As String)
On Error Resume Next

    Dim hFile As Long, result As Long

    '///////////////////////////////////////////////////'
    ' get a handle for the file
    hFile = CreateFile(FileName, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_ALWAYS, 0, 0)
'    hFile = CreateFile(FileName, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)

    SetFilePointer hFile, 0, 0, FILE_BEGIN

    ' write data
    '=========================================='
    WriteFile hFile, ByVal sText, Len(sText), result, ByVal 0&
    '=========================================='

    ' important: truncate file at current position
    SetEndOfFile hFile

    ' close file handle
    CloseHandle hFile
    '///////////////////////////////////////////////////'
End Sub



Public Sub WriteAllText(FileName, text, kind As Integer)
File2Num = FreeFile
If kind = 1 Then
Open FileName For Binary As File2Num
Put #File2Num, , text
Close #File2Num
ElseIf kind = 2 Then
Open FileName For Random As File2Num
Put #File2Num, , text
Close #File2Num
ElseIf kind = 3 Then

Open FileName For Append As #File2Num
Print #File2Num, text
Close #File2Num

ElseIf kind = 4 Then
Open FileName For Append As File2Num
Write #File2Num, text
Close #File2Num
ElseIf kind = 5 Then
Open FileName For Output As File2Num
Print #File2Num, text
Close #File2Num
End If
End Sub

Public Function ReadAllText(ByRef FileName, kind As Integer)
FileNum = FreeFile
If kind = 1 Then
Open FileName For Binary As FileNum
Get #FileNum, , text
Close #FileNum
ElseIf kind = 2 Then
Open FileName For Random As FileNum
Get #FileNum, , text
Close #FileNum
ElseIf kind = 3 Then
Open FileName For Append As FileNum
Input #FileNum, text
Close #FileNum
ElseIf kind = 4 Then
Open FileName For Input As FileNum
Input #FileNum, text
Close #FileNum
ElseIf kind = 5 Then
Open FileName For Input As #FileNum
text = Input(LOF(FileNum), #FileNum)
Close #FileNum
ElseIf kind = 6 Then
Open FileName For Binary As #FileNum
text = Input(LOF(FileNum), #FileNum)
Close #FileNum

End If
ReadAllText = text
End Function
Public Function readLine(ByRef strFilePath, ByRef nLine) As String
Dim NextLine As String
Dim N As Integer
FileNum = FreeFile
Open strFilePath For Input As FileNum
Do Until EOF(FileNum)
Line Input #FileNum, NextLine
N = N + 1
If N = nLine Then readLine = NextLine
Loop
Close
End Function


Public Function KnowNumLine(ByRef strFilePath)
Dim NumLine As Integer
NumLine = 0
FileNum = FreeFile
Open strFilePath For Input As FileNum
Do Until EOF(FileNum)
NumLine = NumLine + 1
Line Input #FileNum, X
Loop
Close
KnowNumLine = NumLine
End Function

Public Sub ChangeLine(strFilePath, NummberLine, Thechange)
Dim KNum, liberation, FFFFF
KNum = KnowNumLine(strFilePath)
FFFFF = Empty
For N = 1 To KNum
'liberation = readLine(strFilePath, n)
If NummberLine <> N Then
liberation = readLine(strFilePath, N)
FFFFF = FFFFF & liberation & vbNewLine
'WriteAllText strFilePath & "_tmp", liberation, 3
Else
'listone.AddItem (Thechange)
FFFFF = FFFFF & Thechange & vbNewLine
'WriteAllText strFilePath & "_tmp", Thechange, 3
End If
Next N
FFFFF = Left(FFFFF, Len(FFFFF) - 2)
Kill strFilePath
WriteAllText strFilePath, FFFFF, 3
'Name strFilePath & "_tmp" As strFilePath
End Sub

Public Sub ChangeLine2(strFilePath As String, NummberLine, Thechange)
Dim KNum, liberation, SySim, Lio, KK
KNum = KnowNumLine(strFilePath)
SySim = OpenFile(strFilePath)
Dim A, B, C
If InStr(1, SySim, readLine(strFilePath, NummberLine)) Then
KK = readLine(strFilePath, NummberLine)
SySim = Replace(SySim, KK, Thechange, , , vbTextCompare)
WriteAllText strFilePath & "_tmp", SySim, 3
End If
Kill strFilePath
Name strFilePath & "_tmp" As strFilePath
End Sub



Public Sub DeleteLine(strFilePath, NummberLine)
Dim KNum, liberation
KNum = KnowNumLine(strFilePath)
For N = 1 To KNum
'liberation = readLine(strFilePath, n)
If NummberLine <> N Then
liberation = readLine(strFilePath, N)
WriteAllText strFilePath & "_tmp", liberation, 3
Else
'listone.AddItem (Thechange)
'WriteAllText strFilePath & "_tmp", Thechange, 3
End If
Next N
Kill strFilePath
Name strFilePath & "_tmp" As strFilePath
End Sub

Public Sub AddUnderLine(strFilePath, NummberLine, text)
Dim KNum, liberation
KNum = KnowNumLine(strFilePath)
For N = 1 To KNum
'liberation = readLine(strFilePath, n)
If NummberLine <> N Then
liberation = readLine(strFilePath, N)
WriteAllText strFilePath & "_tmp", liberation, 3
Else
liberation = readLine(strFilePath, N)
WriteAllText strFilePath & "_tmp", liberation, 3
WriteAllText strFilePath & "_tmp", text, 3
End If
Next N
Kill strFilePath
Name strFilePath & "_tmp" As strFilePath
End Sub

Public Sub AddUpLine(strFilePath, NummberLine, text)
Dim KNum, liberation
KNum = KnowNumLine(strFilePath)
For N = 1 To KNum
'liberation = readLine(strFilePath, n)
If NummberLine <> N + 1 Then
liberation = readLine(strFilePath, N)
WriteAllText strFilePath & "_tmp", liberation, 3
Else
liberation = readLine(strFilePath, N)
WriteAllText strFilePath & "_tmp", liberation, 3
WriteAllText strFilePath & "_tmp", text, 3
End If
Next N
Kill strFilePath
Name strFilePath & "_tmp" As strFilePath
End Sub

Public Sub AddThisLine(strFilePath, NummberLine, text)
Dim KNum, liberation
KNum = KnowNumLine(strFilePath)
For N = 1 To KNum
'liberation = readLine(strFilePath, n)
If NummberLine <> N + 1 Then
liberation = readLine(strFilePath, N)
WriteAllText strFilePath & "_tmp", liberation, 3
Else
liberation = readLine(strFilePath, N)
WriteAllText strFilePath & "_tmp", liberation, 3
WriteAllText strFilePath & "_tmp", text, 3
End If
Next N
Kill strFilePath
Name strFilePath & "_tmp" As strFilePath
End Sub


Public Function FileExists(strPath As String) As Boolean
    strPath = Trim(strPath)
    If strPath = "" Then
        FileExists = False
        Exit Function
    End If
  FileExists = Len(Dir(strPath)) <> 0
End Function
