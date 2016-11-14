VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H00000000&
   Caption         =   "Editor"
   ClientHeight    =   7200
   ClientLeft      =   165
   ClientTop       =   525
   ClientWidth     =   9255
   Icon            =   "FrmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   9255
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Height          =   6975
      Left            =   120
      ScaleHeight     =   6915
      ScaleWidth      =   7875
      TabIndex        =   1
      Top             =   120
      Width           =   7935
      Begin VB.PictureBox Floor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   0
         Left            =   0
         ScaleHeight     =   705
         ScaleWidth      =   705
         TabIndex        =   6
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Tools"
      ForeColor       =   &H00FFFFFF&
      Height          =   6135
      Left            =   8160
      TabIndex        =   0
      Top             =   120
      Width           =   975
      Begin VB.OptionButton Tool 
         Height          =   735
         Index           =   5
         Left            =   120
         Picture         =   "FrmMain.frx":DC7B
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3840
         Width           =   735
      End
      Begin VB.OptionButton Tool 
         Height          =   735
         Index           =   4
         Left            =   120
         Picture         =   "FrmMain.frx":F72D
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3120
         Width           =   735
      End
      Begin VB.OptionButton Tool 
         Height          =   735
         Index           =   3
         Left            =   120
         Picture         =   "FrmMain.frx":1126F
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2400
         Width           =   735
      End
      Begin VB.OptionButton Tool 
         Height          =   735
         Index           =   2
         Left            =   120
         Picture         =   "FrmMain.frx":12DB1
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1680
         Width           =   735
      End
      Begin VB.OptionButton Tool 
         Height          =   735
         Index           =   1
         Left            =   120
         Picture         =   "FrmMain.frx":14863
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   960
         Width           =   735
      End
      Begin VB.OptionButton Tool 
         Height          =   735
         Index           =   0
         Left            =   120
         Picture         =   "FrmMain.frx":16315
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Points"
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   5520
         Width           =   495
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   5760
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Boxes"
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   4920
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   5160
         Width           =   975
      End
   End
   Begin VB.Menu MMFile 
      Caption         =   "File"
      Begin VB.Menu Clear 
         Caption         =   "New"
      End
      Begin VB.Menu OPi 
         Caption         =   "Open"
      End
      Begin VB.Menu MLoad 
         Caption         =   "Load/ Save"
      End
      Begin VB.Menu MSave 
         Caption         =   "Save"
      End
      Begin VB.Menu MAdd 
         Caption         =   "Add"
         Visible         =   0   'False
      End
      Begin VB.Menu MRun 
         Caption         =   "Run"
      End
   End
   Begin VB.Menu CK 
      Caption         =   "Check"
      Begin VB.Menu CA 
         Caption         =   "Check ALL"
      End
      Begin VB.Menu CT 
         Caption         =   "Check this"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Dim GG

Private Sub CA_Click()
Dim NN
FrmList.Show

Do Until FrmList.Timer1.Enabled = False
DoEvents
Loop

For NN = 0 To FrmList.List1.ListCount - 1
FrmList.List1.ListIndex = NN
FrmList.Command1 = True
If CheckGameIS = False Then Exit Sub
DoEvents
Next

End Sub

Private Sub Clear_Click()
Fileef = ""
Dim N
For N = 0 To Floor.UBound
Floor(N).Picture = LoadPicture("")
Next
End Sub

Private Sub CT_Click()
MsgBox "The game Is Correct :- " & CheckGameIS
End Sub

Private Sub Floor_Click(Index As Integer)
With Tool(0)
If .Value = True Then
Floor(Index).Picture = .Picture
End If
End With
With Tool(1)
If .Value = True Then
Floor(Index).Picture = .Picture
End If
End With
With Tool(2)
If .Value = True Then
Floor(Index).Picture = .Picture
End If
End With

With Tool(3)
If .Value = True Then
Floor(Index).Picture = .Picture
End If
End With

With Tool(4)
If .Value = True Then
Floor(Index).Picture = .Picture
End If
End With

With Tool(5)
If .Value = True Then
Floor(Index).Picture = .Picture
End If
End With
End Sub

Private Sub Floor_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Floor_MouseMove(Index, Button, Shift, X, Y)

End Sub

Private Sub Floor_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Dim N
For N = 0 To Floor.UBound
If Tool(1).Value = True Then
If Floor(N).Picture = Tool(1).Picture And N <> Index Then Floor(N).Picture = LoadPicture("")
If Floor(N).Picture = Tool(5).Picture And N <> Index Then Floor(N).Picture = LoadPicture("")

End If
If Tool(5).Value = True Then
If Floor(N).Picture = Tool(1).Picture And N <> Index Then Floor(N).Picture = LoadPicture("")

If Floor(N).Picture = Tool(5).Picture And N <> Index Then Floor(N).Picture = LoadPicture("")
End If
Next
End If

If Button = 1 Then Floor_Click (Index)
If Button = 2 Then Floor(Index).Picture = LoadPicture("")


ReleaseCapture


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then MRun_Click
End Sub

Private Sub Form_Load()
Dim N, M, Num, FloorX

On Error Resume Next

For N = 0 To Screen.Width Step 720
For M = 0 To Screen.Height Step 720
Num = Floor.Count
Load Me.Floor(Num)
Floor(Num).Move N, M

Floor(Num).Visible = True
DoEvents
Next
Next

Dim cc

'CC = 1
'Do Until FileExists(App.Path & "\Data\LEVEL" & CC) = False
'CC = CC * 1 + 1 * 1
'FrmList.List1.AddItem "LEVEL" & CC
'DoEvents
'Loop

End Sub

Private Sub Form_Resize()
If FrmMain.Height <= 3630 Then
FrmMain.Height = 3630
Else
'Form_Load
End If

Frame1.Top = 120
Frame1.Left = Me.Width - 1215
Picture1.Height = Height - 1020
Picture1.Width = Width - 1440

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub MAdd_Click()
'On Error Resume Next
Dim N, cc

cc = 1

Do Until FileExists(App.Path & "\Data\LEVEL" & cc) = False
cc = cc * 1 + 1 * 1
DoEvents
Loop

Fileef = App.Path & "\Data\LEVEL" & cc

SaveFile FileExists(App.Path & "\Data\LEVEL" & cc), ""


For N = 0 To Floor.UBound
'\\\\\\\\\
With Tool(0)
If Floor(N).Picture = .Picture Then
 WriteAllText App.Path & "\Data\LEVEL" & cc, "Wall*" & Floor(N).Top & "*" & Floor(N).Left, 3
End If
End With
'\\\\\\\\\
With Tool(1)
If Floor(N).Picture = .Picture Then

WriteAllText App.Path & "\Data\LEVEL" & cc, "Man*" & Floor(N).Top & "*" & Floor(N).Left, 3
End If
End With

With Tool(5)
If Floor(N).Picture = .Picture Then
WriteAllText App.Path & "\Data\LEVEL" & cc, "Man*" & Floor(N).Top & "*" & Floor(N).Left, 3
WriteAllText App.Path & "\Data\LEVEL" & cc, "XPlace*" & Floor(N).Top & "*" & Floor(N).Left, 3
End If
End With
'Tool(5)
'\\\\\\\\\
'\\\\\\\\\
With Tool(2)
If Floor(N).Picture = .Picture Then

 WriteAllText App.Path & "\Data\LEVEL" & cc, "Box*" & Floor(N).Top & "*" & Floor(N).Left, 3
End If
End With
'\\\\\\\\\
'\\\\\\\\\
With Tool(3)
If Floor(N).Picture = .Picture Then

Call WriteAllText(App.Path & "\Data\LEVEL" & cc, "XPlace*" & Floor(N).Top & "*" & Floor(N).Left, 3)
End If
End With
'\\\\\\\\\



Next
End Sub

Private Sub MLoad_Click()
FrmList.Show
End Sub

Private Sub MRun_Click()
On Error Resume Next
Dim N

SaveFile "C:\LWE", ""

For N = 0 To Floor.UBound
'\\\\\\\\\
With FrmMain.Tool(0)
If Floor(N).Picture = .Picture Then
 WriteAllText "C:\LWE", "Wall*" & Floor(N).Top & "*" & Floor(N).Left, 3
End If
End With
'\\\\\\\\\
With FrmMain.Tool(1)
If Floor(N).Picture = .Picture Then

WriteAllText "C:\LWE", "Man*" & Floor(N).Top & "*" & Floor(N).Left, 3
End If
End With


With FrmMain.Tool(5)
If Floor(N).Picture = .Picture Then
Call WriteAllText("C:\LWE", "XPlace*" & Floor(N).Top & "*" & Floor(N).Left, 3)
WriteAllText "C:\LWE", "Man*" & Floor(N).Top & "*" & Floor(N).Left, 3
End If
End With
'\\\\\\\\\
'\\\\\\\\\
With FrmMain.Tool(2)
If Floor(N).Picture = .Picture Then

 WriteAllText "C:\LWE", "Box*" & Floor(N).Top & "*" & Floor(N).Left, 3
End If
End With
'\\\\\\\\\
'\\\\\\\\\
With FrmMain.Tool(3)
If Floor(N).Picture = .Picture Then

Call WriteAllText("C:\LWE", "XPlace*" & Floor(N).Top & "*" & Floor(N).Left, 3)
End If
End With
'\\\\\\\\\
'\\\\\\\\\
With FrmMain.Tool(4)
If Floor(N).Picture = .Picture Then

Call WriteAllText("C:\LWE", "XPlace*" & Floor(N).Top & "*" & Floor(N).Left, 3)
Call WriteAllText("C:\LWE", "Box*" & Floor(N).Top & "*" & Floor(N).Left, 3)

End If
End With
'\\\\\\\\\


Next

FrmRun.Show
FrmRun.Form_Load
End Sub


Private Sub MSave_Click()
On Error Resume Next

If Fileef = "" Then MAdd_Click

Dim N


SaveFile Fileef, ""

For N = 0 To FrmMain.Floor.UBound
'\\\\\\\\\
With FrmMain.Tool(0)
If FrmMain.Floor(N).Picture = .Picture Then
 WriteAllText Fileef, "Wall*" & FrmMain.Floor(N).Top & "*" & FrmMain.Floor(N).Left, 3
End If
End With
'\\\\\\\\\
With FrmMain.Tool(1)
If FrmMain.Floor(N).Picture = .Picture Then

WriteAllText Fileef, "Man*" & FrmMain.Floor(N).Top & "*" & FrmMain.Floor(N).Left, 3
End If
End With


With Tool(5)
If Floor(N).Picture = .Picture Then
WriteAllText Fileef, "Man*" & Floor(N).Top & "*" & Floor(N).Left, 3
WriteAllText Fileef, "XPlace*" & Floor(N).Top & "*" & Floor(N).Left, 3
End If
End With
'\\\\\\\\\
'\\\\\\\\\
With FrmMain.Tool(2)
If FrmMain.Floor(N).Picture = .Picture Then

 WriteAllText Fileef, "Box*" & FrmMain.Floor(N).Top & "*" & FrmMain.Floor(N).Left, 3
End If
End With
'\\\\\\\\\
'\\\\\\\\\
With FrmMain.Tool(3)
If FrmMain.Floor(N).Picture = .Picture Then

Call WriteAllText(Fileef, "XPlace*" & FrmMain.Floor(N).Top & "*" & FrmMain.Floor(N).Left, 3)
End If
End With
'\\\\\\\\\
With FrmMain.Tool(4)
If FrmMain.Floor(N).Picture = .Picture Then

Call WriteAllText(Fileef, "XPlace*" & FrmMain.Floor(N).Top & "*" & FrmMain.Floor(N).Left, 3)
Call WriteAllText(Fileef, "Box*" & FrmMain.Floor(N).Top & "*" & FrmMain.Floor(N).Left, 3)

End If
End With
With FrmMain.Tool(5)
If FrmMain.Floor(N).Picture = .Picture Then

Call WriteAllText(Fileef, "XPlace*" & FrmMain.Floor(N).Top & "*" & FrmMain.Floor(N).Left, 3)
WriteAllText Fileef, "Man*" & FrmMain.Floor(N).Top & "*" & FrmMain.Floor(N).Left, 3

End If
End With



Next
End Sub

Private Sub OPi_Click()
Dim A
A = ShowOpen(Me)
If A = "" Then Exit Sub
Dim E() As String, E2() As String, TGG, TGG2, GG

For GG = 0 To FrmMain.Floor.UBound
FrmMain.Floor(GG).Picture = LoadPicture("")
Next

Fileef = A

Dim Y
A = ReadAllText(Fileef, 5)

Dim Q, W, R
Y = 0
Q = 0
W = 0
R = 0


TGG = 10000
TGG2 = 10000
E = Split(A, vbNewLine)
For N = 0 To UBound(E) - 1
E2 = Split(E(N), "*")
If E2(0) = "Man" Then
'\
For GG = 0 To FrmMain.Floor.UBound
If FrmMain.Floor(GG).Top = E2(1) And FrmMain.Floor(GG).Left = E2(2) Then
If FrmMain.Floor(GG).Picture = LoadPicture("") Then
FrmMain.Floor(GG).Picture = FrmMain.Tool(1).Picture
ElseIf FrmMain.Floor(GG).Picture = FrmMain.Tool(3).Picture Then
FrmMain.Floor(GG).Picture = FrmMain.Tool(5).Picture
End If
End If
Next
'\
End If
If E2(0) = "Wall" Then
'\
For GG = 0 To FrmMain.Floor.UBound
If FrmMain.Floor(GG).Top = E2(1) And FrmMain.Floor(GG).Left = E2(2) Then
FrmMain.Floor(GG).Picture = FrmMain.Tool(0).Picture
End If
Next
End If
If E2(0) = "XPlace" Then
'\
For GG = 0 To FrmMain.Floor.UBound
If FrmMain.Floor(GG).Top = E2(1) And FrmMain.Floor(GG).Left = E2(2) Then
If FrmMain.Floor(GG).Picture = FrmMain.Tool(3).Picture Then
FrmMain.Floor(GG).Picture = FrmMain.Tool(4).Picture
ElseIf FrmMain.Floor(GG).Picture = LoadPicture("") Then
FrmMain.Floor(GG).Picture = FrmMain.Tool(3).Picture
ElseIf FrmMain.Floor(GG).Picture = FrmMain.Tool(1).Picture Then
FrmMain.Floor(GG).Picture = FrmMain.Tool(5).Picture
End If
End If
Next
End If
If E2(0) = "Box" Then
'\
For GG = 0 To FrmMain.Floor.UBound
If FrmMain.Floor(GG).Top = E2(1) And FrmMain.Floor(GG).Left = E2(2) Then
If FrmMain.Floor(GG).Picture = LoadPicture("") Then
FrmMain.Floor(GG).Picture = FrmMain.Tool(2).Picture

ElseIf FrmMain.Floor(GG).Picture = FrmMain.Tool(3).Picture Then
FrmMain.Floor(GG).Picture = FrmMain.Tool(4).Picture
End If

End If
Next
End If

Next
Waitm 100
End Sub
Public Function Waitm(ByVal TimeToWait As Long) 'Time In seconds
Dim EndTime As Long
EndTime = GetTickCount + TimeToWait  '* 1000 Cause u give seconds and GetTickCount uses Milliseconds
Do Until GetTickCount > EndTime
DoEvents
Loop
End Function


Public Function CheckGameIS() As Boolean
Dim N, BOXX, POINTT
BOXX = 0
POINTT = 0
For N = 0 To Floor.UBound
If Floor(N).Picture = Tool(2).Picture Then BOXX = BOXX * 1 + 1 * 1
If Floor(N).Picture = Tool(3).Picture Then POINTT = POINTT * 1 + 1 * 1
If Floor(N).Picture = Tool(4).Picture Then
BOXX = BOXX * 1 + 1 * 1
POINTT = POINTT * 1 + 1 * 1
End If

If Floor(N).Picture = Tool(5).Picture Then POINTT = POINTT * 1 + 1 * 1


Next
Label1 = BOXX
Label3 = POINTT
If POINTT = BOXX Then
CheckGameIS = True
Else
CheckGameIS = False
End If

End Function

