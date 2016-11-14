VERSION 5.00
Begin VB.Form FrmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Boxes"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6675
   Icon            =   "FrmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmMain.frx":E8EB
   ScaleHeight     =   3915
   ScaleWidth      =   6675
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   840
      Top             =   960
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3240
      Top             =   0
   End
   Begin VB.Timer AreYouWin 
      Interval        =   1000
      Left            =   840
      Top             =   1800
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   1935
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2160
      Top             =   480
   End
   Begin VB.Image Old 
      Height          =   735
      Left            =   600
      Picture         =   "FrmMain.frx":F359
      Stretch         =   -1  'True
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Man 
      Height          =   735
      Left            =   5520
      Picture         =   "FrmMain.frx":1239B
      Top             =   120
      Width           =   720
   End
   Begin VB.Image RP 
      Height          =   735
      Left            =   2880
      Picture         =   "FrmMain.frx":13F6D
      Top             =   2400
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image UP 
      Height          =   735
      Left            =   2880
      Picture         =   "FrmMain.frx":15B3F
      Top             =   1560
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image LP 
      Height          =   735
      Left            =   3720
      Picture         =   "FrmMain.frx":17711
      Top             =   2400
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image DP 
      Height          =   735
      Left            =   3720
      Picture         =   "FrmMain.frx":192E3
      Top             =   1560
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   1
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   1
      Left            =   720
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   1
      Left            =   1440
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      Height          =   735
      Index           =   0
      Left            =   0
      Picture         =   "FrmMain.frx":1AEB5
      Stretch         =   -1  'True
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Floor 
      Height          =   735
      Index           =   0
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image I 
      Height          =   735
      Left            =   3360
      Picture         =   "FrmMain.frx":1C9F7
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image II 
      Height          =   735
      Left            =   4080
      Picture         =   "FrmMain.frx":1FA39
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image III 
      Height          =   735
      Left            =   4800
      Picture         =   "FrmMain.frx":22A7B
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      Height          =   735
      Index           =   0
      Left            =   720
      Picture         =   "FrmMain.frx":25ABD
      Stretch         =   -1  'True
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      Height          =   735
      Index           =   0
      Left            =   1440
      Picture         =   "FrmMain.frx":28AFF
      Stretch         =   -1  'True
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim REE, REE2, REE3, REE3REE3, NumLevel As String, TheRestart, PPOOOX, PPOOOY
Dim BuuB As String
Public Sub CenterForm(frm As Form)
  frm.Left = (Screen.Width - frm.Width) / 2
  frm.Top = (Screen.Height - frm.Height) / 2
End Sub
Private Sub AreYouWin_Timer()

REE = ""
REE2 = ""
REE3 = 0
REE3REE3 = 0

REE3REE3 = XPlace.UBound - 1

For M = 2 To XPlace.UBound
Label1 = M



REE = XPlace(M).Top & ":" & XPlace(M).Left
For N = 2 To Boxes.UBound
If Boxes(N).BorderStyle = 1 Then Exit For
Label2 = N
REE2 = Boxes(N).Top & ":" & Boxes(N).Left

If REE = REE2 Then
REE3 = REE3 * 1 + 1 * 1
Label3 = REE3
End If

Next


Next


If REE3 = REE3REE3 And REE3REE3 <> 0 Then
'Unload Me


NumLevel = NumLevel * 1 + 1 * 1
READGMAE NumLevel

BuuB = ""


For N = 0 To Boxes.UBound
BuuB = BuuB & "Boxes" & "::" & N & "::" & Boxes(N).Top & "::" & Boxes(N).Left & vbNewLine
Next
BuuB = BuuB & "Man" & "::" & Man.Top & "::" & Man.Left & vbNewLine
BuuB = BuuB & "|||" & vbNewLine
Call SaveFile(App.Path & "\GamesUndo.ini", BuuB)

'Me.Hide
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyUp Then
WalkUP
End If
If KeyCode = vbKeyDown Then
WalkDown
End If
If KeyCode = vbKeyRight Then
WalkRight
End If
If KeyCode = vbKeyLeft Then
WalkLeft
End If

If KeyCode = vbKeyZ And Shift = vbCtrlMask Then
BuuB = OpenFile(App.Path & "\GamesUndo.ini")

Dim hhh() As String, hhh2() As String, hhh3() As String, N
hhh = Split(BuuB, "|||")
On Error Resume Next
hhh2 = Split(hhh(UBound(hhh)), vbNewLine)
For N = 0 To UBound(hhh2)
If hhh2(N) <> "" Then
hhh3 = Split(hhh2(N), "::")
If hhh3(0) = "Man" Then Man.Top = hhh3(1): Man.Left = hhh3(2)
If hhh3(0) = "Boxes" Then Boxes(hhh3(1)).Top = hhh3(2): Boxes(hhh3(1)).Left = hhh3(3)
End If
Next
ReDim Preserve hhh(LBound(hhh) To UBound(hhh) - 1) As String
BuuB = Join(hhh, "|||")
Call SaveFile(App.Path & "\GamesUndo.ini", BuuB)

End If

If KeyCode = vbKeyEscape Then Me.Hide: Unload Me
End Sub



Private Sub Rs_Click()
Dim uh() As String
uh = Split(TheRestart, ":")
Man.Top = uh(0)
Man.Left = uh(1)

For N = 2 To UBound(uh) - 1 Step 2
Boxes((N / 2) - 1).Top = uh(N)
Boxes((N / 2) - 1).Left = uh(N + 1)
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub Timer1_Timer()
'Lman.Move Man.Left + (Man.Width / 4), Man.Top + (Man.Height / 3)
'Lman.Caption = "Man"
End Sub

Public Sub WalkUP()
Man.Picture = UP.Picture
Dim BNUm As Integer, NM
'''''''''''''''''''''''''''''''''''''Up'''''''''''''''''''''''''''''''''''
BNUm = 1

For N = 0 To Boxes.UBound
For NM = 0 To Wall.UBound
If Wall(NM).Top = Man.Top - Onelong And Wall(NM).Left = Man.Left Then _
BNUm = BNUm * 1 + 1 * 1
Next

If Boxes(N).Top = Man.Top - Onelong And Boxes(N).Left = Man.Left Then
BNUm = BNUm * 1 + 1 * 1

For NM = 0 To Wall.UBound
If Wall(NM).Top = Man.Top - Onelong And Wall(NM).Left = Man.Left Then _
BNUm = BNUm * 1 + 1 * 1
If Wall(NM).Top = Man.Top - Onelong - Onelong And Wall(NM).Left = Man.Left Then _
BNUm = BNUm * 1 + 1 * 1

Next

For M = 0 To Boxes.UBound
If M <> N Then
If Boxes(M).Top = Man.Top - Onelong - Onelong And Boxes(M).Left = Man.Left Then BNUm = BNUm * 1 + 1 * 1
End If
Next

End If

If Boxes(N).Top = Man.Top - Onelong And Boxes(N).Left = Man.Left And BNUm = 2 Then
Boxes(N).Top = Man.Top - Onelong - Onelong
Man.Top = Man.Top - Onelong

'Exit For
Else
If BNUm = 1 And N = Boxes.UBound Then
Man.Top = Man.Top - Onelong
Exit For
End If
End If

Next
oo
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub WalkDown()
Man.Picture = DP.Picture
Dim BNUm As Integer, NM
'''''''''''''''''''''''''''''''''''''Down'''''''''''''''''''''''''''''''''''
BNUm = 1

For N = 0 To Boxes.UBound
For NM = 0 To Wall.UBound
If Wall(NM).Top = Man.Top + Onelong And Wall(NM).Left = Man.Left Then _
BNUm = BNUm * 1 + 1 * 1

Next
If Boxes(N).Top = Man.Top + Onelong And Boxes(N).Left = Man.Left Then
BNUm = BNUm * 1 + 1 * 1

For NM = 0 To Wall.UBound
If Wall(NM).Top = Man.Top + Onelong And Wall(NM).Left = Man.Left Then _
BNUm = BNUm * 1 + 1 * 1
If Wall(NM).Top = Man.Top + Onelong + Onelong And Wall(NM).Left = Man.Left Then _
BNUm = BNUm * 1 + 1 * 1

Next

For M = 0 To Boxes.UBound
If M <> N Then
If Boxes(M).Top = Man.Top + Onelong + Onelong And Boxes(M).Left = Man.Left Then BNUm = BNUm * 1 + 1 * 1
End If
Next

End If

If Boxes(N).Top = Man.Top + Onelong And Boxes(N).Left = Man.Left And BNUm = 2 Then
Boxes(N).Top = Man.Top + Onelong + Onelong

Man.Top = Man.Top + Onelong
'Exit For
Else
If BNUm = 1 And N = Boxes.UBound Then
Man.Top = Man.Top + Onelong
Exit For
End If
End If

Next
oo
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''j
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub WalkRight()
Man.Picture = RP.Picture
Dim BNUm As Integer, NM
'''''''''''''''''''''''''''''''''''''Right'''''''''''''''''''''''''''''''''''
BNUm = 1

For N = 0 To Boxes.UBound
For NM = 0 To Wall.UBound
If Wall(NM).Top = Man.Top And Wall(NM).Left = Man.Left + Onelong Then _
BNUm = BNUm * 1 + 1 * 1

Next
If Boxes(N).Top = Man.Top And Boxes(N).Left = Man.Left + Onelong Then
BNUm = BNUm * 1 + 1 * 1

For NM = 0 To Wall.UBound
If Wall(NM).Top = Man.Top And Wall(NM).Left = Man.Left + Onelong Then _
BNUm = BNUm * 1 + 1 * 1
If Wall(NM).Top = Man.Top And Wall(NM).Left = Man.Left + Onelong + Onelong Then _
BNUm = BNUm * 1 + 1 * 1

Next

For M = 0 To Boxes.UBound
If M <> N Then
If Boxes(M).Top = Man.Top And Boxes(M).Left = Man.Left + Onelong + Onelong Then BNUm = BNUm * 1 + 1 * 1
End If
Next

End If

If Boxes(N).Top = Man.Top And Boxes(N).Left = Man.Left + Onelong And BNUm = 2 Then
Boxes(N).Left = Man.Left + Onelong + Onelong
Man.Left = Man.Left + Onelong
'Exit For
Else
If BNUm = 1 And N = Boxes.UBound Then
Man.Left = Man.Left + Onelong
Exit For
End If
End If

Next
oo
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''j
Public Sub WalkLeft()
Man.Picture = LP.Picture
Dim BNUm As Integer, NM
'''''''''''''''''''''''''''''''''''''Right'''''''''''''''''''''''''''''''''''
BNUm = 1

For N = 0 To Boxes.UBound

For NM = 0 To Wall.UBound
If Wall(NM).Top = Man.Top And Wall(NM).Left = Man.Left - Onelong Then _
BNUm = BNUm * 1 + 1 * 1

Next

If Boxes(N).Top = Man.Top And Boxes(N).Left = Man.Left - Onelong Then
BNUm = BNUm * 1 + 1 * 1

For NM = 0 To Wall.UBound
If Wall(NM).Top = Man.Top And Wall(NM).Left = Man.Left - Onelong - Onelong Then _
BNUm = BNUm * 1 + 1 * 1

Next

For M = 0 To Boxes.UBound
If M <> N Then
If Boxes(M).Top = Man.Top And Boxes(M).Left = Man.Left - Onelong - Onelong Then BNUm = BNUm * 1 + 1 * 1
End If
Next

End If

If Boxes(N).Top = Man.Top And Boxes(N).Left = Man.Left - Onelong And BNUm = 2 Then
Boxes(N).Left = Man.Left - Onelong - Onelong
Man.Left = Man.Left - Onelong
Exit For
Else
If BNUm = 1 And N = Boxes.UBound Then
Man.Left = Man.Left - Onelong
Exit For
End If
End If

Next
oo
End Sub



Public Sub DrawMan(Value)
If Value = 1 Then
'SMan.Top = Man.Top
'SMan.Left = Man.Left - (SMan.Width / 3)
End If

End Sub



Sub Form_Load()
Icon = FrmMain.Icon
NumLevel = 1
If FileExists(App.Path & "\InLevel.TXT") = True Then
NumLevel = Trim(RemoveChar(RemoveAfterChar(OpenFile(App.Path & "\InLevel.TXT"), "'"), All_Marks))
End If

ADSS
Timer2.Enabled = True
'READGMAE (NumLevel)

PPOOOX = 0
PPOOOY = 0
End Sub

Public Sub READGMAE(LEVEL As String)
'On Error Resume Next
ADSS
Dim A, Y
On Error GoTo ff:
SaveFile App.Path & "\InLevel.TXT", LEVEL

For N = 2 To Boxes.UBound
Unload Boxes(N)
Next
For N = 2 To Wall.UBound
Unload Wall(N)
Next
For N = 2 To XPlace.UBound
Unload XPlace(N)
Next


A = ReadAllText(App.Path & "\Data\Level" & LEVEL, 5)
Dim Q, W, r
Y = 0
Q = 0
W = 0
r = 0

Dim E() As String, E2() As String, TGG, TGG2, num
TGG = 10000
TGG2 = 10000
E = Split(A, vbNewLine)
For N = 0 To UBound(E) - 1
E2 = Split(E(N), "*")
If E2(0) = "Man" Then
Man.Top = E2(1)
Man.Left = E2(2)
End If
If E2(0) = "Wall" Then

num = Me.Wall.UBound + 1

Load Me.Wall(num)
Me.Wall(num).ZOrder
'If Abs(FrmMain.Width) < Abs(E2(1) + 735) Then FrmMain.Width = (E2(1) + 735)
'If Abs(FrmMain.Height) < Abs(E2(2) + 735) Then FrmMain.Height = (E2(2) + 735)
If Abs(TGG) > Abs(E2(1)) Then TGG = (E2(1))
If Abs(TGG2) > Abs(E2(2)) Then TGG2 = (E2(2))

'TGG
Me.Wall(num).Top = E2(1)
Me.Wall(num).Left = E2(2)
Me.Wall(num).Visible = True
Me.Wall(num).Picture = Me.Wall(0)

If PPOOOY < Wall(num).Top Then PPOOOY = Wall(num).Top
If PPOOOX < Wall(num).Left Then PPOOOX = Wall(num).Left

Wall(num).BorderStyle = 0

End If

If E2(0) = "XPlace" Then
num = Me.XPlace.UBound + 1
Load Me.XPlace(num)
Me.XPlace(num).ZOrder
XPlace(num).Visible = True

XPlace(num).Top = E2(1)
XPlace(num).Left = E2(2)

Me.XPlace(num).Picture = Me.XPlace(0)
XPlace(num).BorderStyle = 0

End If

If E2(0) = "Box" Then

num = Me.Boxes.UBound + 1

Load Me.Boxes(num)


Boxes(num).Top = E2(1)
Boxes(num).Left = E2(2)
If PPOOOY < Boxes(num).Top Then PPOOOY = Boxes(num).Top
If PPOOOX < Boxes(num).Left Then PPOOOX = Boxes(num).Left


Boxes(num).ZOrder
Boxes(num).Visible = True

Me.Boxes(num).Picture = Me.Boxes(0)
Boxes(num).BorderStyle = 0

End If

Next



For N = 0 To XPlace.UBound
If XPlace(N).BorderStyle = 1 Then Exit For
XPlace(N).Top = XPlace(N).Top - TGG
XPlace(N).Left = XPlace(N).Left - TGG2
If PPOOOY < XPlace(N).Top Then PPOOOY = XPlace(N).Top
If PPOOOX < XPlace(N).Left Then PPOOOX = XPlace(N).Left

Next


For N = 0 To Boxes.UBound
If Boxes(N).BorderStyle = 1 Then Exit For
Boxes(N).Top = Boxes(N).Top - TGG
Boxes(N).Left = Boxes(N).Left - TGG2
If PPOOOY < Boxes(N).Top Then PPOOOY = Boxes(N).Top
If PPOOOX < Boxes(N).Left Then PPOOOX = Boxes(N).Left

Next

For N = 0 To Wall.UBound
If Wall(N).BorderStyle = 1 Then Exit For
Wall(N).Top = Wall(N).Top - TGG

Wall(N).Left = Wall(N).Left - TGG2
Next


Man.Top = Man.Top - TGG
Man.Left = Man.Left - TGG2
Man.ZOrder



PPOOOY = 0
PPOOOX = 0
For N = 0 To Boxes.UBound
Boxes(N).Top = Boxes(N).Top - TGG
Boxes(N).Left = Boxes(N).Left - TGG2
If PPOOOY < Boxes(N).Top Then PPOOOY = Boxes(N).Top
If PPOOOX < Boxes(N).Left Then PPOOOX = Boxes(N).Left

Boxes(N).ZOrder
Next

For N = 0 To Wall.UBound
Wall(N).Top = Wall(N).Top - TGG
Wall(N).Left = Wall(N).Left - TGG2
If PPOOOY < Wall(N).Top Then PPOOOY = Wall(N).Top
If PPOOOX < Wall(N).Left Then PPOOOX = Wall(N).Left

Next
For N = 0 To XPlace.UBound
XPlace(N).Top = XPlace(N).Top - TGG
XPlace(N).Left = XPlace(N).Left - TGG2
If PPOOOY < XPlace(N).Top Then PPOOOY = XPlace(N).Top
If PPOOOX < XPlace(N).Left Then PPOOOX = XPlace(N).Left

Next
Me.Height = PPOOOY + 1135
Me.Width = PPOOOX + 830
CenterForm Me
Man.ZOrder


For N = 0 To Boxes.UBound
BuuB = BuuB & "Boxes" & "::" & N & "::" & Boxes(N).Top & "::" & Boxes(N).Left & vbNewLine
Next
BuuB = BuuB & "Man" & "::" & Man.Top & "::" & Man.Left & vbNewLine
BuuB = BuuB & "|||" & vbNewLine
Call SaveFile(App.Path & "\GamesUndo.ini", BuuB)



Exit Sub
ff:
End
End Sub





Private Sub Timer2_Timer()
Unload Me
Me.Visible = True
PPOOOX = 55
PPOOOY = 55
NumLevel = 1
If FileExists(App.Path & "\InLevel.TXT") = True Then
NumLevel = Trim(RemoveChar(RemoveAfterChar(OpenFile(App.Path & "\InLevel.TXT"), "'"), All_Marks))
End If
READGMAE NumLevel

CenterForm Me
Timer2.Enabled = False
End Sub





Public Sub ADSS()
Dim num, FloorX
For N = 0 To Me.Width + 5000 Step 720
For M = 0 To Me.Height + 5000 Step 720
num = Floor.Count
Load Me.Floor(num)
Set FloorX = Me.Floor(num)
FloorX.Move N, M
FloorX.Picture = Me.Picture
FloorX.Visible = True
Next
Next
TheRestart = ""
TheRestart = Man.Top & ":" & Man.Left & ":"
For M = 0 To Boxes.UBound
TheRestart = TheRestart & Boxes(M).Top & ":" & Boxes(M).Left & ":"
Next
TheRestart = Left(TheRestart, Len(TheRestart) - 1)

End Sub

Public Sub oo()
BuuB = OpenFile(App.Path & "\GamesUndo.ini")
BuuB = BuuB & "|||" & vbNewLine
For N = 0 To Boxes.UBound
BuuB = BuuB & "Boxes" & "::" & N & "::" & Boxes(N).Top & "::" & Boxes(N).Left & vbNewLine
Next
BuuB = BuuB & "Man" & "::" & Man.Top & "::" & Man.Left & vbNewLine

Call SaveFile(App.Path & "\GamesUndo.ini", BuuB)
End Sub

