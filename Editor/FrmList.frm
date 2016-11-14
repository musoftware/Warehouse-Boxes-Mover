VERSION 5.00
Begin VB.Form FrmList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4335
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1320
      Top             =   2160
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "FrmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long
    
Private Sub Command1_Click()
Dim E() As String, E2() As String, TGG, TGG2, GG
For GG = 0 To FrmMain.Floor.UBound
FrmMain.Floor(GG).Picture = LoadPicture("")
Next

Fileef = App.Path & "\Data\" & List1.text

Dim A, Y
A = ReadAllText(App.Path & "\DATA\" & List1.text, 5)

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

If FrmMain.Floor(GG).Picture = FrmMain.Tool(2).Picture Then
FrmMain.Floor(GG).Picture = FrmMain.Tool(4).Picture
ElseIf FrmMain.Floor(GG).Picture = LoadPicture("") Then
FrmMain.Floor(GG).Picture = FrmMain.Tool(3).Picture
ElseIf FrmMain.Floor(GG).Picture = FrmMain.Tool(1).Picture Then
FrmMain.Floor(GG).Picture = FrmMain.Tool(5).Picture
End If

'FrmMain.Floor(GG).Picture = FrmMain.Tool(3).Picture
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
'Me.Hide
End Sub
Public Function Waitm(ByVal TimeToWait As Long) 'Time In seconds
Dim EndTime As Long
EndTime = GetTickCount + TimeToWait  '* 1000 Cause u give seconds and GetTickCount uses Milliseconds
Do Until GetTickCount > EndTime
DoEvents
Loop
End Function
Private Sub Command2_Click()
On Error Resume Next
Dim N

Fileef = App.Path & "\Data\" & List1.text

SaveFile App.Path & "\Data\" & List1.text, ""

For N = 0 To FrmMain.Floor.UBound
'\\\\\\\\\
With FrmMain.Tool(0)
If FrmMain.Floor(N).Picture = .Picture Then
 WriteAllText App.Path & "\Data\" & List1.text, "Wall*" & FrmMain.Floor(N).Top & "*" & FrmMain.Floor(N).Left, 3
End If
End With
'\\\\\\\\\
With FrmMain.Tool(1)
If FrmMain.Floor(N).Picture = .Picture Then

WriteAllText App.Path & "\Data\" & List1.text, "Man*" & FrmMain.Floor(N).Top & "*" & FrmMain.Floor(N).Left, 3
End If
End With
'\\\\\\\\\
'\\\\\\\\\
With FrmMain.Tool(2)
If FrmMain.Floor(N).Picture = .Picture Then

 WriteAllText App.Path & "\Data\" & List1.text, "Box*" & FrmMain.Floor(N).Top & "*" & FrmMain.Floor(N).Left, 3
End If
End With
'\\\\\\\\\
'\\\\\\\\\
With FrmMain.Tool(3)
If FrmMain.Floor(N).Picture = .Picture Then

Call WriteAllText(App.Path & "\Data\" & List1.text, "XPlace*" & FrmMain.Floor(N).Top & "*" & FrmMain.Floor(N).Left, 3)
End If
End With
'\\\\\\\\\
'\\\\\\\\\
With FrmMain.Tool(4)
If FrmMain.Floor(N).Picture = .Picture Then

Call WriteAllText(App.Path & "\Data\" & List1.text, "XPlace*" & FrmMain.Floor(N).Top & "*" & FrmMain.Floor(N).Left, 3)
Call WriteAllText(App.Path & "\Data\" & List1.text, "Box*" & FrmMain.Floor(N).Top & "*" & FrmMain.Floor(N).Left, 3)

End If
End With

'\\\\\\\\\
'\\\\\\\\\
With FrmMain.Tool(5)
If FrmMain.Floor(N).Picture = .Picture Then

Call WriteAllText(App.Path & "\Data\" & List1.text, "XPlace*" & FrmMain.Floor(N).Top & "*" & FrmMain.Floor(N).Left, 3)
WriteAllText App.Path & "\Data\" & List1.text, "Man*" & FrmMain.Floor(N).Top & "*" & FrmMain.Floor(N).Left, 3

End If
End With

'\\\\\\\\\

Next
End Sub

Private Sub Form_Load()
Icon = FrmMain.Icon
Timer1.Enabled = True

End Sub


Private Sub List1_DblClick()

Command1 = True
End Sub

Private Sub Timer1_Timer()
Dim cc

List1.Clear
cc = 1
Do Until FileExists(App.Path & "\Data\LEVEL" & cc) = False
List1.AddItem "LEVEL" & cc
cc = cc * 1 + 1 * 1
DoEvents
Loop
Form_Load
Timer1.Enabled = False
End Sub
