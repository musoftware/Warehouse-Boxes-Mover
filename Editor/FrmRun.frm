VERSION 5.00
Begin VB.Form FrmRun 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Run"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6390
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmRun.frx":0000
   ScaleHeight     =   5175
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2160
      Top             =   480
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
      Begin VB.Label Label3 
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Timer AreYouWin 
      Interval        =   1000
      Left            =   840
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3240
      Top             =   0
   End
   Begin VB.Shape Man 
      BackColor       =   &H80000003&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   5520
      Shape           =   3  'Circle
      Top             =   120
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   146
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   145
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   144
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   143
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   142
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   141
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   140
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   139
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   138
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   137
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   136
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   135
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   134
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   133
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   132
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   131
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   130
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   129
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   128
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   127
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   126
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   125
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   124
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   123
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   122
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   121
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   120
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   119
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   118
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   117
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   116
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   115
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   114
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   113
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   112
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   111
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   110
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   109
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   108
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   107
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   106
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   105
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   104
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   103
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   102
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   101
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   100
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   99
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   98
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   97
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   96
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   95
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   94
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   93
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   92
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   91
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   90
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   89
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   88
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   87
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   86
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   85
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   84
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   83
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   82
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   81
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   80
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   79
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   78
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   77
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   76
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   75
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   74
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   73
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   72
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   71
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   70
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   69
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   68
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   67
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   66
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   65
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   64
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   63
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   62
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   61
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   60
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   59
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   58
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   57
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   56
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   55
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   54
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   53
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   52
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   51
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   50
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   49
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   48
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   47
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   46
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   45
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   44
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   43
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   42
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   41
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   40
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   39
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   38
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   37
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   36
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   35
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   34
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   33
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   32
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   31
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   30
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   29
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   28
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   27
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   26
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   25
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   24
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   23
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   22
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   21
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   20
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   19
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   18
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   17
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   16
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   15
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   14
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   13
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   12
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   11
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   10
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   9
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   8
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   7
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   6
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   5
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   4
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   3
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   2
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   1
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Boxes 
      Height          =   735
      Index           =   0
      Left            =   1440
      Picture         =   "FrmRun.frx":0A6E
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   132
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   131
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   130
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   129
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   128
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   127
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   126
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   125
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   124
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   123
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   122
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   121
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   120
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   119
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   118
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   117
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   116
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   115
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   114
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   113
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   112
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   111
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   110
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   109
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   108
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   107
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   106
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   105
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   104
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   103
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   102
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   101
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   100
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   99
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   98
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   97
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   96
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   95
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   94
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   93
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   92
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   91
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   90
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   89
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   88
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   87
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   86
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   85
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   84
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   83
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   82
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   81
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   80
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   79
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   78
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   77
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   76
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   75
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   74
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   73
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   72
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   71
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   70
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   69
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   68
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   67
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   66
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   65
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   64
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   63
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   62
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   61
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   60
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   59
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   58
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   57
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   56
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   55
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   54
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   53
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   52
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   51
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   50
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   49
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   48
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   47
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   46
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   45
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   44
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   43
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   42
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   41
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   40
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   39
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   38
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   37
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   36
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   35
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   34
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   33
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   32
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   31
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   30
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   29
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   28
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   27
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   26
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   25
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   24
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   23
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   22
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   21
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   20
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   19
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   18
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   17
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   16
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   15
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   14
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   13
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   12
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   11
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   10
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   9
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   8
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   7
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   6
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   5
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   4
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   3
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   2
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   1
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Wall 
      Height          =   735
      Index           =   0
      Left            =   720
      Picture         =   "FrmRun.frx":3AB0
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   735
   End
   Begin VB.Image III 
      Height          =   735
      Left            =   4800
      Picture         =   "FrmRun.frx":6AF2
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image II 
      Height          =   735
      Left            =   4080
      Picture         =   "FrmRun.frx":9B34
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image I 
      Height          =   735
      Left            =   3360
      Picture         =   "FrmRun.frx":CB76
      Stretch         =   -1  'True
      Top             =   120
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
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   100
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   99
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   98
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   97
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   96
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   95
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   94
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   93
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   92
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   91
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   90
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   89
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   88
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   87
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   86
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   85
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   84
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   83
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   82
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   81
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   80
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   79
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   78
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   77
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   76
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   75
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   74
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   73
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   72
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   71
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   70
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   69
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   68
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   67
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   66
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   65
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   64
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   63
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   62
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   61
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   60
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   59
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   58
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   57
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   56
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   55
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   54
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   53
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   52
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   51
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   50
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   49
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   48
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   47
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   46
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   45
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   44
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   43
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   42
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   41
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   40
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   39
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   38
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   37
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   36
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   35
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   34
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   33
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   32
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   31
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   30
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   29
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   28
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   27
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   26
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   25
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   24
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   23
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   22
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   21
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   20
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   19
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   18
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   17
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   16
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   15
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   14
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   13
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   12
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   11
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   10
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   9
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   8
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   7
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   6
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   5
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   4
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   3
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   2
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   1
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image XPlace 
      Height          =   735
      Index           =   0
      Left            =   0
      Picture         =   "FrmRun.frx":FBB8
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   735
   End
End
Attribute VB_Name = "FrmRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim REE, REE2, REE3, REE3REE3, NumLevel, TheRestart, PPOOOX, PPOOOY
Private Sub AreYouWin_Timer()

REE = ""
REE2 = ""
REE3 = 0
REE3REE3 = 0
For M = 0 To XPlace.UBound
Label1 = M

If XPlace(M).BorderStyle = 1 Then REE3REE3 = M: Exit For

REE = XPlace(M).Top & ":" & XPlace(M).Left
For N = 0 To Boxes.UBound
If Boxes(N).BorderStyle = 1 Then Exit For
Label2 = N
REE2 = Boxes(N).Top & ":" & Boxes(N).Left

If REE = REE2 Then
REE3 = REE3 * 1 + 1 * 1
Label3 = REE3
End If

Next


Next


If REE3 = REE3REE3 Then
Unload Me
Me.Hide
AreYouWin.Enabled = False
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
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub WalkDown()
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
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''j
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub WalkRight()
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
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''j
Public Sub WalkLeft()
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
'Exit For
Else
If BNUm = 1 And N = Boxes.UBound Then
Man.Left = Man.Left - Onelong
Exit For
End If
End If

Next
End Sub



Public Sub DrawMan(Value)
If Value = 1 Then
'SMan.Top = Man.Top
'SMan.Left = Man.Left - (SMan.Width / 3)
End If

End Sub



Sub Form_Load()
Icon = FrmMain.Icon
Dim Num, FloorX
NumLevel = 1


For N = 0 To Me.Width + 1000 Step 720
For M = 0 To Me.Height + 1000 Step 720
Num = Floor.Count
Load Me.Floor(Num)
Set FloorX = Me.Floor(Num)
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
Timer2.Enabled = True
'READGMAE (NumLevel)

End Sub

Public Sub READGMAE(LEVEL)
'On Error Resume Next
Dim A, Y
A = ReadAllText("C:\LWE", 5)
Dim Q, W, R
Y = 0
Q = 0
W = 0
R = 0

Dim E() As String, E2() As String, TGG, TGG2, Num
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
Num = R
'Load Me.Wall(num)
'Set BoxX = Me.Wall(num)

'If Abs(FrmMain.Width) < Abs(E2(1) + 735) Then FrmMain.Width = (E2(1) + 735)
'If Abs(FrmMain.Height) < Abs(E2(2) + 735) Then FrmMain.Height = (E2(2) + 735)
If Abs(TGG) > Abs(E2(1)) Then TGG = (E2(1))
If Abs(TGG2) > Abs(E2(2)) Then TGG2 = (E2(2))

'TGG
Me.Wall(Num).Top = E2(1)
Me.Wall(Num).Left = E2(2)
Me.Wall(Num).Visible = True
Me.Wall(Num).Picture = Me.Wall(0)

If PPOOOY < Wall(Num).Top Then PPOOOY = Wall(Num).Top
If PPOOOX < Wall(Num).Left Then PPOOOX = Wall(Num).Left

Wall(Num).BorderStyle = 0
R = R * 1 + 1 * 1
End If

If E2(0) = "XPlace" Then
Num = Q
'Load Me.XPlace(num)
'Set BoxX = Me.XPlace(num)
XPlace(Num).Top = E2(1)
XPlace(Num).Left = E2(2)
XPlace(Num).Visible = True
Me.XPlace(Num).Picture = Me.XPlace(0)
XPlace(Num).BorderStyle = 0
Q = Q * 1 + 1 * 1
End If

If E2(0) = "Box" Then
Num = W
'Load Me.Boxes(num)
'Set BoxX = Me.Boxes(num)
Boxes(Num).Top = E2(1)
Boxes(Num).Left = E2(2)
If PPOOOY < Boxes(Num).Top Then PPOOOY = Boxes(Num).Top
If PPOOOX < Boxes(Num).Left Then PPOOOX = Boxes(Num).Left


Boxes(Num).ZOrder
Boxes(Num).Visible = True

Me.Boxes(Num).Picture = Me.Boxes(0)
Boxes(Num).BorderStyle = 0
W = W * 1 + 1 * 1
End If

Next

For N = 0 To Boxes.UBound
If Boxes(N).BorderStyle = 1 Then Exit For
Boxes(N).Top = Boxes(N).Top - TGG
Boxes(N).Left = Boxes(N).Left - TGG2
If PPOOOY < Boxes(N).Top Then PPOOOY = Boxes(N).Top
If PPOOOX < Boxes(N).Left Then PPOOOX = Boxes(N).Left

Next

For N = 0 To XPlace.UBound
If XPlace(N).BorderStyle = 1 Then Exit For
XPlace(N).Top = XPlace(N).Top - TGG
XPlace(N).Left = XPlace(N).Left - TGG2
If PPOOOY < XPlace(N).Top Then PPOOOY = XPlace(N).Top
If PPOOOX < XPlace(N).Left Then PPOOOX = XPlace(N).Left

Next
For N = 0 To Wall.UBound
If Wall(N).BorderStyle = 1 Then Exit For
Wall(N).Top = Wall(N).Top - TGG
If PPOOOY < Wall(N).Top Then PPOOOY = Wall(N).Top
If PPOOOX < Wall(N).Left Then PPOOOX = Wall(N).Left

Wall(N).Left = Wall(N).Left - TGG2
Next
Man.Top = Man.Top - TGG
Man.Left = Man.Left - TGG2
Man.ZOrder
End Sub





Private Sub Timer2_Timer()
Unload Me
Me.Visible = True
PPOOOX = 55
PPOOOY = 55
READGMAE 1
Me.Height = PPOOOY + 735
Me.Width = PPOOOX + 735

Timer2.Enabled = False
End Sub


