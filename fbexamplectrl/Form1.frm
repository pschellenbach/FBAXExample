VERSION 5.00
Object = "{C496F38F-3875-4424-B1C9-3F1A79977F4C}#1.0#0"; "FBExampleCtl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2520
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   2520
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2040
      Top             =   1440
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   0
      Top             =   1440
   End
   Begin VB.CheckBox Check2 
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   1920
      Width           =   135
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   1920
      Width           =   135
   End
   Begin FBExampleCtl.xpcmdbutton xpcmdbutton2 
      Default         =   -1  'True
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "Default"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FBExampleCtl.xpcmdbutton xpcmdbutton1 
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "Norm&al"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Theme"
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
      Begin VB.OptionButton optTheme 
         Caption         =   "Metalic"
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton optTheme 
         Caption         =   "Homestead"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton optTheme 
         Caption         =   "Default"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Value           =   -1  'True
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub optTheme_Click(Index As Integer)
   xpcmdbutton1.Theme = Index
   xpcmdbutton2.Theme = Index
End Sub

Private Sub Timer1_Timer()
   Timer1.Enabled = False
   Check1.Value = 0
End Sub

Private Sub Timer2_Timer()
   Timer2.Enabled = False
   Check2.Value = 0
End Sub

Private Sub xpcmdbutton1_Click()
   Check1.Value = 1
   Timer1.Enabled = True
End Sub

Private Sub xpcmdbutton2_Click()
   Check2.Value = 1
   Timer2.Enabled = True
End Sub
