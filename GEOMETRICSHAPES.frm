VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2496
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2496
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "END"
      Height          =   492
      Left            =   2400
      TabIndex        =   6
      Top             =   4200
      Width           =   972
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CLEAR"
      Height          =   492
      Left            =   1080
      TabIndex        =   5
      Top             =   4200
      Width           =   972
   End
   Begin VB.OptionButton Option4 
      Caption         =   "LINE"
      Height          =   612
      Left            =   2640
      TabIndex        =   3
      Top             =   3000
      Width           =   1332
   End
   Begin VB.OptionButton Option3 
      Caption         =   "BI SQUARE"
      Height          =   612
      Left            =   720
      TabIndex        =   2
      Top             =   3000
      Width           =   1332
   End
   Begin VB.OptionButton Option2 
      Caption         =   "SQUARE"
      Height          =   612
      Left            =   2640
      TabIndex        =   1
      Top             =   1800
      Width           =   1332
   End
   Begin VB.OptionButton Option1 
      Caption         =   "CIRCLE"
      Height          =   612
      Left            =   720
      TabIndex        =   0
      Top             =   1800
      Width           =   1332
   End
   Begin VB.Label Label1 
      Caption         =   "GEOMETRIC SHAPES"
      Height          =   492
      Left            =   2160
      TabIndex        =   4
      Top             =   600
      Width           =   1692
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Cls
End Sub
Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
Me.AutoRedraw = True
Me.WindowState = 2
Me.Width = Screen.Width
Me.Height = Screen.Height
Option1.Value = False
Option3.Value = False
Option4.Value = False
End Sub
Private Sub Option1_Click()
Me.Cls
Option1.Value = True
x = Screen.Width / 2
y = Screen.Height / 2
Me.Circle (x, y), (Rnd * 300), QBColor(Rnd * 15)
End Sub
Private Sub Option2_Click()
Me.Cls
Option2.Value = True
x = Screen.Width / 2
y = Screen.Height / 2
Me.Line (5000, 4000)-(10000, 8000), QBColor(Rnd * 15), B
End Sub
Private Sub Option3_Click()
Me.Cls
Option3.Value = True
x = Screen.Width / 2
y = Screen.Height / 2
Me.Line (5000, 4000)-(10000, 8000), QBColor(Rnd * 15), BF
End Sub
Private Sub Option4_Click()
Me.Cls
Option4.Value = True
Me.Line (5000, 4000)-(10000, 8000), QBColor(Rndnn * 15)
End Sub
