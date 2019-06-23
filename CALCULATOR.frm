VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2496
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   8832
   ScaleWidth      =   12192
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command20 
      Caption         =   "END"
      Height          =   492
      Left            =   4560
      TabIndex        =   22
      Top             =   5280
      Width           =   612
   End
   Begin VB.CommandButton Command19 
      Caption         =   "CLEAR"
      Height          =   492
      Left            =   4560
      TabIndex        =   21
      Top             =   4560
      Width           =   732
   End
   Begin VB.CommandButton Command18 
      Caption         =   "="
      Height          =   492
      Left            =   4560
      TabIndex        =   20
      Top             =   3720
      Width           =   612
   End
   Begin VB.CommandButton Command17 
      Caption         =   "EMPTY"
      Height          =   492
      Left            =   4560
      TabIndex        =   19
      Top             =   2880
      Width           =   732
   End
   Begin VB.CommandButton Command16 
      Caption         =   "/"
      Height          =   492
      Left            =   3000
      TabIndex        =   18
      Top             =   5280
      Width           =   612
   End
   Begin VB.CommandButton Command15 
      Caption         =   "*"
      Height          =   492
      Left            =   2160
      TabIndex        =   17
      Top             =   5280
      Width           =   612
   End
   Begin VB.CommandButton Command14 
      Caption         =   "-"
      Height          =   492
      Left            =   1320
      TabIndex        =   16
      Top             =   5280
      Width           =   612
   End
   Begin VB.CommandButton Command13 
      Caption         =   "+"
      Height          =   492
      Left            =   480
      TabIndex        =   15
      Top             =   5280
      Width           =   612
   End
   Begin VB.CommandButton Command12 
      Caption         =   "3"
      Height          =   492
      Left            =   3000
      TabIndex        =   14
      Top             =   4560
      Width           =   612
   End
   Begin VB.CommandButton Command11 
      Caption         =   "0"
      Height          =   492
      Left            =   2160
      TabIndex        =   13
      Top             =   4560
      Width           =   612
   End
   Begin VB.CommandButton Command10 
      Caption         =   "9"
      Height          =   492
      Left            =   1320
      TabIndex        =   12
      Top             =   4560
      Width           =   612
   End
   Begin VB.CommandButton Command9 
      Caption         =   "8"
      Height          =   492
      Left            =   480
      TabIndex        =   11
      Top             =   4560
      Width           =   612
   End
   Begin VB.CommandButton Command8 
      Caption         =   "7"
      Height          =   492
      Left            =   3000
      TabIndex        =   10
      Top             =   3720
      Width           =   612
   End
   Begin VB.CommandButton Command7 
      Caption         =   "6"
      Height          =   492
      Left            =   2160
      TabIndex        =   9
      Top             =   3720
      Width           =   612
   End
   Begin VB.CommandButton Command6 
      Caption         =   "5"
      Height          =   492
      Left            =   1320
      TabIndex        =   8
      Top             =   3720
      Width           =   612
   End
   Begin VB.CommandButton Command5 
      Caption         =   "4"
      Height          =   492
      Left            =   480
      TabIndex        =   7
      Top             =   3720
      Width           =   612
   End
   Begin VB.CommandButton Command4 
      Caption         =   "3"
      Height          =   492
      Left            =   3000
      TabIndex        =   6
      Top             =   2880
      Width           =   612
   End
   Begin VB.CommandButton Command3 
      Caption         =   "2"
      Height          =   492
      Left            =   2160
      TabIndex        =   5
      Top             =   2880
      Width           =   612
   End
   Begin VB.CommandButton Command2 
      Caption         =   "1"
      Height          =   492
      Left            =   1320
      TabIndex        =   4
      Top             =   2880
      Width           =   612
   End
   Begin VB.CommandButton Command1 
      Caption         =   "C"
      Height          =   492
      Left            =   480
      TabIndex        =   3
      Top             =   2880
      Width           =   612
   End
   Begin VB.TextBox Text2 
      Height          =   492
      Left            =   480
      TabIndex        =   2
      Top             =   1920
      Width           =   1332
   End
   Begin VB.TextBox Text1 
      Height          =   492
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   1332
   End
   Begin VB.Label Label1 
      Caption         =   "CALCULATOR"
      Height          =   612
      Left            =   3000
      TabIndex        =   0
      Top             =   360
      Width           =   2652
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b As Integer
Dim s As String
Private Sub Command1_Click()
Text1.Text = Command1(Index).Caption
End Sub
Private Sub Command18_Click()
Text2.Text = Text1.Text
Text1.Text = ""
Command3.Caption = Command2(Index).Caption
c = Command3.Caption
Select Case (c)
Case "%"
Text1.Text = b / 10
End Select
End Sub
Private Sub Command19_Click()
a = Val(Text2.Text)
b = Val(Text2.Text)
c = Command3.Caption
Select Case (c)
Case "+"
Text2.Text = ""
Text1.Text = a + b
Case "-"
Text2.Text = ""
Text1.Text = a - b
Case "*"
Text2.Text = ""
Text1.Text = a * b
Case "/"
Text2.Text = ""
Text1.Text = a / b
End Select
End Sub
Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
End Sub
Private Sub Command20_Click()
End
End Sub
