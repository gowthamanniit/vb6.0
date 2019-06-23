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
   Begin VB.TextBox Text2 
      Height          =   492
      Left            =   2280
      TabIndex        =   7
      Top             =   3600
      Width           =   1932
   End
   Begin VB.TextBox Text1 
      Height          =   492
      Left            =   2280
      TabIndex        =   6
      Top             =   2640
      Width           =   1932
   End
   Begin VB.CommandButton Command5 
      Caption         =   "M"
      Height          =   372
      Left            =   720
      TabIndex        =   5
      Top             =   3600
      Width           =   972
   End
   Begin VB.CommandButton Command4 
      Caption         =   "N"
      Height          =   372
      Left            =   720
      TabIndex        =   4
      Top             =   2640
      Width           =   972
   End
   Begin VB.CommandButton Command3 
      Caption         =   "END"
      Height          =   372
      Left            =   3600
      TabIndex        =   3
      Top             =   1560
      Width           =   972
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CLICK"
      Height          =   372
      Left            =   2280
      TabIndex        =   2
      Top             =   1560
      Width           =   972
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CLEAR"
      Height          =   372
      Left            =   960
      TabIndex        =   1
      Top             =   1560
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "MULTIPLICATION"
      Height          =   492
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   2772
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Text = ""
Text2.Text = ""
Me.Cls
End Sub
Private Sub Command2_Click()
Dim n, m As Integer
n = Val(Text1.Text)
m = Val(Text2.Text)
For i = 1 To n
n = i * m
Print
Print i; "*"; "="; n
Print
Next i
End Sub
Private Sub Command3_Click()
End
End Sub
