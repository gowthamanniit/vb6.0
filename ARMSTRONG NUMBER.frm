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
   Begin VB.CommandButton Command3 
      Caption         =   "END"
      Height          =   612
      Left            =   4680
      TabIndex        =   4
      Top             =   2520
      Width           =   1332
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CLICK"
      Height          =   612
      Left            =   3000
      TabIndex        =   3
      Top             =   2520
      Width           =   1332
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CLEAR"
      Height          =   612
      Left            =   1440
      TabIndex        =   2
      Top             =   2520
      Width           =   1332
   End
   Begin VB.TextBox Text1 
      Height          =   492
      Left            =   1920
      TabIndex        =   1
      Top             =   1320
      Width           =   2532
   End
   Begin VB.Label Label1 
      Caption         =   "ARMSTRONG NUMBER"
      Height          =   492
      Left            =   2040
      TabIndex        =   0
      Top             =   480
      Width           =   2412
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Text = ""
End Sub

Private Sub Command2_Click()
Dim s, r, n, n1 As Integer
n = Val(Text1.Text)
n1 = n
s = 0
Do While (n > 0)
r = n Mod 10
s = s + (r * r * r)
n = n / 10
Loop
If (n1 = s) Then
MsgBox "amstrong "
Else
MsgBox "not amstrong"
End If
End Sub

Private Sub Command3_Click()
End
End Sub
