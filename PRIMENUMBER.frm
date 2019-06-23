VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6468
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   ScaleHeight     =   6468
   ScaleWidth      =   8220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "END"
      Height          =   612
      Left            =   4200
      TabIndex        =   4
      Top             =   3120
      Width           =   1092
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CLICK"
      Height          =   612
      Left            =   2760
      TabIndex        =   3
      Top             =   3120
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CLEAR"
      Height          =   612
      Left            =   1320
      TabIndex        =   2
      Top             =   3120
      Width           =   1092
   End
   Begin VB.TextBox Text1 
      Height          =   732
      Left            =   2160
      TabIndex        =   1
      Top             =   1560
      Width           =   2292
   End
   Begin VB.Label Label1 
      Caption         =   "PRIME NUMBER"
      Height          =   492
      Left            =   2280
      TabIndex        =   0
      Top             =   720
      Width           =   2172
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
Dim a, b, i As Integer
a = Val(Text1.Text)
i = 1
b = 0
While (i <= a)
If ((a Mod i) = 0) Then
b = b + 1
End If
i = i + 1
Wend
If (b = 2) Then
MsgBox "prime"
Else
MsgBox "not prime"
End If
End Sub
Private Sub Command3_Click()
End
End Sub
