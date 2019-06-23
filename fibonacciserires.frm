VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6804
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   ScaleHeight     =   6804
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "end"
      Height          =   492
      Left            =   5880
      TabIndex        =   3
      Top             =   2400
      Width           =   1092
   End
   Begin VB.CommandButton Command2 
      Caption         =   "fibo"
      Height          =   492
      Left            =   4200
      TabIndex        =   2
      Top             =   2400
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "click"
      Height          =   492
      Left            =   2520
      TabIndex        =   1
      Top             =   2400
      Width           =   1092
   End
   Begin VB.TextBox Text1 
      Height          =   732
      Left            =   2520
      TabIndex        =   0
      Top             =   720
      Width           =   4572
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Cls
Text1.Text = ""
End Sub
Private Sub Command2_Click()
Dim a, n, j As Integer
n = Text1.Text
For j = 0 To n - 1
a = fibo(j)
Print a
Next
End Sub
Private Sub Command3_Click()
End
End Sub
Public Function fibo(m As Integer)
If (m = 0) Then
fibo = 0
Else
If (m = 1) Then
fibo = 1
Else
fibo = fibo(m - 1) + fibo(m - 2)
End If
End If
End Function
