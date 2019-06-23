VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5484
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8088
   LinkTopic       =   "Form1"
   ScaleHeight     =   5484
   ScaleWidth      =   8088
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "END"
      Height          =   492
      Left            =   4080
      TabIndex        =   7
      Top             =   4200
      Width           =   1212
   End
   Begin VB.CommandButton Command5 
      Caption         =   "UPPER"
      Height          =   492
      Left            =   2400
      TabIndex        =   6
      Top             =   4200
      Width           =   1212
   End
   Begin VB.CommandButton Command4 
      Caption         =   "LOWER"
      Height          =   492
      Left            =   720
      TabIndex        =   5
      Top             =   4200
      Width           =   1212
   End
   Begin VB.CommandButton Command3 
      Caption         =   "LENGTH"
      Height          =   492
      Left            =   4080
      TabIndex        =   4
      Top             =   3000
      Width           =   1212
   End
   Begin VB.CommandButton Command2 
      Caption         =   "COMP"
      Height          =   492
      Left            =   2400
      TabIndex        =   3
      Top             =   3000
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ASCII"
      Height          =   492
      Left            =   720
      TabIndex        =   2
      Top             =   3000
      Width           =   1212
   End
   Begin VB.TextBox Text1 
      Height          =   732
      Left            =   3120
      TabIndex        =   1
      Top             =   1440
      Width           =   2292
   End
   Begin VB.Label Label1 
      Caption         =   "STRING MANIPULATION"
      Height          =   612
      Left            =   3120
      TabIndex        =   0
      Top             =   480
      Width           =   2892
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s1, s2, s3 As Integer
Dim st1 As Integer
Private Sub Command1_Click()
s1 = InputBox("enter the string")
Text1.Text = Asc(s1)
End Sub
Private Sub Command2_Click()
s1 = InputBox("enter first string")
s2 = InputBox("enter second string")
st1 = StrComp(s1, s2)
If (st1 = 1) Then
Text1.Text = "first string greater than (>) second string"
ElseIf (st1 = -1) Then
Text1.Text = "first string less than (>) second string"
Else
Text1.Text = "both are equal"
End If
End Sub
Private Sub Command3_Click()
s1 = InputBox("enter the string")
Text1.Text = Len(s1)
End Sub
Private Sub Command4_Click()
s1 = InputBox("enter the string")
Text1.Text = LCase(s1)
End Sub
Private Sub Command5_Click()
s1 = InputBox("enter the string")
Text1.Text = UCase(s1)
End Sub
Private Sub Command6_Click()
End
End Sub
