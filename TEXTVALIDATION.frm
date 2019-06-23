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
   Begin VB.TextBox Text10 
      Height          =   492
      Left            =   4800
      TabIndex        =   22
      Top             =   2520
      Width           =   1332
   End
   Begin VB.TextBox Text9 
      Height          =   492
      Left            =   4800
      TabIndex        =   21
      Top             =   1680
      Width           =   1332
   End
   Begin VB.TextBox Text8 
      Height          =   492
      Left            =   1680
      TabIndex        =   20
      Top             =   6240
      Width           =   1332
   End
   Begin VB.TextBox Text7 
      Height          =   492
      Left            =   1680
      TabIndex        =   19
      Top             =   5520
      Width           =   1332
   End
   Begin VB.TextBox Text6 
      Height          =   492
      Left            =   1680
      TabIndex        =   18
      Top             =   4920
      Width           =   1332
   End
   Begin VB.TextBox Text5 
      Height          =   492
      Left            =   1680
      TabIndex        =   17
      Top             =   4320
      Width           =   1332
   End
   Begin VB.TextBox Text4 
      Height          =   492
      Left            =   1680
      TabIndex        =   16
      Top             =   3720
      Width           =   1332
   End
   Begin VB.TextBox Text3 
      Height          =   492
      Left            =   1680
      TabIndex        =   15
      Top             =   3000
      Width           =   1332
   End
   Begin VB.TextBox Text2 
      Height          =   492
      Left            =   1680
      TabIndex        =   14
      Top             =   2400
      Width           =   1332
   End
   Begin VB.TextBox Text1 
      Height          =   492
      Left            =   1680
      TabIndex        =   13
      Top             =   1800
      Width           =   1332
   End
   Begin VB.CommandButton Command2 
      Caption         =   "end"
      Height          =   612
      Left            =   3360
      TabIndex        =   12
      Top             =   5160
      Width           =   1932
   End
   Begin VB.CommandButton Command1 
      Caption         =   "tot/avg/result"
      Height          =   732
      Left            =   3360
      TabIndex        =   11
      Top             =   3840
      Width           =   1932
   End
   Begin VB.Label Label11 
      Caption         =   "result"
      Height          =   492
      Left            =   3480
      TabIndex        =   10
      Top             =   2520
      Width           =   1092
   End
   Begin VB.Label Label10 
      Caption         =   "average"
      Height          =   492
      Left            =   3480
      TabIndex        =   9
      Top             =   1680
      Width           =   1092
   End
   Begin VB.Label Label9 
      Caption         =   "total"
      Height          =   492
      Left            =   360
      TabIndex        =   8
      Top             =   6240
      Width           =   1092
   End
   Begin VB.Label Label8 
      Caption         =   "mark5"
      Height          =   492
      Left            =   360
      TabIndex        =   7
      Top             =   5640
      Width           =   1092
   End
   Begin VB.Label Label7 
      Caption         =   "mark4"
      Height          =   492
      Left            =   360
      TabIndex        =   6
      Top             =   5040
      Width           =   1092
   End
   Begin VB.Label Label6 
      Caption         =   "mark3"
      Height          =   492
      Left            =   360
      TabIndex        =   5
      Top             =   4440
      Width           =   1092
   End
   Begin VB.Label Label5 
      Caption         =   "mark2"
      Height          =   492
      Left            =   360
      TabIndex        =   4
      Top             =   3840
      Width           =   1092
   End
   Begin VB.Label Label4 
      Caption         =   "mark1"
      Height          =   492
      Left            =   360
      TabIndex        =   3
      Top             =   3120
      Width           =   1092
   End
   Begin VB.Label Label3 
      Caption         =   "regno"
      Height          =   492
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   1092
   End
   Begin VB.Label Label2 
      Caption         =   "name"
      Height          =   492
      Left            =   480
      TabIndex        =   1
      Top             =   1800
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "TEXT ATTRIBUTE"
      Height          =   612
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   2052
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a, b As String
Dim c, d, e, f, g, h, i, j, k As Integer
a = Val(Text1.Text)
b = Val(Text2.Text)
c = Val(Text3.Text)
d = Val(Text4.Text)
e = Val(Text5.Text)
f = Val(Text6.Text)
g = Val(Text7.Text)
h = Val(Text8.Text)
i = Val(Text9.Text)
j = Val(Text10.Text)
k = Val(c) + Val(e) + Val(f) + Val(g)
Text8.Text = Val(k)
Text9.Text = Val(i)
If (c > 35 And d > 35 And e > 35 And f > 35 And g > 35) Then
Text10.Text = "PASS"
Else
Text10.Text = "FAIL"
End If
End Sub
Private Sub Command2_Click()
End
End Sub
