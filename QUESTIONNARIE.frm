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
   Begin VB.CommandButton Command3 
      Caption         =   "end"
      Height          =   372
      Left            =   4200
      TabIndex        =   16
      Top             =   4680
      Width           =   1092
   End
   Begin VB.CommandButton Command2 
      Caption         =   "mark"
      Height          =   372
      Left            =   4200
      TabIndex        =   15
      Top             =   3600
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "clear"
      Height          =   372
      Left            =   4200
      TabIndex        =   14
      Top             =   2760
      Width           =   1092
   End
   Begin VB.OptionButton Option9 
      Caption         =   "vabian"
      Height          =   492
      Left            =   2760
      TabIndex        =   13
      Top             =   4920
      Width           =   1332
   End
   Begin VB.OptionButton Option8 
      Caption         =   "australia"
      Height          =   492
      Left            =   1680
      TabIndex        =   12
      Top             =   4920
      Width           =   1332
   End
   Begin VB.OptionButton Option7 
      Caption         =   "greeenland"
      Height          =   492
      Left            =   480
      TabIndex        =   11
      Top             =   4920
      Width           =   1332
   End
   Begin VB.OptionButton Option6 
      Caption         =   "calcuta"
      Height          =   492
      Left            =   2400
      TabIndex        =   9
      Top             =   3840
      Width           =   1092
   End
   Begin VB.OptionButton Option5 
      Caption         =   "mumbai"
      Height          =   492
      Left            =   1440
      TabIndex        =   8
      Top             =   3840
      Width           =   1092
   End
   Begin VB.OptionButton Option4 
      Caption         =   "delhi"
      Height          =   492
      Left            =   600
      TabIndex        =   7
      Top             =   3840
      Width           =   1092
   End
   Begin VB.OptionButton Option3 
      Caption         =   "dennis"
      Height          =   492
      Left            =   2760
      TabIndex        =   5
      Top             =   2520
      Width           =   1092
   End
   Begin VB.OptionButton Option2 
      Caption         =   "bilgates"
      Height          =   492
      Left            =   1680
      TabIndex        =   4
      Top             =   2520
      Width           =   1092
   End
   Begin VB.OptionButton Option1 
      Caption         =   "babbage"
      Height          =   492
      Left            =   600
      TabIndex        =   3
      Top             =   2520
      Width           =   1092
   End
   Begin VB.TextBox Text1 
      Height          =   612
      Left            =   4080
      TabIndex        =   2
      Top             =   1440
      Width           =   1332
   End
   Begin VB.Label Label4 
      Caption         =   "which is the smallest country"
      Height          =   492
      Left            =   600
      TabIndex        =   10
      Top             =   4440
      Width           =   2292
   End
   Begin VB.Label Label3 
      Caption         =   "where is the indegated shifted"
      Height          =   492
      Left            =   720
      TabIndex        =   6
      Top             =   3360
      Width           =   2292
   End
   Begin VB.Label Label2 
      Caption         =   "who is the father of the computer"
      Height          =   492
      Left            =   720
      TabIndex        =   1
      Top             =   1560
      Width           =   2292
   End
   Begin VB.Label Label1 
      Caption         =   "QUESTIONNARIE"
      Height          =   732
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   2292
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer

Private Sub Command1_Click()
Text1.Text = ""
End Sub

Private Sub Command2_Click()
Text1.Text = i
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
i = 0
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False
Option6.Value = False
Option7.Value = False
Option8.Value = False
Option9.Value = False
Option10.Value = False
Option11.Value = False
Option12.Value = False
End Sub

Private Sub Option10_Click()
i = i + 10
End Sub

Private Sub Option12_Click()
i = i + 10
End Sub

Private Sub Option9_Click()
i = i + 10
End Sub

