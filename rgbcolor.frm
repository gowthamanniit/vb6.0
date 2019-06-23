VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6852
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8688
   LinkTopic       =   "Form1"
   ScaleHeight     =   6852
   ScaleWidth      =   8688
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "END"
      Height          =   372
      Left            =   3480
      TabIndex        =   8
      Top             =   5160
      Width           =   1332
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CLEAR"
      Height          =   372
      Left            =   1560
      TabIndex        =   7
      Top             =   5160
      Width           =   1332
   End
   Begin VB.TextBox Text3 
      Height          =   372
      Left            =   3480
      TabIndex        =   6
      Top             =   3960
      Width           =   1332
   End
   Begin VB.TextBox Text2 
      Height          =   372
      Left            =   3480
      TabIndex        =   5
      Top             =   3120
      Width           =   1332
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   3480
      TabIndex        =   4
      Top             =   2280
      Width           =   1332
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   372
      Left            =   360
      TabIndex        =   2
      Top             =   3960
      Width           =   2292
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   372
      Left            =   360
      TabIndex        =   1
      Top             =   3120
      Width           =   2292
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   372
      Left            =   360
      TabIndex        =   0
      Top             =   2280
      Width           =   2292
   End
   Begin VB.Label Label1 
      Caption         =   "RGB COLOR"
      Height          =   852
      Left            =   960
      TabIndex        =   3
      Top             =   600
      Width           =   2292
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
Text3.Text = ""
End Sub
Private Sub Command2_Click()
End
End Sub
Private Sub Form_Load()
HScroll1.Min = 0
HScroll1.Max = 255
HScroll2.Min = 0
HScroll2.Max = 255
HScroll3.Min = 0
HScroll3.Max = 255
End Sub
Private Sub HScroll1_Change()
Form1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Text1.Text = Val(HScroll1.Value)
End Sub
Private Sub HScroll2_Change()
Form1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Text2.Text = Val(HScroll2.Value)
End Sub
Private Sub HScroll3_Change()
Form1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Text3.Text = Val(HScroll3.Value)
End Sub
Private Sub Text1_Change()
HScroll1.Value = Val(Text1.Text)
End Sub
Private Sub Text2_Change()
HScroll2.Value = Val(Text2.Text)
End Sub
Private Sub Text3_Change()
HScroll3.Value = Val(Text3.Text)
End Sub
