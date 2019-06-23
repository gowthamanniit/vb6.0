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
   Begin VB.CommandButton Command4 
      Caption         =   "end"
      Height          =   492
      Left            =   480
      TabIndex        =   3
      Top             =   3840
      Width           =   1092
   End
   Begin VB.CommandButton Command3 
      Caption         =   "stop"
      Height          =   492
      Left            =   480
      TabIndex        =   2
      Top             =   3000
      Width           =   1092
   End
   Begin VB.CommandButton Command2 
      Caption         =   "run"
      Height          =   492
      Left            =   480
      TabIndex        =   1
      Top             =   2160
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "jump"
      Height          =   492
      Left            =   480
      TabIndex        =   0
      Top             =   1440
      Width           =   1092
   End
   Begin VB.Timer Timer3 
      Left            =   2280
      Top             =   360
   End
   Begin VB.Timer Timer2 
      Left            =   1560
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Left            =   840
      Top             =   360
   End
   Begin VB.Image Image1 
      Height          =   3600
      Left            =   2520
      Picture         =   "animationobject.frx":0000
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   4512
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer2.Interval = 1
Timer1.Enabled = False
Timer3.Enabled = False
End Sub
Private Sub Command2_Click()
Timer1.Interval = 1
Timer1.Enabled = True
Timer2.Enabled = False
Timer3.Enabled = False
End Sub
Private Sub Command3_Click()
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
End Sub
Private Sub Command4_Click()
End
End Sub
Private Sub Timer1_Timer()
Image1.Left = Image1.Left + 20
End Sub
Private Sub Timer2_Timer()
Image1.Top = Image1.Top - 50
Image1.Left = Image1.Left + 10
If Image1.Top < 10 Then
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = True
End Sub
Private Sub Timer3_Timer()
Image1.Left = Image1.Top + 20
Image1.Top = Image1.Left + 50
If Image1.Top > 40 Then
Timer2.Enabled = True
Timer3.Enabled = False
End If
End Sub
