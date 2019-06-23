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
   Begin VB.FileListBox File1 
      Height          =   3336
      Left            =   5640
      TabIndex        =   2
      Top             =   840
      Width           =   1932
   End
   Begin VB.DirListBox Dir1 
      Height          =   1584
      Left            =   3120
      TabIndex        =   1
      Top             =   840
      Width           =   1932
   End
   Begin VB.DriveListBox Drive1 
      Height          =   288
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   1692
   End
   Begin VB.Image Image1 
      Height          =   3180
      Left            =   600
      Picture         =   "LOADPICTURE.frx":0000
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   4440
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub
Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub
Private Sub File1_Click()
Image1.Picture = LoadPicture(Dir1.Path + "\" + File1.FileName)
End Sub
