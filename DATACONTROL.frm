VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6276
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7752
   LinkTopic       =   "Form1"
   ScaleHeight     =   6276
   ScaleWidth      =   7752
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "EXIT"
      Height          =   492
      Left            =   5520
      TabIndex        =   14
      Top             =   3960
      Width           =   1092
   End
   Begin VB.CommandButton Command7 
      Caption         =   "NEXT"
      Height          =   492
      Left            =   5520
      TabIndex        =   13
      Top             =   3240
      Width           =   1092
   End
   Begin VB.CommandButton Command6 
      Caption         =   "PREVIOUS"
      Height          =   492
      Left            =   5520
      TabIndex        =   12
      Top             =   2400
      Width           =   1092
   End
   Begin VB.CommandButton Command5 
      Caption         =   "LAST"
      Height          =   492
      Left            =   5520
      TabIndex        =   11
      Top             =   1680
      Width           =   1092
   End
   Begin VB.CommandButton Command4 
      Caption         =   "FIRST"
      Height          =   492
      Left            =   3960
      TabIndex        =   10
      Top             =   3960
      Width           =   1092
   End
   Begin VB.CommandButton Command3 
      Caption         =   "DELETE"
      Height          =   492
      Left            =   3960
      TabIndex        =   9
      Top             =   3120
      Width           =   1092
   End
   Begin VB.CommandButton Command2 
      Caption         =   "UPDATE"
      Height          =   492
      Left            =   3960
      TabIndex        =   8
      Top             =   2400
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      Height          =   492
      Left            =   3960
      TabIndex        =   7
      Top             =   1680
      Width           =   1092
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   372
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4200
      Width           =   1932
   End
   Begin VB.TextBox Text3 
      Height          =   492
      Left            =   1800
      TabIndex        =   6
      Top             =   3120
      Width           =   1452
   End
   Begin VB.TextBox Text2 
      Height          =   492
      Left            =   1800
      TabIndex        =   5
      Top             =   2400
      Width           =   1452
   End
   Begin VB.TextBox Text1 
      Height          =   492
      Left            =   1800
      TabIndex        =   4
      Top             =   1680
      Width           =   1452
   End
   Begin VB.Label Label4 
      Caption         =   "ADDRESS"
      Height          =   372
      Left            =   360
      TabIndex        =   3
      Top             =   3000
      Width           =   1092
   End
   Begin VB.Label Label3 
      Caption         =   "AGE"
      Height          =   372
      Left            =   360
      TabIndex        =   2
      Top             =   2400
      Width           =   1092
   End
   Begin VB.Label Label2 
      Caption         =   "NAME"
      Height          =   372
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "DATA CONTROL"
      Height          =   612
      Left            =   2640
      TabIndex        =   0
      Top             =   360
      Width           =   1812
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Private Sub Command1_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text1.SetFocus
rs.AddNew
End Sub
Private Sub Command2_Click()
rs.Fields(0).Value = Text1.Text
rs.Fields(1).Value = Text2.Text
rs.Fields(2).Value = Text3.Text
rs.Update
MsgBox ("record updated")
End Sub
Private Sub Command3_Click()
Dim n As Integer
n = InputBox("enter the student no to deleted")
rs = MoveFirst
Do While Not rs.EOF
If rs.Fields(1).Value = n Then
rs.Delete
MsgBox ("record deleted")
Exit Sub
End If
rs.MoveLast
Loop
MsgBox ("record not found")
End Sub

Public Function display()
Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(2).Value
End Function
Private Sub Command4_Click()
rs.MovePrevious
If rs.EOF Then
rs.MoveFirst
End If
Call display
End Sub
Private Sub Command5_Click()
Dim n As Integer
n = InputBox("Enter to selected")
rs.MoveFirst
Do While Not rs.EOF
If rs.Fields(1).Value = n Then
Call display
Exit Sub
End If
rs.MoveNext
Loop
MsgBox ("record not found")
End Sub
Private Sub Command6_Click()
rs.MoveFirst
Call display
End Sub
Private Sub Command7_Click()
rs.MoveLast
Call display
End Sub
Private Sub Command8_Click()
rs.MoveNext
If (rs.EOF) Then
rs.MoveLast
End If
Call display
End Sub
Private Sub Form_Load()
'Set db = OpenDatabase("d:\gfolder\DAO.mdb")
'Set rs = db.OpenRecordset("stu")
End Sub


