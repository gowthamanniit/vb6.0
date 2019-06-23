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
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   372
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5880
      Width           =   912
   End
   Begin VB.CommandButton Command9 
      Caption         =   "SAVE"
      Height          =   612
      Left            =   6600
      TabIndex        =   25
      Top             =   960
      Width           =   1092
   End
   Begin VB.CommandButton Command8 
      Caption         =   "EXIT"
      Height          =   612
      Left            =   6840
      TabIndex        =   24
      Top             =   4440
      Width           =   1092
   End
   Begin VB.CommandButton Command7 
      Caption         =   "NEXT"
      Height          =   612
      Left            =   6840
      TabIndex        =   23
      Top             =   3720
      Width           =   1092
   End
   Begin VB.CommandButton Command6 
      Caption         =   "LAST"
      Height          =   612
      Left            =   6840
      TabIndex        =   22
      Top             =   2880
      Width           =   1092
   End
   Begin VB.CommandButton Command5 
      Caption         =   "FIRST"
      Height          =   612
      Left            =   6840
      TabIndex        =   21
      Top             =   2040
      Width           =   1092
   End
   Begin VB.CommandButton Command4 
      Caption         =   "FIND"
      Height          =   612
      Left            =   5520
      TabIndex        =   20
      Top             =   4440
      Width           =   1092
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PREVIOUS"
      Height          =   612
      Left            =   5520
      TabIndex        =   19
      Top             =   3720
      Width           =   1092
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DELETE"
      Height          =   612
      Left            =   5520
      TabIndex        =   18
      Top             =   2880
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      Height          =   612
      Left            =   5520
      TabIndex        =   17
      Top             =   2040
      Width           =   1092
   End
   Begin VB.TextBox Text8 
      Height          =   492
      Left            =   2520
      TabIndex        =   16
      Top             =   5640
      Width           =   2052
   End
   Begin VB.TextBox Text7 
      Height          =   492
      Left            =   2520
      TabIndex        =   15
      Top             =   5040
      Width           =   2052
   End
   Begin VB.TextBox Text6 
      Height          =   492
      Left            =   2520
      TabIndex        =   14
      Top             =   4440
      Width           =   2052
   End
   Begin VB.TextBox Text5 
      Height          =   492
      Left            =   2520
      TabIndex        =   13
      Top             =   3840
      Width           =   2052
   End
   Begin VB.TextBox Text4 
      Height          =   492
      Left            =   2520
      TabIndex        =   12
      Top             =   3240
      Width           =   2052
   End
   Begin VB.TextBox Text3 
      Height          =   492
      Left            =   2520
      TabIndex        =   11
      Top             =   2640
      Width           =   2052
   End
   Begin VB.TextBox Text2 
      Height          =   492
      Left            =   2520
      TabIndex        =   10
      Top             =   2040
      Width           =   2052
   End
   Begin VB.TextBox Text1 
      Height          =   492
      Left            =   2520
      TabIndex        =   9
      Top             =   1440
      Width           =   2052
   End
   Begin VB.Label Label9 
      Caption         =   "CHARGE"
      Height          =   492
      Left            =   240
      TabIndex        =   8
      Top             =   5640
      Width           =   852
   End
   Begin VB.Label Label8 
      Caption         =   "UNIT"
      Height          =   492
      Left            =   240
      TabIndex        =   7
      Top             =   5040
      Width           =   852
   End
   Begin VB.Label Label7 
      Caption         =   "LAST MONTH READING"
      Height          =   492
      Left            =   240
      TabIndex        =   6
      Top             =   4440
      Width           =   2172
   End
   Begin VB.Label Label6 
      Caption         =   "CURRENT MONTH READING"
      Height          =   492
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   2652
   End
   Begin VB.Label Label5 
      Caption         =   "METER NUMBER"
      Height          =   492
      Left            =   360
      TabIndex        =   4
      Top             =   3360
      Width           =   1812
   End
   Begin VB.Label Label4 
      Caption         =   "ADD"
      Height          =   492
      Left            =   360
      TabIndex        =   3
      Top             =   2640
      Width           =   852
   End
   Begin VB.Label Label3 
      Caption         =   "PHONE NO"
      Height          =   492
      Left            =   360
      TabIndex        =   2
      Top             =   2160
      Width           =   1212
   End
   Begin VB.Label Label2 
      Caption         =   "NAME"
      Height          =   492
      Left            =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "EBBILL CALCULATION USING DAO"
      Height          =   252
      Left            =   2400
      TabIndex        =   0
      Top             =   600
      Width           =   3972
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
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
rs.addnew
End Sub
Private Sub Command2_Click()
Dim n As Integer
n = InputBox("Enter the meter no. to be deleted")
rs.MoveFirst
Do While Not rs.EOF
If rs.Fields(3).Value = n Then
rs.Delete
MsgBox ("record deleted")
Exit Sub
End If
rs.MoveLast
Loop
MsgBox ("record not found")
End Sub
Private Sub Command3_Click()
rs.MovePrevious
If (rs.EOF) Then
rs.MoveFirst
End If
Call display
End Sub
Private Sub Command4_Click()
Dim n As Integer
n = InputBox("Enter to seleted")
rs.MoveFirst
Do While Not rs.EOF
If rs.Fields(3).Value = n Then
Call display
Exit Sub
End If
rs.MoveNext
Loop
MsgBox ("record not found")
End Sub

Private Sub Command5_Click()
rs.MoveFirst
Call display
End Sub

Private Sub Command6_Click()
rs.MoveLast
Call display
End Sub

Private Sub Command7_Click()
rs.MoveNext
If (rs.EOF) Then
rs.MoveLast
End If
Call display
End Sub

Private Sub Command8_Click()
Dim a As Integer
a = MsgBox("do you want to exit", vbQuestion - vbYesNo, "exit")
If a = vbYes Then
End
Else
MsgBox ("exit cancelled")
End If
End Sub

Private Sub Command9_Click()
Dim a As Integer
a = MsgBox("saved document?", vbQuestion + vbYesNo, "save")
If a = vbYes Then
rs.Fields(0).Value = Text1.Text
rs.Fields(1).Value = Text2.Text
rs.Fields(2).Value = Text3.Text
rs.Fields(3).Value = Text4.Text
rs.Fields(4).Value = Text5.Text
rs.Fields(5).Value = Text6.Text
rs.Update
MsgBox ("saved")
Else
MsgBox ("cancelled")
End If
End Sub
Public Function display()
Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(2).Value
Text4.Text = rs.Fields(3).Value
Text5.Text = rs.Fields(4).Value
Text6.Text = rs.Fields(5).Value
End Function

Private Sub Form_Load()
'Set db = OpenDatabase("e:/gfolder/eb.mdb")
'Set rs = db.OpenRecordset("ebb")
End Sub

Public Function calculate()
Text7.Text = Val(Text5.Text) - Val(Text6.Text)
Text8.Text = Val(Text7.Text) * 0.075
End Function

Private Sub Text5_Change()
Call calculate
End Sub

Private Sub Text6_Change()
Call calculate
End Sub
