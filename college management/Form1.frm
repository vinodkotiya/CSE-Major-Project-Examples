VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Course Information"
   ClientHeight    =   4590
   ClientLeft      =   3750
   ClientTop       =   2580
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo3 
      Height          =   288
      ItemData        =   "Form1.frx":0000
      Left            =   5760
      List            =   "Form1.frx":000A
      Sorted          =   -1  'True
      TabIndex        =   25
      Top             =   360
      Width           =   1212
   End
   Begin VB.CommandButton Last 
      Caption         =   "&Last"
      Height          =   372
      Left            =   5400
      TabIndex        =   11
      Top             =   3480
      Width           =   972
   End
   Begin VB.CommandButton Back 
      Caption         =   "&Back"
      Height          =   372
      Left            =   4200
      TabIndex        =   10
      Top             =   3480
      Width           =   972
   End
   Begin VB.CommandButton Next 
      Caption         =   "&Next"
      Height          =   372
      Left            =   2880
      TabIndex        =   9
      Top             =   3480
      Width           =   972
   End
   Begin VB.CommandButton First 
      Caption         =   "&First"
      Height          =   372
      Left            =   1560
      TabIndex        =   8
      Top             =   3480
      Width           =   972
   End
   Begin VB.CommandButton Update 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   372
      Left            =   2880
      TabIndex        =   13
      Top             =   3960
      Width           =   972
   End
   Begin VB.TextBox Text7 
      Height          =   288
      Left            =   2520
      TabIndex        =   6
      Top             =   2760
      Width           =   1212
   End
   Begin VB.TextBox Text6 
      Height          =   288
      Left            =   5760
      TabIndex        =   7
      Text            =   "as"
      Top             =   2760
      Width           =   1212
   End
   Begin VB.TextBox Text5 
      Height          =   288
      Left            =   5760
      TabIndex        =   5
      Top             =   2040
      Width           =   1212
   End
   Begin VB.CommandButton Delete 
      Caption         =   "&Delete"
      Height          =   372
      Left            =   5400
      TabIndex        =   15
      Top             =   3960
      Width           =   972
   End
   Begin VB.CommandButton Edit 
      Caption         =   "&Edit"
      Height          =   372
      Left            =   4200
      TabIndex        =   14
      Top             =   3960
      Width           =   972
   End
   Begin VB.CommandButton Add 
      Caption         =   "&Add"
      Height          =   372
      Left            =   1560
      TabIndex        =   12
      Top             =   3960
      Width           =   972
   End
   Begin VB.TextBox Text4 
      Height          =   288
      Left            =   2520
      TabIndex        =   4
      Top             =   2160
      Width           =   1212
   End
   Begin VB.TextBox Text3 
      Height          =   288
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   1212
   End
   Begin VB.ComboBox Combo2 
      Height          =   288
      ItemData        =   "Form1.frx":0016
      Left            =   5760
      List            =   "Form1.frx":0026
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1440
      Width           =   1212
   End
   Begin VB.ComboBox Combo1 
      Height          =   288
      ItemData        =   "Form1.frx":0056
      Left            =   2520
      List            =   "Form1.frx":0060
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1440
      Width           =   1212
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   2520
      TabIndex        =   0
      Top             =   360
      Width           =   1212
   End
   Begin VB.Shape Shape1 
      Height          =   3012
      Left            =   360
      Top             =   240
      Width           =   6972
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Course Admission Date"
      Height          =   192
      Index           =   8
      Left            =   600
      TabIndex        =   24
      Top             =   2880
      Width           =   1692
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Course Session Date"
      Height          =   192
      Index           =   7
      Left            =   4080
      TabIndex        =   23
      Top             =   2880
      Width           =   1524
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Course Fees"
      Height          =   192
      Index           =   6
      Left            =   4080
      TabIndex        =   22
      Top             =   2160
      Width           =   924
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Course Seats"
      Height          =   192
      Index           =   5
      Left            =   600
      TabIndex        =   21
      Top             =   2160
      Width           =   972
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Course Eligiblity"
      Height          =   192
      Index           =   4
      Left            =   4080
      TabIndex        =   20
      Top             =   1560
      Width           =   1152
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Course Duration"
      Height          =   192
      Index           =   3
      Left            =   600
      TabIndex        =   19
      Top             =   1560
      Width           =   1152
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Course Name"
      Height          =   192
      Index           =   2
      Left            =   600
      TabIndex        =   18
      Top             =   840
      Width           =   996
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Course Type"
      Height          =   192
      Index           =   1
      Left            =   4080
      TabIndex        =   17
      Top             =   360
      Width           =   936
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Course Id"
      Height          =   192
      Index           =   0
      Left            =   600
      TabIndex        =   16
      Top             =   360
      Width           =   684
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub add_Click()
rs1.AddNew
Text1.Text = ""
Combo3.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Text1.SetFocus
update.Enabled = True
End Sub

Private Sub Close_Click()
Unload Me
End Sub

Private Sub back_Click()
On Error GoTo err
rs1.MovePrevious
If rs1.BOF Then
   rs1.MoveFirst
End If
ShowRecord
Exit Sub
err:
MsgBox "No Records"
End Sub

Private Sub delete_Click()
On Error GoTo err
rs1.delete
rs1.MoveNext
If rs1.EOF Then
   rs1.MoveLast
End If
If rs1.BOF Then
   rs1.MoveFirst
End If
MsgBox "Record Deleted", , "Course"
ShowRecord
Exit Sub
err:
Text1.Text = ""
Combo3.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Combo1.Text = ""
Combo2.Text = ""
MsgBox "No Records", , "Course"
End Sub

Private Sub edit_Click()
rs1.edit
Text1.SetFocus
update.Enabled = True
End Sub

Private Sub first_Click()
On Error GoTo err
rs1.MoveFirst
ShowRecord
Exit Sub
err:
MsgBox "No Records"
End Sub

Private Sub Form_Load()
rs1.MoveFirst
ShowRecord
End Sub

Private Sub last_Click()
On Error GoTo err
rs1.MoveLast
ShowRecord
Exit Sub
err:
MsgBox "No Records"
End Sub

Private Sub next_Click()
On Error GoTo err
rs1.MoveNext
If rs1.EOF Then
   rs1.MoveLast
End If
ShowRecord
Exit Sub
err:
MsgBox "No Records"
End Sub

Private Sub update_Click()
rs1.Fields(0).Value = Text1.Text
rs1.Fields(1).Value = Combo3.Text
rs1.Fields(2).Value = Text3.Text
rs1.Fields(3).Value = Combo1.Text
rs1.Fields(4).Value = Combo2.Text
rs1.Fields(5).Value = Text4.Text
rs1.Fields(6).Value = Text5.Text
rs1.Fields(7).Value = Text6.Text
rs1.Fields(8).Value = Text7.Text
rs1.update
MsgBox "Record Updated", , "Course"
rs1.MoveLast
update.Enabled = False
End Sub

Public Sub ShowRecord()
Text1.Text = rs1.Fields(0).Value
Combo3.Text = rs1.Fields(1).Value
Text3.Text = rs1.Fields(2).Value
Combo1.Text = rs1.Fields(3).Value
Combo2.Text = rs1.Fields(4).Value
Text4.Text = rs1.Fields(5).Value
Text5.Text = rs1.Fields(6).Value
Text6.Text = rs1.Fields(7).Value
Text7.Text = rs1.Fields(8).Value
End Sub
