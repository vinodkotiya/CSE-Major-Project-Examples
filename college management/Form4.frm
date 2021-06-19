VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Student Admission"
   ClientHeight    =   5505
   ClientLeft      =   4305
   ClientTop       =   2790
   ClientWidth     =   7215
   LinkTopic       =   "Form4"
   ScaleHeight     =   5505
   ScaleWidth      =   7215
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton edit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   3720
      TabIndex        =   21
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton delete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   5040
      TabIndex        =   22
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton update 
      Caption         =   "&Update"
      Height          =   375
      Left            =   2400
      TabIndex        =   20
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton add 
      Caption         =   "&Add"
      Height          =   375
      Left            =   1080
      TabIndex        =   19
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton last 
      Caption         =   "&Last"
      Height          =   375
      Left            =   5040
      TabIndex        =   18
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton back 
      Caption         =   "&Back"
      Height          =   375
      Left            =   3720
      TabIndex        =   17
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton next 
      Caption         =   "&Next"
      Height          =   375
      Left            =   2400
      TabIndex        =   16
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton first 
      Caption         =   "&First"
      Height          =   375
      Left            =   1080
      TabIndex        =   15
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   5385
      TabIndex        =   14
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   1545
      TabIndex        =   13
      Top             =   3600
      Width           =   1215
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   5385
      TabIndex        =   12
      Top             =   2760
      Width           =   1215
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   825
      TabIndex        =   11
      Top             =   3120
      Width           =   1215
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "Form4.frx":0000
      Left            =   825
      List            =   "Form4.frx":000D
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   2760
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1665
      TabIndex        =   8
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5385
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   5385
      TabIndex        =   9
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1665
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   5385
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1665
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   5385
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1665
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   5385
      TabIndex        =   1
      Top             =   330
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   3855
      Left            =   105
      Top             =   120
      Width           =   6975
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Stream"
      Height          =   195
      Left            =   3945
      TabIndex        =   37
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Last Qualification"
      Height          =   195
      Left            =   225
      TabIndex        =   36
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Section"
      Height          =   195
      Left            =   3945
      TabIndex        =   35
      Top             =   2760
      Width           =   540
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Year"
      Height          =   195
      Left            =   225
      TabIndex        =   34
      Top             =   3120
      Width           =   330
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Course"
      Height          =   195
      Left            =   225
      TabIndex        =   33
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Nationality"
      Height          =   195
      Left            =   3945
      TabIndex        =   32
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Marital Status"
      Height          =   195
      Left            =   225
      TabIndex        =   31
      Top             =   2400
      Width           =   960
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Sex"
      Height          =   195
      Left            =   3945
      TabIndex        =   30
      Top             =   1800
      Width           =   270
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Date of Birth"
      Height          =   195
      Left            =   225
      TabIndex        =   29
      Top             =   1800
      Width           =   885
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Permanent Address"
      Height          =   195
      Left            =   3945
      TabIndex        =   28
      Top             =   1080
      Width           =   1380
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Postal Address"
      Height          =   195
      Left            =   225
      TabIndex        =   27
      Top             =   1080
      Width           =   1050
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Mother's Name"
      Height          =   195
      Left            =   3945
      TabIndex        =   26
      Top             =   720
      Width           =   1065
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Father's Name"
      Height          =   195
      Left            =   225
      TabIndex        =   25
      Top             =   720
      Width           =   1020
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Student's - Id"
      Height          =   195
      Left            =   3945
      TabIndex        =   24
      Top             =   360
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Student's Name"
      Height          =   195
      Left            =   225
      TabIndex        =   23
      Top             =   360
      Width           =   1125
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub add_Click()
rs2.AddNew
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Combo1.Text = ""
Combo4.Text = ""
Combo5.Text = ""
Text1.SetFocus
update.Enabled = True
End Sub

Private Sub back_Click()
On Error GoTo err
rs2.MovePrevious
If rs2.BOF Then
   rs2.MoveFirst
End If
ShowStdRecord
Exit Sub
err:
MsgBox "No Records"
End Sub

Private Sub delete_Click()
On Error GoTo err
rs2.delete
rs2.MoveNext
If rs2.EOF Then
   rs2.MoveLast
End If
If rs2.BOF Then
   rs2.MoveFirst
End If
MsgBox "Record Deleted", , "Student"
ShowStdRecord
Exit Sub
err:
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Text8.Text = ""
Combo3.Text = ""
Combo4.Text = ""
Combo5.Text = ""
Text9.Text = ""
Text10.Text = ""
MsgBox "No Records", , "Student"
End Sub

Private Sub edit_Click()
rs2.edit
Text1.SetFocus
update.Enabled = True
End Sub

Private Sub first_Click()
On Error GoTo err
rs2.MoveFirst
ShowStdRecord
Exit Sub
err:
MsgBox "No Records"
End Sub

Private Sub Form_Load()
rs2.MoveFirst
ShowStdRecord
End Sub

Private Sub last_Click()
On Error GoTo err
rs2.MoveLast
ShowStdRecord
Exit Sub
err:
MsgBox "No Records"
End Sub

Private Sub next_Click()
On Error GoTo err
rs2.MoveNext
If rs2.EOF Then
   rs2.MoveLast
End If
ShowStdRecord
Exit Sub
err:
MsgBox "No Records"
End Sub

Private Sub update_Click()
rs2.update
rs2.Fields(0).Value = Text1.Text
rs2.Fields(1).Value = Text2.Text
rs2.Fields(2).Value = Text3.Text
rs2.Fields(3).Value = Text4.Text
rs2.Fields(4).Value = Text5.Text
rs2.Fields(5).Value = Text6.Text
rs2.Fields(6).Value = Text7.Text
rs2.Fields(7).Value = Combo1.Text
rs2.Fields(8).Value = Combo2.Text
rs2.Fields(9).Value = Text8.Text
rs2.Fields(10).Value = Combo3.Text
rs2.Fields(11).Value = Combo4.Text
rs2.Fields(12).Value = Combo5.Text
rs2.Fields(13).Value = Text9.Text
rs2.Fields(14).Value = Text10.Text
rs2.update
MsgBox "Record Updated", , "Student"
rs2.MoveLast
update.Enabled = False
End Sub

Public Sub ShowStdRecord()
Text1.Text = rs2.Fields(0).Value
Text2 = rs2.Fields(1).Value
Text3.Text = rs2.Fields(2).Value
Text4.Text = rs2.Fields(3).Value
Text5.Text = rs2.Fields(4).Value
Text6.Text = rs2.Fields(5).Value
Text7.Text = rs2.Fields(6).Value
Combo1.Text = rs2.Fields(7).Value
Combo2.Text = rs2.Fields(8).Value
Text8.Text = rs2.Fields(9).Value
Combo3.Text = rs2.Fields(10).Value
Combo4.Text = rs2.Fields(11).Value
Combo5.Text = rs2.Fields(12).Value
Text9.Text = rs2.Fields(13).Value
Text10.Text = rs2.Fields(14).Value
End Sub
