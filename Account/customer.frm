VERSION 5.00
Begin VB.Form form4 
   Caption         =   "Form4"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   6480
      TabIndex        =   24
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   9600
      Top             =   6360
   End
   Begin VB.CommandButton cmclear 
      Caption         =   "clear"
      Height          =   495
      Left            =   8880
      TabIndex        =   22
      Top             =   4680
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   5760
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdlast 
      Caption         =   "last record"
      Height          =   495
      Left            =   8880
      TabIndex        =   21
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdprev 
      Caption         =   "previous"
      Height          =   495
      Left            =   8880
      TabIndex        =   20
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdnextfo 
      Caption         =   "next form"
      Height          =   495
      Left            =   8880
      TabIndex        =   19
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdfirst 
      Caption         =   "first record"
      Height          =   495
      Left            =   8880
      TabIndex        =   18
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "save"
      Height          =   495
      Left            =   8880
      TabIndex        =   17
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "next"
      Height          =   495
      Left            =   8880
      TabIndex        =   16
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "exit"
      Height          =   495
      Left            =   8880
      TabIndex        =   15
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox txtdue 
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   4080
      Width           =   1335
   End
   Begin VB.TextBox txtlimit 
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox txtphone 
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txtadd 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox txtname 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox txtcode 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "amount paid after due period"
      Height          =   240
      Left            =   3480
      TabIndex        =   23
      Top             =   4200
      Width           =   3000
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "cust_code"
      Height          =   240
      Left            =   240
      TabIndex        =   14
      Top             =   1200
      Width           =   1080
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "cust_nm"
      Height          =   240
      Left            =   240
      TabIndex        =   13
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "cust_add"
      Height          =   240
      Left            =   240
      TabIndex        =   12
      Top             =   3000
      Width           =   960
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "cust_phone"
      Height          =   240
      Left            =   4320
      TabIndex        =   11
      Top             =   1080
      Width           =   1200
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "cr_limit"
      Height          =   240
      Left            =   4680
      TabIndex        =   10
      Top             =   3240
      Width           =   750
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "cr_rate"
      Height          =   240
      Left            =   4440
      TabIndex        =   9
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "amt_due"
      Height          =   240
      Left            =   360
      TabIndex        =   8
      Top             =   4200
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CUSTOMER MASTER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2280
      TabIndex        =   7
      Top             =   120
      Width           =   5340
   End
End
Attribute VB_Name = "form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Private Sub cmclear_Click()
txtcode.Text = ""
txtadd.Text = ""
txtname.Text = ""
txtphone.Text = ""
Combo1.Text = ""
txtlimit.Text = ""
txtdue.Text = ""
txtcode.SetFocus
End Sub

Private Sub cmdexit_Click()
x = MsgBox("YOU WANT TO QUIT", vbYesNo, message)
If x = vbYes Then
Unload Me
ElseIf x = vbNo Then
txtcode.SetFocus
End If
End Sub
Private Sub cmdfirst_Click()
rs.MoveFirst
txtcode.Text = rs!cust_code
txtadd.Text = rs!cust_add
txtname.Text = rs!cust_nm
txtphone.Text = rs!cust_phone
Combo1.Text = rs!cr_rate
txtlimit.Text = rs!cr_limit
txtdue.Text = rs!amt_due
Text1.Text = rs!amt_due_after


End Sub

Private Sub cmdlast_Click()
rs.MoveLast
txtcode.Text = rs!cust_code
txtadd.Text = rs!cust_add
txtname.Text = rs!cust_nm
txtphone.Text = rs!cust_phone
Combo1.Text = rs!cr_rate
txtlimit.Text = rs!cr_limit
txtdue.Text = rs!amt_due
Text1.Text = rs!amt_due_after


End Sub

Private Sub cmdnext_Click()
rs.MoveNext
If rs.EOF = True Then
x = MsgBox("THIS IS THE LAST RECORD", vbOKOnly, message)
rs.MoveLast
End If
txtcode.Text = rs!cust_code
txtadd.Text = rs!cust_add
txtname.Text = rs!cust_nm
txtphone.Text = rs!cust_phone
Combo1.Text = rs!cr_rate
txtlimit.Text = rs!cr_limit
txtdue.Text = rs!amt_due
Text1.Text = rs!amt_due_after


End Sub

Private Sub cmdnextfo_Click()
Form2.Show
End Sub

Private Sub cmdprev_Click()
rs.MovePrevious
If rs.BOF = True Then
x = MsgBox("THIS IS THE FIRST RECORD", vbOKOnly, message)
rs.MoveFirst
End If
txtcode.Text = rs!cust_code
txtadd.Text = rs!cust_add
txtname.Text = rs!cust_nm
txtphone.Text = rs!cust_phone
Combo1.Text = rs!cr_rate
txtlimit.Text = rs!cr_limit
txtdue.Text = rs!amt_due
Text1.Text = rs!amt_due_after

End Sub

Private Sub cmdsave_Click()
con.Execute ("insert into customer values('" & txtcode.Text & "','" & txtname.Text & "','" & txtadd.Text & "','" & txtphone.Text & "','" & Combo1.Text & "','" & txtlimit.Text & "','" & txtdue.Text & "','" & Text1.Text & "')")
con.Execute ("commit")
MsgBox ("data seved")

End Sub

Private Sub Combo1_click()
If Combo1.Text = "A" Then
txtlimit.Text = 100000
ElseIf Combo1.Text = "B" Then
txtlimit.Text = 60000
ElseIf Combo1.Text = "C" Then
txtlimit.Text = 40000
End If

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Form_Activate()
form4.Text1.Text = Form6.Text4.Text
End Sub

Private Sub Form_Load()
Combo1.AddItem "A"
Combo1.AddItem "B"
Combo1.AddItem "C"
Text1.Enabled = False
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
con.open ("Provider=MSDAORA.1;User ID=scott;password=tiger;Persist Security Info=False")
rs.ActiveConnection = con
rs.CursorType = adOpenDynamic
rs.open ("select * from customer")



End Sub

Private Sub Text1_gotfo()

End Sub

Private Sub Timer1_Timer()
'Label1.Caption = (Mid(Label1.Caption, 2, Len(Label1.Caption) - 1) + Mid(Label1.Caption, 1, 1))
End Sub




Private Sub txtcode_gotfocus()
form4.txtcode.Text = Form1.txtcust.Text
form4.txtname.Text = Form1.txtnam.Text

End Sub

Private Sub txtdue_gotfocus()
form4.txtdue.Text = Val(Form1.txtamt.Text) - Val(Form1.txtcash.Text)


End Sub

Private Sub txtphone_keypress(KeyAscii As Integer)
If Not Chr(KeyAscii) Like "#" Then
x = MsgBox("PLEASE ENTER NUMBER ONLY", vbOKOnly, "ERROR!")
KeyAscii = 0
txtphone.Text = ""
txtphone.SetFocus
End If
End Sub

Private Sub txtphone_gotFocus()
If txtadd.Text = "" Then
x = MsgBox("field should not be left blank", vbOKOnly, Error)
txtadd.SetFocus
End If

End Sub
