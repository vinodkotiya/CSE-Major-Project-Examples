VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton Command2 
      Caption         =   "exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   26
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   25
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   24
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdfirst 
      Caption         =   "first record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   23
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdnextfo 
      Caption         =   "next form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   22
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdprev 
      Caption         =   "previous"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   21
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdlast 
      Caption         =   "last record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   20
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   19
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   16
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   480
      Top             =   120
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   5
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtper 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox txtint 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox txtlimit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2160
      TabIndex        =   2
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "cr_int @5%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      TabIndex        =   18
      Top             =   4440
      Width           =   1170
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "cr_int @4%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      TabIndex        =   17
      Top             =   4440
      Width           =   1050
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "total amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4080
      TabIndex        =   15
      Top             =   4560
      Width           =   1275
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "amount due"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4080
      TabIndex        =   14
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "INTEREST CACULATION FORM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   13
      Top             =   120
      Width           =   8055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "cust name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4560
      TabIndex        =   12
      Top             =   1320
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "cust code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   11
      Top             =   1200
      Width           =   1020
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "cr_limit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4560
      TabIndex        =   10
      Top             =   2400
      Width           =   750
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "cr_int @2%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      TabIndex        =   9
      Top             =   4440
      Width           =   1170
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "cr_period"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   8
      Top             =   3480
      Width           =   1005
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "cr_rating"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      TabIndex        =   7
      Top             =   2400
      Width           =   915
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub cmdfirst_Click()
rs.MoveFirst
Text1.Text = rs!cust_code
Combo1.Text = rs!cr_rating
txtint.Text = rs!cr_int
txtper.Text = rs!cr_period
Text2.Text = rs!cust_name
txtlimit.Text = rs!cr_limit
Text3.Text = rs!amt_due
Text4.Text = rs!tot_amt
End Sub

Private Sub cmdlast_Click()
rs.MoveLast
Text1.Text = rs!cust_code
Combo1.Text = rs!cr_rating
txtint.Text = rs!cr_int
txtper.Text = rs!cr_period
Text2.Text = rs!cust_name
txtlimit.Text = rs!cr_limit
Text3.Text = rs!amt_due
Text4.Text = rs!tot_amt
End Sub

Private Sub cmdnext_Click()
rs.MoveNext
If rs.EOF = True Then
x = MsgBox("THIS IS THE LAST RECORD", vbOKOnly, message)
rs.MoveLast
End If
Text1.Text = rs!cust_code
Combo1.Text = rs!cr_rating
txtint.Text = rs!cr_int
txtper.Text = rs!cr_period
Text2.Text = rs!cust_name
txtlimit.Text = rs!cr_limit
Text3.Text = rs!amt_due
Text4.Text = rs!tot_amt
End Sub

Private Sub cmdnextfo_Click()
Form3.Show
End Sub

Private Sub cmdprev_Click()
rs.MovePrevious
If rs.BOF = True Then
x = MsgBox("THIS IS THE FIRST RECORD", vbOKOnly, message)
rs.MoveFirst
End If
Text1.Text = rs!cust_code
Combo1.Text = rs!cr_rating
txtint.Text = rs!cr_int
txtper.Text = rs!cr_period
Text2.Text = rs!cust_name
txtlimit.Text = rs!cr_limit
Text3.Text = rs!amt_due
Text4.Text = rs!tot_amt
End Sub

Private Sub cmdsave_Click()
con.Execute ("insert into interest values('" & Text1.Text & "','" & Combo1.Text & "','" & txtint.Text & "','" & txtper.Text & "','" & Text2.Text & "','" & txtlimit.Text & "','" & Text3.Text & "','" & Text4.Text & "')")
con.Execute ("commit")
MsgBox "data saved"

End Sub

Private Sub Combo1_GotFocus()
Form6.Combo1.Text = form4.Combo1.Text
If Combo1.Text = "A" Then
txtlimit.Text = 100000
ElseIf Combo1.Text = "B" Then
txtlimit.Text = 60000
ElseIf Combo1.Text = "C" Then
txtlimit.Text = 40000
End If
If Combo1.Text = "A" Then
txtper.Text = "15 days"
ElseIf Combo1.Text = "B" Then
txtper.Text = "25 days"
ElseIf Combo1.Text = "C" Then
txtper.Text = "30 days"
End If
KeyAscii = 0
End Sub

Private Sub Command1_Click()
Text1.Text = ""
Combo1.Text = ""
txtint.Text = ""
txtper.Text = ""
Text2.Text = ""
txtlimit.Text = ""
Text3.Text = ""
Text4.Text = ""
Text1.SetFocus
End Sub

Private Sub Command2_Click()
x = MsgBox("YOU WANT TO QUIT", vbYesNo, message)
If x = vbYes Then
Unload Me
ElseIf x = vbNo Then
Text1.SetFocus
End If
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
con.open ("Provider=MSDAORA.1;User ID=scott;password=tiger;Persist Security Info=False")
rs.ActiveConnection = con
rs.CursorType = adOpenDynamic
rs.open ("select * from interest")
End Sub

Private Sub Text1_gotfocus()
Form6.Text1.Text = Form1.txtcust.Text
Form6.Text2.Text = Form1.txtnam.Text
End Sub

Private Sub Text3_gotfocus()
Form6.Text3.Text = Form2.Text1.Text
End Sub

Private Sub Text4_gotfocus()
Text4.Text = Val(Text3.Text) + Val(txtint.Text)
End Sub

Private Sub Timer1_Timer()
'Label3.Caption = (Mid(Label3.Caption, 2, Len(Label3.Caption) - 1) + Mid(Label3.Caption, 1, 1))
End Sub

Private Sub txtint_gotfocus()
If Combo1.Text = "A" Then
Label6.Visible = True
Label10.Visible = False
Label11.Visible = False
ElseIf Combo1.Text = "B" Then
Label6.Visible = False
Label10.Visible = True
Label11.Visible = False
ElseIf Combo1.Text = "C" Then
Label6.Visible = False
Label10.Visible = False
Label11.Visible = True
End If
If Combo1.Text = "A" Then
txtint.Text = 0.02 * Val(Text3.Text)
ElseIf Combo1.Text = "B" Then
txtint.Text = 0.04 * Val(Text3.Text)
ElseIf Combo1.Text = "C" Then
txtint.Text = 0.05 * Val(Text3.Text)
End If
End Sub

