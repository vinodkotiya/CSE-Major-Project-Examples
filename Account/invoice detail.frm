VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   Begin VB.CommandButton Command6 
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
      Left            =   7920
      TabIndex        =   22
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
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
      Left            =   7920
      TabIndex        =   21
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
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
      Left            =   7920
      TabIndex        =   20
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
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
      Left            =   7920
      TabIndex        =   19
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
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
      Left            =   7920
      TabIndex        =   18
      Top             =   5760
      Width           =   1215
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
      Left            =   6120
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
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
      Left            =   3600
      TabIndex        =   7
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox txt_no 
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
      Left            =   1440
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txt_date 
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
      Left            =   6000
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmd4 
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
      Left            =   7920
      TabIndex        =   13
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmd5 
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
      Left            =   7920
      TabIndex        =   12
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmd6 
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
      Left            =   7920
      TabIndex        =   11
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox txtamt 
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
      Left            =   2040
      TabIndex        =   3
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox txtnm 
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
      Left            =   6000
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox txtcode 
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
      Left            =   1560
      TabIndex        =   2
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "cash recieved"
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
      Left            =   4440
      TabIndex        =   17
      Top             =   3360
      Width           =   1485
   End
   Begin VB.Label Label5 
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
      Left            =   1920
      TabIndex        =   16
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "inv_date"
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
      TabIndex        =   15
      Top             =   1320
      Width           =   900
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "inv_no"
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
      Left            =   240
      TabIndex        =   14
      Top             =   1200
      Width           =   690
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "cust_code"
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
      Left            =   240
      TabIndex        =   10
      Top             =   2280
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "cust_nm"
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
      Left            =   4680
      TabIndex        =   9
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "tot_amt"
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
      Top             =   3360
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "invoice  detail  form"
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
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   4725
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub cmd4_Click()
con.Execute ("insert into invoice_detail values('" & txt_no.Text & "','" & txt_date.Text & "','" & txtcode.Text & "','" & txtnm.Text & "','" & txtamt.Text & "','" & Text2.Text & "','" & Text1.Text & "')")
con.Execute ("commit")
MsgBox "data saved"
End Sub

Private Sub cmd5_Click()
txt_no.Text = ""
txt_date.Text = ""
txtcode.Text = ""
txtnm.Text = ""
txtamt.Text = ""

Text2.Text = ""
Text1.Text = ""
txt_no.SetFocus
End Sub

Private Sub cmd6_Click()
x = MsgBox("YOU WANT TO QUIT", vbYesNo, message)
If x = vbYes Then
Unload Me
ElseIf x = vbNo Then
txtno.SetFocus
End If
End Sub

Private Sub Command1_Click()
Form6.Show
End Sub
Private Sub Combo1_click()
If Combo1.Text = "A" Then
txtper.Text = "15 days"
ElseIf Combo1.Text = "B" Then
txtper.Text = "25 days"
ElseIf Combo1.Text = "C" Then
txtper.Text = "30 days"
End If
End Sub
Private Sub Command3_Click()
rs.MoveNext
If rs.EOF = True Then
x = MsgBox("THIS IS THE LAST RECORD", vbOKOnly, message)
rs.MoveLast
End If
txt_no.Text = rs!inv_no
txt_date.Text = rs!inv_date
txtcode.Text = rs!cust_code
txtnm.Text = rs!cust_name
txtamt.Text = rs!tot_amt
Text2.Text = rs!cash_rec
Text1.Text = rs!amt_due

End Sub

Private Sub Command4_Click()
rs.MovePrevious
If rs.BOF = True Then
x = MsgBox("THIS IS THE FIRST RECORD", vbOKOnly, message)
rs.MoveFirst
End If
txt_no.Text = rs!inv_no
txt_date.Text = rs!inv_date
txtcode.Text = rs!cust_code
txtnm.Text = rs!cust_name
txtamt.Text = rs!tot_amt
Text2.Text = rs!cash_rec
Text1.Text = rs!amt_due
End Sub

Private Sub Command5_Click()
rs.MoveLast
txt_no.Text = rs!inv_no
txt_date.Text = rs!inv_date
txtcode.Text = rs!cust_code
txtnm.Text = rs!cust_name
txtamt.Text = rs!tot_amt
Text2.Text = rs!cash_rec
Text1.Text = rs!amt_due

End Sub

Private Sub Command6_Click()
rs.MoveFirst
txt_no.Text = rs!inv_no
txt_date.Text = rs!inv_date
txtcode.Text = rs!cust_code
txtnm.Text = rs!cust_name
txtamt.Text = rs!tot_amt
Text2.Text = rs!cash_rec
Text1.Text = rs!amt_due

End Sub
Private Sub Form_Load()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
con.open ("Provider=MSDAORA.1;User ID=scott;password=tiger;Persist Security Info=False")
rs.ActiveConnection = con
rs.CursorType = adOpenDynamic
rs.open ("select * from invoice_detail")

End Sub

Private Sub Text1_gotfocus()
Form2.Text1.Text = form4.txtdue.Text

End Sub

Private Sub Text2_gotfocus()
Form2.Text2.Text = Form1.txtcash.Text
End Sub

Private Sub txt_date_gotfocus()
Form2.txt_date.Text = Form1.txtdate.Text


End Sub

Private Sub txt_no_gotfocus()
Form2.txt_no.Text = Form1.txtno.Text

End Sub

Private Sub txtamt_gotfocus()
Form2.txtamt.Text = Form1.txtamt.Text
End Sub

Private Sub txtcode_gotfocus()
Form2.txtcode.Text = Form1.txtcust.Text
Form2.txtnm.Text = Form1.txtnam.Text
End Sub
