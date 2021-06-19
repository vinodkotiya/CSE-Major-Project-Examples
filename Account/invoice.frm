VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   Begin VB.CommandButton cmclear 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   7440
      TabIndex        =   32
      Top             =   6960
      Width           =   1245
   End
   Begin VB.CommandButton cmdnextfo 
      Caption         =   "NEXT FORM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   8760
      TabIndex        =   31
      Top             =   6960
      Width           =   1245
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   10080
      TabIndex        =   30
      Top             =   6960
      Width           =   1245
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6120
      TabIndex        =   29
      Top             =   6960
      Width           =   1245
   End
   Begin VB.CommandButton cmdlast 
      Caption         =   "LAST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4680
      TabIndex        =   28
      Top             =   6960
      Width           =   1245
   End
   Begin VB.CommandButton cmdprev 
      Caption         =   "PREVIOUS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3120
      TabIndex        =   27
      Top             =   6960
      Width           =   1395
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1800
      TabIndex        =   26
      Top             =   6960
      Width           =   1245
   End
   Begin VB.CommandButton cmdfirst 
      Caption         =   "FIRST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   360
      TabIndex        =   25
      Top             =   6960
      Width           =   1320
   End
   Begin VB.TextBox txtcash 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3360
      TabIndex        =   6
      Top             =   6510
      Width           =   1845
   End
   Begin VB.TextBox txtnounit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8460
      TabIndex        =   5
      Top             =   3750
      Width           =   1845
   End
   Begin VB.TextBox txtnam 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3360
      TabIndex        =   4
      Top             =   5520
      Width           =   1845
   End
   Begin VB.TextBox txttax 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8460
      TabIndex        =   10
      Top             =   5520
      Width           =   1845
   End
   Begin VB.TextBox txtinamt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8460
      TabIndex        =   9
      Top             =   4560
      Width           =   1845
   End
   Begin VB.TextBox txtunit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8460
      TabIndex        =   8
      Top             =   2850
      Width           =   1845
   End
   Begin VB.TextBox txtdes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8460
      TabIndex        =   7
      Top             =   2010
      Width           =   1845
   End
   Begin VB.TextBox txtcust 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3360
      TabIndex        =   3
      Top             =   4620
      Width           =   1845
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
      Height          =   405
      Left            =   8460
      TabIndex        =   11
      Top             =   6480
      Width           =   1845
   End
   Begin VB.TextBox txtprocode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3360
      TabIndex        =   2
      Top             =   3720
      Width           =   1845
   End
   Begin VB.TextBox txtdate 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3360
      TabIndex        =   1
      Top             =   2850
      Width           =   1845
   End
   Begin VB.TextBox txtno 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3360
      TabIndex        =   0
      Top             =   1980
      Width           =   1845
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "CASH RECEIVED"
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
      Left            =   900
      TabIndex        =   24
      Top             =   6600
      Width           =   1830
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "NO OF UNITS"
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
      Left            =   5460
      TabIndex        =   23
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "CUSTOMER NAME "
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
      Left            =   900
      TabIndex        =   22
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "INVOICE NO"
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
      Left            =   900
      TabIndex        =   21
      Top             =   1980
      Width           =   1305
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "INVOICE DATE"
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
      Left            =   900
      TabIndex        =   20
      Top             =   2910
      Width           =   1590
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "PRODUCT CODE"
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
      Left            =   900
      TabIndex        =   19
      Top             =   3870
      Width           =   1815
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "TOTAL AMOUNT(Rs.)"
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
      Left            =   5460
      TabIndex        =   18
      Top             =   6600
      Width           =   2280
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "CUSTOMER CODE"
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
      Left            =   900
      TabIndex        =   17
      Top             =   4740
      Width           =   1980
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "PRODUCT DESCRIPTION"
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
      Left            =   5460
      TabIndex        =   16
      Top             =   2070
      Width           =   2715
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "UNIT PRICE"
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
      Left            =   5460
      TabIndex        =   15
      Top             =   2910
      Width           =   1290
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "INVOICE AMOUNT"
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
      Left            =   5490
      TabIndex        =   14
      Top             =   4680
      Width           =   1950
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "SALES TAX(Rs.)"
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
      Left            =   5460
      TabIndex        =   13
      Top             =   5640
      Width           =   1725
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CUSTOMER INVOICE FORM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3540
      TabIndex        =   12
      Top             =   750
      Width           =   6735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Private Sub cmclear_Click()
    txtno.Text = ""
    txtdate.Text = ""
    txtprocode.Text = ""
    txtcust.Text = ""
    txtnam.Text = ""
    txtunit.Text = ""
    txtdes.Text = ""
    txtnounit.Text = ""
    txtinamt.Text = ""
    txttax.Text = ""
    txtamt.Text = ""
    txtcash.Text = ""
    txtno.SetFocus
End Sub

Private Sub cmd4_Click()
On Error Resume Next
    Set con = New ADODB.Connection
    con.open ("Provider=MSDAORA.1;User ID=scott;password=tiger;Persist Security Info=False")
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = con
    rs.open ("SELECT * FROM CINVOICE_DETAIL")
    rs.AddNew
         rs!inv_no = txtno.Text
         rs!inv_date = txtdate.Text
         rs!product_code = txtprocode.Text
         rs!cust_code = txtcust.Text
         rs!cust_name = txtnam.Text
         rs!product_desc = txtdes.Text
         rs!unit_price = txtunit.Text
         rs!no_units = txtnounit.Text
         rs!inv_amt = txtinamt.Text
         rs!sale_tax = txttax.Text
         rs!cash_rec = txtcash.Text
         rs!tot_amt = txtamt.Text
    rs.Update

End Sub

Private Sub cmdexit_Click()
        x = MsgBox("YOU WANT TO QUIT", vbYesNo, "MESSAGE")
        If x = vbYes Then
        Unload Me
        ElseIf x = vbNo Then
        txtno.SetFocus
        End If
End Sub

Private Sub cmdfirst_Click()
    On Error Resume Next
    rs.MoveFirst
        txtno.Text = rs!inv_no
        txtdate.Text = rs!inv_date
        txtprocode.Text = rs!product_code
        txtcust.Text = rs!cust_code
        txtnam.Text = rs!cust_name
        txtdes.Text = rs!product_desc
        txtunit.Text = rs!unit_price
        txtnounit.Text = rs!no_units
        txtinamt.Text = rs!inv_amt
        txttax.Text = rs!sale_tax
        txtcash.Text = rs!cash_rec
        txtamt.Text = rs!tot_amt

End Sub

Private Sub cmdlast_Click()
On Error Resume Next
    rs.MoveLast
        txtno.Text = rs!inv_no
        txtdate.Text = rs!inv_date
        txtprocode.Text = rs!product_code
        txtcust.Text = rs!cust_code
        txtnam.Text = rs!cust_name
        txtdes.Text = rs!product_desc
        txtunit.Text = rs!unit_price
        txtnounit.Text = rs!no_units
        txtinamt.Text = rs!inv_amt
        txttax.Text = rs!sale_tax
        txtcash.Text = rs!cash_rec
        txtamt.Text = rs!tot_amt

End Sub

Private Sub cmdnext_Click()
On Error Resume Next
    rs.MoveNext
        If rs.EOF = True Then
        x = MsgBox("THIS IS THE LAST RECORD", vbOKOnly, "MESSAGE")
        rs.MoveLast
        End If
    txtno.Text = rs!inv_no
    txtdate.Text = rs!inv_date
    txtprocode.Text = rs!product_code
    txtcust.Text = rs!cust_code
    txtnam.Text = rs!cust_name
    txtdes.Text = rs!product_desc
    txtunit.Text = rs!unit_price
    txtnounit.Text = rs!no_units
    txtinamt.Text = rs!inv_amt
    txttax.Text = rs!sale_tax
    txtcash.Text = rs!cash_rec
    txtamt.Text = rs!tot_amt
End Sub

Private Sub cmdnextfo_Click()
    form4.Show
End Sub
Private Sub cmdprev_Click()
On Error Resume Next
    rs.MovePrevious
        If rs.BOF = True Then
        x = MsgBox("THIS IS THE FIRST RECORD", vbOKOnly, "MESSAGE")
        rs.MoveFirst
        End If
    txtno.Text = rs!inv_no
    txtdate.Text = rs!inv_date
    txtprocode.Text = rs!product_code
    txtcust.Text = rs!cust_code
    txtnam.Text = rs!cust_name
    txtdes.Text = rs!product_desc
    txtunit.Text = rs!unit_price
    txtnounit.Text = rs!no_units
    txtinamt.Text = rs!inv_amt
    txttax.Text = rs!sale_tax
    txtcash.Text = rs!cash_rec
    txtamt.Text = rs!tot_amt
End Sub


Private Sub Form_Load()
    Set con = New ADODB.Connection
    Set rs = New ADODB.Recordset
    con.open ("Provider=MSDAORA.1;User ID=scott;password=tiger;Persist Security Info=False")
    rs.ActiveConnection = con
    rs.CursorType = adOpenDynamic
    rs.open ("select * from invoice_detail")
End Sub

Private Sub txtcash_GotFocus()
    If txtnounit.Text = "" Then
    x = MsgBox("field should not be left blank", vbOKOnly, "ERROR!")
    txtnounit.SetFocus
    End If
        txtinamt.Text = Val(txtunit.Text) * Val(txtnounit.Text)
        txttax.Text = 0.1 * Val(txtinamt)
        txtamt.Text = Val(txtinamt.Text) + Val(txttax.Text)
End Sub
Private Sub txtcash_LostFocus()
    If txtcash.Text = "" Then
    txtcash.Text = 0
    txtcash.SetFocus
    End If
End Sub

Private Sub txtcust_GotFocus()
    Form1.txtcust.Text = Form5.Text1.Text
End Sub

Private Sub txtcust_lostfocus()
    'If txtcust.Text = "" Then
   ' x = MsgBox("field should not be left blank", vbOKOnly, "ERROR!")
    'txtcust.SetFocus
   ' ElseIf Not txtcust.Text Like "C###" Then
   ' x = MsgBox("customer code should be like C001", vbOKOnly, "ERROR!")
   ' txtcust.Text = ""
   ' txtcust.SetFocus
  '  End If
 '
End Sub

Private Sub txtcust_KeyPress(KeyAscii As Integer)
    x = Chr(KeyAscii) Like "A-Z"
    If x = False Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtdate_GotFocus()
    If txtno.Text = "" Then
    x = MsgBox("field should not be left blank", vbOKOnly, "ERROR!")
    txtno.SetFocus
    ElseIf Not txtno.Text Like "I###" Then
    x = MsgBox("invoice code should be like I001", vbOKOnly, "ERROR!")
    txtno.Text = ""
    txtno.SetFocus
    End If
End Sub
Private Sub txtnam_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) Like "#" Then
    x = MsgBox("name should not be numeric", vbOKOnly, "ERROR!")
    KeyAscii = 0
    txtnam.Text = ""
    txtnam.SetFocus
    End If
End Sub
Private Sub txtno_KeyPress(KeyAscii As Integer)
    x = Chr(KeyAscii) Like "A-Z"
    If x = False Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub
Private Sub txtnounit_GotFocus()
    If txtnam.Text = "" Then
    x = MsgBox("field should not be left blank", vbOKOnly, "ERROR!")
    txtnam.SetFocus
    End If
End Sub
Private Sub txtnounit_KeyPress(KeyAscii As Integer)
    If Not Chr(KeyAscii) Like "#" Then
    x = MsgBox("PLEASE ENTER NUMBER ONLY", vbOKOnly, "ERROR!")
    KeyAscii = 0
    txtnounit.Text = ""
    txtnounit.SetFocus
    End If
End Sub
Private Sub txtprocode_gotfocus()
    If Not txtdate.Text Like "##/##/##" Then
    x = MsgBox("Enter in format of dd/mm/yy", vbOKOnly, "ERROR!")
    txtdate.Text = ""
    txtdate.SetFocus
    End If
        Form1.txtprocode.Text = Form3.Combo1.Text
            If txtprocode.Text = "P001" Then
            txtdes.Text = "AIR CONDITION"
                If txtdes.Text = "AIR CONDITION" Then
                txtunit.Text = 40000
                End If
            ElseIf txtprocode.Text = "P002" Then
            txtdes.Text = "T.V"
                If txtdes.Text = "T.V" Then
                txtunit.Text = 25000
                End If
            ElseIf txtprocode.Text = "P003" Then
            txtdes.Text = "COMPUTER"
                If txtdes.Text = "COMPUTER" Then
                txtunit.Text = 42000
                End If
            ElseIf txtprocode.Text = "P004" Then
            txtdes.Text = "MUSIC SYSTEM"
                If txtdes.Text = "MUSIC SYSTEM" Then
               txtunit.Text = 30000
                End If
            ElseIf txtprocode.Text = "P005" Then
            txtdes.Text = "VEDIO C.D. PLAYER"
                If txtdes.Text = "VEDIO C.D. PLAYER" Then
                txtunit.Text = 40000
                End If
            End If
End Sub

Private Sub txtprocode_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


