VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "PRODUCT"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdlast 
      Caption         =   "last"
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
      Left            =   6240
      TabIndex        =   13
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdfirst 
      Caption         =   "first"
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
      Left            =   6240
      TabIndex        =   12
      Top             =   1680
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
      Left            =   6240
      TabIndex        =   11
      Top             =   2880
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
      Left            =   6240
      TabIndex        =   10
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   8160
      Top             =   6480
   End
   Begin VB.CommandButton cmdclear 
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
      Left            =   6240
      TabIndex        =   9
      Top             =   4680
      Width           =   1215
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
      Left            =   2520
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdnextfrm 
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
      Left            =   6240
      TabIndex        =   8
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdexit 
      Appearance      =   0  'Flat
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
      Left            =   6240
      TabIndex        =   7
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox txtprice 
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
      Left            =   2640
      TabIndex        =   3
      Top             =   3240
      Width           =   1455
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
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "product_code"
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
      TabIndex        =   6
      Top             =   1560
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "description"
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
      TabIndex        =   5
      Top             =   2400
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "product_price"
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
      TabIndex        =   4
      Top             =   3360
      Width           =   1440
   End
   Begin VB.Label Label1 
      Caption         =   "PRODUCT FORM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   360
      Width           =   5175
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Private Sub cmdexit_Click()
x = MsgBox("YOU WANT TO QUIT", vbYesNo, message)
If x = vbYes Then
Unload Me
ElseIf x = vbNo Then
Combo1.SetFocus
End If
End Sub

Private Sub cmdfirst_Click()
rs.MoveFirst

Combo1.Text = rs!product_code
txtdes.Text = rs!Description
txtprice.Text = rs!product_price
End Sub

Private Sub cmdlast_Click()
rs.MoveLast
Combo1.Text = rs!product_code
txtdes.Text = rs!Description
txtprice.Text = rs!product_price

End Sub

Private Sub cmdnext_Click()

rs.MoveNext
If rs.EOF = True Then
x = MsgBox("THIS IS THE LAST RECORD", vbOKOnly, message)
rs.MoveLast
End If
Combo1.Text = rs!product_code
txtdes.Text = rs!Description
txtprice.Text = rs!product_price

End Sub

Private Sub cmdnextfrm_Click()
Form1.Show
End Sub
'Private Sub Timer1_Timer()
'Label1.Caption = (Mid(Label1.Caption, 2, Len(Label1.Caption) - 1) + Mid(Label1.Caption, 1, 1))
'End Sub

Private Sub cmdclear_Click()
Combo1.Text = ""
txtdes.Text = ""
txtprice.Text = ""
Combo1.SetFocus
End Sub

Private Sub cmdprev_Click()
rs.MovePrevious
If rs.BOF = True Then
x = MsgBox("THIS IS THE FIRST RECORD", vbOKOnly, message)
rs.MoveFirst


End If
Combo1.Text = rs!product_code
txtdes.Text = rs!Description
txtprice.Text = rs!product_price
End Sub



Private Sub Combo1_click()
If Combo1.Text = "P001" Then
txtdes.Text = "AIR CONDITION"
    If txtdes.Text = "AIR CONDITION" Then
    txtprice.Text = 40000
    End If
ElseIf Combo1.Text = "P002" Then
txtdes.Text = "T.V"
    If txtdes.Text = "T.V" Then
    txtprice.Text = 25000
    End If
ElseIf Combo1.Text = "P003" Then
txtdes.Text = "COMPUTER"
    If txtdes.Text = "COMPUTER" Then
    txtprice.Text = 42000
    End If
ElseIf Combo1.Text = "P004" Then
txtdes.Text = "MUSIC SYSTEM"
    If txtdes.Text = "MUSIC SYSTEM" Then
    txtprice.Text = 30000
    End If
ElseIf Combo1.Text = "P005" Then
txtdes.Text = "VEDIO C.D. PLAYER"
    If txtdes.Text = "VEDIO C.D. PLAYER" Then
    txtprice.Text = 40000
    End If

End If
cmdnextfrm.Enabled = True
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub Form_Load()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
con.open ("Provider=MSDAORA.1;User ID=scott;password=tiger;Persist Security Info=False")
rs.ActiveConnection = con
rs.CursorType = adOpenDynamic
rs.open ("select * from product")

Combo1.AddItem "P001"
Combo1.AddItem "P002"
Combo1.AddItem "P003"
Combo1.AddItem "P004"
Combo1.AddItem "P005"
cmdnextfrm.Enabled = False
End Sub

Private Sub txtdes_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub txtprice_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


