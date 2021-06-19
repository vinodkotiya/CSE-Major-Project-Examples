VERSION 5.00
Begin VB.Form frmAccCheck 
   Caption         =   "Account No.Checking"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton CommandCancle 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton CommandOk 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox txtName 
      Height          =   405
      Left            =   2040
      TabIndex        =   1
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox txtAcc_No 
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   4560
      X2              =   4560
      Y1              =   600
      Y2              =   2880
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   120
      X2              =   4560
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   120
      X2              =   120
      Y1              =   600
      Y2              =   2880
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   4560
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Account No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "frmAccCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Private Sub CommandCancle_Click()
Unload Me
frmmain1.Show
frmmain1.SetFocus
End Sub

Private Sub CommandOk_Click()
Dim SQL
Dim today
'today = Format(Now, "short date")
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\bankproject\bank.mdb;Persist Security Info=False"
Set Rs.ActiveConnection = db
SQL = "Select * from initial where acc_no = '" & txtAcc_No.Text & "' and name = '" & txtName.Text & "'"
'MsgBox SQL
Rs.Open SQL

If Rs.EOF Or Rs.BOF Then
    MsgBox "Invalid Account No or Name.....", vbCritical
    txtAcc_No.SetFocus
Else
    'MsgBox "Success", vbInformation
    frmTranDaily.Show
    frmTranDaily.txttrandate.Text = frmTranDaily.DTPicker1.Value
    today = Format(Now, "short date")
    frmTranDaily.txtaccno.Text = frmAccCheck.txtAcc_No.Text
    Unload Me
End If
Rs.Close
db.Close
Set db = Nothing
Set Rs = Nothing
End Sub




