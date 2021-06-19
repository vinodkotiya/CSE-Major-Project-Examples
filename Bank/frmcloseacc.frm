VERSION 5.00
Begin VB.Form frmcloseacc 
   Caption         =   "Form1"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5220
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcancle 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdcloseacc 
      Caption         =   "Close "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox txtname 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox txtaccno 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   960
      Width           =   1815
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   120
      X2              =   5040
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   5040
      X2              =   5040
      Y1              =   600
      Y2              =   2760
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   120
      X2              =   120
      Y1              =   600
      Y2              =   2760
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   5040
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
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Account No"
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
      Left            =   600
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
End
Attribute VB_Name = "frmcloseacc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As New ADODB.Connection
Dim Rs As New ADODB.Recordset

Private Sub cmdcancle_Click()
Unload Me
End Sub

Private Sub cmdcloseacc_Click()
Dim balance1 As Long
Dim cb1 As Long
Dim acc1 As String
Dim name1 As String
Dim address1 As String
MsgBox " Are you sure want to Close the Account"
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\bankproject\bank.mdb;Persist Security Info=False"
Set Rs.ActiveConnection = db


 
SQL = "Select * from initial where acc_no = '" & txtaccno.Text & "' and name = '" & txtname.Text & "'"
'MsgBox SQL
Rs.Open SQL
'If Rs.BOF Or Rs.EOF Then


If Rs.EOF Or Rs.BOF Then
    MsgBox "Invalid Account No or Name.....", vbCritical
    txtaccno.SetFocus
Else
    'MsgBox "Success", vbInformation
    
acc1 = Rs.Fields("acc_no")
name1 = Rs.Fields("name")
address1 = Rs.Fields("address")
balance1 = Rs.Fields("balance")
cb1 = Rs.Fields("currentbalance")
'MsgBox "cb " & cb1
     
 SQL1 = " delete from initial where acc_no = '" & txtaccno.Text & "'"
 db.Execute SQL1
 SQL2 = "delete from tran where acc_no='" & txtaccno.Text & " '"
 db.Execute SQL2
 MsgBox "Account No." & txtaccno.Text & " is Closed "
    
    SQL6 = "insert into delete1(acc_no, name, address, " & _
      "balance,currentbalance)" & _
      " values('" & _
      acc1 & "', '" & _
      name1 & "', '" & _
      address1 & "', '" & _
      balance1 & "', '" & _
      cb1 & "')"
        db.Execute SQL6
    Unload Me
End If
Rs.Close
db.Close
Set db = Nothing
Set Rs = Nothing

End Sub

