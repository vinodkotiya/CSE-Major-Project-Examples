VERSION 5.00
Begin VB.Form payform 
   BackColor       =   &H00FFC0C0&
   Caption         =   "payment form"
   ClientHeight    =   5550
   ClientLeft      =   2205
   ClientTop       =   2025
   ClientWidth     =   7530
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   7530
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcancle 
      Caption         =   "CANCEL"
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
      Left            =   3218
      TabIndex        =   6
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton CMDSAVE 
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
      Height          =   495
      Left            =   938
      TabIndex        =   5
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton CMDEXIT 
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
      Height          =   495
      Left            =   5498
      TabIndex        =   7
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   2895
      Left            =   878
      TabIndex        =   8
      Top             =   1320
      Width           =   5775
      Begin VB.TextBox Txtdate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3000
         TabIndex        =   3
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox Txtamt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   2
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox Txtreason 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   1
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtpaidto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   0
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   12
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   11
         Top             =   1760
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Reason for payment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   10
         Top             =   1240
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Paid to"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   9
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Payment form"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   975
      Left            =   2018
      TabIndex        =   4
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "payform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim d, m As Integer

Private Sub Cmdcancle_Click()
txtpaidto = ""
Txtamt = ""
Txtdate = ""
Txtreason = ""

End Sub

Private Sub CMDEXIT_Click()
main.Enabled = True

If rs.State = 1 Then rs.Close
If cn.State = 1 Then cn.Close

Unload Me
Load main
main.Show

End Sub

Private Sub Cmdsave_Click()
If txtpaidto.Text = "" Or Txtamt.Text = "" Or Txtreason.Text = "" Then
MsgBox "PLEASE ENTER THE VALUES", vbOKOnly, "USER INFORMATION"
txtpaidto.SetFocus
Exit Sub
End If
rs.AddNew

rs!paidto = "" & txtpaidto
rs!amount = Val(Txtamt)
rs!reason = "" & Txtreason
rs!paiddate = Date
rs.Update

txtpaidto = ""
Txtreason = ""
Txtamt = ""



End Sub

Private Sub Form_Load()
main.Enabled = False

cn.CursorLocation = adUseClient
cn.Open " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\agrawal.mdb;Persist Security Info=False"
If rs.State = 1 Then rs.Close
rs.Open "select * from expense", cn, adOpenDynamic, adLockOptimistic
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text3_Change()

End Sub

Private Sub Form_Unload(Cancel As Integer)
If rs.State = 1 Then rs.Close
If cn.State = 1 Then cn.Close
End Sub

Private Sub Txtdate_Change()
Txtdate = Date

End Sub
