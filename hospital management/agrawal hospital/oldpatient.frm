VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form oldpatient 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9030
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   9030
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmddel 
      Caption         =   "Delete"
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
      Left            =   3720
      TabIndex        =   32
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton Cmdcancle 
      Caption         =   "Cancle"
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
      Left            =   1320
      TabIndex        =   4
      Top             =   6720
      Width           =   1440
   End
   Begin VB.CommandButton CMDEXIT 
      Caption         =   "Exit"
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
      Left            =   6390
      TabIndex        =   5
      Top             =   6720
      Width           =   1440
   End
   Begin VB.TextBox txtdis 
      Enabled         =   0   'False
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
      Left            =   5685
      TabIndex        =   21
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Patient selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   0
      TabIndex        =   16
      Top             =   960
      Width           =   3255
      Begin VB.CommandButton cmdsearch 
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdclear 
         Caption         =   "CLEAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtsearch 
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
         Left            =   2040
         TabIndex        =   1
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox cmbsearch 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "oldpatient.frx":0000
         Left            =   2040
         List            =   "oldpatient.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid dgsearch 
         Height          =   3255
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   5741
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   16777215
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "search "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "select serch option"
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
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.TextBox txp 
      Enabled         =   0   'False
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
      Left            =   5685
      TabIndex        =   15
      Top             =   5400
      Width           =   1455
   End
   Begin VB.TextBox txag 
      Enabled         =   0   'False
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
      Left            =   5685
      TabIndex        =   14
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox txtsex 
      Enabled         =   0   'False
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
      Left            =   5685
      TabIndex        =   13
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox txtpcode 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   12
      Top             =   1080
      Width           =   150
   End
   Begin VB.TextBox txtname 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Left            =   5685
      TabIndex        =   11
      Top             =   1080
      Width           =   3015
   End
   Begin VB.TextBox txtatten 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Left            =   5685
      TabIndex        =   10
      Top             =   1680
      Width           =   3015
   End
   Begin VB.TextBox txtadd 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Left            =   5685
      TabIndex        =   9
      Top             =   2280
      Width           =   3015
   End
   Begin VB.TextBox txtrfa 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Left            =   5685
      TabIndex        =   8
      Top             =   2880
      Width           =   3015
   End
   Begin VB.TextBox txtedu 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Left            =   5685
      TabIndex        =   7
      Top             =   3390
      Width           =   1455
   End
   Begin VB.TextBox txtdoa 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Left            =   5685
      TabIndex        =   6
      Top             =   3915
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3825
      TabIndex        =   31
      Top             =   2295
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Patient Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3825
      TabIndex        =   30
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "So/Do/Wo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3825
      TabIndex        =   29
      Top             =   1740
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Age"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3825
      TabIndex        =   28
      Top             =   4545
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sex"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3825
      TabIndex        =   27
      Top             =   4965
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Diagnosis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3825
      TabIndex        =   26
      Top             =   2835
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Education"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3825
      TabIndex        =   25
      Top             =   3450
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Phone no."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3825
      TabIndex        =   24
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Date of Admission"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3825
      TabIndex        =   23
      Top             =   3990
      Width           =   1695
   End
   Begin VB.Label Lbldis 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Discharge date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3825
      TabIndex        =   22
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Old Patient Record"
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
      Height          =   735
      Left            =   848
      TabIndex        =   20
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "oldpatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim st As String
Private Sub cmbsearch_Click()
   Select Case cmbsearch.ListIndex
        Case 0
            st = "pcode"
        Case 1
            st = "pname"
        Case Else
            st = ""
        End Select

End Sub

Private Sub Cmdcancle_Click()
txtempty

End Sub

Private Sub cmdclear_Click()
txtsearch.Text = ""
 
 st = ""
 rs.Filter = 0
 rs.Requery
 Set dgsearch.DataSource = rs
 dgsearch.Enabled = True

End Sub

Private Sub cmddel_Click()
If txtpcode = "" Then
MsgBox "NO RECOED TO DELETE", vbCritical
Exit Sub
End If
If MsgBox("CONFIRM RECOED DELETE", vbYesNo) = vbYes Then
If rs.RecordCount > 0 Then

rs.Delete
rs.Update
txtempty
End If
End If
End Sub

Private Sub CMDEXIT_Click()
main.Enabled = True

If rs.State = 1 Then rs.Close
If cn.State = 1 Then cn.Close
Unload Me
Load main
main.Show

End Sub
Private Sub txtempty()
txtrfa = ""
txtedu = ""
txtdoa = ""
txtatten = ""
txtname = ""
txag = ""
txtdoa = ""
txp = ""
txtsex = ""
txtward = ""
txtadd = ""
txtpcode = ""
txtdis = ""

End Sub
Private Sub rs2txt()
If rs.EOF Then
txtempty
Exit Sub
End If
txtrfa = rs!pccomplaint
txtedu = rs!education
txtdoa = rs!doa
txtatten = rs!paname
txtname = rs!pname
txtadd = rs!address
txag = rs!age
txtdoa = rs!doa
txp = rs!ph
txtsex = rs!sex
txtpcode = rs!pcode
txtdis = rs!ddate
End Sub

Private Sub cmdsearch_Click()
If cmbsearch.Text = "" Then
  MsgBox "PLEASE SELECT  FIELD TO SEARCH", vbInformation, "USER INFORMATION"
 Exit Sub
 End If
If txtsearch.Text = "" Then
   MsgBox "PLEASE ENTER TEXT TO SEARCH", vbInformation, "USER INFORMATION"
   Exit Sub
End If
If txtsearch = "" And cmbsearch.Text = "pcode" Then
        MsgBox "Please Enter code to Search", vbInformation, "User Information"
         txtsearch.SetFocus
 ElseIf txtsearch = "" And cmbsearch.ListIndex = 1 Then
        MsgBox "Please Enter name to Search", vbInformation, "User Information"
        txtsearch.SetFocus
 Else
        Set dgsearch.DataSource = Nothing
        txtempty
        rs.Filter = st & " like '" & txtsearch & "*'"
        Set dgsearch.DataSource = rs
        If rs.RecordCount = 0 Then
            MsgBox "No Records Found", vbInformation, "User Information"
            cmdclear_Click
        Else
              rs2txt
        End If
 End If

End Sub

Private Sub Command1_Click()

End Sub

Private Sub dgsearch_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If rs.RecordCount > 0 And Not (rs.BOF Or rs.EOF) Then
  rs2txt
End If
Exit Sub
End Sub

Private Sub Form_Load()
main.Enabled = False

cn.CursorLocation = adUseClient
If cn.State = 1 Then cn.Close

cn.Open " Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\agrawal.mdb;Persist Security Info=False"
If rs.State = 1 Then rs.Close
rs.Open "select * from oldpatient", cn, adOpenDynamic, adLockOptimistic
Set dgsearch.DataSource = rs
txtpcode.Visible = False
rs2txt

End Sub

Private Sub Form_Unload(Cancel As Integer)
If rs.State = 1 Then rs.Close
If cn.State = 1 Then cn.Close
End Sub
