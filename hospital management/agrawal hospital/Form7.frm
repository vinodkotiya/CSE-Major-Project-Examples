VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form treatment_entry 
   BackColor       =   &H00FFC0C0&
   Caption         =   "treatment entry form"
   ClientHeight    =   7530
   ClientLeft      =   420
   ClientTop       =   945
   ClientWidth     =   11040
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   ScaleHeight     =   7530
   ScaleWidth      =   11040
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdnew 
      Caption         =   "NEW"
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
      Left            =   1726
      TabIndex        =   11
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmclear 
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
      Height          =   315
      Left            =   2400
      TabIndex        =   6
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmsearch 
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
      Height          =   315
      Left            =   960
      TabIndex        =   5
      Top             =   4440
      Width           =   975
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
      Height          =   495
      Left            =   7606
      TabIndex        =   12
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
      Height          =   2655
      Left            =   3713
      TabIndex        =   32
      Top             =   3720
      Width           =   6855
      Begin VB.TextBox txtrate 
         Height          =   375
         Left            =   2880
         TabIndex        =   40
         Top             =   1440
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtref 
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
         Left            =   2520
         TabIndex        =   8
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtcat 
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
         Left            =   2520
         TabIndex        =   14
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   " CANCLE"
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
         Left            =   5400
         TabIndex        =   10
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "ADD"
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
         Left            =   5400
         TabIndex        =   9
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtqty 
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
         Left            =   2520
         TabIndex        =   7
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtitem 
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
         Left            =   2520
         TabIndex        =   15
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Reffered by :"
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
         Left            =   720
         TabIndex        =   37
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Categery"
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
         Left            =   720
         TabIndex        =   35
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Quantity"
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
         Left            =   720
         TabIndex        =   34
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Name of item"
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
         Left            =   720
         TabIndex        =   33
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Height          =   2775
      Left            =   3713
      TabIndex        =   21
      Top             =   840
      Width           =   6855
      Begin VB.TextBox txtpcode 
         Height          =   375
         Left            =   2640
         TabIndex        =   38
         Text            =   "Text11"
         Top             =   2160
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtward 
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
         Left            =   2160
         TabIndex        =   31
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtdoa 
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
         Left            =   2160
         TabIndex        =   29
         Top             =   1395
         Width           =   1335
      End
      Begin VB.TextBox txtsex 
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
         Left            =   5040
         TabIndex        =   27
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtage 
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
         Left            =   5040
         TabIndex        =   25
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtname 
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
         Left            =   2160
         TabIndex        =   23
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ward"
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
         Left            =   240
         TabIndex        =   30
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label6 
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
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   1440
         Width           =   1815
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
         Height          =   375
         Left            =   3840
         TabIndex        =   26
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label4 
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
         Height          =   375
         Left            =   3840
         TabIndex        =   24
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Name"
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
         Left            =   360
         TabIndex        =   22
         Top             =   720
         Width           =   615
      End
   End
   Begin MSDataGridLib.DataGrid dg2s 
      Height          =   1335
      Left            =   480
      TabIndex        =   20
      Top             =   4920
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   2355
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
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
   Begin VB.ComboBox cmbcat 
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
      ItemData        =   "Form7.frx":0000
      Left            =   2160
      List            =   "Form7.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Service selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   473
      TabIndex        =   18
      Top             =   3720
      Width           =   3255
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Select Category"
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
         Top             =   240
         Width           =   1455
      End
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
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
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
      ItemData        =   "Form7.frx":0004
      Left            =   2400
      List            =   "Form7.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid dgs1 
      Height          =   1215
      Left            =   480
      TabIndex        =   17
      Top             =   2280
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   2143
      _Version        =   393216
      AllowUpdate     =   0   'False
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Patient Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   473
      TabIndex        =   13
      Top             =   840
      Width           =   3255
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
         Height          =   315
         Left            =   2040
         TabIndex        =   3
         Top             =   1080
         Width           =   855
      End
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
         Height          =   315
         Left            =   360
         TabIndex        =   2
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Search "
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
         TabIndex        =   36
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Select serch Option"
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
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Treatment Entry form"
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
      Left            =   2873
      TabIndex        =   39
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "treatment_entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset
'Dim rs5 As New ADODB.Recordset
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
Private Sub cmbcat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 cmsearch_Click
End If
End Sub
Private Sub cmdadd_Click()
If txtitem = "" Then
MsgBox "Enter Item name", vbOKOnly
txtitem.SetFocus
Exit Sub
End If
If txtqty.Text = "" Then
MsgBox "Enter Quantity name", vbOKOnly
txtqty.SetFocus
Exit Sub
End If
If txtpcode = "" Then
MsgBox "select patient"
Exit Sub
End If
rs22txt
emp1
End Sub
Private Sub CMDEXIT_Click()
main.Enabled = True

If rs.State = 1 Then rs.Close
If rs1.State = 1 Then rs1.Close
If rs2.State = 1 Then rs2.Close
If rs3.State = 1 Then rs3.Close
If rs4.State = 1 Then rs4.Close
If cn.State = 1 Then cn.Close
Unload Me
Load main
main.Show
End Sub
Private Sub cmdsearch_Click()
If cmbsearch.Text = "" Then
MsgBox "PLEASE SELECT  FIELD TO SEARCH", vbInformation, "USER INFORMATION"
Exit Sub
End If
If txtsearch.Text = "" Then
MsgBox "PLEASE  ENTER TEXT TO SEARCH", vbInformation, "USER INFORMATION"
Exit Sub
End If
If txtsearch = "" And cmbsearch.Text = "pcode" Then
MsgBox "Please Enter code to Search", vbInformation, "User Information"
        Text1.SetFocus
  ElseIf txtsearch = "" And cmbsearch.ListIndex = 1 Then
         MsgBox "Please Enter name to Search", vbInformation, "User Information"
        txtsearch.SetFocus
  Else
       Set dgs1.DataSource = Nothing
        txtempty
        rs.Filter = st & " like '" & txtsearch & "*'"
        Set dgs1.DataSource = rs
          
        If rs.RecordCount = 0 Then
            MsgBox "No Records Found", vbInformation, "User Information"
            cmdclear_Click
        Else
              rs2text
        End If
    End If
    

End Sub
Private Sub cmdclear_Click()
txtsearch.Text = ""
st = ""
    rs.Filter = 0
    rs.Requery
    Set dgs1.DataSource = rs
  
End Sub
Private Sub cmdnew_Click()
txtempty
emp1
End Sub
Private Sub cmsearch_Click()
If cmbcat.Text = "" Then
        MsgBox "Please select category  to Search", vbInformation, "User Information"
        cmbcat.SetFocus
        
   Else
        Set dg2s.DataSource = Nothing
        txtcat = ""
        txtitem = ""
        txtno = ""
        rs1.Filter = "category" & " like '" & cmbcat.Text & "*'"
        Set dg2s.DataSource = rs1
        If rs1.RecordCount = 0 Then
            MsgBox "No Records Found", vbInformation, "User Information"
            cmclear_Click
        Else
              rs12text
        End If
    End If
End Sub
Private Sub cmclear_Click()

rs1.Filter = 0
rs1.Requery
Set dg2s.DataSource = rs1
End Sub
Private Sub dgs1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If rs.RecordCount > 0 And Not (rs.BOF Or rs.EOF) Then
 rs2text
End If
End Sub
Private Sub dg2s_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If rs1.RecordCount > 0 And Not (rs1.BOF Or rs1.EOF) Then
 rs12text
End If
End Sub
Private Sub Form_Load()
main.Enabled = False

cn.CursorLocation = adUseClient
If cn.State = 1 Then cn.Close
cn.Open " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\agrawal.mdb;Persist Security Info=False"
If rs.State = 1 Then rs.Close
rs.Open "select * from patient", cn, adOpenDynamic, adLockOptimistic
Set dgs1.DataSource = rs
If rs3.State = 1 Then rs3.Close
rs3.Open "select pcode,category from bed_tra", cn, adOpenDynamic, adLockOptimistic
rs2text
If rs1.State = 1 Then rs1.Close
rs1.Open "select * from service_master", cn, adOpenDynamic, adLockOptimistic
Set dg2s.DataSource = rs1
rs12text
If rs2.State = 1 Then
rs2.Close
End If
If rs2.State = 1 Then rs2.Close
rs2.Open "select * from treatment", cn, adOpenDynamic, adLockOptimistic
If rs4.State = 1 Then rs4.Close
rs4.Open "select cat from catgri", cn, adOpenDynamic, adLockOptimistic
rs2cmb
End Sub
Private Sub rs2text()
If rs.EOF Then
txtempty
Exit Sub
End If
txtname = rs!pname
txtage = rs!age
txtdoa = rs!doa
txtsex = rs!sex
txtpcode = rs!pcode
rs3.Filter = "pcode" & " like '" & txtpcode & "*'"
rs3.MoveLast
txtward = rs3!category
End Sub
Private Sub txtempty()
txtname = ""
txtsex = ""
txtdoa = ""
txtage = ""
txtward = ""
txtpcode = ""
End Sub
Private Sub rs12text()
If rs1.RecordCount = 0 Then
txtitem = ""
txtqty = ""
txtno = ""
Else
txtno = rs1!itemno
txtcat = rs1!category
txtitem = rs1!itemname
txtrate = rs1!Rate
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If rs.State = 1 Then rs.Close
If rs1.State = 1 Then rs1.Close
If rs2.State = 1 Then rs2.Close
If rs3.State = 1 Then rs3.Close
If rs4.State = 1 Then rs4.Close
If cn.State = 1 Then cn.Close
End Sub

Private Sub txtqty_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
 MsgBox "PLEASE ENTER QUANTITY IN NUMBERS", vbOKOnly
 KeyAscii = 0
 txtqty = ""
 txtqty.SetFocus
End If
End Sub
Private Sub txtsearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 cmdsearch_Click
End If
End Sub
Private Sub rs2cmb()
If rs4.RecordCount > 0 And Not (rs4.BOF Or rs4.EOF) Then
rs4.MoveFirst
While Not (rs4.EOF)
cmbcat.AddItem rs4!cat
rs4.MoveNext
Wend
End If
Exit Sub
End Sub
Private Sub rs22txt()
rs2.AddNew
rs2!pcode = txtpcode & ""
rs2!itemname = txtitem & ""
rs2!qty = Val(txtqty)
rs2!Date = Date
rs2!refby = txtref & ""
rs2!qr = (Val(txtqty) * Val(txtrate))
rs2.Update
End Sub
Private Sub emp1()
txtitem = ""
txtqty = ""
txtref = ""
txtcat = ""
End Sub


