VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmoutput 
   AutoRedraw      =   -1  'True
   Caption         =   "Bhopal To Delhi"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "Detail..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7920
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
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
      Left            =   8280
      MaxLength       =   5
      TabIndex        =   10
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7920
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7920
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmout.frx":0000
      Height          =   4575
      Left            =   4800
      TabIndex        =   13
      Top             =   840
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   8070
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "Train_no"
         Caption         =   "Train_no"
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
         DataField       =   "Train_name"
         Caption         =   "Train_name"
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
      BeginProperty Column02 
         DataField       =   "Departure (Bhopal)"
         Caption         =   "Departure (Bhopal)"
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
      BeginProperty Column03 
         DataField       =   "Arrival (Delhi)"
         Caption         =   "Arrival (Delhi)"
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
      BeginProperty Column04 
         DataField       =   "Day"
         Caption         =   "Day"
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
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   11520
      Top             =   3480
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\Microsoft Visual Studio\VB98\btod.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\Microsoft Visual Studio\VB98\btod.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "btod"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Trains : Bhopal To Delhi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   6000
      TabIndex        =   12
      Top             =   360
      Width           =   6615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Train No. :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   9
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "131 - General Enquiry (Manual) "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4800
      TabIndex        =   5
      Top             =   8880
      Width           =   6735
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "1332 - Trains Coming From Itarsi Side and  Going towards Bina and Ujjain "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   9360
      Width           =   6495
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "1331 - Trains Coming from Bina and Ujjain Side and Going towards Itarsi Side"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4680
      TabIndex        =   3
      Top             =   9600
      Width           =   6855
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "1335 - Reservation Enquiry (IVRS) ( 05.30 to 22.30 Hrs.)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4680
      TabIndex        =   2
      Top             =   9840
      Width           =   6855
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Note : Mon.-1, Tues.-2, Wed.-3, Thur.-4, Fri.-5, Sat.-6, Sun.-7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4680
      TabIndex        =   1
      Top             =   10200
      Width           =   6975
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "1334 - Railway Enquiry Ph. No. : Interactive Voice Response System (IVRS)       "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4680
      TabIndex        =   0
      Top             =   9120
      Width           =   7215
   End
   Begin VB.Image Image1 
      Height          =   11295
      Left            =   -120
      Picture         =   "frmout.frx":0015
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   15375
   End
End
Attribute VB_Name = "frmoutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
frminput.Show
End Sub
Private Sub Command2_Click()
cn.Close
cn1.Close
Unload Me
MDIForm1.Show
End Sub

Private Sub Command3_Click()
frmoutput.PrintForm
End Sub

Private Sub Command4_Click()
If Text1.Text = "" Then
MsgBox "Please Enter Valid No.", vbCritical + vbDefaultButton3, OK
frmoutput.Show
frmoutput.Text1.SetFocus
Else
If frmoutput.Text1.Text = 9367 Then
frmdetail.Show
frmdetail.Label2.Caption = 9367
frmdetail.Label4.Caption = "Indore-Jammu Tawi Malwa Exp."
frmdetail.Label6.Caption = "Daily"
frmdetail.grdtable.Row = 0
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "AC1"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "AC2"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "AC3"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "Sleeper"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "F Class"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "CC"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "Sec Class"

frmdetail.grdtable.Row = 1
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "YES"

frmdetail.grdtable.Row = 2
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "1074"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "234"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "765"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "140"

frmdetail.Label8.Caption = "Bhopal"
frmdetail.Label10.Caption = "18:45"
frmdetail.Label12.Caption = "19:05"
frmdetail.Label14.Caption = "New Delhi"
frmdetail.Label16.Caption = "07:50"
frmdetail.Label18.Caption = "08:10"
frmdetail.Label20.Caption = "701 Kms."
Else

If frmoutput.Text1.Text = 2137 Then
frmdetail.Show
frmdetail.Label2.Caption = 2137
frmdetail.Label4.Caption = "Mumbai CST-Firozpur Punjabmail"
frmdetail.Label6.Caption = "Daily"
frmdetail.grdtable.Row = 0
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "AC1"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "AC2"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "AC3"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "Sleeper"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "F Class"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "CC"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "Sec Class"

frmdetail.grdtable.Row = 1
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "YES"

frmdetail.grdtable.Row = 2
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "2145"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "1103"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "699"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "241"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "144"

frmdetail.Label8.Caption = "Bhopal"
frmdetail.Label10.Caption = "09:05"
frmdetail.Label12.Caption = "09:15"
frmdetail.Label14.Caption = "New Delhi"
frmdetail.Label16.Caption = "20:20"
frmdetail.Label18.Caption = "20:55"
frmdetail.Label20.Caption = "701 Kms."
Else

If frmoutput.Text1.Text = 2615 Then
frmdetail.Show
frmdetail.Label2.Caption = 2615
frmdetail.Label4.Caption = "Chennai New Delhi G.T. Exp."
frmdetail.Label6.Caption = "Daily"
frmdetail.grdtable.Row = 0
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "AC1"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "AC2"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "AC3"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "Sleeper"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "F Class"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "CC"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "Sec Class"

frmdetail.grdtable.Row = 1
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "YES"

frmdetail.grdtable.Row = 2
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "2145"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "1103"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "699"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "241"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "144"

frmdetail.Label8.Caption = "Bhopal"
frmdetail.Label10.Caption = "17:30"
frmdetail.Label12.Caption = "17:40"
frmdetail.Label14.Caption = "New Delhi"
frmdetail.Label16.Caption = "05:00"
frmdetail.Label18.Caption = "STN of Destination"
frmdetail.Label20.Caption = "701 Kms."
Else

If frmoutput.Text1.Text = 2621 Then
frmdetail.Show
frmdetail.Label2.Caption = 2621
frmdetail.Label4.Caption = "Chennai New Delhi Tamil Nadu Exp."
frmdetail.Label6.Caption = "Daily"
frmdetail.grdtable.Row = 0
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "AC1"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "AC2"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "AC3"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "Sleeper"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "F Class"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "CC"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "Sec Class"

frmdetail.grdtable.Row = 1
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "YES"

frmdetail.grdtable.Row = 2
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "2145"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "1103"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "699"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "241"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "144"

frmdetail.Label8.Caption = "Bhopal"
frmdetail.Label10.Caption = "20:55"
frmdetail.Label12.Caption = "21:05"
frmdetail.Label14.Caption = "New Delhi"
frmdetail.Label16.Caption = "07:30"
frmdetail.Label18.Caption = "STN of Destination"
frmdetail.Label20.Caption = "701 Kms."
Else

If frmoutput.Text1.Text = 2625 Then
frmdetail.Show
frmdetail.Label2.Caption = 2625
frmdetail.Label4.Caption = "Tiruvananthapuram New Delhi Kerla Exp."
frmdetail.Label6.Caption = "Daily"
frmdetail.grdtable.Row = 0
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "AC1"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "AC2"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "AC3"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "Sleeper"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "F Class"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "CC"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "Sec Class"

frmdetail.grdtable.Row = 1
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "YES"

frmdetail.grdtable.Row = 2
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "1103"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "699"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "241"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "144"

frmdetail.Label8.Caption = "Bhopal"
frmdetail.Label10.Caption = "04:45"
frmdetail.Label12.Caption = "04:55"
frmdetail.Label14.Caption = "New Delhi"
frmdetail.Label16.Caption = "15:45"
frmdetail.Label18.Caption = "STN of Destination"
frmdetail.Label20.Caption = "701 Kms."
Else

If frmoutput.Text1.Text = 2627 Then
frmdetail.Show
frmdetail.Label2.Caption = 2627
frmdetail.Label4.Caption = "Banglore New Delhi Karnatka Exp."
frmdetail.Label6.Caption = "Daily"
frmdetail.grdtable.Row = 0
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "AC1"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "AC2"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "AC3"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "Sleeper"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "F Class"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "CC"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "Sec Class"

frmdetail.grdtable.Row = 1
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "YES"

frmdetail.grdtable.Row = 2
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "1103"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "699"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "241"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "144"

frmdetail.Label8.Caption = "Bhopal"
frmdetail.Label10.Caption = "00:55"
frmdetail.Label12.Caption = "01:05"
frmdetail.Label14.Caption = "New Delhi"
frmdetail.Label16.Caption = "12:05"
frmdetail.Label18.Caption = "STN of Destination"
frmdetail.Label20.Caption = "701 Kms."
Else

If frmoutput.Text1.Text = 2715 Then
frmdetail.Show
frmdetail.Label2.Caption = 2715
frmdetail.Label4.Caption = "Nanded Amritsar Sachkhand Exp."
frmdetail.Label6.Caption = "1,2,3,6,7"
frmdetail.grdtable.Row = 0
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "AC1"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "AC2"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "AC3"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "Sleeper"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "F Class"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "CC"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "Sec Class"

frmdetail.grdtable.Row = 1
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "YES"

frmdetail.grdtable.Row = 2
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "1103"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "699"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "241"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "144"

frmdetail.Label8.Caption = "Bhopal"
frmdetail.Label10.Caption = "00:15"
frmdetail.Label12.Caption = "00:25"
frmdetail.Label14.Caption = "New Delhi"
frmdetail.Label16.Caption = "13:20"
frmdetail.Label18.Caption = "14:00"
frmdetail.Label20.Caption = "701 Kms."
Else

If frmoutput.Text1.Text = 2723 Then
frmdetail.Show
frmdetail.Label2.Caption = 2723
frmdetail.Label4.Caption = "Hyderabad New Delhi A.P. Exp."
frmdetail.Label6.Caption = "Daily"
frmdetail.grdtable.Row = 0
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "AC1"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "AC2"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "AC3"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "Sleeper"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "F Class"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "CC"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "Sec Class"

frmdetail.grdtable.Row = 1
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "YES"

frmdetail.grdtable.Row = 2
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "2145"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "1103"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "699"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "241"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "144"

frmdetail.Label8.Caption = "Bhopal"
frmdetail.Label10.Caption = "22:20"
frmdetail.Label12.Caption = "22:30"
frmdetail.Label14.Caption = "New Delhi"
frmdetail.Label16.Caption = "08:40"
frmdetail.Label18.Caption = "STN of Destination"
frmdetail.Label20.Caption = "701 Kms."
Else

If frmoutput.Text1.Text = 6317 Then
frmdetail.Show
frmdetail.Label2.Caption = 6317
frmdetail.Label4.Caption = "Kanniyakumari Jammu Tawi Himsagar Exp."
frmdetail.Label6.Caption = "7"
frmdetail.grdtable.Row = 0
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "AC1"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "AC2"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "AC3"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "Sleeper"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "F Class"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "CC"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "Sec Class"

frmdetail.grdtable.Row = 1
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "YES"

frmdetail.grdtable.Row = 2
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "1103"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "241"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "786"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "144"

frmdetail.Label8.Caption = "Bhopal"
frmdetail.Label10.Caption = "11:15"
frmdetail.Label12.Caption = "11:25"
frmdetail.Label14.Caption = "New Delhi"
frmdetail.Label16.Caption = "23:45"
frmdetail.Label18.Caption = "00:15"
frmdetail.Label20.Caption = "701 Kms."
Else

If frmoutput.Text1.Text = 8237 Then
frmdetail.Show
frmdetail.Label2.Caption = 8237
frmdetail.Label4.Caption = "Bilaspur Amritsar Chattisgarh Exp."
frmdetail.Label6.Caption = "Daily"
frmdetail.grdtable.Row = 0
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "AC1"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "AC2"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "AC3"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "Sleeper"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "F Class"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "CC"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "Sec Class"

frmdetail.grdtable.Row = 1
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "YES"

frmdetail.grdtable.Row = 2
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "1103"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "241"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "144"

frmdetail.Label8.Caption = "Bhopal"
frmdetail.Label10.Caption = "06:30"
frmdetail.Label12.Caption = "06:40"
frmdetail.Label14.Caption = "New Delhi"
frmdetail.Label16.Caption = "20:30"
frmdetail.Label18.Caption = "21:05"
frmdetail.Label20.Caption = "701 Kms."
Else

If frmoutput.Text1.Text = 1057 Then
frmdetail.Show
frmdetail.Label2.Caption = 1057
frmdetail.Label4.Caption = "Dadar Amritsar Exp."
frmdetail.Label6.Caption = "Daily"
frmdetail.grdtable.Row = 0
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "AC1"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "AC2"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "AC3"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "Sleeper"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "F Class"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "CC"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "Sec Class"

frmdetail.grdtable.Row = 1
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "YES"

frmdetail.grdtable.Row = 2
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "1103"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "699"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "241"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "144"

frmdetail.Label8.Caption = "Bhopal"
frmdetail.Label10.Caption = "15:00"
frmdetail.Label12.Caption = "15:10"
frmdetail.Label14.Caption = "New Delhi"
frmdetail.Label16.Caption = "04:30"
frmdetail.Label18.Caption = "05:10"
frmdetail.Label20.Caption = "701 Kms."
Else

If frmoutput.Text1.Text = 2441 Then
frmdetail.Show
frmdetail.Label2.Caption = 2441
frmdetail.Label4.Caption = "Bilaspur New Delhi Rajdhani Exp."
frmdetail.Label6.Caption = "7"
frmdetail.grdtable.Row = 0
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "AC1"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "AC2"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "AC3"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "Sleeper"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "F Class"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "CC"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "Sec Class"

frmdetail.grdtable.Row = 1
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "NO"

frmdetail.grdtable.Row = 2
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "0"

frmdetail.Label8.Caption = "Bhopal"
frmdetail.Label10.Caption = "20:45"
frmdetail.Label12.Caption = "20:50"
frmdetail.Label14.Caption = "New Delhi"
frmdetail.Label16.Caption = "05:15"
frmdetail.Label18.Caption = "STN of Destination"
frmdetail.Label20.Caption = "701 Kms."
Else

If frmoutput.Text1.Text = 6687 Then
frmdetail.Show
frmdetail.Label2.Caption = 6687
frmdetail.Label4.Caption = "Mangalore Jammu Tawi Navyug Exp."
frmdetail.Label6.Caption = "3"
frmdetail.grdtable.Row = 0
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "AC1"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "AC2"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "AC3"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "Sleeper"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "F Class"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "CC"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "Sec Class"

frmdetail.grdtable.Row = 1
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "YES"

frmdetail.grdtable.Row = 2
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "1103"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "241"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "144"

frmdetail.Label8.Caption = "Bhopal"
frmdetail.Label10.Caption = "11:15"
frmdetail.Label12.Caption = "11:25"
frmdetail.Label14.Caption = "New Delhi"
frmdetail.Label16.Caption = "23:45"
frmdetail.Label18.Caption = "00:15"
frmdetail.Label20.Caption = "701 Kms."
Else

If frmoutput.Text1.Text = 2001 Then
frmdetail.Show
frmdetail.Label2.Caption = 2001
frmdetail.Label4.Caption = "Bhopal New Delhi Shatabdi Exp."
frmdetail.Label6.Caption = "Daily"
frmdetail.grdtable.Row = 0
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "AC1"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "AC2"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "ACCC"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "AC3"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "CC"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "EACCC"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = ""

frmdetail.grdtable.Row = 1
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = ""

frmdetail.grdtable.Row = 2
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "850"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "1800"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = ""

frmdetail.Label8.Caption = "Bhopal"
frmdetail.Label10.Caption = "14:30"
frmdetail.Label12.Caption = "14:50"
frmdetail.Label14.Caption = "New Delhi"
frmdetail.Label16.Caption = "22:50"
frmdetail.Label18.Caption = "STN of Destination"
frmdetail.Label20.Caption = "701 Kms."
Else

If frmoutput.Text1.Text = 1077 Then
frmdetail.Show
frmdetail.Label2.Caption = 1077
frmdetail.Label4.Caption = "Pune Jammu Tawi Jhelum Exp."
frmdetail.Label6.Caption = "Daily"
frmdetail.grdtable.Row = 0
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "AC1"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "AC2"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "AC3"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "Sleeper"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "F Class"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "CC"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "Sec Class"

frmdetail.grdtable.Row = 1
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "YES"

frmdetail.grdtable.Row = 2
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "1103"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "699"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "241"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "144"

frmdetail.Label8.Caption = "Bhopal"
frmdetail.Label10.Caption = "09:20"
frmdetail.Label12.Caption = "09:30"
frmdetail.Label14.Caption = "New Delhi"
frmdetail.Label16.Caption = "21:20"
frmdetail.Label18.Caption = "21:50"
frmdetail.Label20.Caption = "701 Kms."
Else

If frmoutput.Text1.Text = 6031 Then
frmdetail.Show
frmdetail.Label2.Caption = 6031
frmdetail.Label4.Caption = "Chennai Jammu Tawi Andoman Exp."
frmdetail.Label6.Caption = "1,4,5"
frmdetail.grdtable.Row = 0
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "AC1"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "AC2"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "AC3"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "Sleeper"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "F Class"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "CC"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "Sec Class"

frmdetail.grdtable.Row = 1
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "YES"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "NO"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "YES"

frmdetail.grdtable.Row = 2
frmdetail.grdtable.Col = 0
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 1
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 2
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 3
frmdetail.grdtable.Text = "241"
frmdetail.grdtable.Col = 4
frmdetail.grdtable.Text = "786"
frmdetail.grdtable.Col = 5
frmdetail.grdtable.Text = "0"
frmdetail.grdtable.Col = 6
frmdetail.grdtable.Text = "144"

frmdetail.Label8.Caption = "Bhopal"
frmdetail.Label10.Caption = "10:15"
frmdetail.Label12.Caption = "10:25"
frmdetail.Label14.Caption = "New Delhi"
frmdetail.Label16.Caption = "22:45"
frmdetail.Label18.Caption = "00:15"
frmdetail.Label20.Caption = "701 Kms."
End If
MsgBox "Try Again", vbAbortRetryIgnore + vbApplicationModal, OK
frmoutput.Show
Text1.Text = " "
Text1.SetFocus
Text1.Text = " "
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End Sub
