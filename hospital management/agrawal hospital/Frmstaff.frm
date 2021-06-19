VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Frmstaff 
   Caption         =   "Form1"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmdcancle 
      Caption         =   "Cancle"
      Height          =   495
      Index           =   5
      Left            =   4560
      TabIndex        =   34
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton Cmdsave 
      Caption         =   "Save"
      Height          =   495
      Index           =   4
      Left            =   5880
      TabIndex        =   33
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton Cmdexit 
      Caption         =   "Exit"
      Height          =   495
      Index           =   3
      Left            =   7200
      TabIndex        =   32
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton Cmdupdate 
      Caption         =   "Update"
      Height          =   495
      Index           =   2
      Left            =   3240
      TabIndex        =   31
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton Cmddelete 
      Caption         =   "Delete"
      Height          =   495
      Index           =   1
      Left            =   1920
      TabIndex        =   30
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton Cmdnew 
      Caption         =   "New"
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   29
      Top             =   5520
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   4335
      Left            =   3960
      TabIndex        =   8
      Top             =   840
      Width           =   4215
      Begin VB.ComboBox Cmbsex 
         Height          =   315
         ItemData        =   "Frmstaff.frx":0000
         Left            =   2520
         List            =   "Frmstaff.frx":000A
         TabIndex        =   28
         Top             =   1560
         Width           =   735
      End
      Begin VB.ComboBox Cmbdegic 
         Height          =   315
         Left            =   1320
         TabIndex        =   26
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox Txtphone 
         Height          =   285
         Index           =   6
         Left            =   1320
         TabIndex        =   25
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox Txtaddress 
         Height          =   285
         Index           =   5
         Left            =   1320
         TabIndex        =   24
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox Txtedu 
         Height          =   285
         Index           =   4
         Left            =   1320
         TabIndex        =   23
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox Txtahe 
         Height          =   285
         Index           =   3
         Left            =   1320
         TabIndex        =   22
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Txtdob 
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   21
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox Txtfname 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   20
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox Txtname 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   19
         Top             =   480
         Width           =   1935
      End
      Begin VB.Frame Frame3 
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   3600
         Width           =   3495
         Begin VB.Label Label11 
            Caption         =   "Employee code"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.Label Label12 
         Caption         =   "Sex"
         Height          =   255
         Left            =   2040
         TabIndex        =   27
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "Phone no"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Address"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Degicnation"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Education"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Age"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Date of birth"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Father's name"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Name"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   3015
      Begin VB.CommandButton Cmdclear 
         Caption         =   "clear"
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton Cmdserch 
         Caption         =   "serch"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox Txtserch 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox Cmbserch 
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2415
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   4260
         _Version        =   393216
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
      Begin VB.Label Label2 
         Caption         =   "select text"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "select serch option"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "Staff information"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   35
      Top             =   0
      Width           =   6495
   End
End
Attribute VB_Name = "Frmstaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
