VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form uudpatient 
   Caption         =   "Form1"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   11610
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtname 
      BackColor       =   &H00FFFFFF&
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
      Left            =   7680
      TabIndex        =   17
      Top             =   240
      Width           =   3015
   End
   Begin VB.TextBox txtatten 
      BackColor       =   &H00FFFFFF&
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
      Left            =   7440
      TabIndex        =   16
      Top             =   1080
      Width           =   3015
   End
   Begin VB.TextBox txtage 
      BackColor       =   &H00FFFFFF&
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
      Left            =   9240
      TabIndex        =   15
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtadd 
      BackColor       =   &H00FFFFFF&
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
      Left            =   7440
      TabIndex        =   14
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox txtrfa 
      BackColor       =   &H00FFFFFF&
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
      Left            =   7440
      TabIndex        =   13
      Top             =   2640
      Width           =   3015
   End
   Begin VB.ComboBox cmbward 
      BackColor       =   &H00FFFFFF&
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
      ItemData        =   "uudpatient.frx":0000
      Left            =   6720
      List            =   "uudpatient.frx":0010
      TabIndex        =   12
      Top             =   4920
      Width           =   1575
   End
   Begin VB.ComboBox cmbsex 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "uudpatient.frx":0038
      Left            =   9240
      List            =   "uudpatient.frx":0045
      TabIndex        =   11
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox txtedu 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4920
      TabIndex        =   10
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox txtph 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   9240
      TabIndex        =   9
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox txtdoa 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4920
      TabIndex        =   8
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin VB.ComboBox cmbsearch 
         Height          =   315
         ItemData        =   "uudpatient.frx":005F
         Left            =   1800
         List            =   "uudpatient.frx":0069
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtsearch 
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdsearch 
         Caption         =   "SEARCH"
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton cmdclear 
         Caption         =   "CLEAR"
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Top             =   3360
         Width           =   855
      End
      Begin MSDataGridLib.DataGrid dgs1 
         Height          =   1935
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   3413
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
      Begin VB.Label Label1 
         Caption         =   "select serch option"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "search text"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Label Label12 
      Caption         =   "PATIENT NAME"
      Height          =   255
      Left            =   3840
      TabIndex        =   27
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "PARENTS NAME"
      Height          =   255
      Left            =   3840
      TabIndex        =   26
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "AGE"
      Height          =   375
      Left            =   12240
      TabIndex        =   25
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "SEX"
      Height          =   375
      Left            =   12120
      TabIndex        =   24
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "ADDRESS"
      Height          =   255
      Left            =   3840
      TabIndex        =   23
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "REASON FOR ADMIT"
      Height          =   255
      Left            =   3840
      TabIndex        =   22
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "WARD"
      Height          =   255
      Left            =   3840
      TabIndex        =   21
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "EDUCATION"
      Height          =   255
      Left            =   3840
      TabIndex        =   20
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "PHONE NO."
      Height          =   375
      Left            =   7920
      TabIndex        =   19
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "DATE OF ADMISSION"
      Height          =   255
      Left            =   3840
      TabIndex        =   18
      Top             =   4560
      Width           =   1695
   End
End
Attribute VB_Name = "uudpatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
