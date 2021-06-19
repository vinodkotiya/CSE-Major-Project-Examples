VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H008080FF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9990
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraMainFrame 
      BackColor       =   &H00C0C0FF&
      Height          =   5310
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9780
      Begin VB.Timer Timer2 
         Interval        =   50
         Left            =   240
         Top             =   840
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   1080
         TabIndex        =   8
         Top             =   3960
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.PictureBox picLogo 
         BackColor       =   &H00C0C0FF&
         Height          =   2025
         Left            =   240
         Picture         =   "frmSplash.frx":0000
         ScaleHeight     =   1965
         ScaleWidth      =   2115
         TabIndex        =   2
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         Caption         =   "LicenseTo : Unique College Bhopal"
         Height          =   255
         Left            =   2640
         TabIndex        =   1
         Tag             =   "LicenseTo"
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "Railway Enquiry"
         BeginProperty Font 
            Name            =   "Myriad Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3360
         TabIndex        =   7
         Tag             =   "Product"
         Top             =   1320
         Width           =   4095
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "CompanyProduct"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2505
         TabIndex        =   6
         Tag             =   "CompanyProduct"
         Top             =   765
         Width           =   3000
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "Platform Windows Version"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3450
         TabIndex        =   5
         Tag             =   "Platform"
         Top             =   2160
         Width           =   3675
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "Version 1.0.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5520
         TabIndex        =   4
         Tag             =   "Version"
         Top             =   2520
         Width           =   1605
      End
      Begin VB.Label lblCompany 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Company : Sun Software Ltd."
         Height          =   255
         Left            =   5160
         TabIndex        =   3
         Tag             =   "Company"
         Top             =   2880
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
End Sub

Private Sub Timer2_Timer()
ProgressBar1.Value = ProgressBar1.Value + 2
If ProgressBar1.Value = ProgressBar1.Max Then
Unload Me
frmAbout.Show
End If
End Sub
