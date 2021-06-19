VERSION 5.00
Begin VB.Form frmfare 
   Caption         =   "Fares....."
   ClientHeight    =   9465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12540
   LinkTopic       =   "Form5"
   ScaleHeight     =   9465
   ScaleWidth      =   12540
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FF8080&
      Caption         =   "&Home"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6840
      Width           =   6255
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H008080FF&
      Caption         =   "Jan Shatabdi Exp. K/m."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6120
      Width           =   3135
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H008080FF&
      Caption         =   "Jan Shatabdi Express Pairs of Station Wise Fares"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4680
      Width           =   6255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H008080FF&
      Caption         =   "Rajdhani Express Pairs of Station Wise Fares"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3960
      Width           =   6255
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "Shatabdi Express Pairs of Station Wise Fares"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      Width           =   6255
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "Rajdhani Express K/m."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Shatabdi Express K/m."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5400
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "Mail/Express K/m."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   3135
   End
   Begin VB.OLE OLE7 
      Class           =   "Package"
      Height          =   735
      Left            =   7680
      OleObjectBlob   =   "Fares.frx":0000
      SourceDoc       =   "D:\janshatabdi_km_files\jan_shatabdi_fares_km.htm"
      TabIndex        =   15
      Top             =   6120
      Width           =   3135
   End
   Begin VB.OLE OLE6 
      Class           =   "Package"
      Height          =   735
      Left            =   4560
      OleObjectBlob   =   "Fares.frx":5FA18
      SourceDoc       =   "D:\mail_1_to 250.htm"
      TabIndex        =   14
      Top             =   6120
      Width           =   3135
   End
   Begin VB.OLE OLE5 
      Class           =   "Package"
      Height          =   735
      Left            =   7680
      OleObjectBlob   =   "Fares.frx":10B030
      SourceDoc       =   "D:\rajdhani_km_files\rajdhani_fares_km.htm"
      TabIndex        =   13
      Top             =   5400
      Width           =   3135
   End
   Begin VB.OLE OLE4 
      Class           =   "Package"
      Height          =   735
      Left            =   4560
      OleObjectBlob   =   "Fares.frx":1D7848
      SourceDoc       =   "D:\shatabadi_fares_km.htm"
      TabIndex        =   12
      Top             =   5400
      Width           =   3135
   End
   Begin VB.OLE OLE3 
      Class           =   "Package"
      Height          =   615
      Left            =   4560
      OleObjectBlob   =   "Fares.frx":23C060
      SourceDoc       =   "D:\Janshatabdi_fares.htm"
      TabIndex        =   11
      Top             =   4800
      Width           =   6255
   End
   Begin VB.OLE OLE2 
      Class           =   "Package"
      Height          =   735
      Left            =   4560
      OleObjectBlob   =   "Fares.frx":24C278
      SourceDoc       =   "D:\rajdhani_1.htm"
      TabIndex        =   10
      Top             =   3960
      Width           =   6255
   End
   Begin VB.OLE OLE1 
      Class           =   "Package"
      Height          =   735
      Left            =   4560
      OleObjectBlob   =   "Fares.frx":30D690
      SourceDoc       =   "C:\My Documents\Indian Railways Online Passenger Reservation Site Providing Availability_files\shatabadi%20fares.htm"
      TabIndex        =   9
      Top             =   3240
      Width           =   6255
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   6
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FARES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   3240
      TabIndex        =   5
      Top             =   2160
      Width           =   8535
   End
   Begin VB.Image Image1 
      Height          =   11415
      Left            =   -120
      Picture         =   "Fares.frx":34CAA8
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   15420
   End
End
Attribute VB_Name = "frmfare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
OLE6.Visible = True
Command1.Enabled = False
End Sub
Private Sub Command2_Click()
OLE4.Visible = True
Command2.Enabled = False
End Sub
Private Sub Command3_Click()
OLE5.Visible = True
Command3.Enabled = False
End Sub
Private Sub Command4_Click()
OLE1.Visible = True
Command4.Enabled = False
End Sub
Private Sub Command5_Click()
OLE2.Visible = True
Command5.Enabled = False
End Sub
Private Sub Command6_Click()
OLE3.Visible = True
Command6.Enabled = False
End Sub
Private Sub Command7_Click()
OLE7.Visible = True
Command7.Enabled = False
End Sub
Private Sub Command8_Click()
cn.Close
cn1.Close
Me.Hide
MDIForm1.Show
End Sub

