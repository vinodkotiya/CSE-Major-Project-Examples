VERSION 5.00
Begin VB.Form background 
   BackColor       =   &H00C0E0FF&
   Caption         =   "AGRAWAL HOPITAL'S INFORMATION SYSTEM"
   ClientHeight    =   6915
   ClientLeft      =   660
   ClientTop       =   1185
   ClientWidth     =   10695
   LinkTopic       =   "Form2"
   ScaleHeight     =   6915
   ScaleWidth      =   10695
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6840
      Top             =   120
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "SYSTEM"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4380
      TabIndex        =   2
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   " HOSPITAL INFORMATION"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2340
      TabIndex        =   1
      Top             =   3720
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "WELCOME  TO  AGRAWAL "
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2340
      TabIndex        =   0
      Top             =   2520
      Width           =   6015
   End
End
Attribute VB_Name = "background"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Unload Me
Load main
main.Show

End Sub
