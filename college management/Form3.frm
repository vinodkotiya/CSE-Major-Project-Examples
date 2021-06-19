VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Course Information U.G. / Course Information P.G."
   ClientHeight    =   3195
   ClientLeft      =   4305
   ClientTop       =   2595
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Text            =   "(None)"
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Course - Id"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Mainform.Show
End Sub
