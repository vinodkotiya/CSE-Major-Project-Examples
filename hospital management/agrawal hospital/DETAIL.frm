VERSION 5.00
Begin VB.Form DETAIL 
   Caption         =   "DETAIL OF TREATEMENT"
   ClientHeight    =   4110
   ClientLeft      =   2085
   ClientTop       =   2745
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   8295
   Begin VB.CommandButton Cmdprint 
      Caption         =   "PRINT"
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Cmdback 
      Caption         =   "BACK"
      Height          =   495
      Left            =   6960
      TabIndex        =   0
      Top             =   3600
      Width           =   1095
   End
End
Attribute VB_Name = "DETAIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Cmdback_Click()
Unload Me
Load billing
billing.Show

End Sub

Private Sub Form_Load()

End Sub
