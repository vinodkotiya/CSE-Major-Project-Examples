VERSION 5.00
Begin VB.Form main 
   BackColor       =   &H00C0E0FF&
   Caption         =   "AGRAWAL HOSPITAL'S INFORMATION SYSTEM"
   ClientHeight    =   4740
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5280
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   4740
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuPatient 
      Caption         =   "&Patient"
      Begin VB.Menu mnuPatientAdmitPatient 
         Caption         =   "&Admit Patient"
         Shortcut        =   ^A
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu mnumodificationupdate_patient 
         Caption         =   "&Update Patient"
         Shortcut        =   ^U
      End
      Begin VB.Menu b 
         Caption         =   "-"
      End
      Begin VB.Menu mnuModificationShiftPatient 
         Caption         =   "&Shift Patient"
         Shortcut        =   ^S
      End
      Begin VB.Menu c 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPatientRecievePayment 
         Caption         =   "&Recieve Payment"
         Shortcut        =   ^R
      End
      Begin VB.Menu d 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdmitTreatment 
         Caption         =   "&Treatment"
         Shortcut        =   ^T
      End
      Begin VB.Menu e 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPatientPatientStatus 
         Caption         =   "&Patient Status"
         Shortcut        =   ^P
      End
      Begin VB.Menu f 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPatientDischarge 
         Caption         =   "&Discharge"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu info 
      Caption         =   "&Information"
      Begin VB.Menu infohistory 
         Caption         =   "&Patient History"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu mnuinformation 
      Caption         =   "&Bed"
      Begin VB.Menu mnuBedBedPosition 
         Caption         =   "&Bed Position"
         Shortcut        =   ^B
      End
      Begin VB.Menu i 
         Caption         =   "-"
      End
      Begin VB.Menu mnubedbedmodification 
         Caption         =   "&Bed Modification"
         Shortcut        =   ^N
      End
   End
   Begin VB.Menu MNUExpence 
      Caption         =   "&Expence"
      Begin VB.Menu MNUEXPENCEEXPENCE 
         Caption         =   "&Expence"
         Shortcut        =   ^E
      End
      Begin VB.Menu expr 
         Caption         =   "&Expense Report"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu mnustaff 
      Caption         =   "&Staff"
      Begin VB.Menu mnustaffstaffinformation 
         Caption         =   "&Staff Infomation"
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu mnuService 
      Caption         =   "&Service"
      Begin VB.Menu mnuServiceAddservice 
         Caption         =   "&Add Service"
         Shortcut        =   +{F1}
      End
      Begin VB.Menu j 
         Caption         =   "-"
      End
      Begin VB.Menu mnuServiceUpdateservice 
         Caption         =   "&Update Service"
         Shortcut        =   +{F2}
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuAdmitGivenTreatment_Click()

End Sub

Private Sub expr_Click()
Load expancerepo
expancerepo.Show
End Sub

Private Sub Form_Load()

End Sub

Private Sub infohistory_Click()
Load oldpatient
oldpatient.Show

End Sub

Private Sub mnuAdmitTreatment_Click()

  Load treatment_entry
  treatment_entry.Show
  
End Sub

Private Sub mnubedbedmodification_Click()
Load bedmodify
bedmodify.Show

End Sub

Private Sub mnuBedBedPosition_Click()
Load BEDINFO
BEDINFO.Show

End Sub

Private Sub mnuExit_Click()
Unload Me
Load login
login.Show
End
End Sub

Private Sub MNUEXPENCEEXPENCE_Click()
Load payform
payform.Show

End Sub

Private Sub mnuInformationBedPosition_Click()
End Sub

Private Sub mnumodification_Click()

End Sub

Private Sub mnuModificationShiftPatient_Click()
Load shiftpatient
shiftpatient.Show

End Sub

Private Sub mnumodificationupdate_patient_Click()
Load updpatient
updpatient.Show

End Sub

Private Sub mnuPatientAdmitPatient_Click()

Load entry
entry.Show

End Sub

Private Sub mnuPatientDischarge_Click()
Load billing
billing.Show

End Sub

Private Sub mnuPatientPatientStatus_Click()
Load status
status.Show

End Sub

Private Sub mnuPatientRecievePayment_Click()

Load payment
payment.Show

End Sub

Private Sub mnuServiceAddservice_Click()

Load addservice
addservice.Show



End Sub

Private Sub mnuServiceUpdateservice_Click()

Load updatservice
updatservice.Show

End Sub

Private Sub mnustaffstaffinformation_Click()
Load staff
staff.Show

End Sub
