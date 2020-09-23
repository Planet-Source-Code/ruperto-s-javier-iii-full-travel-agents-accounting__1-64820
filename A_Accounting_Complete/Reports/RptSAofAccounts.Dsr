VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RptSAofAccounts 
   Caption         =   "Statement of Accounts"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "RptSAofAccounts.dsx":0000
End
Attribute VB_Name = "RptSAofAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
Me.lblAsOf = Format(Now, "mm/dd/yyyy")
End Sub

Private Sub Detail_Format()
Static ctr As Long
On Error GoTo FailSafe_Error
ctr = ctr + 1
Me.lblCtr = ctr
Exit Sub
FailSafe_Error:

End Sub
