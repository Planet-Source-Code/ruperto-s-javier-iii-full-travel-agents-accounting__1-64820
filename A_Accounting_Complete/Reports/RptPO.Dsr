VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RptPO 
   Caption         =   "Purchase Order"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "RptPO.dsx":0000
End
Attribute VB_Name = "RptPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
Me.lblDate = RetDate(Now)
Me.lblPOnumber = frmPO_Domestic.txtPONumber
End Sub

Private Sub Detail_Format()
Static iRow As Integer
    If iRow Mod 2 = 0 Then
        Detail.BackColor = &HE0E0E0
    Else
        Detail.BackColor = &HC0C0FF
    End If
    iRow = iRow + 1
   
End Sub
