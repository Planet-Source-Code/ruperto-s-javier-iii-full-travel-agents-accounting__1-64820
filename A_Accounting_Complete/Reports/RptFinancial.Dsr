VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RptFinancial 
   Caption         =   "FINANCIAL"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "RptFinancial.dsx":0000
End
Attribute VB_Name = "RptFinancial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
Set SubReport1.object = New RptFinancialSub
    'SubReport1.object.RecordSource = "SELECT * FROM qryBankPassbook"
End Sub

Private Sub Detail_Format()
On Error Resume Next
Me.lblTotLine = Format(CDbl(Me.Field3) + CDbl(Me.Field4) + CDbl(Me.Field5) + CDbl(Me.Field6), "###,##0.00")

End Sub

Private Sub GroupFooter1_Format()
On Error Resume Next
Me.lblTotal = Format(CDbl(Me.Field8) + CDbl(Me.Field9) + CDbl(Me.Field10) + CDbl(Me.Field11), "###,##0.00")
End Sub

