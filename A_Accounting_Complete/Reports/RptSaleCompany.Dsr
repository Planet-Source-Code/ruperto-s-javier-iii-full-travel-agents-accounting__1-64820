VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RptSaleCompany 
   Caption         =   "Sales Report Company"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "RptSaleCompany.dsx":0000
End
Attribute VB_Name = "RptSaleCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private SumSales As Double


Private Sub ActiveReport_NoData()
MsgBox "There are no records to print!", vbInformation, "ELS"
Unload Me
End Sub

Private Sub ActiveReport_ReportStart()
Me.lblRundate = Format(Now, "mm/dd/yyyy")
Me.lblControlNum = Year(Now) & Month(Now) & Day(Now) & Hour(Now) & Minute(Now) & "SRE"
End Sub


Private Sub GroupFooter3_Format()
Me.lblCashDue = Format(CDbl(Me.lblCashDue) + CDbl(Me.txtCashDue), "###,##0.00")
Me.lblSuTotal = Format(CDbl(Me.lblSuTotal) + CDbl(Me.txtSubTotals), "###,##0.00")
Me.lblnsurance = Format(CDbl(Me.lblnsurance) + CDbl(Me.txtSumOfInsurance), "###,##0.00")
Me.lblASF = Format(CDbl(Me.lblASF) + CDbl(Me.txtASF), "###,##0.00")
Me.lblTF = Format(CDbl(Me.lblTF) + CDbl(Me.txtTF), "###,##0.00")
Me.lblMeals = Format(CDbl(Me.lblMeals) + CDbl(Me.txtMeals), "###,##0.00")
Me.lblGross = Format(CDbl(Me.lblGross) + CDbl(Me.txtGross), "###,##0.00")
End Sub

Private Sub GroupHeader3_Format()
Static counter As Double

'counter = CDbl(counter) + CDbl(txtCashDue)
'SumSales = counter

'MsgBox txtCashDue
'lblCashDue.Caption = SumSales
End Sub

