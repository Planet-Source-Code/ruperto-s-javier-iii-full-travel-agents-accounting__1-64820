VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RptOR 
   Caption         =   "OR"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "RptOR.dsx":0000
End
Attribute VB_Name = "RptOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private SumSales As Double

Private Sub ActiveReport_ReportStart()
Me.lblDate = Format(Now, "mm/dd/yyyy") & " " & Format(Now, "hh:mm AMPM")
Me.Printer.PaperSize = 256
Me.Printer.PaperWidth = 1440 * 5
Me.Printer.PaperHeight = 1440 * 6
Me.lblCustName = frmCashier.txtReceivedFrom
Me.lblAddress = frmCashier.txtAddress
End Sub

Private Sub GroupHeader2_Format()
SumSales = SumSales + CDbl(Me.txtAmount)
Me.lblSum = Format(SumSales, "###,##0.00")
End Sub

