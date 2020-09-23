VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RptProfit 
   Caption         =   "Profit Report"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "RptProfit.dsx":0000
End
Attribute VB_Name = "RptProfit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ActiveReport_NoData()
MsgBox "Nothing to print!!!", vbInformation + vbOKOnly
Unload Me
End Sub

Private Sub ActiveReport_ReportStart()
Me.lblRundate = Format(Now, "d-mmmm-yyyy")
Me.lblFrom = frmReportSel.DTPicker1
Me.lblTo = frmReportSel.DTPicker2



End Sub


Private Sub GroupFooter3_Format()
Static ctr As Long
'On Error GoTo FailSafe_Error
On Error Resume Next

If FindAirline(frmReportSel.Combo1) = 36 Then
        Me.txtEEvat = Format(roundUp(Me.txtEEvat), "###,##0.00")
    Else
        Me.txtEEvat = Format(Me.txtEEvat, "###,##0.00")
End If

'Me.txtCashDue = Format(CDbl(Me.txtEEvat) + CDbl(Me.txtSubTotals), "###,##0.00")
'Me.lblCashDue = Format(CDbl(Me.lblCashDue) + CDbl(Me.txtCashDue), "###,##0.00")
Me.lblSuTotal = Format(CDbl(Me.lblSuTotal) + CDbl(Me.txtSubTotals), "###,##0.00")

Me.lblEvat = Format(roundUp(Me.lblEvat) + roundUp(Me.txtEEvat), "###,##0.00")
Me.lblnsurance = Format(CDbl(Me.lblnsurance) + CDbl(Me.txtSumOfInsurance), "###,##0.00")
Me.lblASF = Format(CDbl(Me.lblASF) + CDbl(Me.txtASF), "###,##0.00")
Me.lblTF = Format(CDbl(Me.lblTF) + CDbl(Me.txtTF), "###,##0.00")
Me.lblMeals = Format(CDbl(Me.lblMeals) + CDbl(Me.txtMeals), "###,##0.00")
Me.lblGross = Format(CDbl(Me.lblGross) + CDbl(Me.txtGross), "###,##0.00")
Me.lblVat = Format(CDbl(Me.lblVat) + CDbl(Me.txtVat), "###,##0.00")
Me.lblComm = Format(CDbl(Me.lblComm) + CDbl(Me.txtCommAmt), "###,##0.00")
'Me.lblNetProfit = Format(CDbl(Me.Field2) - CDbl(Me.txtCashDue), "###,##0.00")
ctr = ctr + 1
Me.lblCtr = ctr
Exit Sub
FailSafe_Error:
End Sub

Function FindAirline(Param) As Long
Dim Tmp As ADODB.Recordset
Set Tmp = New ADODB.Recordset
SQL = "SELECT * FROM tbl_Airline WHERE [AirlineName]='" & Param & "'"
With Tmp
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
    If .RecordCount > 0 Then
        FindAirline = .Fields(0).Value
      Else
        FindAirline = -1
    End If
    .Close
End With
Set Tmp = Nothing
End Function

