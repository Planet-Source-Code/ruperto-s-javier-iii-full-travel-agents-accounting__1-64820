VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RptAging 
   Caption         =   "Aging Accounts Report"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "RptAging.dsx":0000
End
Attribute VB_Name = "RptAging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function CountDays(ByVal userAccount As String, ByVal UserLowLimit As Long, ByVal UserUpLimit As Long) As Long
Dim Rs As New ADODB.Recordset
Dim SQL As String

SQL = "SELECT * FROM qryAging WHERE [AccountNo]='" & userAccount & "' AND [ElapsedDay]>=" & UserLowLimit & " AND [ElapsedDay]<=" & UserUpLimit
With Rs
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            CountDays = .RecordCount
        Else
            CountDays = 0
        End If
        
        .Close
     Set Rs = Nothing
     
End With

End Function

Private Sub GroupHeader1_Format()
Me.lbl_0_15 = CountDays(Me.txtAccountNo, 0, 15)
Me.lbl_16_30 = CountDays(Me.txtAccountNo, 16, 30)
Me.lbl_31_60 = CountDays(Me.txtAccountNo, 31, 60)
Me.lbl_61_90 = CountDays(Me.txtAccountNo, 61, 90)
Me.lbl_91_180 = CountDays(Me.txtAccountNo, 91, 180)
Me.lbl_181_360 = CountDays(Me.txtAccountNo, 181, 360)
Me.lbl_over_360 = CountDays(Me.txtAccountNo, 361, 9000)
Me.lbl_total = CDbl(Me.lbl_0_15) + CDbl(Me.lbl_16_30) + CDbl(Me.lbl_31_60) + CDbl(Me.lbl_61_90) + CDbl(Me.lbl_91_180) + CDbl(Me.lbl_181_360) + CDbl(Me.lbl_over_360)
End Sub
