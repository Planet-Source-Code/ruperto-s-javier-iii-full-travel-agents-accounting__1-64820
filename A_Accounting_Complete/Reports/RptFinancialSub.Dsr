VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RptFinancialSub 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "RptFinancialSub.dsx":0000
End
Attribute VB_Name = "RptFinancialSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SQL As String
Dim myDate As Date

Function ReturnBal(ByVal sDate As Date, ByVal Desc As String, ByVal nBank, Optional flag As String, Optional nCard As String, Optional nPDC As Boolean) As Double
Dim Rst As New ADODB.Recordset
Dim tmpBal As Double

If Desc = "Cash" Then
    If flag = 1 Then
    SQL = "SELECT * FROM qryBankPassbook " & _
          "WHERE  ( (([Deposit Date])=#" & sDate & "#) AND ((Description)='" & Desc _
          & "') AND (([Bank1])='" & nBank & "'))"
    End If
    If flag = 2 Then
    SQL = "SELECT * FROM qryBankPassbook " & _
          "WHERE  ( (([Deposit Date])=#" & sDate & "#) AND ((Description)='" & Desc _
          & "') AND (([Bank2])='" & nBank & "'))"
    End If
    If flag = 3 Then
    SQL = "SELECT * FROM qryBankPassbook " & _
          "WHERE  ( (([Deposit Date])=#" & sDate & "#) AND ((Description)='" & Desc _
          & "') AND (([Bank3])='" & nBank & "'))"
    End If

End If
      
If Desc = "Check" Then

    If flag = 1 Then
    SQL = "SELECT * FROM qryBankPassbook " & _
          "WHERE  ( (([Deposit Date])=#" & sDate & "#) AND ((Description)='" & Desc _
          & "') AND (([Bank1])='" & nBank & "') AND (([Post Dated])=" & nPDC & "))"
    End If
    
    If flag = 2 Then
    SQL = "SELECT * FROM qryBankPassbook " & _
          "WHERE  ( (([Deposit Date])=#" & sDate & "#) AND ((Description)='" & Desc _
          & "') AND (([Bank2])='" & nBank & "') AND (([Post Dated])=" & nPDC & "))"
    End If
    
    If flag = 3 Then
    SQL = "SELECT * FROM qryBankPassbook " & _
          "WHERE  ( (([Deposit Date])=#" & sDate & "#) AND ((Description)='" & Desc _
          & "') AND (([Bank3])='" & nBank & "') AND (([Post Dated])=" & nPDC & "))"
    End If
End If
      
      
If Desc = "Card" Then
    If flag = 1 Then
    SQL = "SELECT * FROM qryBankPassbook " & _
          "WHERE  ( (([Deposit Date])=#" & sDate & "#) AND ((Description)='" & Desc _
          & "') AND (([Bank1])='" & nBank & "') AND (([Card Name])='" & nCard & "') )"
    End If
    If flag = 2 Then
    SQL = "SELECT * FROM qryBankPassbook " & _
          "WHERE  ( (([Deposit Date])=#" & sDate & "#) AND ((Description)='" & Desc _
          & "') AND (([Bank2])='" & nBank & "') AND (([Card Name])='" & nCard & "') )"
    End If
    If flag = 3 Then
    SQL = "SELECT * FROM qryBankPassbook " & _
          "WHERE  ( (([Deposit Date])=#" & sDate & "#) AND ((Description)='" & Desc _
          & "') AND (([Bank3])='" & nBank & "') AND (([Card Name])='" & nCard & "') )"
    End If
End If
      
      
With Rst

        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            
            If Desc = "Card" Then
            tmpBal = 0
              .MoveFirst
              Do While Not .EOF
                    tmpBal = tmpBal + .Fields("Card Amount").Value
                    .MoveNext
              Loop
                ReturnBal = tmpBal
            End If
            
'//=======================================
            If Desc = "Cash" Then
            tmpBal = 0
              .MoveFirst
              Do While Not .EOF
                    tmpBal = tmpBal + .Fields("Cash Amount").Value
                    .MoveNext
              Loop
                ReturnBal = tmpBal
            End If
'//=======================================

            If Desc = "Check" Then
                tmpBal = 0
                .MoveFirst
                    Do While Not .EOF
                          tmpBal = tmpBal + .Fields("Check Amount").Value
                          .MoveNext
                    Loop
                ReturnBal = tmpBal
                
            End If
        Else
            ReturnBal = 0
        End If
        .Close
     Set Rst = Nothing
End With

End Function

Private Sub ActiveReport_ReportStart()
myDate = Format(frmReportRange.DTPicker2.Value, "mm/dd/yyyy")
End Sub

Private Sub GroupHeader2_Format()

Dim mytmpCash As Double
Dim mytmpCheck As Double
Dim mytmpCard As Double
Dim mytotCashCheck As Double

'myDate = Format(Now, "MM/DD/YYYY")

'//-------------------------------------------------------
'// For mbtc-1 pal
'//-------------------------------------------------------
mytmpCash = ReturnBal(myDate, "Cash", "PAL-MBTC", 1)
mytmpCheck = ReturnBal(myDate, "Check", "PAL-MBTC", 2, , False)
mytotCashCheck = CDbl(mytmpCash) + CDbl(mytmpCheck)
Me.lblCashCheck1 = Format(mytotCashCheck, "###,##0.00")

'//-------------------------------------------------------
'//CP
'//-------------------------------------------------------
mytmpCash = ReturnBal(myDate, "Cash", "CP-MBTC", 1)
mytmpCheck = ReturnBal(myDate, "Check", "CP-MBTC", 2, , False)
mytotCashCheck = CDbl(mytmpCash) + CDbl(mytmpCheck)
Me.lblCashCheck2 = Format(mytotCashCheck, "###,##0.00")



'//-------------------------------------------------------
'//AP
'//-------------------------------------------------------
mytmpCash = ReturnBal(myDate, "Cash", "AP/AS-EPCI", 1)
mytmpCheck = ReturnBal(myDate, "Check", "AP/AS-EPCI", 2, , False)
mytotCashCheck = CDbl(mytmpCash) + CDbl(mytmpCheck)
Me.lblCashCheck3 = Format(mytotCashCheck, "###,##0.00")

'//-------------------------------------------------------
'//CK,NN,TP/TA
'//-------------------------------------------------------
mytmpCash = ReturnBal(myDate, "Cash", "CK/NN/WGA/TA-EPCI", 1)
mytmpCheck = ReturnBal(myDate, "Check", "CK/NN/WGA/TA-EPCI", 2, , False)
mytotCashCheck = CDbl(mytmpCash) + CDbl(mytmpCheck)
Me.lblCashCheck4 = Format(mytotCashCheck, "###,##0.00")





'//-----------------------------------------------------------------------------------
'//FOR CARD PAYMENT
'//-----------------------------------------------------------------------------------
'//                             ==>MASTERCARD<==
'// 1.  PAL
mytmpCard = ReturnBal(myDate, "Card", "PAL-MBTC", 3, "MASTERCARD")
Me.lblEPCIcard1 = Format(mytmpCard, "###,##0.00")
'// 2.  CEBU-PAC
mytmpCard = ReturnBal(myDate, "Card", "CP-MBTC", 3, "MASTERCARD")
Me.lblEPCIcard2 = Format(mytmpCard, "###,##0.00")
'// 3.  AP-AS
mytmpCard = ReturnBal(myDate, "Card", "AP/AS-EPCI", 3, "MASTERCARD")
Me.lblEPCIcard3 = Format(mytmpCard, "###,##0.00")
'// 4.  CK-NN
mytmpCard = ReturnBal(myDate, "Card", "CK/NN/WGA/TA-EPCI", 3, "MASTERCARD")
Me.lblEPCIcard4 = Format(mytmpCard, "###,##0.00")
'//-----------------------------------------------------------------------------------
'//                               ==>DINERS<==
mytmpCard = ReturnBal(myDate, "Card", "PAL-MBTC", 3, "DINERS")
Me.lblDiners1 = Format(mytmpCard, "###,##0.00")
mytmpCard = ReturnBal(myDate, "Card", "CP-MBTC", 3, "DINERS")
Me.lblDiners2 = Format(mytmpCard, "###,##0.00")
mytmpCard = ReturnBal(myDate, "Card", "AP/AS-EPCI", 3, "DINERS")
Me.lblDiners3 = Format(mytmpCard, "###,##0.00")
mytmpCard = ReturnBal(myDate, "Card", "CK/NN/WGA/TA-EPCI", 3, "DINERS")
Me.lblDiners4 = Format(mytmpCard, "###,##0.00")


'//-----------------------------------------------------------------------------------
'//FOR CARD PAYMENT
'//-----------------------------------------------------------------------------------
'//For Postaded check
mytmpCheck = ReturnBal(myDate, "Check", "PAL-MBTC", 2, , True)
Me.lblPostDated1 = Format(mytmpCheck, "###,##0.00")

mytmpCheck = ReturnBal(myDate, "Check", "CP-MBTC", 2, , True)
Me.lblPostDated2 = Format(mytmpCheck, "###,##0.00")

mytmpCheck = ReturnBal(myDate, "Check", "AP/AS-EPCI", 2, , True)
Me.lblPostDated3 = Format(mytmpCheck, "###,##0.00")

mytmpCheck = ReturnBal(myDate, "Check", "CK/NN/WGA/TA-EPCI", 2, , True)
Me.lblPostDated4 = Format(mytmpCheck, "###,##0.00")



'//For pal HSBC only
mytmpCard = ReturnBal(myDate, "Card", "PAL-HSBC", 3, "PAL-HSBC")
Me.lblEPCIcard5 = Format(mytmpCard, "###,##0.00")


End Sub

