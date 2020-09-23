VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RptPOOTHERS 
   Caption         =   "Purchase Order OTHERS"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "RptPOOthers.dsx":0000
End
Attribute VB_Name = "RptPOOTHERS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
Me.lblDate = RetDate(Now)

Me.lblPOnumber = frmPO_Domestic.txtPONumber
cGroupIT
End Sub

Function cGroupIT() As String
Dim i As Integer

With frmPO_Domestic
        If .ListView1.ListItems.Count > 0 Then
                For i = 1 To .ListView1.ListItems.Count
                    If i = 1 Then
                                 Me.lblFrom1 = .ListView1.ListItems(i).SubItems(3)
                                 Me.lblto1 = .ListView1.ListItems(i).SubItems(4)
                                 Me.lblQty1 = .ListView1.ListItems(i).SubItems(5)
                    End If
                    If i = 2 Then
                                 Me.lblFrom2 = .ListView1.ListItems(i).SubItems(3)
                                 Me.lblto2 = .ListView1.ListItems(i).SubItems(4)
                                 Me.lblQty2 = .ListView1.ListItems(i).SubItems(5)
                    End If
                    If i = 3 Then
                                 Me.lblFrom3 = .ListView1.ListItems(i).SubItems(3)
                                 Me.lblto3 = .ListView1.ListItems(i).SubItems(4)
                                 Me.lblQty3 = .ListView1.ListItems(i).SubItems(5)
                    End If
                    If i = 4 Then
                                 Me.lblFrom4 = .ListView1.ListItems(i).SubItems(3)
                                 Me.lblto4 = .ListView1.ListItems(i).SubItems(4)
                                 Me.lblQty4 = .ListView1.ListItems(i).SubItems(5)
                    End If
                    If i = 5 Then
                                 Me.lblFrom5 = .ListView1.ListItems(i).SubItems(3)
                                 Me.lblto5 = .ListView1.ListItems(i).SubItems(4)
                                 Me.lblQty5 = .ListView1.ListItems(i).SubItems(5)
                    End If
                    
                    
                    
                Next i
        End If
End With

End Function

