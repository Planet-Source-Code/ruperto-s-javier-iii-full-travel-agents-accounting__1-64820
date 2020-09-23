VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RptStatement 
   Caption         =   "Statement"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "RptStatement.dsx":0000
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "RptStatement.dsx":27A2
End
Attribute VB_Name = "RptStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Integer
Dim PreviousVal As String

Private Sub ActiveReport_ReportStart()
Me.lblRundate = Me.lblRundate & " " & Format(Now, "mm/dd/yyyy")

End Sub



Private Sub Detail_Format()

Me.Field5.Text = UCase(Mid(Me.Field5, 1, 3)) ' Ticket type
If Me.Field7 = PreviousVal Then
    
    Me.Field8.Visible = False
    Me.Field7.Visible = False
    Me.Field6.Visible = False
    
Else
    PreviousVal = Me.Field7.Text
    

    Me.Field8.Visible = True
    Me.Field7.Visible = True
    Me.Field6.Visible = True
    
    Me.Field5.Text = UCase(Mid(Me.Field5, 1, 3)) ' Ticket type
    Me.Field6.Text = UCase(Me.Field6.Text)  ' Pax Name
    Me.txtRoute.Text = UCase(Me.txtRoute.Text)
End If

End Sub

