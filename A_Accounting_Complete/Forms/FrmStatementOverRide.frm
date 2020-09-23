VERSION 5.00
Begin VB.Form FrmStatementOverRide 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Statement Over-Ride"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "What would you like to do?"
      Height          =   2520
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   3510
      Begin VB.OptionButton Option3 
         Caption         =   "Set Commission"
         Height          =   330
         Left            =   465
         TabIndex        =   5
         Top             =   1365
         Value           =   -1  'True
         Width           =   2820
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   450
         Left            =   1935
         TabIndex        =   4
         Top             =   1860
         Width           =   1410
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         Height          =   450
         Left            =   300
         TabIndex        =   3
         Top             =   1860
         Width           =   1410
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Edit the Transaction (Update Only)"
         Height          =   330
         Left            =   465
         TabIndex        =   2
         Top             =   870
         Width           =   2820
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cancel the Transaction (Delete)"
         Height          =   330
         Left            =   465
         TabIndex        =   1
         Top             =   435
         Width           =   2640
      End
   End
End
Attribute VB_Name = "FrmStatementOverRide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim SQL As String

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
'On Error GoTo ErrExit
Dim Rstmp           As New ADODB.Recordset
Dim ask             As Integer
Dim strSa           As String
Dim numAirline      As Long
Dim myStrMsg        As String

Set Rs = New ADODB.Recordset
strSa = frmSelectStatement.DataGrid1.Columns(1).Text
numAirline = frmSelectStatement.DataGrid1.Columns(2).Text


myStrMsg = "Warning!!!!" & Chr(13) & _
                            "Statement #:" & strSa & " are aboout to be cancelled! " & Chr(13) & _
                            "Cancelling the statement will return the tickets" & Chr(13) & _
                            "to its Un-Sold state and will remove the statement" & Chr(13) & _
                            "from the list of sales prior to its date" & Chr(13) & _
                            "continue?"

'//-----------------------------------------------------------------------------------------------------------
'//For Statement Cancellation
'//-----------------------------------------------------------------------------------------------------------
If Me.Option1 Then
    ask = MsgBox(myStrMsg, vbCritical + vbYesNo, "Confirm")
    If ask = vbNo Then
        MsgBox "The action was aborted...", vbInformation
        Exit Sub
    End If


        SQL = "SELECT * FROM tbl_Statement WHERE [sNumber]='" & strSa & "'"
    With Rs
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
       
                SQL = "SELECT * from tbl_StatementDetail WHERE [TransID]= " & FindTransID(strSa)
                With Rstmp
                        .Open SQL, cn, adOpenKeyset, adLockOptimistic
                        If .RecordCount > 0 Then
                        
                            .MoveFirst
                            Do While Not .EOF
                                Call Return_Ticket(.Fields("Ticket No").Value)
                            .MoveNext
                            Loop
                                MsgBox .RecordCount & " tickets mark as Un-Sold", vbInformation
                        End If
                End With
        
      ''GoTo xxx
                SQL = "DELETE * FROM tbl_Statement WHERE [sNumber]='" & strSa & "'"
                cn.BeginTrans
                    cn.Execute SQL
                cn.CommitTrans
''xxx:
                .Close
                Set Rs = Nothing
                MsgBox "Statement successfully cancelled...", vbInformation
                Else
                MsgBox "This statement already deleted or does not exist", vbCritical
        End If
        
    End With
End If

'//-----------------------------------------------------------------------------------------------------------
'//For Statement edit
'//-----------------------------------------------------------------------------------------------------------

If Me.Option2 Then
Call frmStatement.Clear

        SQL = "SELECT * FROM tbl_Statement WHERE [sNumber]='" & strSa & "'"
    With Rs
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            With frmStatement
                    .txtNo = strSa
                    .txtAgencyName = IIf(IsNull(Rs.Fields("AgencyName").Value), "", Rs.Fields("AgencyName").Value)
                    .CboAccountName = Rs.Fields("AccountNo").Value
                    .Text2 = IIf(IsNull(Rs.Fields("Telno").Value), "", Rs.Fields("Telno").Value)
                    .CboAccountName = Rs.Fields("AccountNo").Value
                    .Combo1 = IIf(IsNull(Rs.Fields("Airline").Value), "", FindAirline(Rs.Fields("Airline").Value))
                    .txtMisc = IIf(IsNull(Rs.Fields("MiscAmt").Value), "0.00", Rs.Fields("MiscAmt").Value)
                    .txtIssuedBy = IIf(IsNull(Rs.Fields("Issued by").Value), "", Rs.Fields("Issued by").Value)
            End With
            frmStatement.Tag = "over_ride"
            Call Return_Tickets_COmm(strSa)
            'MsgBox "Tickets succesfully retrieve", vbInformation
        End If
    End With
End If
'//-----------------------------------------------------------------------------------------------------------
'//For Statement edit
'//-----------------------------------------------------------------------------------------------------------

If Me.Option3 Then
        frmStatement.txtCommPercent.Enabled = True
        Call kulotHL(frmStatement.txtCommPercent)
End If

Unload Me
Exit Sub
ErrExit:
cn.RollbackTrans
MsgBox "An Error occured while deleting the statement.." & Chr(13) & _
        Err.Description, vbCritical
End Sub

Function Return_Ticket(ByVal strTicket As String) As Boolean
Dim RsCheck As New ADODB.Recordset

SQL = "SELECT * FROM tbl_Tickets WHERE [Ticket No]='" & strTicket & "' AND [Status]='Sold'"
With RsCheck
            .Open SQL, cn, adOpenKeyset, adLockOptimistic

                    If .RecordCount > 0 Then
                      .Fields("Status").Value = "Un-Sold"
                      .Update
                    End If
          
           .Close
        Set RsCheck = Nothing
End With
End Function

Function Return_Tickets_COmm(ByVal param As String) As Double
Dim RsLook As New ADODB.Recordset
Dim Rstmp As New ADODB.Recordset
Dim ctr As Long
Dim STRSQL As String
Dim mytempDate As String


SQL = "SELECT * FROM tbl_Statement WHERE [sNumber]='" & param & "'"

ctr = 0
With RsLook
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
        mytempDate = .Fields("Date").Value
            SQL = "SELECT * from tbl_StatementDetail WHERE [TransID]= " & FindTransID(param)
                With Rstmp
                        .Open SQL, cn, adOpenKeyset, adLockOptimistic
                        If .RecordCount > 0 Then
                            .MoveFirst
                            frmStatement.txtCommPercent = .Fields("Commision").Value
                            Do While Not .EOF
                               frmStatement.txtTicketNo(ctr).Text = .Fields("Ticket No").Value
                               frmStatement.txtPassengerName(ctr).Text = .Fields("Name").Value
                               frmStatement.cboTicketType(ctr).Text = .Fields("Ticket Type").Value
                               frmStatement.cboPassengerType(ctr).Text = IIf(.Fields("Void").Value = True, "VOID", "")
                               Call frmStatement.FillRoutes(frmStatement.FindTicketTypeID(frmStatement.txtTicketNo(ctr)))
                               Call Return_Routes(.Fields("StatementDetails").Value, ctr)
                               ctr = ctr + 1
                            .MoveNext
                            Loop
                          frmStatement.txtDate = mytempDate
                          MsgBox ctr & " tickets succesfully retrieved...", vbInformation
                          frmStatement.Caption = "over_ride"
                        End If
                End With
        End If
        .Close
End With
Set RsLook = Nothing

End Function

Function Return_Routes(ByVal nDetailsID As Long, ByVal Index As Integer) As Boolean
Dim RsRoute As New ADODB.Recordset
Dim SQL As String
Dim ctr As Long

SQL = "SELECT * FROM tbl_StatementTickets WHERE [StatementDetails] = " & nDetailsID
ctr = 0
With RsRoute
      .Open SQL, cn, adOpenKeyset, adLockOptimistic
      If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                
                If ctr = 0 Then
                    frmStatement.cboFrom(Index).Text = .Fields("Route").Value
                End If
                
                If ctr = 1 Then
                    frmStatement.cboDest1(Index).Text = .Fields("Route").Value
                End If
                
                If ctr = 2 Then
                    frmStatement.cboDest2(Index).Text = .Fields("Route").Value
                End If
                
                If ctr = 3 Then
                    frmStatement.cboDest3(Index).Text = .Fields("Route").Value
                End If
                
                
                ctr = ctr + 1
            .MoveNext
            Loop
      End If
End With

End Function

Function FindTransID(ByVal param As String) As Long
Dim RsLook As New ADODB.Recordset
Dim STRSQL As String
SQL = "SELECT * FROM tbl_Statement WHERE [sNumber]='" & param & "'"
With RsLook
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            FindTransID = .Fields(0).Value
            Else
            FindTransID = -1
        End If
        .Close
End With
Set RsLook = Nothing
End Function

Function FindAirline(ByVal param As Long) As String
Dim Tmp As ADODB.Recordset
Set Tmp = New ADODB.Recordset
SQL = "SELECT * FROM tbl_Airline WHERE [AirlineID]=" & param
With Tmp
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
    If .RecordCount > 0 Then
        FindAirline = .Fields(1).Value
      Else
        FindAirline = "none"
    End If
    .Close
End With
Set Tmp = Nothing
End Function

