VERSION 5.00
Begin VB.Form frmSetEVAT 
   Caption         =   "Set EVAT"
   ClientHeight    =   2070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5925
   Icon            =   "frmSetEVAT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set"
      Height          =   555
      Left            =   2475
      TabIndex        =   4
      Top             =   1485
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   555
      Left            =   4140
      TabIndex        =   3
      Top             =   1485
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Set Evat for Shipping/Airline"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1350
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   5790
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2340
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   705
         Width           =   3300
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2340
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   3345
      End
      Begin VB.Label Label2 
         Caption         =   "EVAT %"
         Height          =   285
         Left            =   180
         TabIndex        =   6
         Top             =   750
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Shippling/Airline Name :"
         Height          =   330
         Left            =   210
         TabIndex        =   5
         Top             =   300
         Width           =   1980
      End
   End
End
Attribute VB_Name = "frmSetEVAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RsEvat As ADODB.Recordset
Dim SQL As String

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSet_Click()
Call CheckStat
End Sub

Private Sub Form_Load()
Call FillCombo
Me.Combo1.ListIndex = 0
End Sub

Sub EvatUpdate1(ByVal strShip As String, ByVal cAmt As Double)
'On Error GoTo FailSafe_Error
SQL = "UPDATE tbl_RoutePricing SET [EVAT] = " & cAmt & " WHERE [AirlineID]= " & FindAirline(strShip)
cn.BeginTrans
    cn.Execute SQL
cn.CommitTrans
MsgBox "EVAT for " & Me.Combo1 & " successfully updated!!!", vbInformation
Exit Sub
FailSafe_Error:
cn.RollbackTrans
MsgBox "Error updating for " & Me.Combo1 & " please try again...", vbInformation
End Sub


Sub EvatUpdate(ByVal strShip As String, ByVal cAmt As Double)
'On Error GoTo FailSafe_Error
Dim Rs As New ADODB.Recordset

SQL = "SELECT * FROM tbl_RoutePricing WHERE [AirlineID]= " & FindAirline(strShip)
With Rs
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
        cn.BeginTrans
            
            .MoveFirst
            Do While Not .EOF
            .Fields("EVAT").Value = (.Fields("Net Fare").Value) * (0 / 100)
            .Fields("Net Fare").Value = .Fields("Gross Fare").Value + .Fields("Insurance").Value + _
                                        .Fields("ASF").Value + .Fields("Terminal Fee").Value + _
                                        .Fields("Meals").Value + IIf(IsNull(.Fields("Misc").Value), 0, .Fields("Misc").Value) + .Fields("EVAT").Value
            .Update
            .MoveNext
            Loop
            
            .MoveFirst
            Do While Not .EOF
            If FindAirline(strShip) = 36 Or FindAirline(strShip) = 50 Or FindAirline(strShip) = 47 Then
                    .Fields("EVAT").Value = roundUp((.Fields("Net Fare").Value) * (cAmt / 100))
            Else
                    .Fields("EVAT").Value = (.Fields("Net Fare").Value) * (cAmt / 100)
            End If
            
            .Fields("PercentEvat").Value = CDbl(Me.Text1)
            .Fields("Net Fare").Value = .Fields("Gross Fare").Value + .Fields("Insurance").Value + _
                                        .Fields("ASF").Value + .Fields("Terminal Fee").Value + _
                                        .Fields("Meals").Value + IIf(IsNull(.Fields("Misc").Value), 0, .Fields("Misc").Value) + .Fields("EVAT").Value
            .Update
            .MoveNext
            Loop
            
        cn.CommitTrans
        End If
        
If Me.Tag = "-drp-" Then
    Call frmDisplayRoutePricing.Refresh_Grid
End If
        
End With
MsgBox "EVAT for " & Me.Combo1 & " successfully updated!!!", vbInformation
Exit Sub
FailSafe_Error:
cn.RollbackTrans
MsgBox "Error updating for " & Me.Combo1 & " please try again...", vbInformation
End Sub


Sub CheckStat()
Dim Rst As New ADODB.Recordset
Dim ask As Integer
SQL = "SELECT * FROM tbl_RoutePricing WHERE [AirlineID]= " & FindAirline(Me.Combo1)
With Rst
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
               ask = MsgBox("You are about to update " & .RecordCount & " record(s) continue?", vbInformation + vbYesNo)
               If ask = vbYes Then
                Call EvatUpdate(Me.Combo1, CDbl(Me.Text1))
               End If
            Else
                MsgBox "There are no records that match your selection", vbInformation
            End If
End With
End Sub

Sub FillCombo()
Dim tmp As New ADODB.Recordset
SQL = "SELECT * FROM tbl_Airline"
With tmp
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
           If .RecordCount > 0 Then
                Me.Combo1.Clear
                Do While Not .EOF
                    Me.Combo1.AddItem .Fields(1).Value
                .MoveNext
                Loop
           End If
End With
End Sub

Function FindAirline(Param) As Long
Dim tmp As ADODB.Recordset
Set tmp = New ADODB.Recordset
SQL = "SELECT * FROM tbl_Airline WHERE [AirlineName]='" & Param & "'"
With tmp
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
    If .RecordCount > 0 Then
        FindAirline = .Fields(0).Value
      Else
        FindAirline = -1
    End If
    .Close
End With
Set tmp = Nothing
End Function

Private Sub Text1_Change()
If Not IsNumeric(Me.Text1) Then
            Me.Text1 = "0.00"
            Call kulotHL(Me.Text1)
End If
End Sub


Public Function RoundToNext(Value As Long, Interval As Long)
    Dim lngRemainder As Long
    ' Rounds a given Value to the
    'closest matching interval.
    
    lngRemainder = Value Mod (Interval)


    If lngRemainder >= Interval / 2 Then
        RoundToNext = Value + (Interval - lngRemainder)
    Else
        RoundToNext = Value - lngRemainder
    End If
End Function


