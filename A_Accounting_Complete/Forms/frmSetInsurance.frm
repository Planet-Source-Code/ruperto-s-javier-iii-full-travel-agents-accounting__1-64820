VERSION 5.00
Begin VB.Form frmSetInsurance 
   Caption         =   "Set Insurance"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5925
   Icon            =   "frmSetInsurance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set"
      Height          =   555
      Left            =   2475
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   555
      Left            =   4170
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Set Insurance for Shipping/Airline"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   5790
      Begin VB.PictureBox Picture1 
         Enabled         =   0   'False
         Height          =   840
         Left            =   90
         ScaleHeight     =   780
         ScaleWidth      =   5520
         TabIndex        =   5
         Top             =   210
         Width           =   5580
         Begin VB.TextBox txtRouteTo 
            Height          =   315
            Left            =   3615
            TabIndex        =   10
            Top             =   375
            Width           =   1860
         End
         Begin VB.TextBox txtRouteFrom 
            Height          =   315
            Left            =   1785
            TabIndex        =   9
            Top             =   375
            Width           =   1815
         End
         Begin VB.TextBox txtAirline 
            Height          =   315
            Left            =   1785
            TabIndex        =   8
            Top             =   30
            Width           =   3690
         End
         Begin VB.Label Label3 
            Caption         =   "Route :"
            Height          =   285
            Left            =   75
            TabIndex        =   7
            Top             =   375
            Width           =   840
         End
         Begin VB.Label Label1 
            Caption         =   "Airline/ Shipping Line :"
            Height          =   285
            Left            =   75
            TabIndex        =   6
            Top             =   90
            Width           =   1680
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1905
         TabIndex        =   1
         Text            =   "0.00"
         Top             =   1140
         Width           =   3705
      End
      Begin VB.Label Label2 
         Caption         =   "Insurance :"
         Height          =   285
         Left            =   195
         TabIndex        =   4
         Top             =   1185
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmSetInsurance"
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
With Me
        .txtAirline = frmDisplayRoutePricing.Combo1
        .txtRouteFrom = frmDisplayRoutePricing.Combo2
        .txtRouteTo = frmDisplayRoutePricing.Combo3
End With
End Sub

Sub EvatUpdate1(ByVal strShip As String, ByVal cAmt As Double)
'On Error GoTo FailSafe_Error
SQL = "UPDATE tbl_RoutePricing SET [EVAT] = " & cAmt & " WHERE [AirlineID]= " & FindAirline(strShip)
cn.BeginTrans
    cn.Execute SQL
cn.CommitTrans
'MsgBox "EVAT for " & Me.Combo1 & " successfully updated!!!", vbInformation
Exit Sub
FailSafe_Error:
cn.RollbackTrans
'MsgBox "Error updating for " & Me.Combo1 & " please try again...", vbInformation
End Sub


Sub InsUpdate(ByVal strShip As String, ByVal cAmt As Double)
'On Error GoTo FailSafe_Error
Dim Rs As New ADODB.Recordset

'SQL = "SELECT * FROM tbl_RoutePricing WHERE [AirlineID]= " & FindAirline(strShip)
SQL = "SELECT * FROM tbl_RoutePricing WHERE [AirlineID]= " & FindAirline(strShip) & " AND [RouteID]= " & ReturnRouteID(Me.txtRouteFrom, Me.txtRouteTo)
With Rs
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
       
        cn.BeginTrans
            
            .MoveFirst
            Do While Not .EOF
            .Fields("EVAT").Value = 0
            .Fields("Net Fare").Value = .Fields("Gross Fare").Value + .Fields("Insurance").Value + _
                                        .Fields("ASF").Value + .Fields("Terminal Fee").Value + _
                                        .Fields("Meals").Value + .Fields("EVAT").Value
            .Update
            .MoveNext
            Loop
            
            .MoveFirst
            Do While Not .EOF
            If FindAirline(strShip) = 36 Then
                    .Fields("EVAT").Value = roundUp((.Fields("Net Fare").Value) * (IIf(IsNull(.Fields("PercentEvat").Value), 0, .Fields("PercentEvat").Value) / 100))
            Else
                    .Fields("EVAT").Value = (.Fields("Net Fare").Value) * (IIf(IsNull(.Fields("PercentEvat").Value), 0, .Fields("PercentEvat").Value) / 100)
            End If
            
            .Fields("Insurance").Value = CDbl(Me.Text1)
            .Fields("Net Fare").Value = .Fields("Gross Fare").Value + .Fields("Insurance").Value + _
                                        .Fields("ASF").Value + .Fields("Terminal Fee").Value + _
                                        .Fields("Meals").Value + .Fields("EVAT").Value
            .Update
            .MoveNext
            Loop
        cn.CommitTrans
        End If
        
End With
MsgBox "INSURANCE for " & Me.txtAirline & " successfully updated!!!", vbInformation
Call frmDisplayRoutePricing.Refresh_Grid
Exit Sub
FailSafe_Error:
cn.RollbackTrans
MsgBox "Error updating for " & Me.txtAirline & " please try again...", vbInformation
End Sub

Function ReturnRouteID(ByVal strFrom As String, ByVal strTo As String) As Long
Dim Rst As New ADODB.Recordset
SQL = "SELECT * FROM tbl_Routes WHERE [From]= '" & strFrom & "' AND [To] ='" & strTo & "'"
With Rst
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
       If .RecordCount > 0 Then
            ReturnRouteID = .Fields(0).Value
       Else
            ReturnRouteID = -1
       End If
       .Close
    Set Rst = Nothing
End With
End Function


Sub CheckStat()
Dim Rst As New ADODB.Recordset
Dim ask As Integer
SQL = "SELECT * FROM tbl_RoutePricing WHERE [AirlineID]= " & FindAirline(Me.txtAirline) & " AND [RouteID]= " & ReturnRouteID(Me.txtRouteFrom, Me.txtRouteTo)
With Rst
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
               ask = MsgBox("You are about to update " & .RecordCount & " record(s) continue?", vbInformation + vbYesNo, "Update Insurance")
               If ask = vbYes Then
                Call InsUpdate(Me.txtAirline, CDbl(Me.Text1))
               End If
            Else
                MsgBox "There are no records that match your selection", vbInformation
            End If
End With
End Sub

Sub FillCombo()
Dim Tmp As New ADODB.Recordset
SQL = "SELECT * FROM tbl_Airline"
With Tmp
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
           If .RecordCount > 0 Then
                'Me.Combo1.Clear
                Do While Not .EOF
                    'Me.Combo1.AddItem .Fields(1).Value
                .MoveNext
                Loop
           End If
End With
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


