VERSION 5.00
Begin VB.Form frmSetRefundFee 
   Caption         =   "Set Refund Fee"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5925
   Icon            =   "frmSetRefundFee.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Edit"
      Height          =   555
      Left            =   1755
      TabIndex        =   27
      Top             =   4125
      Width           =   1695
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set"
      Height          =   555
      Left            =   75
      TabIndex        =   11
      Top             =   4125
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   555
      Left            =   4170
      TabIndex        =   12
      Top             =   4125
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Set Refund/Service/Void fee for Shipping/Airline"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   5790
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2340
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   690
         Width           =   3345
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Important Fill this for Refund"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2595
         Left            =   165
         TabIndex        =   14
         Top             =   1200
         Width           =   5490
         Begin VB.TextBox txtBaseNSF 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   1365
            TabIndex        =   9
            Text            =   "0.00"
            Top             =   1920
            Width           =   1035
         End
         Begin VB.TextBox txtBaseVF 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   1365
            TabIndex        =   7
            Text            =   "0.00"
            Top             =   1575
            Width           =   1035
         End
         Begin VB.TextBox txtBaseSF 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   1365
            TabIndex        =   3
            Text            =   "0.00"
            Top             =   825
            Width           =   1035
         End
         Begin VB.TextBox txtBaseRf 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   1365
            TabIndex        =   5
            Text            =   "0.00"
            Top             =   1200
            Width           =   1035
         End
         Begin VB.TextBox txtEvatNSF 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   2415
            TabIndex        =   10
            Text            =   "0.00"
            Top             =   1920
            Width           =   1035
         End
         Begin VB.TextBox txtEvatVF 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   2415
            TabIndex        =   8
            Text            =   "0.00"
            Top             =   1575
            Width           =   1035
         End
         Begin VB.TextBox txtEvatSF 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   2415
            TabIndex        =   4
            Text            =   "0.00"
            Top             =   825
            Width           =   1035
         End
         Begin VB.TextBox txtEvatRF 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   2415
            TabIndex        =   6
            Text            =   "0.00"
            Top             =   1200
            Width           =   1035
         End
         Begin VB.TextBox txtRefund 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   3855
            TabIndex        =   18
            Text            =   "0.00"
            Top             =   1200
            Width           =   1035
         End
         Begin VB.TextBox txtServiceFee 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   3855
            TabIndex        =   17
            Text            =   "0.00"
            Top             =   825
            Width           =   1035
         End
         Begin VB.TextBox txtVoidFee 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   3855
            TabIndex        =   16
            Text            =   "0.00"
            Top             =   1575
            Width           =   1035
         End
         Begin VB.TextBox txtNoShowFee 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   3855
            TabIndex        =   15
            Text            =   "0.00"
            Top             =   1920
            Width           =   1035
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Fee(s)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3975
            TabIndex        =   26
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "EVAT %"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   2475
            TabIndex        =   25
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Base Fee"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1395
            TabIndex        =   24
            Top             =   465
            Width           =   1020
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Refund Fee:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   90
            TabIndex        =   22
            Top             =   1170
            Width           =   1260
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Service Fee:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   105
            TabIndex        =   21
            Top             =   810
            Width           =   1260
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Void Fee :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   105
            TabIndex        =   20
            Top             =   1560
            Width           =   1260
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "NoShow Fee :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   105
            TabIndex        =   19
            Top             =   1935
            Width           =   1260
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2340
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   3345
      End
      Begin VB.Label Label4 
         Caption         =   "Fare Basis :"
         Height          =   330
         Left            =   210
         TabIndex        =   23
         Top             =   660
         Width           =   1980
      End
      Begin VB.Label Label1 
         Caption         =   "Shippling/Airline Name :"
         Height          =   330
         Left            =   210
         TabIndex        =   13
         Top             =   300
         Width           =   1980
      End
   End
End
Attribute VB_Name = "frmSetRefundFee"
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

Private Sub Combo1_Change()
LoadValues
End Sub

Private Sub Combo1_Click()
LoadValues
End Sub

Private Sub Combo2_Change()
LoadValues
End Sub

Private Sub Combo2_Click()
LoadValues
End Sub

Private Sub Command1_Click()
Me.Frame4.Enabled = True
Call kulotHL(txtBaseSF)
End Sub

Private Sub Command2_Click()
Call Recalc
End Sub

Private Sub Form_Load()
Call FillCombo
Me.Combo1.ListIndex = 0
FillComboFareBasis
Me.Combo2.ListIndex = 0
End Sub

Sub UpdateFees(ByVal usrShip As Long, ByVal usrFareBasis As Long)
'On Error GoTo FailSafe_Error

Dim Rst As New ADODB.Recordset
Dim ask As Integer
SQL = "SELECT * FROM tbl_Fees WHERE [AirlineID]= " & FindAirline(Me.Combo1) & " AND [FareBasisID]=" & FindFareBasis(Me.Combo2)
With Rst
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .Fields("Service Fee").Value = CDbl(Me.txtServiceFee)
                .Fields("Refund Fee").Value = CDbl(Me.txtRefund)
                .Fields("Void Fee").Value = CDbl(Me.txtVoidFee)
                .Fields("No Show Fee").Value = CDbl(Me.txtNoShowFee)
                
                .Fields("Service Fee evat").Value = CDbl(Me.txtEvatSF)
                .Fields("Refund Fee evat").Value = CDbl(Me.txtEvatRF)
                .Fields("Void Fee evat").Value = CDbl(Me.txtEvatVF)
                .Fields("No Show Fee evat").Value = CDbl(Me.txtEvatNSF)
                    
                .Fields("BaseSF").Value = CDbl(Me.txtBaseSF)
                .Fields("BaseRF").Value = CDbl(Me.txtBaseRf)
                .Fields("BaseVF").Value = CDbl(Me.txtBaseVF)
                .Fields("BaseNSF").Value = CDbl(Me.txtBaseNSF)
                    
                .Update
            MsgBox "Fee(s) for " & Me.Combo1 & " successfully updated!!!", vbInformation
            Else
               .AddNew
                .Fields(1).Value = FindAirline(Me.Combo1)
                .Fields(2).Value = FindFareBasis(Me.Combo2)
                .Fields(3).Value = Me.Combo2
                
                .Fields("Service Fee").Value = CDbl(Me.txtServiceFee)
                .Fields("Refund Fee").Value = CDbl(Me.txtRefund)
                .Fields("Void Fee").Value = CDbl(Me.txtVoidFee)
                .Fields("No Show Fee").Value = CDbl(Me.txtNoShowFee)
                
                .Fields("Service Fee evat").Value = CDbl(Me.txtEvatSF)
                .Fields("Refund Fee evat").Value = CDbl(Me.txtEvatRF)
                .Fields("Void Fee evat").Value = CDbl(Me.txtEvatVF)
                .Fields("No Show Fee evat").Value = CDbl(Me.txtEvatNSF)
                
                
                .Fields("BaseSF").Value = CDbl(Me.txtBaseSF)
                .Fields("BaseRF").Value = CDbl(Me.txtBaseRf)
                .Fields("BaseVF").Value = CDbl(Me.txtBaseVF)
                .Fields("BaseNSF").Value = CDbl(Me.txtBaseNSF)
                
                
               .Update
            
            End If
End With

            
Exit Sub
FailSafe_Error:
cn.RollbackTrans
MsgBox "Error updating for " & Me.Combo1 & " please try again...", vbInformation
End Sub


Sub LoadValues()
Dim Rst As New ADODB.Recordset
Dim ask As Integer
SQL = "SELECT * FROM tbl_Fees WHERE [AirlineID]= " & FindAirline(Me.Combo1) & " AND [FareBasisID]=" & FindFareBasis(Me.Combo2)
With Rst
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                Me.txtServiceFee = Format(.Fields("Service Fee").Value, "###,##0.00")
                Me.txtRefund = Format(.Fields("Refund Fee").Value, "###,##0.00")
                Me.txtVoidFee = Format(.Fields("Void Fee").Value, "###,##0.00")
                Me.txtNoShowFee = Format(.Fields("No Show Fee").Value, "###,##0.00")
                
                Me.txtEvatSF = Format(.Fields("Service Fee evat").Value, "###,##0.00")
                Me.txtEvatRF = Format(.Fields("Refund Fee evat").Value, "###,##0.00")
                Me.txtEvatVF = Format(.Fields("Void Fee evat").Value, "###,##0.00")
                Me.txtEvatNSF = Format(.Fields("No Show Fee evat").Value, "###,##0.00")
                
                Me.txtBaseSF = Format(.Fields("BaseSF").Value, "###,##0.00")
                Me.txtBaseRf = Format(.Fields("BaseRF").Value, "###,##0.00")
                Me.txtBaseVF = Format(.Fields("BaseVF").Value, "###,##0.00")
                Me.txtBaseNSF = Format(.Fields("BaseNSF").Value, "###,##0.00")
                
             Else
                Me.txtServiceFee = "0.00"
                Me.txtRefund = "0.00"
                Me.txtVoidFee = "0.00"
                Me.txtNoShowFee = "0.00"
                
                Me.txtEvatSF = "0.00"
                Me.txtEvatRF = "0.00"
                Me.txtEvatVF = "0.00"
                Me.txtEvatNSF = "0.00"
             
                Me.txtBaseSF = "0.00"
                Me.txtBaseRf = "0.00"
                Me.txtBaseVF = "0.00"
                Me.txtBaseNSF = "0.00"
             
             
            End If
End With
End Sub

Sub UpdateRoutePrice(ByVal usrShip As Long, ByVal usrFareBasis As Long)
'On Error GoTo FailSafe_Error
Dim Rs As New ADODB.Recordset

SQL = "SELECT * FROM tbl_RoutePricing WHERE [AirlineID]= " & FindAirline(Me.Combo1) & " AND [FareBasisID]=" & FindFareBasis(Me.Combo2)

With Rs
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
        cn.BeginTrans
            
            .MoveFirst
            Do While Not .EOF
            
                .Fields("Service Fee").Value = CDbl(Me.txtServiceFee)
                .Fields("Refund Fee").Value = CDbl(Me.txtRefund)
                .Fields("Void Fee").Value = CDbl(Me.txtVoidFee)
                .Fields("Noshow Fee").Value = CDbl(Me.txtNoShowFee)
                
            .Update
            .MoveNext
            Loop
        cn.CommitTrans
        End If
        
If Me.Tag = "-drp-" Then
    Call frmDisplayRoutePricing.Refresh_Grid
End If
        
End With

Exit Sub
FailSafe_Error:
cn.RollbackTrans
MsgBox "Error updating for " & Me.Combo1 & " please try again...", vbInformation
End Sub

Sub CheckStat()

Dim Rst As New ADODB.Recordset
Dim ask As Integer
SQL = "SELECT * FROM tbl_RoutePricing WHERE [AirlineID]= " & FindAirline(Me.Combo1) & " AND [FareBasisID]=" & FindFareBasis(Me.Combo2)
With Rst
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
               ask = MsgBox("You are about to update " & .RecordCount & " record(s) continue?", vbInformation + vbYesNo)
               If ask = vbYes Then
               Call UpdateFees(FindAirline(Me.Combo1), FindFareBasis(Me.Combo2))
               Call UpdateRoutePrice(FindAirline(Me.Combo1), FindFareBasis(Me.Combo2))
               MsgBox "Fees updated", vbInformation
               End If
            Else
                MsgBox "There are no records that match your selection", vbInformation
            End If
End With
End Sub

Sub FillCombo()
Dim Tmp As New ADODB.Recordset
SQL = "SELECT * FROM tbl_Airline ORDER by [AirlineName] ASC"
With Tmp
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

Sub FillComboFareBasis()
Dim Tmp As New ADODB.Recordset
SQL = "SELECT * FROM tbl_FareBasis ORDER by [FareBasis] ASC"
With Tmp
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
           If .RecordCount > 0 Then
                Me.Combo2.Clear
                Do While Not .EOF
                    Me.Combo2.AddItem .Fields(1).Value
                .MoveNext
                Loop
           End If
End With
End Sub

Function FindAirline(param) As Long
Dim Tmp As ADODB.Recordset
Set Tmp = New ADODB.Recordset
SQL = "SELECT * FROM tbl_Airline WHERE [AirlineName]='" & param & "'"
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

Function FindFareBasis(param) As Long
Dim Tmp As ADODB.Recordset
Set Tmp = New ADODB.Recordset
SQL = "SELECT * FROM tbl_FareBasis WHERE [FareBasis]='" & param & "'"
With Tmp
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
    If .RecordCount > 0 Then
        FindFareBasis = .Fields(0).Value
      Else
        FindFareBasis = -1
    End If
    .Close
End With
Set Tmp = Nothing
End Function




Private Sub txtBaseNSF_Change()
Recalc
End Sub

Private Sub txtBaseNSF_GotFocus()
Call kulotHL(Me.txtBaseNSF)
End Sub

Private Sub txtBaseNSF_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtBaseRf_Change()
Recalc
End Sub

Private Sub txtBaseRf_GotFocus()
Call kulotHL(Me.txtBaseRf)
End Sub


Private Sub txtBaseRf_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtBaseSF_Change()
Recalc
End Sub

Private Sub txtBaseSF_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtBaseVF_Change()
Recalc
End Sub

Private Sub txtBaseVF_GotFocus()
Call kulotHL(Me.txtBaseVF)
End Sub


Private Sub txtBaseVF_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtEvatNSF_Change()
Recalc
End Sub

Private Sub txtEvatNSF_GotFocus()
Call kulotHL(Me.txtEvatNSF)
End Sub

Private Sub txtEvatNSF_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtEvatRF_Change()
Recalc
End Sub

Private Sub txtEvatRF_GotFocus()
Call kulotHL(Me.txtEvatRF)
End Sub

Private Sub txtEvatRF_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtEvatSF_Change()
Recalc
End Sub

Private Sub txtEvatSF_GotFocus()
Call kulotHL(Me.txtEvatSF)
End Sub

Private Sub txtEvatSF_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtEvatVF_Change()
Recalc
End Sub

Private Sub txtEvatVF_GotFocus()
Call kulotHL(Me.txtEvatVF)
End Sub

Private Sub txtEvatVF_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtNoShowFee_Change()
If Not IsNumeric(Me.txtNoShowFee) Then
            Me.txtNoShowFee = "0.00"
            Call kulotHL(Me.txtNoShowFee)
End If
End Sub

Private Sub txtRefund_Change()
If Not IsNumeric(Me.txtRefund) Then
            Me.txtRefund = "0.00"
            Call kulotHL(Me.txtRefund)
End If
End Sub

Private Sub txtServiceFee_Change()
If Not IsNumeric(Me.txtServiceFee) Then
            Me.txtServiceFee = "0.00"
            Call kulotHL(Me.txtServiceFee)
End If
End Sub

Private Sub txtVoidFee_Change()
If Not IsNumeric(Me.txtVoidFee) Then
            Me.txtVoidFee = "0.00"
            Call kulotHL(Me.txtVoidFee)
End If
End Sub


Sub Recalc()
Dim myEvat(1 To 4) As Double

myEvat(1) = 0
myEvat(2) = 0
myEvat(3) = 0
myEvat(4) = 0

myEvat(1) = CDbl(Me.txtBaseSF) * (CDbl(Me.txtEvatSF) / 100)
myEvat(2) = CDbl(Me.txtBaseRf) * (CDbl(Me.txtEvatRF) / 100)
myEvat(3) = CDbl(Me.txtBaseVF) * (CDbl(Me.txtEvatVF) / 100)
myEvat(4) = CDbl(Me.txtBaseNSF) * (CDbl(Me.txtEvatNSF) / 100)

Me.txtServiceFee = CDbl(Me.txtBaseSF) + myEvat(1)
Me.txtRefund = CDbl(Me.txtBaseRf) + myEvat(2)
Me.txtVoidFee = CDbl(Me.txtBaseVF) + myEvat(3)
Me.txtNoShowFee = CDbl(Me.txtBaseNSF) + myEvat(4)


End Sub
