VERSION 5.00
Begin VB.Form frmVoucherBuySellDollar 
   Caption         =   "Buy / Sell Dollar"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8190
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Exchange Rate"
      ForeColor       =   &H80000008&
      Height          =   2220
      Left            =   120
      TabIndex        =   26
      Top             =   5580
      Width           =   8010
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   420
         Left            =   4335
         TabIndex        =   30
         Text            =   "0.00"
         Top             =   1110
         Width           =   3360
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Height          =   420
         Left            =   4335
         TabIndex        =   29
         Text            =   "0.00"
         Top             =   690
         Width           =   3360
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   420
         Left            =   4335
         TabIndex        =   27
         Text            =   "0.00"
         Top             =   255
         Width           =   3360
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Peso Equivalent :"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1860
         TabIndex        =   32
         Top             =   1140
         Width           =   1455
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Exchange Rate :"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1860
         TabIndex        =   31
         Top             =   690
         Width           =   1455
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Dollar Amount :"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1875
         TabIndex        =   28
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Buy Sell"
      ForeColor       =   &H80000008&
      Height          =   2685
      Left            =   4125
      TabIndex        =   23
      Top             =   2805
      Width           =   3975
      Begin VB.OptionButton Option4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sell Dollar"
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   495
         TabIndex        =   25
         Top             =   990
         Width           =   2985
      End
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Buy Dollar"
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   495
         TabIndex        =   24
         Top             =   495
         Value           =   -1  'True
         Width           =   2985
      End
   End
   Begin VB.TextBox txtVoucherID 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5595
      TabIndex        =   12
      Top             =   30
      Width           =   2550
   End
   Begin VB.CommandButton cmdTransfer 
      Caption         =   "Buy / Sell"
      Height          =   510
      Left            =   105
      TabIndex        =   11
      Top             =   7980
      Width           =   1920
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Destination"
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   4125
      TabIndex        =   4
      Top             =   630
      Width           =   3975
      Begin VB.TextBox txtBal2 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   135
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   1545
         Width           =   3705
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   630
         Width           =   3720
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Current Account Balance :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   150
         TabIndex        =   8
         Top             =   1200
         Width           =   2805
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Account Number :"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   165
         TabIndex        =   6
         Top             =   330
         Width           =   3390
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Source"
      ForeColor       =   &H80000008&
      Height          =   4875
      Left            =   105
      TabIndex        =   1
      Top             =   630
      Width           =   3975
      Begin VB.Frame Frame4 
         Caption         =   "Check Details"
         Height          =   1965
         Left            =   120
         TabIndex        =   16
         Top             =   2790
         Width           =   3780
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   1290
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   1155
            Width           =   2385
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   1290
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   735
            Width           =   2385
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   1290
            TabIndex        =   17
            Text            =   "Text1"
            Top             =   300
            Width           =   2385
         End
         Begin VB.Label Label9 
            Caption         =   "CheckBank"
            Height          =   210
            Left            =   165
            TabIndex        =   22
            Top             =   1155
            Width           =   1035
         End
         Begin VB.Label Label7 
            Caption         =   "Check Amnt"
            Height          =   210
            Left            =   150
            TabIndex        =   20
            Top             =   735
            Width           =   1035
         End
         Begin VB.Label Label6 
            Caption         =   "Check #"
            Height          =   210
            Left            =   150
            TabIndex        =   18
            Top             =   300
            Width           =   1035
         End
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Check"
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   375
         TabIndex        =   15
         Top             =   2460
         Value           =   -1  'True
         Width           =   3300
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Cash "
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   360
         TabIndex        =   14
         Top             =   2115
         Width           =   3300
      End
      Begin VB.TextBox txtBal1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   150
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   1545
         Width           =   3705
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   630
         Width           =   3720
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Current Account Balance :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   135
         TabIndex        =   7
         Top             =   1215
         Width           =   2805
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Account Number :"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   165
         TabIndex        =   2
         Top             =   330
         Width           =   3390
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   510
      Left            =   6195
      TabIndex        =   0
      Top             =   7980
      Width           =   1920
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "VOUCHER # :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3885
      TabIndex        =   13
      Top             =   45
      Width           =   1515
   End
End
Attribute VB_Name = "frmVoucherBuySellDollar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RsVoucher As ADODB.Recordset
Dim RsVoucherDetails As ADODB.Recordset
Dim SQL As String

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdTransfer_Click()
On Error GoTo ErrExit

Dim ask As Integer

If CheckNull(Me.Combo1) Then
    MsgBox "[Source] Account Name should not be blank", vbCritical
    Exit Sub
End If

If CheckNull(Me.Combo2) Then
    MsgBox "[Destination] Account Name should not be blank", vbCritical
    Exit Sub
End If


If Me.Combo1 = Me.Combo2 Then
MsgBox "[Source] account should not be the same to [Destination] account", vbInformation
Exit Sub
End If

ask = MsgBox("Are you sure you want to transfer this amount?", vbInformation + vbYesNo)
If ask = vbNo Then: Exit Sub



cn.BeginTrans
    With RsVoucher
            .AddNew
            .Fields(1).Value = "FUND TRANSFER"
            .Fields(2).Value = "ELS"
            .Fields(3).Value = Format(Now, "mm/dd/yyyy")
            .Fields(4).Value = FindBankID(Me.Combo1)
            .Fields(5).Value = "N/A"
            .Fields("Cash").Value = False
            .Fields("Check").Value = False
            .Fields("Electronic Transfer").Value = True
            
            .Fields(7).Value = False
        .Update
        Me.Tag = .Fields(0).Value
        Me.txtVoucherID = Me.Tag
    End With
    
    '// Now debut this voucher to bank
'    Call UpdatePassbook(CDbl(Me.txtAmount), _
'    "N/A", _
'    "N/A", _
'    "Fund Transfer from acc#[" & Me.Combo1 & "]", _
'    Me.Combo1, "n/a", _
'    0, 0, _
'    0, CDbl(Me.txtAmount), "", "", "", "", "", "", "", "debit")
    
    '// Now add this voucher to bank
'    Call UpdatePassbook(CDbl(Me.txtAmount), _
'    "N/A", _
'    "N/A", _
'    "Fund Transfer to acc#[" & Me.Combo2 & "]", _
'    Me.Combo2, "n/a", _
'    0, 0, _
'    0, CDbl(Me.txtAmount), "", "", "", "", "", "", "", "credit")
cn.CommitTrans
'MsgBox Me.txtAmount & " successfully transfered from acc#" & Me.Combo1 & " to acc# " & Me.Combo2, vbInformation
Exit Sub
ErrExit:
cn.RollbackTrans
MsgBox Err.Description, vbCritical, "Contact system developer"
End Sub


Sub UpdatePassbook(ByVal nAmt As Double, _
    ByVal CheckNo As String, ByVal CheckDate As String, _
    Optional Desc As String, Optional ByVal AccNo As String, _
    Optional strAir As String, _
    Optional nCash, Optional nCard, Optional nCheck, _
    Optional nOthers, Optional nCardName, _
    Optional nCardNumber, Optional nCardHolder, _
    Optional nBank1, Optional nBank2, _
    Optional nBank3, Optional nBank4, Optional usrCriteria)
    
'On Error Resume Next
Dim RsPassbk As New ADODB.Recordset
Dim RsBalance As New ADODB.Recordset
Dim TempBal As Double
Dim Tmp As Double
Dim myTempAccno As String


If usrCriteria = "credit" Then
    myTempAccno = Me.Combo2
Else
    myTempAccno = Me.Combo1
End If


SQL = "SELECT * FROM tbl_AccountsSetting WHERE [Account Number]='" & myTempAccno & "'"

With RsPassbk
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                    .MoveFirst
                 TempBal = .Fields("Current Balance").Value
                Else
                 TempBal = 0
            End If
            .Close
      Set RsPassbk = Nothing
End With




SQL = "SELECT * FROM tbl_BankPassbook"
With RsPassbk
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        'If IsExist_Voucher(Me.txtVoucherID) Then
        '    .MoveFirst
        '    .Find "[Voucher No]='" & Me.txtVoucherID & "'"
        'Else
            .AddNew
        'End If
            .Fields("Deposit Date").Value = Format(Now, "mm/dd/yyyy")
            .Fields("Check No").Value = CheckNo
            .Fields("Check Date").Value = CheckDate
            .Fields("Voucher No").Value = Me.txtVoucherID
            .Fields("Description").Value = Desc
            
            If usrCriteria = "credit" Then
                .Fields("Credit").Value = nAmt
                .Fields("Debit").Value = 0
            Else
                .Fields("Credit").Value = 0
                .Fields("Debit").Value = nAmt
            
            End If
            
            .Fields("Account Number").Value = myTempAccno
            
            If usrCriteria = "credit" Then
                    .Fields("Balance").Value = TempBal + nAmt
                Else
                        If CDbl(TempBal) <= 0 Then
                            Tmp = 0
                        Else
                            Tmp = CDbl(TempBal) - nAmt
                        End If
                
                    .Fields("Balance").Value = Tmp
            End If
            
            
            .Fields("Cash Amount").Value = CDbl(nCash)
            .Fields("Card Amount").Value = CDbl(nCard)
            .Fields("Check Amount").Value = CDbl(nCheck)
            .Fields("Others Amount").Value = CDbl(nOthers)
            .Fields("ORno").Value = "n/a"
            .Fields("Airline").Value = -1   'strAir
            .Fields("Card Name").Value = nCardName
            .Fields("Card Number").Value = nCardNumber
            .Fields("Card Holder").Value = nCardHolder
            .Fields("Bank1").Value = nBank1
            .Fields("Bank2").Value = nBank2
            .Fields("Bank3").Value = nBank3
            .Fields("Bank4").Value = nBank4
        
            .Update
End With

If usrCriteria = "credit" Then
SQL = "UPDATE tbl_AccountsSetting SET [Current Balance] = " & _
              CDbl(TempBal + nAmt) _
              & " WHERE [Account Number]= '" & _
              UCase(Me.Combo2) & "'"
              cn.BeginTrans
                    cn.Execute SQL
              cn.CommitTrans

End If

If usrCriteria = "debit" Then


If CDbl(TempBal) <= 0 Then
    Tmp = 0
Else
    Tmp = CDbl(TempBal) - nAmt
End If
SQL = "UPDATE tbl_AccountsSetting SET [Current Balance] = " & _
              CDbl(Tmp) _
              & " WHERE [Account Number]= '" & _
              UCase(Me.Combo1) & "'"
              cn.BeginTrans
                    cn.Execute SQL
              cn.CommitTrans

End If


Exit Sub
FailSafe_Error:
cn.RollbackTrans
End Sub

Function IsExist_Voucher(param) As Boolean
Dim Rst         As New ADODB.Recordset
Dim SQL         As String

SQL = "SELECT * FROM qryBankPassbook WHERE [Voucher No]='" & param & "'"
With Rst
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
                    IsExist_Voucher = True
            Else
                    IsExist_Voucher = False
        End If
        .Close
      Set Rst = Nothing
End With
End Function


Function FindBankID(param) As Long
Dim Rst As New ADODB.Recordset
SQL = "SELECT * FROM tbl_AccountsSetting WHERE [Account Number]='" & param & "'"
With Rst
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
              FindBankID = .Fields(1).Value
          Else
              FindBankID = -1
        End If
     .Close
   Set Rst = Nothing
End With
End Function

Private Sub Combo1_Click()
Me.txtBal1 = DisplayBalance(Me.Combo1)
End Sub

Private Sub Combo2_Click()
Me.txtBal2 = DisplayBalance(Me.Combo2)
End Sub

Private Sub Form_Load()
Call FillAccount
Set RsVoucher = New ADODB.Recordset
Set RsVoucherDetails = New ADODB.Recordset

SQL = "SELECT * FROM tbl_Voucher"
With RsVoucher
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
End With
End Sub

Sub FillAccount()
Dim Rst As New ADODB.Recordset
SQL = "SELECT DISTINCT  [Account Number] FROM tbl_AccountsSetting ORDER by [Account Number] ASC"
With Rst
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            Me.Combo1.Clear
            Me.Combo2.Clear
                    .MoveFirst
               Do While Not .EOF
                    Me.Combo1.AddItem .Fields("Account Number").Value
                    Me.Combo2.AddItem .Fields("Account Number").Value
                    .MoveNext
               Loop
            
        End If
End With

End Sub

Function DisplayBalance(param) As String
Dim Rst As New ADODB.Recordset
Dim RsPassbk As New ADODB.Recordset

Dim SQL As String

SQL = "SELECT * FROM tbl_AccountsSetting WHERE [Account Number]='" & param & "'"
With RsPassbk
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                    .MoveFirst
                 DisplayBalance = Format(.Fields("Current Balance").Value, "###,##0.00")
            End If
            .Close
      Set RsPassbk = Nothing
End With

End Function

