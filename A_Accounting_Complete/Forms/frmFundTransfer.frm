VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFundTransfer 
   Caption         =   "Fund Transfer (Electronic)"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8190
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTotAmount 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   5655
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "0.00"
      Top             =   7245
      Width           =   2430
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
      Left            =   5580
      TabIndex        =   16
      Top             =   75
      Width           =   2550
   End
   Begin VB.CommandButton cmdTransfer 
      Caption         =   "Transfer Now!"
      Height          =   510
      Left            =   75
      TabIndex        =   4
      Top             =   7995
      Width           =   1920
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Voucher Details"
      ForeColor       =   &H80000008&
      Height          =   4185
      Left            =   90
      TabIndex        =   10
      Top             =   3015
      Width           =   8010
      Begin VB.CommandButton cmdInsert 
         Caption         =   "Insert"
         Height          =   420
         Left            =   6600
         TabIndex        =   3
         Top             =   330
         Width           =   1305
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3120
         Left            =   90
         TabIndex        =   18
         Top             =   960
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   5503
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "VoucherID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "From Account"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "To Account"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
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
         Left            =   2880
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   375
         Width           =   3705
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Amount to be Transfered :"
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
         TabIndex        =   15
         Top             =   435
         Width           =   2805
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Destination"
      ForeColor       =   &H80000008&
      Height          =   2310
      Left            =   4125
      TabIndex        =   8
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
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   1545
         Width           =   3705
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   1
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
         TabIndex        =   12
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
         TabIndex        =   9
         Top             =   330
         Width           =   3390
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Source"
      ForeColor       =   &H80000008&
      Height          =   2310
      Left            =   105
      TabIndex        =   6
      Top             =   630
      Width           =   3975
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
         TabIndex        =   13
         Text            =   "0.00"
         Top             =   1545
         Width           =   3705
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   0
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
         TabIndex        =   11
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
         TabIndex        =   7
         Top             =   330
         Width           =   3390
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   510
      Left            =   6165
      TabIndex        =   5
      Top             =   7995
      Width           =   1920
   End
   Begin VB.Label Label6 
      Caption         =   "Total Amount :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   20
      Top             =   7245
      Width           =   1410
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
      Left            =   3870
      TabIndex        =   17
      Top             =   90
      Width           =   1515
   End
End
Attribute VB_Name = "frmFundTransfer"
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

Private Sub cmdInsert_Click()
Dim myList As ListItem


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

If CDbl(Me.txtAmount) <= 0 Then
    MsgBox "Amount should not be less than or equal to zero", vbInformation
    Exit Sub
End If


            Set myList = ListView1.ListItems.Add(, , "")
            myList.SubItems(1) = Me.Combo1
            myList.SubItems(2) = Me.Combo2
            myList.SubItems(3) = Format(Me.txtAmount, "###,##0.00")
            
            Me.txtTotAmount = Format(SumListView, "###,##0.00")

End Sub

Private Sub cmdTransfer_Click()
On Error GoTo ErrExit

Dim ask                     As Integer
Dim y                       As Integer
Dim myTempAmount            As Double
Dim myTempAccount           As String
Dim myTempFrom_Account      As String
Dim myTempTo_Account        As String

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

If CDbl(Me.txtAmount) <= 0 Then
MsgBox "Amount to be transfered should not be less than or equal to zero", vbInformation
Call kulotHL(Me.txtAmount)
Exit Sub
End If


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
            .Fields(6).Value = Format(Me.txtTotAmount, "###,##0.00")
            .Fields(7).Value = False
        .Update
        Me.Tag = .Fields(0).Value
        Me.txtVoucherID = Me.Tag
    End With
    
    
If Me.ListView1.ListItems.Count > 0 Then

 For y = 1 To Me.ListView1.ListItems.Count
    '// Now debut this voucher to bank
    
    myTempAmount = CDbl(ListView1.ListItems.Item(y).SubItems(3))
    
    myTempFrom_Account = ListView1.ListItems.Item(y).SubItems(1)
    myTempTo_Account = ListView1.ListItems.Item(y).SubItems(2)
    
    Call UpdatePassbook(myTempAmount, _
    "N/A", _
    "N/A", _
    "Fund Transfer TO acc#[" & myTempTo_Account & "]", _
    , "n/a", _
    0, 0, _
    0, CDbl(Me.txtAmount), "", "", "", "", "", "", "", "debit", myTempFrom_Account, myTempTo_Account)
    
    '// Now add this voucher to bank
    Call UpdatePassbook(myTempAmount, _
    "N/A", _
    "N/A", _
    "Fund Transfer FROM acc#[" & myTempFrom_Account & "]", _
    , "n/a", _
    0, 0, _
    0, myTempAmount, "", "", "", "", "", "", "", "credit", myTempFrom_Account, myTempTo_Account)
    
    MsgBox myTempAmount & " successfully transfered from acc#" & ListView1.ListItems.Item(y).SubItems(1) & " to acc# " & ListView1.ListItems.Item(y).SubItems(2), vbInformation
 Next
    Call UpDateVoucherDetails
End If
    
cn.CommitTrans
Exit Sub
ErrExit:
cn.RollbackTrans
MsgBox Err.Description, vbCritical, "Contact system developer"
End Sub


Sub UpDateVoucherDetails()
Dim y               As Integer
Dim RstInsert       As New ADODB.Recordset
Dim mySQL           As String

mySQL = "SELECT * FROM tbl_VoucherDetails"

If Me.ListView1.ListItems.Count > 0 Then

        RstInsert.Open mySQL, cn, adOpenKeyset, adLockOptimistic

        For y = 1 To Me.ListView1.ListItems.Count
        '//if voucher details does not exist  add
        
            If Not IsNumeric(Me.ListView1.ListItems(y).Text) Then
                RstInsert.AddNew
                        RstInsert.Fields("VoucherID").Value = Me.txtVoucherID
                        RstInsert.Fields("Particulars").Value = "Fund Transfer from acc# :" & ListView1.ListItems.Item(y).SubItems(1) & " to acc# :" & ListView1.ListItems.Item(y).SubItems(2)
                        RstInsert.Fields("Amount").Value = CDbl(ListView1.ListItems.Item(y).SubItems(3))
                RstInsert.Update
            End If
            
        Next y
End If
End Sub

Sub UpdatePassbook(ByVal nAmt As Double, _
    ByVal CheckNo As String, ByVal CheckDate As String, _
    Optional Desc As String, Optional ByVal AccNo As String, _
    Optional strAir As String, _
    Optional nCash, Optional nCard, Optional nCheck, _
    Optional nOthers, Optional nCardName, _
    Optional nCardNumber, Optional nCardHolder, _
    Optional nBank1, Optional nBank2, _
    Optional nBank3, Optional nBank4, Optional usrCriteria, Optional usrCredit_Acc, Optional usrDebit_Acc)
    
'On Error Resume Next
Dim RsPassbk As New ADODB.Recordset
Dim RsBalance As New ADODB.Recordset
Dim TempBal As Double
Dim Tmp As Double
Dim myTempAccno As String


If usrCriteria = "credit" Then
    myTempAccno = usrDebit_Acc
Else
    myTempAccno = usrCredit_Acc
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
              UCase(myTempAccno) & "'"
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
              UCase(myTempAccno) & "'"
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

Function SumListView() As Double
Dim y               As Integer
Dim Tmp             As Double

If Me.ListView1.ListItems.Count > 0 Then
        For y = 1 To Me.ListView1.ListItems.Count
            If Me.ListView1.ListItems(y).SubItems(3) <> "" Then
                Tmp = Tmp + CDbl(Me.ListView1.ListItems(y).SubItems(3))
            End If
        Next y
Else
        Tmp = 0
End If
        SumListView = Tmp
End Function

Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub
