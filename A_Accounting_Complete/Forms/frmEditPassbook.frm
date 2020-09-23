VERSION 5.00
Begin VB.Form frmEditPassbook 
   Caption         =   "Edit Passbook"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10380
   Icon            =   "frmEditPassbook.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   10380
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   555
      Left            =   8325
      TabIndex        =   24
      Top             =   5370
      Width           =   1980
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   555
      Left            =   75
      TabIndex        =   23
      Top             =   5385
      Width           =   1980
   End
   Begin VB.Frame Frame1 
      Caption         =   "Editing of Bank Passbook"
      Height          =   5265
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   10245
      Begin VB.PictureBox Picture4 
         Enabled         =   0   'False
         Height          =   4890
         Left            =   5490
         ScaleHeight     =   4830
         ScaleWidth      =   4575
         TabIndex        =   16
         Top             =   210
         Width           =   4635
         Begin VB.TextBox txtBalance 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   195
            TabIndex        =   22
            Text            =   "0.00"
            Top             =   3240
            Visible         =   0   'False
            Width           =   4320
         End
         Begin VB.TextBox txtDebit 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   195
            TabIndex        =   19
            Text            =   "0.00"
            Top             =   1680
            Width           =   4320
         End
         Begin VB.TextBox txtCredit 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   180
            TabIndex        =   18
            Text            =   "0.00"
            Top             =   570
            Width           =   4320
         End
         Begin VB.Label Label9 
            Caption         =   "CURRENT BALANCE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   195
            TabIndex        =   21
            Top             =   2850
            Width           =   3825
         End
         Begin VB.Label Label8 
            Caption         =   "DEBIT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   195
            TabIndex        =   20
            Top             =   1290
            Width           =   2775
         End
         Begin VB.Label Label7 
            Caption         =   "CREDIT "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   180
            TabIndex        =   17
            Top             =   210
            Width           =   2775
         End
      End
      Begin VB.PictureBox Picture3 
         Height          =   1995
         Left            =   105
         ScaleHeight     =   1935
         ScaleWidth      =   5265
         TabIndex        =   13
         Top             =   3105
         Width           =   5325
         Begin VB.TextBox txtDesc 
            Height          =   1395
            Left            =   90
            MultiLine       =   -1  'True
            TabIndex        =   15
            Text            =   "frmEditPassbook.frx":1272
            Top             =   375
            Width           =   5070
         End
         Begin VB.Label Label6 
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   105
            TabIndex        =   14
            Top             =   75
            Width           =   1230
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   2790
         Left            =   2835
         ScaleHeight     =   2730
         ScaleWidth      =   2535
         TabIndex        =   6
         Top             =   240
         Width           =   2595
         Begin VB.TextBox txtVoucherNo 
            Height          =   435
            Left            =   90
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   2175
            Width           =   2370
         End
         Begin VB.TextBox txtCheckNo 
            Height          =   435
            Left            =   90
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   435
            Width           =   2370
         End
         Begin VB.TextBox txtCheckDate 
            Height          =   435
            Left            =   90
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   1290
            Width           =   2370
         End
         Begin VB.Label Label5 
            Caption         =   "Voucher #"
            Height          =   330
            Left            =   105
            TabIndex        =   12
            Top             =   1950
            Width           =   1635
         End
         Begin VB.Label Label4 
            Caption         =   "Check #"
            Height          =   330
            Left            =   90
            TabIndex        =   10
            Top             =   135
            Width           =   1635
         End
         Begin VB.Label Label3 
            Caption         =   "Check Date"
            Height          =   330
            Left            =   90
            TabIndex        =   9
            Top             =   1050
            Width           =   1635
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   2760
         Left            =   135
         ScaleHeight     =   2700
         ScaleWidth      =   2535
         TabIndex        =   1
         Top             =   255
         Width           =   2595
         Begin VB.TextBox txtTransCode 
            Enabled         =   0   'False
            Height          =   435
            Left            =   90
            TabIndex        =   25
            Text            =   "Text1"
            Top             =   420
            Width           =   2370
         End
         Begin VB.TextBox txtSAno 
            Height          =   435
            Left            =   90
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   2160
            Width           =   2370
         End
         Begin VB.TextBox txtDepDate 
            Height          =   435
            Left            =   90
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   1275
            Width           =   2370
         End
         Begin VB.Label Label10 
            Caption         =   "Transaction Code"
            Height          =   270
            Left            =   105
            TabIndex        =   26
            Top             =   75
            Width           =   1590
         End
         Begin VB.Label Label2 
            Caption         =   "Statement #"
            Height          =   330
            Left            =   90
            TabIndex        =   5
            Top             =   1920
            Width           =   1635
         End
         Begin VB.Label Label1 
            Caption         =   "Deposit Date"
            Height          =   330
            Left            =   90
            TabIndex        =   3
            Top             =   1035
            Width           =   1635
         End
      End
   End
End
Attribute VB_Name = "frmEditPassbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SQL As String

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Dim ask As Integer

ask = MsgBox("Sure you want to update changes?", vbInformation + vbYesNo)
If ask = vbYes Then
            If CommitChanges Then
                        Call frmPassbook.AccPassbook(frmPassbook.Combo1)
                        MsgBox "Changes successfull", vbInformation
                  Else
                        MsgBox "Changes not successfull", vbCritical
            End If
    Else
            MsgBox "Changes cancelled"
End If

End Sub

Private Sub Form_Load()
Call BindCtl
End Sub


Sub BindCtl()
With frmPassbook.DataGrid1
        Me.txtTransCode = .Columns(0).Text
        Me.txtDepDate = .Columns(2).Text
        Me.txtSAno = .Columns(3).Text
        Me.txtCheckNo = .Columns(4).Text
        Me.txtCheckDate = .Columns(5).Text
        Me.txtVoucherNo = .Columns(6).Text
        Me.txtDesc = .Columns(7).Text
        Me.txtCredit = .Columns(12).Text
        Me.txtDebit = .Columns(13).Text
        Me.txtBalance = .Columns(14).Text
End With
End Sub

Function CommitChanges() As Boolean
'On Error GoTo FailSafe_Error
Dim Rst As New ADODB.Recordset
SQL = "SELECT * FROM tbl_BankPassbook WHERE [TransID]=" & CLng(Me.txtTransCode)

cn.BeginTrans
With Rst
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
      If .RecordCount > 0 Then
             .Fields("Deposit Date").Value = Me.txtDepDate
             .Fields("SA no").Value = Me.txtSAno
             .Fields("Check No").Value = Me.txtCheckNo
             .Fields("Check Date").Value = Me.txtCheckDate
             .Fields("Voucher No").Value = Me.txtVoucherNo
             .Fields("Description").Value = Me.txtDesc
             '.Fields("Credit").Value = Me.txtCredit
             '.Fields("Debit").Value = Me.txtDebit
             '.Fields("Balance").Value = Me.txtBalance
             .Update

             CommitChanges = True
      Else
             CommitChanges = False
      End If
End With

If CommitChanges Then
          cn.CommitTrans
End If

Exit Function

FailSafe_Error:
    cn.RollbackTrans
    CommitChanges = False
End Function
