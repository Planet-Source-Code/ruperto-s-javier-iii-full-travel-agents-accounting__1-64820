VERSION 5.00
Begin VB.Form FrmVoucherSelect 
   Caption         =   "Select Voucher type"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6810
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Close"
      Height          =   585
      Left            =   4935
      TabIndex        =   9
      Top             =   5325
      Width           =   1770
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Height          =   585
      Left            =   3180
      TabIndex        =   8
      Top             =   5325
      Width           =   1770
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Voucher"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5130
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.OptionButton Option7 
         Caption         =   "7 . ) OTHER(S)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   525
         TabIndex        =   7
         Top             =   3945
         Width           =   5325
      End
      Begin VB.OptionButton Option6 
         Caption         =   "6 . ) BUY/SELL DOLLAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   540
         TabIndex        =   6
         Top             =   3375
         Width           =   5325
      End
      Begin VB.OptionButton Option5 
         Caption         =   "5 . ) EXPENSES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   525
         TabIndex        =   5
         Top             =   2820
         Width           =   5325
      End
      Begin VB.OptionButton Option4 
         Caption         =   "4 . ) COMMISSION(S)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   510
         TabIndex        =   4
         Top             =   2250
         Width           =   5325
      End
      Begin VB.OptionButton Option3 
         Caption         =   "3 . ) REFUND(S)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   495
         TabIndex        =   3
         Top             =   1695
         Width           =   5325
      End
      Begin VB.OptionButton Option2 
         Caption         =   "2 . ) PAYMENT TO SHIP/AIRLINE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   495
         TabIndex        =   2
         Top             =   1125
         Width           =   5325
      End
      Begin VB.OptionButton Option1 
         Caption         =   "1 . ) FUND(S) TRANSFER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   495
         TabIndex        =   1
         Top             =   600
         Width           =   5325
      End
   End
End
Attribute VB_Name = "FrmVoucherSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
'//Fund Transfer Voucher
Me.Hide
If Me.Option1 Then
        frmFundTransfer.Show 1
End If

'//payment to airline
If Me.Option2 Then
        frmVoucherPayShipAirSel.Show 1
End If

'//Fund Transfer Voucher
If Me.Option3 Then
        frmVoucherRefund.Show 1
End If

'//Commission
If Me.Option4 Then
        frmCashVoucher.Show 1
End If

'//Expenses
If Me.Option5 Then
        frmCashVoucher.Show 1
End If

'//Buy Dollar
If Me.Option6 Then
        frmVoucherBuySellDollar.Show 1
End If

'//Others
If Me.Option7 Then
        frmCashVoucher.Show 1
End If


'frmCashVoucher.Show 1

End Sub

