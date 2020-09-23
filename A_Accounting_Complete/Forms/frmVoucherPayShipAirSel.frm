VERSION 5.00
Begin VB.Form frmVoucherPayShipAirSel 
   Caption         =   "Select Type of Voucher"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   1845
      Width           =   3930
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Voucher w/ out PO"
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   1110
      Width           =   3930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Voucher with PO"
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   375
      Width           =   3930
   End
End
Attribute VB_Name = "frmVoucherPayShipAirSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmVoucherPayShipAirPO.Show 1
End Sub

Private Sub Command2_Click()
frmVoucherPayShipAir.Show 1
End Sub

Private Sub Command3_Click()
Unload Me
End Sub
