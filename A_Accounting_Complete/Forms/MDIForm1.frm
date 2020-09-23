VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0660518B-D86B-4046-BC56-427D4BD8270D}#1.0#0"; "kulotMenu.ocx"
Begin VB.MDIForm MDImain 
   BackColor       =   &H8000000C&
   Caption         =   "Els Travel and Tours"
   ClientHeight    =   6270
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11325
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin KulotMenu.imgKulotMenu imgKulotMenu1 
      Left            =   1005
      Top             =   2910
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   90
      Top             =   375
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   36
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":000C
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0CE6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1B38
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":298A
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3264
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3B3E
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4418
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4DE2
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":56BC
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":59D6
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":62B0
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":6B8A
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":7464
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":777E
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":8058
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":8932
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":920C
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":9AE6
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A3C0
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":AC9A
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B574
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":BE4E
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":C728
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":D002
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":D8DC
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":E1B6
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":EA90
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":F36A
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":FC44
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1051E
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":10DD4
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":116AE
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":11B00
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":11F52
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":14704
            Key             =   "IMG35"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":15986
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   5910
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList itb32x32 
      Left            =   750
      Top             =   375
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":18138
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":19ACA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1B45C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1CDEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1CE4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1CEAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1CF08
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1CF66
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1CFC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1D022
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1D080
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1D0DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1D13C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1D19A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1D1F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1D256
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1D2B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFiles 
      Caption         =   "File"
      Begin VB.Menu mnuf1 
         Caption         =   "Ship/Airline"
      End
      Begin VB.Menu mnudiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRoutes 
         Caption         =   "Add Routes"
      End
      Begin VB.Menu mnuf2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuf3 
         Caption         =   "Add Ticket Type"
      End
      Begin VB.Menu mnuf4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuF5 
         Caption         =   "Add Tickets"
      End
      Begin VB.Menu mnuf6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuf7 
         Caption         =   "Add Passenger Type"
      End
      Begin VB.Menu mnuf8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuf9 
         Caption         =   "Set Ticket Routes and Pricing"
      End
      Begin VB.Menu mnuf10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBankSet 
         Caption         =   "Bank Account Settings"
      End
      Begin VB.Menu mnuaddcheck1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddChecks 
         Caption         =   "Add Cheques"
      End
      Begin VB.Menu mnuCust1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCust 
         Caption         =   "Customer"
      End
   End
   Begin VB.Menu mnuTrans 
      Caption         =   "Transaction"
      Begin VB.Menu mnuState 
         Caption         =   "Statement of Accounts"
         Begin VB.Menu mnuDomestic 
            Caption         =   "Domestic"
         End
         Begin VB.Menu mnuInternational 
            Caption         =   "International"
         End
         Begin VB.Menu mnuExchangeDoc 
            Caption         =   "Exchange Document"
         End
         Begin VB.Menu mnuvoid 
            Caption         =   "Void Ticket"
         End
      End
      Begin VB.Menu mnuPassporting 
         Caption         =   "Passporting"
      End
      Begin VB.Menu mnuTrans2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrans3 
         Caption         =   "Cashier"
         Begin VB.Menu mnuCashierPayments 
            Caption         =   "Payments"
         End
         Begin VB.Menu mnuCashierVoucher 
            Caption         =   "Voucher"
         End
      End
      Begin VB.Menu mnuTrans4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrans6 
         Caption         =   "Ticket Refund"
      End
      Begin VB.Menu mnuTrans5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrans7 
         Caption         =   "Purchase Order"
         Begin VB.Menu mnuPDomesticPO 
            Caption         =   "PO Domestic"
         End
         Begin VB.Menu mnuInternationalPO 
            Caption         =   "PO International"
         End
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Reports"
      Begin VB.Menu mnuSales1 
         Caption         =   "ELS Sales Report Detailed"
      End
      Begin VB.Menu mnuSales2 
         Caption         =   "ELS Sales Report"
      End
      Begin VB.Menu r3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBankDeposit 
         Caption         =   "Bank Deposit Report"
      End
      Begin VB.Menu r5 
         Caption         =   "Bank Deposit HSBC Credit Card"
      End
      Begin VB.Menu R66 
         Caption         =   "Bank Deposit VISA & MASTERCARD"
      End
      Begin VB.Menu R777 
         Caption         =   "Bank Deposit DINERS CARD"
      End
      Begin VB.Menu R7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSoldUnSold 
         Caption         =   "Sold / Unsold Tickets Report"
      End
      Begin VB.Menu R8 
         Caption         =   "Company Sales Report"
      End
      Begin VB.Menu R87 
         Caption         =   "-"
      End
      Begin VB.Menu mnucamrpt 
         Caption         =   "Cancelled Ticket"
      End
      Begin VB.Menu mnuStatement 
         Caption         =   "Statement"
      End
      Begin VB.Menu mnuProfit 
         Caption         =   "Profit Report"
      End
      Begin VB.Menu mnuFinancial 
         Caption         =   "Financial Report"
      End
      Begin VB.Menu mnuRefund 
         Caption         =   "Refund Report"
      End
      Begin VB.Menu mnuRoutePricing 
         Caption         =   "Route Pricing"
      End
      Begin VB.Menu mnuPurchaseORder 
         Caption         =   "Purchase Order Report"
         Begin VB.Menu mnuPO_local 
            Caption         =   "Domestic PO"
         End
         Begin VB.Menu mnuPO_Intl 
            Caption         =   "International PO"
         End
      End
      Begin VB.Menu mnuVoucherReport 
         Caption         =   "Voucher Report"
      End
      Begin VB.Menu mnuRptAcc 
         Caption         =   "Accounting Reports"
         Begin VB.Menu mnuAcc1 
            Caption         =   "Statement of Account"
            Begin VB.Menu mnuRptDomestic 
               Caption         =   "Domestic"
            End
            Begin VB.Menu mnuRptIntern 
               Caption         =   "International"
            End
            Begin VB.Menu mnuRptDocument 
               Caption         =   "Documentation"
            End
         End
         Begin VB.Menu mnuAcc2 
            Caption         =   "Accounts Ledger"
         End
         Begin VB.Menu mnuAcc3 
            Caption         =   "Daily Statement"
         End
         Begin VB.Menu mnuAcc4 
            Caption         =   "Aging "
         End
         Begin VB.Menu mnuAcc5 
            Caption         =   "List of Accounts"
         End
      End
   End
   Begin VB.Menu mnuUtilities 
      Caption         =   "Utilities"
      Begin VB.Menu mnuBackup 
         Caption         =   "Back-Up"
      End
      Begin VB.Menu mnuUtils1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuUitls2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSysteminfo 
         Caption         =   "System Information"
      End
      Begin VB.Menu mnuUitls3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCalc 
         Caption         =   "Calculator"
      End
      Begin VB.Menu mnuUtils4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEXPLORER 
         Caption         =   "Windows Explorer"
      End
      Begin VB.Menu mnuUtils5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKeyboard 
         Caption         =   "On Screen Keyboard"
      End
   End
   Begin VB.Menu mnuSystem 
      Caption         =   "System"
      Begin VB.Menu mnuUSer 
         Caption         =   "Add / Set User Accounts"
      End
      Begin VB.Menu mnuSetBranch 
         Caption         =   "Set Branch"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Window"
      Begin VB.Menu mnuCascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu mnuTileHor 
         Caption         =   "Tile Horizontally"
      End
      Begin VB.Menu mnuTileVert 
         Caption         =   "Tile Vertically"
      End
      Begin VB.Menu mnuArrangeICO 
         Caption         =   "Arrange Icons"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuContents 
         Caption         =   "Contents"
      End
      Begin VB.Menu mnuIndex 
         Caption         =   "Index"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "Search"
      End
      Begin VB.Menu mnuTech 
         Caption         =   "Technical Support"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "MDImain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Load()
Me.StatusBar1.Panels(1).Text = "Date :" & Format(Now, "mm/dd/yyyy")
Me.Caption = AppCaption
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'If RsBackUpSettings!BackUpOnClose = True Then
'    Frm_AutoBackUp.Show 1
'End If
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show 1
End Sub

Private Sub mnuAcc2_Click()
frmReportRange.lblReportName = "AR"
frmReportRange.Show 1

End Sub

Private Sub mnuAcc3_Click()
frmReportRange.lblReportName = "LIST OF SA"
frmReportRange.Show 1
End Sub

Private Sub mnuAcc4_Click()
'RptAging.Show 1
frmReportRange.lblReportName = "AGING"
frmReportRange.Show 1

End Sub

Private Sub mnuAcc5_Click()
SQL = "select * from qryListofAccounts"
Dim Rpt As New RptListOfAccounts
With Rpt
            .DataControl1.Connection = cn
            .DataControl1.Source = SQL
            .Show 1
End With
End Sub

Private Sub mnuAddChecks_Click()
frmAddChecks.Show 1
End Sub

Private Sub MNUAR_Click()
'Dim Rpt As New RptAR
'SQL = "SELECT * FROM qryAR"
'With Rpt
'            .DataControl1.Connection = cn
'            .DataControl1.Source = SQL
'            .Show 1
'End With

End Sub

Private Sub mnuBackup_Click()
frmBackUpData.Show 1
End Sub

Private Sub mnuBankDeposit_Click()
With frmReportSel
        .lblReportName.Caption = "Bank Deposits Report"
        .Show
End With
End Sub

Private Sub mnuBankSet_Click()
frmBankAccSettings.Show 1
'frmPassbook.Show 1
End Sub

Private Sub mnuCalc_Click()
On Error Resume Next
Shell ("calc"), vbMinimizedFocus
Exit Sub
End Sub

Private Sub mnucamrpt_Click()
With frmReportSel
        .lblReportName.Caption = "Cancelled Report"
        '.Show 1
End With
End Sub

Private Sub mnuCashierPayments_Click()
frmCashier.Show 1
End Sub

Private Sub mnuCashierVoucher_Click()
'frmCashVoucher.Show 1
FrmVoucherSelect.Show 1
End Sub

Private Sub mnuCust_Click()
frmCustomerAccounts.Show 1
End Sub

Private Sub mnuDomestic_Click()
frmStatement.Show
End Sub

Private Sub mnuExchangeDoc_Click()
frmExchangeDoc.Show 1
End Sub

Private Sub mnuEXPLORER_Click()
On Error Resume Next
Shell ("explorer"), vbMaximizedFocus
Exit Sub

End Sub

Private Sub mnuf1_Click()
frmShipAirline.Show 1
End Sub

Private Sub mnuf3_Click()
frmAddTicketType.Show 1
End Sub

Private Sub mnuF5_Click()
'frmAddPurchaseOrder.Show 1
frmAddTickets.Show
Me.Arrange 1
End Sub

Private Sub mnuf7_Click()
frmAddPassengerType.Show 1
End Sub

Private Sub mnuf9_Click()
frmSetTicketPricing.Show
Me.Arrange 1
End Sub

Private Sub mnuFinancial_Click()
'RptFinancial.Show 1

frmReportRange.lblReportName = "Financial"
frmReportRange.Show 1
End Sub

Private Sub mnuInternational_Click()
frmStatement_INTL.Show 1
End Sub

Private Sub mnuInternationalPO_Click()
frmPO_INTL.Show 1
End Sub

Private Sub mnuKeyboard_Click()
On Error Resume Next
    Shell "osk", vbNormalFocus
Exit Sub
End Sub


Private Sub mnuPassporting_Click()
frmPassporting.Show 1
End Sub

Private Sub mnuPDomesticPO_Click()
frmPO_Domestic.Show 1
End Sub

Private Sub mnuPO_Intl_Click()
frmReportRange.lblReportName = "Purchase Order International"
frmReportRange.Show 1
End Sub

Private Sub mnuPO_local_Click()
frmReportRange.lblReportName = "Purchase Order"
frmReportRange.Show 1
End Sub

Private Sub mnuProfit_Click()
With frmReportSel
        .lblReportName.Caption = "Company Profit Report"
        .Show
End With
End Sub


Private Sub mnuRefund_Click()
frmReportRange.lblReportName = "Refund"
frmReportRange.Show 1
End Sub

Private Sub mnuRestore_Click()
frmRestore.Show 1
End Sub

Private Sub mnuRoutePricing_Click()
frmReportRoute.Show 1
End Sub

Private Sub mnuRoutes_Click()
frmAddRoutes.Show 1
End Sub

Private Sub mnuRptDomestic_Click()
frmReportRange.lblReportName = "STATEMENT OF ACCOUNTS"
frmReportRange.Show 1

End Sub

Private Sub mnuSales1_Click()
With frmReportSel
        .lblReportName.Caption = "Detailed Sales Report"
        .Show
End With
End Sub

Private Sub mnuSales2_Click()
'With frmReportSel
'        .lblReportName.Caption = "Sales Report"
'        .Show 1
'End With
                    Dim Rpt As New RptStatement
                    'AskPrint = MsgBox("PLEASE INSERT PAPER AND CLICK OK TO START PRINTING...", vbOKCancel + vbExclamation)
                    'If AskPrint = vbOK Then
                    
                               With Rpt
                                    .DataControl1.Connection = cn
                                    .DataControl1.Source = "SELECT * FROM qryStatement" ' WHERE [sNumber]='" & Me.txtNo & "'"
                                    .Show 1
                                End With
                                Set Rpt = Nothing
                   ' End If


End Sub

Private Sub mnuSoldUnSold_Click()
frmReportSelSoldUnsold.Show 1
End Sub

Private Sub mnuStatement_Click()
frmSelectStatement.Tag = "view_sa"
frmSelectStatement.Show 1
End Sub

Private Sub mnuSysteminfo_Click()
FRM_SYS_INFO.Show
End Sub

Private Sub mnuTrans6_Click()
frmRefund.Show 1
End Sub

Private Sub mnuUSer_Click()
frmUserVerifySetAcc.Show 1
End Sub

Private Sub mnuvoid_Click()
frmVoidTicket.Show 1
End Sub

Private Sub mnuVoucherReport_Click()
frmReportRange.lblReportName = "Voucher"
frmReportRange.Show 1
End Sub

Private Sub r5_Click()
With frmReportSel
        .lblReportName.Caption = "Bank Deposits PAL Credit Card"
        .Show 1
End With
End Sub

Private Sub R8_Click()
With frmReportSel
        .lblReportName.Caption = "Company Sales Report"
        .Show
End With
End Sub
