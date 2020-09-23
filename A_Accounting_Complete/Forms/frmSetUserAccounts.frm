VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "Lvbuttons.ocx"
Begin VB.Form frmSetUserAccounts 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set User Accounts"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10740
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   10740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   9690
      Top             =   15
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
            Picture         =   "frmSetUserAccounts.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":0CDA
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":1B2C
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":297E
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":3258
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":3B32
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":440C
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":4DD6
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":56B0
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":59CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":62A4
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":6B7E
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":7458
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":7772
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":804C
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":8926
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":9200
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":9ADA
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":A3B4
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":AC8E
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":B568
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":BE42
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":C71C
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":CFF6
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":D8D0
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":E1AA
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":EA84
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":F35E
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":FC38
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":10512
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":10DC8
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":116A2
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":11AF4
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":11F46
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":146F8
            Key             =   "IMG35"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetUserAccounts.frx":1597A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Set Form / Reports to be Opened"
      ForeColor       =   &H80000008&
      Height          =   4950
      Left            =   30
      TabIndex        =   12
      Top             =   4230
      Width           =   10635
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Reports"
         ForeColor       =   &H80000008&
         Height          =   4530
         Left            =   5280
         TabIndex        =   32
         Top             =   330
         Width           =   2550
         Begin VB.CheckBox chkCancelledTickets 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Cancelled Tickets"
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   90
            TabIndex        =   39
            Top             =   2505
            Width           =   2220
         End
         Begin VB.CheckBox chkAR 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Accounts Receivables"
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   105
            TabIndex        =   38
            Top             =   2145
            Width           =   2220
         End
         Begin VB.CheckBox chkCompanySalesRpt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Company Sales Report"
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   120
            TabIndex        =   37
            Top             =   1755
            Width           =   2310
         End
         Begin VB.CheckBox chkSold_unsold 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Sold/Unsold Tickets"
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   105
            TabIndex        =   36
            Top             =   1380
            Width           =   2340
         End
         Begin VB.CheckBox chkBankDeposit 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Bank Deposit Report"
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   105
            TabIndex        =   35
            Top             =   1005
            Width           =   2280
         End
         Begin VB.CheckBox chkSalesRpt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Sales Report"
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   105
            TabIndex        =   34
            Top             =   645
            Width           =   2310
         End
         Begin VB.CheckBox chkSalesRptDetailed 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Sales Report Detailed"
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   105
            TabIndex        =   33
            Top             =   300
            Width           =   2355
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Transaction"
         ForeColor       =   &H80000008&
         Height          =   4530
         Left            =   2700
         TabIndex        =   23
         Top             =   330
         Width           =   2550
         Begin VB.CheckBox chkStatementDomestic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Statement Domestic"
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   105
            TabIndex        =   31
            Top             =   285
            Width           =   2355
         End
         Begin VB.CheckBox chkStatementInter 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Statement International"
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   105
            TabIndex        =   30
            Top             =   645
            Width           =   2310
         End
         Begin VB.CheckBox chkExchangeDoc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Exchange Document"
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   105
            TabIndex        =   29
            Top             =   1005
            Width           =   2280
         End
         Begin VB.CheckBox ChkVoidTicket 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Void Ticket"
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   105
            TabIndex        =   28
            Top             =   1380
            Width           =   1590
         End
         Begin VB.CheckBox chkPassporting 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Passporting"
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   105
            TabIndex        =   27
            Top             =   1755
            Width           =   2310
         End
         Begin VB.CheckBox chkPayments 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Payments"
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   105
            TabIndex        =   26
            Top             =   2145
            Width           =   2220
         End
         Begin VB.CheckBox chkVoucher 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Voucher"
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   105
            TabIndex        =   25
            Top             =   2505
            Width           =   2220
         End
         Begin VB.CheckBox chkTicketRefund 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Ticket Refund"
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   105
            TabIndex        =   24
            Top             =   2865
            Width           =   2220
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "File"
         ForeColor       =   &H80000008&
         Height          =   4530
         Left            =   120
         TabIndex        =   13
         Top             =   330
         Width           =   2550
         Begin VB.CheckBox chkAddCust 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Add Customer(s)"
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   105
            TabIndex        =   22
            Top             =   3240
            Width           =   2220
         End
         Begin VB.CheckBox chkAddChecks 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Add Checks"
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   105
            TabIndex        =   21
            Top             =   2880
            Width           =   2220
         End
         Begin VB.CheckBox ChkBankAccSettings 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Bank Account Settings"
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   105
            TabIndex        =   20
            Top             =   2505
            Width           =   2220
         End
         Begin VB.CheckBox chkSetRoutePrice 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Set Ticket Routes/Price"
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   105
            TabIndex        =   19
            Top             =   2145
            Width           =   2220
         End
         Begin VB.CheckBox chkPassgrType 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Add Passgr type"
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   105
            TabIndex        =   18
            Top             =   1755
            Width           =   1590
         End
         Begin VB.CheckBox chkAddTickets 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Add Tickets"
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   105
            TabIndex        =   17
            Top             =   1380
            Width           =   1590
         End
         Begin VB.CheckBox ChkAddTicketType 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Add Ticket Type"
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   105
            TabIndex        =   16
            Top             =   1005
            Width           =   1590
         End
         Begin VB.CheckBox chkaddRoutes 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Add Routes"
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   105
            TabIndex        =   15
            Top             =   645
            Width           =   1590
         End
         Begin VB.CheckBox chkAddShipAirline 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Add Ship/Airline"
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   105
            TabIndex        =   14
            Top             =   285
            Width           =   1590
         End
      End
      Begin LVbuttons.LaVolpeButton cmdAccessRights 
         Height          =   435
         Left            =   7905
         TabIndex        =   40
         Top             =   420
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   767
         BTYPE           =   3
         TX              =   "Set Access Rights"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   14215660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmSetUserAccounts.frx":15DCC
         ALIGN           =   1
         IMGLST          =   "SmallImages"
         IMGICON         =   "19"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4125
      Left            =   30
      TabIndex        =   1
      Top             =   105
      Width           =   4215
      Begin VB.TextBox txtUserName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1425
         TabIndex        =   9
         Top             =   555
         Width           =   2415
      End
      Begin VB.ComboBox cboAccessType 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmSetUserAccounts.frx":15DE8
         Left            =   1410
         List            =   "frmSetUserAccounts.frx":15DFB
         TabIndex        =   7
         Top             =   1485
         Width           =   2415
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1425
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1020
         Width           =   2415
      End
      Begin LVbuttons.LaVolpeButton cmdSet 
         Height          =   435
         Left            =   150
         TabIndex        =   3
         Top             =   3585
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   767
         BTYPE           =   3
         TX              =   "Save"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   14215660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmSetUserAccounts.frx":15E29
         ALIGN           =   1
         IMGLST          =   "SmallImages"
         IMGICON         =   "12"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton cmdExit 
         Height          =   435
         Left            =   2040
         TabIndex        =   4
         Top             =   3585
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   767
         BTYPE           =   3
         TX              =   "Exit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   14215660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmSetUserAccounts.frx":15E45
         ALIGN           =   1
         IMGLST          =   "SmallImages"
         IMGICON         =   "5"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton cmdNew 
         Height          =   435
         Left            =   150
         TabIndex        =   10
         Top             =   3105
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   767
         BTYPE           =   3
         TX              =   "New"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   14215660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmSetUserAccounts.frx":15E61
         ALIGN           =   1
         IMGLST          =   "SmallImages"
         IMGICON         =   "23"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton cmdRemove 
         Height          =   435
         Left            =   2025
         TabIndex        =   11
         Top             =   3105
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   767
         BTYPE           =   3
         TX              =   "Remove"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   14215660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmSetUserAccounts.frx":15E7D
         ALIGN           =   1
         IMGLST          =   "SmallImages"
         IMGICON         =   "36"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin VB.Image Image2 
         Height          =   1365
         Left            =   105
         Picture         =   "frmSetUserAccounts.frx":15E99
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Access Type"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   105
         TabIndex        =   8
         Top             =   1500
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   105
         TabIndex        =   6
         Top             =   1020
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   105
         TabIndex        =   5
         Top             =   540
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4080
      Left            =   4305
      TabIndex        =   0
      Top             =   150
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   7197
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "UserID"
         Caption         =   "UserID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "UserName"
         Caption         =   "UserName"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "UserPass"
         Caption         =   "UserPass"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "*"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "USerAccesLevel"
         Caption         =   "USerAccesLevel"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3525.166
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSetUserAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As ADODB.Recordset
Dim RsUser As ADODB.Recordset

Dim SQL As String


Function LoadUserAcc(param)
Dim Rst As New ADODB.Recordset
SQL = "SELECT * FROM tbl_Users WHERE [UserID]=" & param
With Rst
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
    
    If .RecordCount > 0 Then
      
      If Len(.Fields("form1").Value) > 8 Then
            chkAddShipAirline.Value = 1
          Else
            chkAddShipAirline.Value = 0
      End If
      
      If Len(.Fields("form2").Value) > 8 Then
            chkaddRoutes.Value = 1
          Else
            chkaddRoutes.Value = 0
      End If
      
      If Len(.Fields("form3").Value) > 8 Then
            ChkAddTicketType.Value = 1
          Else
          ChkAddTicketType.Value = 0
      End If
      
      If Len(.Fields("form4").Value) > 8 Then
            chkAddTickets.Value = 1
            Else
            chkAddTickets.Value = 0
      End If
      If Len(.Fields("form5").Value) > 8 Then
            chkPassgrType.Value = 1
            Else
            chkPassgrType.Value = 0
      End If
      If Len(.Fields("form6").Value) > 8 Then
            chkSetRoutePrice.Value = 1
            Else
            chkSetRoutePrice.Value = 0
      End If
      If Len(.Fields("form7").Value) > 8 Then
            ChkBankAccSettings.Value = 1
            Else
            ChkBankAccSettings.Value = 0
      End If
      If Len(.Fields("form8").Value) > 8 Then
            chkAddChecks.Value = 1
            Else
            chkAddChecks.Value = 0
      End If
      If Len(.Fields("form9").Value) > 8 Then
            chkAddCust.Value = 1
            Else
            chkAddCust.Value = 0
      End If
      If Len(.Fields("form10").Value) > 8 Then
            chkStatementDomestic.Value = 1
            Else
            chkStatementDomestic.Value = 0
      End If
      
      If Len(.Fields("form11").Value) > 8 Then
            chkStatementInter.Value = 1
            Else
            chkStatementInter.Value = 0
      End If
      
      If Len(.Fields("form12").Value) > 8 Then
            chkExchangeDoc.Value = 1
            Else
            chkExchangeDoc.Value = 0
      End If
      If Len(.Fields("form13").Value) > 8 Then
            ChkVoidTicket.Value = 1
            Else
            ChkVoidTicket.Value = 0
      End If
      If Len(.Fields("form14").Value) > 8 Then
            chkPassporting.Value = 1
            Else
            chkPassporting.Value = 0
      End If
      If Len(.Fields("form15").Value) > 8 Then
            chkPayments.Value = 1
            Else
            chkPayments.Value = 0
      End If
      If Len(.Fields("form16").Value) > 8 Then
            chkVoucher.Value = 1
            Else
            chkVoucher.Value = 0
      End If
      If Len(.Fields("form17").Value) > 8 Then
            chkTicketRefund.Value = 1
            Else
            chkTicketRefund.Value = 0
      End If
      
      
    End If
    Rst.Close
    Set Rst = Nothing
End With

End Function


Private Sub cmdAccessRights_Click()
Dim ask
ask = MsgBox("Sure you want to save?", vbInformation + vbYesNo, "ELS")
If ask = vbYes Then
SQL = "UPDATE tbl_Users SET [form1] = " & IIf(Me.chkAddShipAirline = 1, "'frmShipAirline'", "''") & " WHERE (((UserID)=" & Me.DataGrid1.Columns(0).Text & ")) "
cn.Execute SQL

SQL = "UPDATE tbl_Users SET [form2] = " & IIf(Me.chkaddRoutes = 1, "'frmAddRoutes'", "''") & " WHERE (((UserID)=" & Me.DataGrid1.Columns(0).Text & ")) "
cn.Execute SQL

SQL = "UPDATE tbl_Users SET [form3] = " & IIf(Me.ChkAddTicketType = 1, "'frmAddTicketType'", "''") & " WHERE (((UserID)=" & Me.DataGrid1.Columns(0).Text & ")) "
cn.Execute SQL

SQL = "UPDATE tbl_Users SET [form4] = " & IIf(Me.chkAddTickets = 1, "'frmAddTickets'", "''") & " WHERE (((UserID)=" & Me.DataGrid1.Columns(0).Text & ")) "
cn.Execute SQL

SQL = "UPDATE tbl_Users SET [form5] = " & IIf(Me.chkPassgrType = 1, "'frmAddPassengerType'", "''") & " WHERE (((UserID)=" & Me.DataGrid1.Columns(0).Text & ")) "
cn.Execute SQL

SQL = "UPDATE tbl_Users SET [form6] = " & IIf(Me.chkSetRoutePrice = 1, "'frmSetTicketPricing'", "''") & " WHERE (((UserID)=" & Me.DataGrid1.Columns(0).Text & ")) "
cn.Execute SQL

SQL = "UPDATE tbl_Users SET [form7] = " & IIf(Me.ChkBankAccSettings = 1, "'frmBankAccSettings'", "''") & " WHERE (((UserID)=" & Me.DataGrid1.Columns(0).Text & ")) "
cn.Execute SQL

SQL = "UPDATE tbl_Users SET [form8] = " & IIf(Me.chkAddChecks = 1, "'frmAddChecks'", "''") & " WHERE (((UserID)=" & Me.DataGrid1.Columns(0).Text & ")) "
cn.Execute SQL

SQL = "UPDATE tbl_Users SET [form9] = " & IIf(Me.chkAddCust = 1, "'frmCustomerAccounts'", "''") & " WHERE (((UserID)=" & Me.DataGrid1.Columns(0).Text & ")) "
cn.Execute SQL

SQL = "UPDATE tbl_Users SET [form10] = " & IIf(Me.chkStatementDomestic = 1, "'frmStatement'", "''") & " WHERE (((UserID)=" & Me.DataGrid1.Columns(0).Text & ")) "
cn.Execute SQL

SQL = "UPDATE tbl_Users SET [form11] = " & IIf(Me.chkStatementInter = 1, "'frmStatementInter'", "''") & " WHERE (((UserID)=" & Me.DataGrid1.Columns(0).Text & ")) "
cn.Execute SQL

SQL = "UPDATE tbl_Users SET [form12] = " & IIf(Me.chkExchangeDoc = 1, "'frmExchangeDoc'", "''") & " WHERE (((UserID)=" & Me.DataGrid1.Columns(0).Text & ")) "
cn.Execute SQL

SQL = "UPDATE tbl_Users SET [form13] = " & IIf(Me.ChkVoidTicket = 1, "'frmVoidTicket'", "''") & " WHERE (((UserID)=" & Me.DataGrid1.Columns(0).Text & ")) "
cn.Execute SQL

SQL = "UPDATE tbl_Users SET [form14] = " & IIf(Me.chkPassporting = 1, "'frmPassporting'", "''") & " WHERE (((UserID)=" & Me.DataGrid1.Columns(0).Text & ")) "
cn.Execute SQL

SQL = "UPDATE tbl_Users SET [form15] = " & IIf(Me.chkPayments = 1, "'frmCashier'", "''") & " WHERE (((UserID)=" & Me.DataGrid1.Columns(0).Text & ")) "
cn.Execute SQL

SQL = "UPDATE tbl_Users SET [form16] = " & IIf(Me.chkVoucher = 1, "'frmCashVoucher'", "''") & " WHERE (((UserID)=" & Me.DataGrid1.Columns(0).Text & ")) "
cn.Execute SQL

SQL = "UPDATE tbl_Users SET [form17] = " & IIf(Me.chkTicketRefund = 1, "'frmRefund'", "''") & " WHERE (((UserID)=" & Me.DataGrid1.Columns(0).Text & ")) "
cn.Execute SQL

Else

End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdNew_Click()
 Me.txtUserName = ""
 Me.txtPassword = ""
 Me.cmdNew.Enabled = False
 Me.cmdSet.Enabled = True
End Sub

Private Sub cmdRemove_Click()
On Error GoTo ErrExit
Dim ask As Integer
ask = MsgBox("Remove the selected user?", vbInformation + vbYesNo)
If ask = vbYes Then
    SQL = "DELETE * FROM tbl_Users WHERE [UserID]=" & Me.DataGrid1.Columns(0).Text
    cn.BeginTrans
    cn.Execute SQL
    Rs.Requery
    cn.CommitTrans
End If
Exit Sub
ErrExit:
cn.RollbackTrans
MsgBox "Delete cancelled"
End Sub

Private Sub cmdSet_Click()
On Error GoTo FailSafe_Error
If Me.txtUserName = "" Then
    MsgBox "User Name should not be blank!"
    Exit Sub
End If


If Me.txtPassword = "" Then
    MsgBox "You forgot your password!"
    Exit Sub
End If

Set Me.DataGrid1.DataSource = Nothing
With Rs
If Me.cmdNew.Enabled = False Then
        .AddNew
End If
        .Fields(1).Value = Me.txtUserName
        .Fields(2).Value = Encrypt(Me.txtPassword)
        .Fields(3).Value = Me.cboAccessType
        .Update
        .Requery
     Me.cmdNew.Enabled = True
Set Me.DataGrid1.DataSource = Rs
End With
 Me.cmdNew.Enabled = False
 Me.cmdSet.Enabled = True
 Exit Sub
FailSafe_Error:
End Sub

Private Sub DataGrid1_Click()
Call LoadUserAcc(Me.DataGrid1.Columns(0).Text)
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
With Me
        .txtUserName = Me.DataGrid1.Columns(1).Text
        .txtPassword = Me.DataGrid1.Columns(2).Text
        .cboAccessType = Me.DataGrid1.Columns(3).Value
        
End With


End Sub

Private Sub Form_Load()

Set Rs = New ADODB.Recordset
SQL = "SELECT * FROM tbl_Users ORDER BY [UserName] ASC"
With Rs
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        
    Set Me.DataGrid1.DataSource = Rs
End With
End Sub
