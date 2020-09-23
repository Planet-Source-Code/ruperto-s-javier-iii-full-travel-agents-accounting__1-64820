VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "Lvbuttons.ocx"
Begin VB.Form frmVoucherPayShipAirPO 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Voucher Purchase Order"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11940
   Icon            =   "frmVoucherPayShipAirPO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   11940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Select"
      ForeColor       =   &H80000008&
      Height          =   1305
      Left            =   75
      TabIndex        =   29
      Top             =   7005
      Width           =   5460
      Begin VB.OptionButton OptDOMESTIC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "DOMESTIC PO"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   960
         TabIndex        =   31
         Top             =   270
         Value           =   -1  'True
         Width           =   4395
      End
      Begin VB.OptionButton OptINTL 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "INTERNATIONAL PO"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   975
         TabIndex        =   30
         Top             =   795
         Width           =   3690
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3240
      Left            =   45
      TabIndex        =   16
      Top             =   3690
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   5715
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "POID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "PO #"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "PO Date"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Payto"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Address"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Posted"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Amount"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   150
      Top             =   75
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   37
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":08CA
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":15A4
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":23F6
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":3248
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":3B22
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":43FC
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":4CD6
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":56A0
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":5F7A
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":6294
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":6B6E
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":7448
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":7D22
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":803C
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":8916
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":91F0
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":9ACA
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":A3A4
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":AC7E
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":B558
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":BE32
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":C70C
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":CFE6
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":D8C0
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":E19A
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":EA74
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":F34E
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":FC28
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":10502
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":10DDC
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":11692
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":11F6C
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":123BE
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":12810
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":14FC2
            Key             =   "IMG35"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":16244
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoucherPayShipAirPO.frx":1655E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3600
      Left            =   15
      TabIndex        =   5
      Top             =   -30
      Width           =   11895
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Voucher Details"
         ForeColor       =   &H80000008&
         Height          =   2325
         Left            =   7410
         TabIndex        =   20
         Top             =   195
         Width           =   4395
         Begin VB.ComboBox cboAccount 
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
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   705
            Width           =   2220
         End
         Begin VB.ComboBox cboBank 
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
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   195
            Width           =   2220
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
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
            Left            =   2175
            Locked          =   -1  'True
            TabIndex        =   22
            Text            =   "0.00"
            Top             =   1755
            Width           =   2175
         End
         Begin VB.TextBox txtCheckNo 
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
            Left            =   2175
            TabIndex        =   21
            Top             =   1230
            Width           =   2175
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "BANK NAME :"
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
            Left            =   525
            TabIndex        =   28
            Top             =   240
            Width           =   1590
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "ACCOUNT # :"
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
            Left            =   555
            TabIndex        =   27
            Top             =   690
            Width           =   1845
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "SUB TOTAL :"
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
            Left            =   75
            TabIndex        =   26
            Top             =   1815
            Width           =   2055
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "CHECK # :"
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
            Left            =   885
            TabIndex        =   25
            Top             =   1275
            Width           =   1110
         End
      End
      Begin VB.TextBox txtVoucherNo 
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
         Height          =   465
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1545
         Width           =   3690
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   45
         ScaleHeight     =   705
         ScaleWidth      =   11760
         TabIndex        =   11
         Top             =   2655
         Width           =   11790
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   405
            Left            =   3030
            TabIndex        =   19
            Top             =   150
            Value           =   -1  'True
            Width           =   1800
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Cash"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   405
            Left            =   615
            TabIndex        =   18
            Top             =   165
            Width           =   1725
         End
         Begin VB.TextBox txtTotalAmount 
            Alignment       =   1  'Right Justify
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
            Left            =   9165
            Locked          =   -1  'True
            TabIndex        =   13
            Text            =   "0.00"
            Top             =   135
            Width           =   2565
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL AMOUNT :"
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
            Left            =   7005
            TabIndex        =   12
            Top             =   135
            Width           =   2055
         End
      End
      Begin VB.TextBox txtAddress 
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
         Left            =   1545
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   840
         Width           =   4515
      End
      Begin VB.TextBox txtDate 
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
         Left            =   1530
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   2160
         Width           =   2550
      End
      Begin VB.TextBox txtPayto 
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
         Left            =   1545
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   270
         Width           =   4515
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher # :"
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
         Left            =   120
         TabIndex        =   14
         Top             =   1575
         Width           =   1515
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS :"
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
         Left            =   135
         TabIndex        =   9
         Top             =   855
         Width           =   1245
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "DATE :"
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
         Left            =   555
         TabIndex        =   7
         Top             =   2115
         Width           =   1110
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PAY TO :"
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
         Left            =   135
         TabIndex        =   6
         Top             =   255
         Width           =   1110
      End
   End
   Begin LVbuttons.LaVolpeButton cmdNew 
      Height          =   480
      Left            =   60
      TabIndex        =   0
      Top             =   8400
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   847
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
      COLTYPE         =   2
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmVoucherPayShipAirPO.frx":1CDC0
      ALIGN           =   1
      IMGLST          =   "SmallImages"
      IMGICON         =   "4"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton cmdOverRide 
      Height          =   480
      Left            =   2790
      TabIndex        =   10
      Top             =   8400
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "Over ride"
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
      MICON           =   "frmVoucherPayShipAirPO.frx":1CDDC
      ALIGN           =   1
      IMGLST          =   "SmallImages"
      IMGICON         =   "29"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton cmdPost 
      Height          =   480
      Left            =   1425
      TabIndex        =   3
      Top             =   8400
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   847
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
      MICON           =   "frmVoucherPayShipAirPO.frx":1CDF8
      ALIGN           =   1
      IMGLST          =   "SmallImages"
      IMGICON         =   "9"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton cmdExit 
      Height          =   480
      Left            =   10575
      TabIndex        =   4
      Top             =   8415
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   847
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
      MICON           =   "frmVoucherPayShipAirPO.frx":1CE14
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
   Begin LVbuttons.LaVolpeButton cmdFind 
      Height          =   480
      Left            =   4155
      TabIndex        =   17
      Top             =   8400
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "Find"
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
      MICON           =   "frmVoucherPayShipAirPO.frx":1CE30
      ALIGN           =   1
      IMGLST          =   "SmallImages"
      IMGICON         =   "21"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
End
Attribute VB_Name = "frmVoucherPayShipAirPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsPO As ADODB.Recordset
Dim RsPODetails As ADODB.Recordset


Private Sub cboAccount_Click()
Me.txtCheckNo = ""
'Return Check if check is selected
Me.txtCheckNo = ReturnCheck()
End Sub

Private Sub cboBank_Click()
Call FillAccount
End Sub

Private Sub cmdFind_Click()

If Me.OptDOMESTIC Then
    frmPO_DomesticFind.Tag = "voucher"
    frmPO_DomesticFind.Show 1
End If

If Me.OptINTL Then
    frmPO_INTL_Find.Tag = "voucher"
    frmPO_INTL_Find.Show 1
End If



End Sub

Private Sub cmdInsert_Click()
frmPO_DomesticDetails.Tag = Me.Tag
frmPO_DomesticDetails.Show 1

End Sub

Private Sub cmdOverRide_Click()
If Me.cmdOVerRide.Caption = "Over ride" Then
     Me.cmdOVerRide.Caption = "Cancel"
     Me.cmdNew.Enabled = False
     Me.cmdOVerRide.Enabled = True
     Me.cmdPost.Enabled = True
     Me.cmdFind.Enabled = True
     Me.Frame1.Enabled = True
Else
    Me.cmdOVerRide.Caption = "Over ride"
     Me.cmdNew.Enabled = True
     Me.cmdOVerRide.Enabled = True
     Me.cmdPost.Enabled = False
     Me.cmdFind.Enabled = True
    Me.Frame1.Enabled = False
End If
End Sub

Private Sub cmdRemove_Click()
On Error GoTo ErrExit
Dim ask As Integer

ask = MsgBox("Are you sure you want to remove this details?", vbCritical + vbYesNo)
If ask = vbYes Then
        MsgBox "Under construction"
End If
Exit Sub
ErrExit:
Select Case Err.Number
Case 7005
        MsgBox "There are no details to remove", vbInformation
End Select
End Sub

Private Sub Form_Load()
Call ClearTxt

Set RsPO = New ADODB.Recordset
Set RsPODetails = New ADODB.Recordset

SQL = "SELECT * FROM Tbl_PO_Domestic"

RsPO.Open SQL, cn, adOpenKeyset, adLockOptimistic

   Call FillBank
    
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdNew_Click()
Me.cmdNew.Enabled = False
Me.cmdPost.Enabled = True
'Me.cmdInsert.Enabled = False
'Me.cmdRemove.Enabled = False
Me.Frame1.Enabled = True
Me.Picture1.Enabled = True
Call ClearTxt
Me.txtVoucherNo = RetVoucherNo
Me.txtPayto.SetFocus
End Sub

Function RetVoucherNo() As String
Dim Rst As New ADODB.Recordset
Dim SQL As String

SQL = "SELECT * FROM tbl_Voucher"
With Rst
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        .MoveLast
     RetVoucherNo = CDbl(.Fields(0).Value) + 1
     .Close
    Set Rst = Nothing
End With

End Function

Function SumListView() As Double
Dim Y               As Integer
Dim Tmp             As Double

If Me.ListView1.ListItems.Count > 0 Then
        For Y = 1 To Me.ListView1.ListItems.Count
            If Me.ListView1.ListItems(Y).SubItems(6) <> "" Then
                Tmp = Tmp + CDbl(Me.ListView1.ListItems(Y).SubItems(6))
            End If
        Next Y
Else
        Tmp = 0
End If
        SumListView = Tmp
End Function


Function CheckExist(Param) As Boolean
Dim Rst     As New ADODB.Recordset
Dim mySQL   As String

mySQL = "SELECT * FROM Tbl_PO_Domestic WHERE [Po Number]='" & Param & "'"
With Rst
    .Open mySQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            CheckExist = True
        Else
            CheckExist = False
        End If
        .Close
     Set Rst = Nothing
End With

End Function

Private Sub cmdPost_Click()
'On Error GoTo FailSafe_Err
Dim Rst             As New ADODB.Recordset
Dim SQL             As String
Dim i               As Integer
Dim myParticular    As String
Dim ask             As Integer
Dim myAmount        As Double
Dim myPOnumber      As String

SQL = "SELECT * FROM tbl_Voucher ORDER by VoucherID ASC"

If ZeroBal Then
    MsgBox "cannot continue insufficient amount for current account..." & Me.cboAccount, vbInformation
    Exit Sub
End If

If CheckNull(Me.cboBank) Then: MsgBox "Please select Bank!", vbInformation: Exit Sub
If CheckNull(Me.cboAccount) Then: MsgBox "Please select account!", vbInformation: Exit Sub
If Me.txtCheckNo = "No Checks Available!" Then: MsgBox "Invalid check!", vbInformation: Exit Sub
If Me.ListView1.ListItems.Count = 0 Then: MsgBox "No Po to voucher!", vbInformation: Exit Sub

ask = MsgBox("Sure to save this?", vbInformation + vbYesNo)
If ask = vbNo Then: Exit Sub
With Rst
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
         cn.BeginTrans
        .AddNew
                .Fields("Payto").Value = Me.txtPayto
                .Fields("Date").Value = RetDate(Now)
                .Fields("BankID").Value = FindBankID(Me.cboBank)
                .Fields("CheckNo").Value = Me.txtCheckNo
                .Fields("TotalAmount").Value = Me.txtTotalAmount
                .Fields("Has Issued Check").Value = True
                .Fields("Cash").Value = False
                .Fields("Check").Value = True
                .Fields("For Refund").Value = False
                .Fields("Electronic Transfer").Value = False
        .Update
        
        If Me.ListView1.ListItems.Count > 0 Then
        
                For i = 1 To Me.ListView1.ListItems.Count
               '     myParticular = Me.ListView1.ListItems(i).SubItems(1)
                    myAmount = CDbl(Me.ListView1.ListItems(i).SubItems(6))
                    myPOnumber = myPOnumber & "," & Me.ListView1.ListItems(i).SubItems(1)
               '         Call FilterPoDetails(Me.ListView1.ListItems(i).Text, .Fields("VoucherID").Value)
                myParticular = Me.ListView1.ListItems(i).Text & "-" & Me.ListView1.ListItems(i).SubItems(1) & " " & Me.ListView1.ListItems(i).SubItems(2)
                Call VoucherDetailsAdd(.Fields("VoucherID").Value, Me.ListView1.ListItems(i).SubItems(1), myAmount)
                Next i
        End If
        
        
'// Now Deduct this voucher to bank
    Call UpdatePassbook(CDbl(Me.txtTotalAmount), _
    IIf(Me.Option2 = True, Me.txtCheckNo, ""), _
    Format(Now, "mm/dd/yyyy"), _
    "Issued Voucher to :" & Me.txtPayto & " as payment(s) for PO  " & _
                            myPOnumber, _
    Me.cboAccount, "n/a", _
    IIf(Me.Option1 = True, CDbl(Me.txtTotalAmount), 0), 0, _
    IIf(Me.Option2 = True, CDbl(Me.txtTotalAmount), 0), 0, "", "", "", "", "", "", "", .Fields("VoucherID").Value)
    Call UpdateCheck
        
        
        cn.CommitTrans
        MsgBox "Voucher Save", vbInformation
       .Close
     Set Rst = Nothing
End With
Exit Sub
FailSafe_Err:
cn.RollbackTrans
MsgBox "There was an error while saving the voucher", vbInformation

End Sub

Sub FilterPoDetails(Param, usrVID)
Dim Rst As New ADODB.Recordset
Dim mySQL As String

mySQL = "SELECT * FROM tbl_PODetails_Domestic WHERE [POid]=" & Param
Me.ListView1.ListItems.Clear
         With Rst
                .Open mySQL, cn, adOpenKeyset, adLockOptimistic
               If .RecordCount > 0 Then
                .MoveFirst
                    Do While Not .EOF
                        Call VoucherDetailsAdd(usrVID, .Fields("Particulars").Value, .Fields("Amount").Value)
                        .MoveNext
                    Loop
               End If
        End With

End Sub

Sub VoucherDetailsAdd(Param, usrParticulars, usrAmount)
Dim Rst As New ADODB.Recordset
Dim SQL As String

SQL = "SELECT * FROM tbl_VoucherDetails"
With Rst
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
        .AddNew
                .Fields("VoucherID").Value = Param
                .Fields("Particulars").Value = usrParticulars
                .Fields("Amount").Value = CDbl(usrAmount)
                .Fields("Payto").Value = Me.txtPayto
        .Update
       .Close
     Set Rst = Nothing
End With


End Sub

Sub UpDatePODetails(usrPOID)
Dim Y               As Integer
Dim Tmp             As Double
Dim RstInsert       As New ADODB.Recordset
Dim mySQL           As String

mySQL = "SELECT * FROM tbl_PODetails_Domestic"

If Me.ListView1.ListItems.Count > 0 Then

        RstInsert.Open mySQL, cn, adOpenKeyset, adLockOptimistic
        For Y = 1 To Me.ListView1.ListItems.Count
                RstInsert.AddNew
                        RstInsert.Fields("POID").Value = usrPOID
                        RstInsert.Fields("Particulars").Value = ListView1.ListItems.Item(Y).SubItems(2)
                        RstInsert.Fields("Amount").Value = CDbl(ListView1.ListItems.Item(Y).SubItems(3))
                RstInsert.Update
        Next Y
End If
End Sub

Sub ClearTxt()
With Me
        .txtPayto = Empty
        .txtAddress = Empty
        
        .txtTotalAmount = "0.00"
        '.txtAccName = Empty
        .txtVoucherNo = Empty
        
       
        .txtDate = Format(Now, "mm/dd/yyyy")
End With
End Sub


Sub LoadValues(Param)
Dim Rst         As New ADODB.Recordset
Dim mySQL               As String
Dim ctr                 As Integer
Dim mylist  As ListItem


With Me

'//=======================================================================
'//pull out data from po and load it to list view
'//=======================================================================
mySQL = "SELECT * FROM Tbl_PO_Domestic WHERE [POid]=" & Param
'Me.ListView1.ListItems.Clear
         With Rst
                .Open mySQL, cn, adOpenKeyset, adLockOptimistic
               If .RecordCount > 0 Then
               
                    Set mylist = ListView1.ListItems.Add(, , .Fields("POID").Value)
                        mylist.SubItems(1) = .Fields("Po Number").Value
                        mylist.SubItems(2) = .Fields("PO Date").Value
                        mylist.SubItems(3) = .Fields("Pay to").Value
                        mylist.SubItems(4) = .Fields("Address").Value
                        mylist.SubItems(5) = .Fields("Posted").Value
                        mylist.SubItems(6) = .Fields("Amount").Value
               End If
               .Close
             Set Rst = Nothing
        End With

 
'//=======================================================================
End With

End Sub

Sub LoadValuesINTL(Param)
Dim Rst                 As New ADODB.Recordset
Dim mySQL               As String
Dim ctr                 As Integer
Dim mylist              As ListItem


With Me

'//=======================================================================
'//pull out data from po and load it to list view
'//=======================================================================
mySQL = "SELECT * FROM Tbl_PO_INTL WHERE [POid]=" & Param
'Me.ListView1.ListItems.Clear
         With Rst
                .Open mySQL, cn, adOpenKeyset, adLockOptimistic
               If .RecordCount > 0 Then
               
                    Set mylist = ListView1.ListItems.Add(, , .Fields("POID").Value)
                        mylist.SubItems(1) = .Fields("Po Number").Value
                        mylist.SubItems(2) = .Fields("PO Date").Value
                        mylist.SubItems(3) = .Fields("Pay to").Value
                        mylist.SubItems(4) = .Fields("Route").Value
                        mylist.SubItems(5) = .Fields("Posted").Value
                        mylist.SubItems(6) = .Fields("Grand Total Peso").Value
                        
                        
               End If
               .Close
             Set Rst = Nothing
        End With

 
'//=======================================================================
End With

End Sub

Private Sub cmdPost_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub


Private Sub txtAddress_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtAddress_LostFocus()
txtAddress = UCase(txtAddress)
End Sub

Private Sub txtCheckNo_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub


Private Sub txtPayto_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtPayto_LostFocus()
txtPayto = UCase(txtPayto)
End Sub

Function IsExist_Voucher(Param) As Boolean
Dim Rst         As New ADODB.Recordset
Dim SQL         As String

SQL = "SELECT * FROM qryBankPassbook WHERE [Voucher No]='" & Param & "'"
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


Function GetLastNumber() As String
Dim RsFnumber       As ADODB.Recordset
Dim SQL             As String
Dim Tmp             As String
Dim myTmpPos        As Integer

Set RsFnumber = New ADODB.Recordset
SQL = "SELECT [Po Number] from Tbl_PO_Domestic ORDER by [Po Number] ASC"
RsFnumber.Open SQL, cn, adOpenKeyset, adLockOptimistic

With RsFnumber
        If .RecordCount > 0 Then
                .MoveLast
               Tmp = RsFnumber("Po Number").Value
               myTmpPos = Int(ReturnFirst(Tmp)) - (Int(ReturnFirst(Tmp)) - Int(Return_1stDash(Tmp)))
               Tmp = Mid(Tmp, Return_1stDash(Tmp) + 5, myTmpPos - 1)
               GetLastNumber = "PO" & WhichBranch.Fields(2).Value & Mid(Year(Now), 3, 2) & "-" & returnMon & returnDay & AutoIncrement(Tmp)
        Else
               GetLastNumber = "PO" & WhichBranch.Fields(2).Value & Mid(Year(Now), 3, 2) & "-" & returnMon & returnDay & "000000000"
        End If
End With

End Function



Sub UpdatePassbook(ByVal nAmt As Double, _
    ByVal CheckNo As String, ByVal CheckDate As String, _
    Optional Desc As String, Optional ByVal AccNo As String, _
    Optional strAir As String, _
    Optional nCash, Optional nCard, Optional nCheck, _
    Optional nOthers, Optional nCardName, _
    Optional nCardNumber, Optional nCardHolder, _
    Optional nBank1, Optional nBank2, _
    Optional nBank3, Optional nBank4, Optional usrVID)
    
'On Error Resume Next
Dim RsPassbk As New ADODB.Recordset
Dim RsBalance As New ADODB.Recordset
Dim TempBal As Double

SQL = "SELECT * FROM tbl_AccountsSetting WHERE [Account Number]='" & AccNo & "'"
With RsPassbk
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                    .MoveFirst
                 TempBal = .Fields("Current Balance").Value
            End If
            .Close
      Set RsPassbk = Nothing
End With




SQL = "SELECT * FROM tbl_BankPassbook"
With RsPassbk
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If IsExist_Voucher(usrVID) Then
            .MoveFirst
            .Find "[Voucher No]='" & usrVID & "'"
        Else
            .AddNew
        End If
            .Fields("Deposit Date").Value = Format(Now, "mm/dd/yyyy")
            .Fields("Check No").Value = CheckNo
            .Fields("Check Date").Value = CheckDate
            .Fields("Voucher No").Value = usrVID
            .Fields("Description").Value = Desc
            .Fields("Credit").Value = 0
            .Fields("Debit").Value = nAmt
            .Fields("Account Number").Value = AccNo
            .Fields("Balance").Value = TempBal - nAmt
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

SQL = "UPDATE tbl_AccountsSetting SET [Current Balance] = " & _
              CDbl(TempBal - nAmt) & " WHERE [Account Number]= '" & UCase(AccNo) & "'"
              cn.BeginTrans
                    cn.Execute SQL
              cn.CommitTrans
Exit Sub
FailSafe_Error:
cn.RollbackTrans
End Sub

Function ZeroBal() As Boolean
Dim RsPassbk As New ADODB.Recordset
Dim RsBalance As New ADODB.Recordset
Dim TempBal As Double

SQL = "SELECT * FROM tbl_AccountsSetting WHERE [Account Number]='" & Me.cboAccount & "'"
With RsPassbk
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                    .MoveFirst
                 TempBal = .Fields("Current Balance").Value
                 
                 If CDbl(TempBal) <= 0 Then
                     ZeroBal = True
                 Else
                     ZeroBal = False
                 End If
            End If
            .Close
      Set RsPassbk = Nothing
End With
End Function


Sub UpdateCheck()
cn.BeginTrans
SQL = "UPDATE tbl_checks SET [Status] = 'issued' WHERE [CheckNo]='" & Me.txtCheckNo & "' AND [AccNo]='" & Me.cboAccount & "' AND [Status]='Un-Used' "
cn.Execute SQL
cn.CommitTrans
End Sub

Sub FillAccount()
Dim Rst As New ADODB.Recordset
SQL = "SELECT DISTINCT  [Account Number] FROM tbl_AccountsSetting WHERE [BankID]=" & FindBankID(Me.cboBank) & " ORDER by [Account Number] ASC"
With Rst
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            Me.cboAccount.Clear
                    .MoveFirst
               Do While Not .EOF
                    Me.cboAccount.AddItem .Fields("Account Number").Value
                    .MoveNext
               Loop
        Else
            Me.cboAccount.Clear
            Me.txtCheckNo = ""
        End If
End With

End Sub

Sub FillBank()
Dim Rst As New ADODB.Recordset
SQL = "SELECT * FROM tbl_Banks"
        Me.cboBank.Clear
With Rst
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            Do While Not .EOF
                Me.cboBank.AddItem .Fields(1).Value
                .MoveNext
            Loop
        End If
       .Close
     Set Rst = Nothing
End With
End Sub

Function ReturnCheck() As String
Dim Tmp As ADODB.Recordset
Set Tmp = New ADODB.Recordset
SQL = "SELECT  * FROM tbl_checks WHERE [BankID]=" & FindBankID(Me.cboBank) & " AND [Status]='Un-Used' AND [AccNo]='" & Me.cboAccount & "'ORDER by [CheckNo] ASC"
With Tmp
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            ReturnCheck = .Fields("CheckNo").Value
            
        Else
            ReturnCheck = "No Checks Available!"
        End If
      .Close
    Set Tmp = Nothing
End With

End Function

Function FindBankID(Param) As Long
Dim Rst As New ADODB.Recordset
SQL = "SELECT * FROM tbl_Banks WHERE [Bank Name]='" & UCase(Param) & "'"
With Rst
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
              FindBankID = .Fields(0).Value
          Else
              FindBankID = -1
        End If
     .Close
   Set Rst = Nothing
End With
End Function


Function CheckVoucher(usrPart) As Boolean
Dim Rst As New ADODB.Recordset
Dim SQL As String

SQL = "SELECT * FROM tbl_VoucherDetails WHERE [Particulars]='" & Trim(usrPart) & "'"
With Rst
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            
            If .RecordCount > 0 Then
                CheckVoucher = True
                     Else
                CheckVoucher = False
            End If
       .Close
     Set Rst = Nothing
End With

End Function
