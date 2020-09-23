VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "Lvbuttons.ocx"
Begin VB.Form frmCashVoucher 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cash Voucher"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11940
   Icon            =   "frmCashVoucher.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   11940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ListView1 
      Height          =   2865
      Left            =   15
      TabIndex        =   29
      Top             =   4635
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   5054
      View            =   3
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "VoucherID"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Particulars"
         Object.Width           =   10583
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Amount"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   15
      Top             =   60
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
            Picture         =   "frmCashVoucher.frx":08CA
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":15A4
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":23F6
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":3248
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":3B22
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":43FC
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":4CD6
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":56A0
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":5F7A
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":6294
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":6B6E
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":7448
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":7D22
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":803C
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":8916
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":91F0
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":9ACA
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":A3A4
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":AC7E
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":B558
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":BE32
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":C70C
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":CFE6
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":D8C0
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":E19A
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":EA74
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":F34E
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":FC28
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":10502
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":10DDC
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":11692
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":11F6C
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":123BE
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":12810
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":14FC2
            Key             =   "IMG35"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":16244
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucher.frx":1655E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Enabled         =   0   'False
      Height          =   4230
      Left            =   15
      TabIndex        =   9
      Top             =   360
      Width           =   11895
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Select"
         ForeColor       =   &H80000008&
         Height          =   1470
         Left            =   30
         TabIndex        =   26
         Top             =   1440
         Width           =   11820
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1170
            Left            =   6045
            TabIndex        =   30
            Top             =   165
            Width           =   5670
            Begin VB.OptionButton OptDomestic 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "DOMESTIC"
               ForeColor       =   &H80000008&
               Height          =   360
               Left            =   1035
               TabIndex        =   32
               Top             =   195
               Value           =   -1  'True
               Width           =   3690
            End
            Begin VB.OptionButton OptInternational 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "INTERNATIONAL"
               ForeColor       =   &H80000008&
               Height          =   360
               Left            =   1035
               TabIndex        =   31
               Top             =   735
               Width           =   3690
            End
         End
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
            Left            =   1110
            TabIndex        =   28
            Top             =   840
            Value           =   -1  'True
            Width           =   3255
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
            Left            =   1110
            TabIndex        =   27
            Top             =   345
            Width           =   3255
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
         Left            =   9270
         TabIndex        =   23
         Top             =   825
         Width           =   2550
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1200
         Left            =   30
         ScaleHeight     =   1170
         ScaleWidth      =   11790
         TabIndex        =   15
         Top             =   2970
         Width           =   11820
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
            Left            =   2175
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   570
            Width           =   3795
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
            Left            =   2175
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   60
            Width           =   3795
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
            TabIndex        =   19
            Text            =   "0.00"
            Top             =   600
            Width           =   2565
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
            Left            =   9165
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   75
            Width           =   2565
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
            Left            =   150
            TabIndex        =   21
            Top             =   120
            Width           =   1590
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
            TabIndex        =   18
            Top             =   600
            Width           =   2055
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
            Left            =   120
            TabIndex        =   17
            Top             =   570
            Width           =   2040
         End
         Begin VB.Label Label4 
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
            Left            =   7815
            TabIndex        =   16
            Top             =   90
            Width           =   1110
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
         Left            =   9270
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   285
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
      Begin LVbuttons.LaVolpeButton cmdFind 
         Height          =   480
         Left            =   6105
         TabIndex        =   24
         Top             =   270
         Width           =   1200
         _ExtentX        =   2117
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
         MICON           =   "frmCashVoucher.frx":1CDC0
         ALIGN           =   1
         IMGLST          =   "SmallImages"
         IMGICON         =   "37"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
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
         Left            =   7560
         TabIndex        =   22
         Top             =   840
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
         TabIndex        =   13
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
         Left            =   8295
         TabIndex        =   11
         Top             =   240
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
         TabIndex        =   10
         Top             =   255
         Width           =   1110
      End
   End
   Begin LVbuttons.LaVolpeButton cmdNew 
      Height          =   480
      Left            =   30
      TabIndex        =   0
      Top             =   8280
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
      MICON           =   "frmCashVoucher.frx":1CDDC
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
      Left            =   2760
      TabIndex        =   14
      Top             =   8295
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
      MICON           =   "frmCashVoucher.frx":1CDF8
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
      Left            =   1395
      TabIndex        =   5
      Top             =   8280
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "Post"
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
      MICON           =   "frmCashVoucher.frx":1CE14
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
      Left            =   10545
      TabIndex        =   7
      Top             =   8295
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
      MICON           =   "frmCashVoucher.frx":1CE30
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
   Begin LVbuttons.LaVolpeButton cmdInsert 
      Height          =   480
      Left            =   4110
      TabIndex        =   6
      Top             =   7590
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "Insert Details"
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
      MICON           =   "frmCashVoucher.frx":1CE4C
      ALIGN           =   1
      IMGLST          =   "SmallImages"
      IMGICON         =   "22"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton cmdRemove 
      Height          =   480
      Left            =   6135
      TabIndex        =   20
      Top             =   7590
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "Remove Details"
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
      MICON           =   "frmCashVoucher.frx":1CE68
      ALIGN           =   1
      IMGLST          =   "SmallImages"
      IMGICON         =   "14"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "VOUCHER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   -1500
      TabIndex        =   8
      Top             =   0
      Width           =   14490
   End
End
Attribute VB_Name = "frmCashVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsVoucher As ADODB.Recordset
Dim RsVoucherDetails As ADODB.Recordset



Private Sub cboAccount_Click()
Me.txtCheckNo = ""

'Return Check if check is selected
If Me.Option2 Then
    Me.txtCheckNo = ReturnCheck()
End If

End Sub

Private Sub cboBank_Click()
Call FillAccount
End Sub

Private Sub cmdFind_Click()
frmCashVoucherFind.Show 1
End Sub

Private Sub cmdInsert_Click()
frmCashVoucherDetails.Tag = Me.Tag
frmCashVoucherDetails.Show 1
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
Call FillBank
Set RsVoucher = New ADODB.Recordset
Set RsVoucherDetails = New ADODB.Recordset

SQL = "SELECT * FROM tbl_Voucher"
With RsVoucher
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
End With
 

End Sub


Private Sub cboBank_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
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
Me.txtPayto.SetFocus
End Sub

Function SumListView() As Double
Dim Y               As Integer
Dim Tmp             As Double

If Me.ListView1.ListItems.Count > 0 Then
        For Y = 1 To Me.ListView1.ListItems.Count
            If Me.ListView1.ListItems(Y).SubItems(2) <> "" Then
                Tmp = Tmp + CDbl(Me.ListView1.ListItems(Y).SubItems(2))
            End If
        Next Y
Else
        Tmp = 0
End If
        SumListView = Tmp
End Function


Sub UpDateVoucherDetails()
Dim Y               As Integer
Dim Tmp             As Double
Dim RstInsert       As New ADODB.Recordset
Dim mySQL           As String

mySQL = "SELECT * FROM tbl_VoucherDetails"

If Me.ListView1.ListItems.Count > 0 Then

        RstInsert.Open mySQL, cn, adOpenKeyset, adLockOptimistic

        For Y = 1 To Me.ListView1.ListItems.Count
        '//if voucher details does not exist  add
            If Not IsNumeric(Me.ListView1.ListItems(Y).Text) Then
                RstInsert.AddNew
                        RstInsert.Fields("VoucherID").Value = Me.txtVoucherID
                        RstInsert.Fields("Particulars").Value = ListView1.ListItems.Item(Y).SubItems(1)
                        RstInsert.Fields("Amount").Value = CDbl(ListView1.ListItems.Item(Y).SubItems(2))
                RstInsert.Update
            End If
        Next Y
Else
        Tmp = 0
End If
End Sub


Sub UpDateVoucher(Param)
On Error GoTo ErrExit
cn.BeginTrans
SQL = "UPDATE tbl_Voucher SET [TotalAmount] =" & CDbl(Me.txtTotalAmount) & ", [Has Issued Check]=TRUE WHERE VoucherID= " & Param
cn.Execute SQL
cn.CommitTrans
Exit Sub
ErrExit:
cn.RollbackTrans
End Sub


Sub UpdateCheck()
cn.BeginTrans
SQL = "UPDATE tbl_checks SET [Status] = 'issued' WHERE [CheckNo]='" & Me.txtCheckNo & "' AND [AccNo]='" & Me.cboAccount & "' AND [Status]='Un-Used' "
cn.Execute SQL
cn.CommitTrans
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

Private Sub cmdPost_Click()
'On Error GoTo ErrExit

Dim ask As Integer

ask = MsgBox("Are you sure you want to save this Voucher?", vbInformation + vbYesNo)
If ask = vbNo Then: Exit Sub

If ZeroBal Then
    MsgBox "cannot continue insufficient amount for current account..." & Me.cboAccount, vbInformation
    Exit Sub
End If

If CheckNull(Me.cboBank) Then
MsgBox "Bank Name should not be blank", vbCritical
Exit Sub
End If

If CheckNull(Me.txtCheckNo) Then
MsgBox "Check NO should not be blank", vbCritical
Me.txtCheckNo.SetFocus
Exit Sub
End If

cn.BeginTrans
    With RsVoucher
        If CheckNull(Me.txtVoucherID) Then
            .AddNew
        Else
            .MoveFirst
            .Find "[VoucherID]=" & CLng(Me.txtVoucherID)
        End If
            .Fields(1).Value = UCase(Me.txtPayto)
            .Fields(2).Value = UCase(Me.txtAddress)
            .Fields(3).Value = Format(Now, "mm/dd/yyyy")
            .Fields(4).Value = FindBankID(Me.cboBank)
            If Me.Option2 Then
                .Fields(5).Value = UCase(Me.txtCheckNo)
                .Fields("Cash").Value = False
                .Fields("Check").Value = True
            Else
                .Fields("Cash").Value = True
                .Fields("Check").Value = False
            End If
            .Fields(6).Value = Format(Me.txtTotalAmount, "###,##0.00")
            .Fields(7).Value = True
        .Update
        Me.Tag = .Fields(0).Value
        Me.txtVoucherID = Me.Tag
    End With
    
    Dim myTmp
    Dim i
    myTmp = ""
    If Me.ListView1.ListItems.Count > 0 Then
        For i = 1 To Me.ListView1.ListItems.Count
            myTmp = myTmp & "," & Me.ListView1.ListItems(i).SubItems(1)
        Next i
    End If
    
    '// Now Deduct this voucher to bank
    Call UpdatePassbook(CDbl(Me.txtTotalAmount), _
    IIf(Me.Option2 = True, Me.txtCheckNo, ""), _
    Format(Now, "mm/dd/yyyy"), _
    "Issued Expense Voucher to :" & Me.txtPayto & " as payment for :" & myTmp, _
    Me.cboAccount, "n/a", _
    IIf(Me.Option1 = True, CDbl(Me.txtTotalAmount), 0), 0, _
    IIf(Me.Option2 = True, CDbl(Me.txtTotalAmount), 0), 0, "", "", "", "", "", "", "")
    
    Call UpDateVoucherDetails
    Call UpdateCheck
cn.CommitTrans
    MsgBox "Voucher Expense successfully created", vbInformation
'Call RefreshGrid(Me.Tag)

Me.cmdNew.Enabled = True
Me.cmdPost.Enabled = False
Me.cmdInsert.Enabled = True
Me.cmdRemove.Enabled = True
Me.Frame1.Enabled = False
Me.Picture1.Enabled = False
Me.cmdInsert.SetFocus
Exit Sub
ErrExit:
cn.RollbackTrans
MsgBox "There was an error while saving Voucher expense REASON:" & Err.Description, vbCritical
End Sub

Sub ClearTxt()
With Me
        .txtPayto = Empty
        .txtAddress = Empty
        .txtCheckNo = Empty
        .txtTotalAmount = "0.00"
        '.txtAccName = Empty
        .txtVoucherID = Empty
        
       ' .cboBank.Clear
        .txtDate = Format(Now, "mm/dd/yyyy")
End With
End Sub


Sub LoadValues(Param)
Dim RstVoucherDetails   As New ADODB.Recordset
Dim mySQL               As String

    With RsVoucher
        .MoveFirst
        .Find "[VoucherID]=" & Param
        If .EOF Or .BOF Then
            MsgBox "Not Found!!!!", vbCritical
            Exit Sub
        End If
    End With
    
mySQL = "SELECT * FROM tbl_VoucherDetails WHERE [VoucherID]=" & Param
With Me
        .txtVoucherID = IIf(Not IsNull(RsVoucher.Fields(0).Value), RsVoucher.Fields(0).Value, "")
        .txtPayto = IIf(Not IsNull(RsVoucher.Fields(1).Value), RsVoucher.Fields(1).Value, "")
        .txtAddress = IIf(Not IsNull(RsVoucher.Fields(2).Value), RsVoucher.Fields(2).Value, "")
        .txtDate = IIf(Not IsNull(RsVoucher.Fields(3).Value), RsVoucher.Fields(3).Value, "")
        .txtTotalAmount = IIf(Not IsNull(RsVoucher.Fields(6).Value), Format(RsVoucher.Fields(6).Value, "###,##0.00"), "")
        .Tag = .txtVoucherID
'Call RefreshGrid(.txtVoucherID)
'//=======================================================================
'//pull out data from details and load it to list view
'//=======================================================================

 
         With RstVoucherDetails
                .Open mySQL, cn, adOpenKeyset, adLockOptimistic
               If .RecordCount > 0 Then
                .MoveFirst
                    Me.ListView1.ListItems.Clear
                    Do While Not .EOF
                        ListView1.ListItems.Add , , IIf(Not IsNull(.Fields("VoucherIDDetails").Value), .Fields("VoucherIDDetails").Value, "")
                        ListView1.ListItems.Item(1).SubItems(1) = IIf(Not IsNull(.Fields("Particulars").Value), .Fields("Particulars").Value, "")
                        ListView1.ListItems.Item(1).SubItems(2) = IIf(Not IsNull(.Fields("Amount").Value), Format(.Fields("Amount").Value, "###,##0.00"), "")
                        .MoveNext
                    Loop
               End If
        End With
'//=======================================================================
End With

End Sub
Private Sub cmdPost_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
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

Private Sub txtAddress_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub


Private Sub txtAddress_LostFocus()
txtAddress = UCase(txtAddress)
End Sub

Private Sub txtCheckNo_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub


Private Sub txtCheckNo_LostFocus()
'If txtCheckNo <> "" Then
 '   FindCheck (Me.txtCheckNo)
'End If
End Sub

Private Sub txtPayto_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub


Sub FindCheck(check)
Dim Tmp As ADODB.Recordset
Set Tmp = New ADODB.Recordset
SQL = "SELECT  * FROM tbl_checks WHERE [CheckNo]='" & UCase(check) & "' AND [BankID]=" & FindBankID(Me.cboBank) & " AND [Status]='Un-Used'"
With Tmp
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            'Me.txtAccName = .Fields("AccNo").Value
            'Me.txtBankName = FindBankID(.Fields("BankID").Value)
        Else
            MsgBox "This check # does not exist or already used!", vbCritical
            With Me.txtCheckNo
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
            End With
        End If
      .Close
    Set Tmp = Nothing
End With

End Sub

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


Private Sub txtPayto_LostFocus()
txtPayto = UCase(txtPayto)
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
            
        End If
End With

End Sub


Sub UpdatePassbook(ByVal nAmt As Double, _
    ByVal CheckNo As String, ByVal CheckDate As String, _
    Optional Desc As String, Optional ByVal AccNo As String, _
    Optional strAir As String, _
    Optional nCash, Optional nCard, Optional nCheck, _
    Optional nOthers, Optional nCardName, _
    Optional nCardNumber, Optional nCardHolder, _
    Optional nBank1, Optional nBank2, _
    Optional nBank3, Optional nBank4)
    
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
        If IsExist_Voucher(Me.txtVoucherID) Then
            .MoveFirst
            .Find "[Voucher No]='" & Me.txtVoucherID & "'"
        Else
            .AddNew
        End If
            .Fields("Deposit Date").Value = Format(Now, "mm/dd/yyyy")
            .Fields("Check No").Value = CheckNo
            .Fields("Check Date").Value = CheckDate
            .Fields("Voucher No").Value = Me.txtVoucherID
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

