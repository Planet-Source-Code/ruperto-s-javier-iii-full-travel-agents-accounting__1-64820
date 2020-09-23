VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "Lvbuttons.ocx"
Begin VB.Form frmPO_Domestic 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Purchase Order"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11940
   Icon            =   "frmPO_Domestic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   11940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ListView1 
      Height          =   2865
      Left            =   30
      TabIndex        =   18
      Top             =   3675
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "POIDDetails"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "POID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Particulars"
         Object.Width           =   10583
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "From"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "To"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Qty"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Amount"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   6195
      Top             =   210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   39
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":08CA
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":15A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":3D56
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":4BA8
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":59FA
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":62D4
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":6BAE
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":7488
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":7E52
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":872C
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":8A46
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":9320
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":9BFA
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":A4D4
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":A7EE
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":B0C8
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":B9A2
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":C27C
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":CB56
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":D430
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":DD0A
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":E5E4
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":EEBE
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":F798
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":10072
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":1094C
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":11226
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":11B00
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":123DA
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":12CB4
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":1358E
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":13E44
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":1471E
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":14B70
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":14FC2
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":17774
            Key             =   "IMG35"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":189F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":18D10
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_Domestic.frx":1F572
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3660
      Left            =   15
      TabIndex        =   6
      Top             =   -15
      Width           =   11895
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Note"
         Height          =   975
         Left            =   6480
         TabIndex        =   23
         Top             =   1830
         Width           =   5355
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Posted means PO exist. Simply add tickets that arrive"
            Height          =   285
            Left            =   450
            TabIndex        =   25
            Top             =   615
            Width           =   4725
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Not Posted means PO was created before actual tickets arrived."
            Height          =   285
            Left            =   450
            TabIndex        =   24
            Top             =   345
            Width           =   4665
         End
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "frmPO_Domestic.frx":21D24
         Left            =   8130
         List            =   "frmPO_Domestic.frx":21D2E
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1365
         Width           =   3720
      End
      Begin VB.TextBox txtPONumber 
         BackColor       =   &H00C0FFC0&
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
         Left            =   8130
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   825
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
         TabIndex        =   12
         Top             =   2835
         Width           =   11790
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
            TabIndex        =   14
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
            TabIndex        =   13
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
         Left            =   9270
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   285
         Width           =   2550
      End
      Begin VB.TextBox txtPayto 
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
         Left            =   1545
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   270
         Width           =   4515
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "STATUS :"
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
         Left            =   6510
         TabIndex        =   22
         Top             =   1410
         Width           =   1515
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "PO # :"
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
         Left            =   6885
         TabIndex        =   16
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
         TabIndex        =   10
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   255
         Width           =   1110
      End
   End
   Begin LVbuttons.LaVolpeButton cmdNew 
      Height          =   480
      Left            =   30
      TabIndex        =   0
      Top             =   7395
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
      MICON           =   "frmPO_Domestic.frx":21D46
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
      TabIndex        =   11
      Top             =   7395
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
      MICON           =   "frmPO_Domestic.frx":21D62
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
   Begin LVbuttons.LaVolpeButton cmdPost 
      Height          =   480
      Left            =   1395
      TabIndex        =   3
      Top             =   7395
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
      MICON           =   "frmPO_Domestic.frx":21D7E
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
      TabIndex        =   5
      Top             =   7410
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
      MICON           =   "frmPO_Domestic.frx":21D9A
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
      Left            =   3990
      TabIndex        =   4
      Top             =   6735
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
      MICON           =   "frmPO_Domestic.frx":21DB6
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
      Left            =   6015
      TabIndex        =   15
      Top             =   6735
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
      MICON           =   "frmPO_Domestic.frx":21DD2
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
   Begin LVbuttons.LaVolpeButton cmdFind 
      Height          =   480
      Left            =   4125
      TabIndex        =   20
      Top             =   7395
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
      MICON           =   "frmPO_Domestic.frx":21DEE
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
   Begin LVbuttons.LaVolpeButton cmdPrint 
      Height          =   480
      Left            =   5490
      TabIndex        =   21
      Top             =   7395
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "Print"
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
      MICON           =   "frmPO_Domestic.frx":21E0A
      ALIGN           =   1
      IMGLST          =   "SmallImages"
      IMGICON         =   "39"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
End
Attribute VB_Name = "frmPO_Domestic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsPO As ADODB.Recordset
Dim RsPODetails As ADODB.Recordset
Dim lngIndex As Long

Private Sub cmdFind_Click()
frmPO_DomesticFind.Show 1
End Sub

Private Sub cmdInsert_Click()
frmPO_DomesticDetails.Tag = Me.Tag
frmPO_DomesticDetails.Show 1

End Sub

Private Sub cmdOverRide_Click()
If Me.cmdOverRide.Caption = "Over ride" Then
     Me.cmdOverRide.Caption = "Cancel"
     Me.cmdNew.Enabled = False
     Me.cmdOverRide.Enabled = True
     Me.cmdPost.Enabled = True
     Me.cmdFind.Enabled = True
     Me.Frame1.Enabled = True
Else
    Me.cmdOverRide.Caption = "Over ride"
     Me.cmdNew.Enabled = True
     Me.cmdOverRide.Enabled = True
     Me.cmdPost.Enabled = False
     Me.cmdFind.Enabled = True
    Me.Frame1.Enabled = False
End If
End Sub

Private Sub cmdPrint_Click()

frmAuthorized.Show 1
'        RptPOOTHERS.lblPayto = Me.txtPayto
'        RptPOOTHERS.lblAuthorized.Caption = "HADJIE VILLANUEVA"
'        RptPOOTHERS.Show 1

End Sub

Private Sub cmdRemove_Click()
On Error GoTo ErrExit
Dim ask As Integer

ask = MsgBox("Are you sure you want to remove this details?", vbCritical + vbYesNo)
If ask = vbYes Then
        Me.ListView1.ListItems.Remove lngIndex
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
Me.ListView1.ListItems.Clear
If ReturnFirst(GetLastNumber) = 0 Then
    Me.txtPONumber = GetLastNumber & "-" & kulotRead(App.Path & "\Settings.txt")
Else
    Me.txtPONumber = GetLastNumber & "-" & kulotRead(App.Path & "\Settings.txt")
End If

End Sub

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


Function CheckExist(param) As Boolean
Dim Rst     As New ADODB.Recordset
Dim mySQL   As String

mySQL = "SELECT * FROM Tbl_PO_Domestic WHERE [Po Number]='" & param & "'"
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
'On Error GoTo ErrExit
Dim ask         As Integer
Dim myList      As ListItem

If CheckNull(Me.Combo1) Then
        MsgBox "Select if POSTED or NOT POSTED", vbCritical
        Exit Sub
End If

ask = MsgBox("Are you sure you want to save this PO?", vbInformation + vbYesNo)
If ask = vbNo Then: Exit Sub
    cn.BeginTrans
        With RsPO
        
        If CheckExist(Me.txtPONumber) Then
        'Just to ensure no dups im too lazy
                cn.Execute "DELETE * FROM Tbl_PO_Domestic WHERE [Po Number]='" & Me.txtPONumber & "'"
        End If
        
                .AddNew
                  .Fields("Po Number").Value = Me.txtPONumber
                  .Fields("PO Date").Value = CDate(Me.txtDate)
                  .Fields("Pay to").Value = Me.txtPayto
                  .Fields("Address").Value = Me.txtAddress
                  .Fields("Posted").Value = Me.Combo1
                  .Fields("Amount").Value = CDbl(Me.txtTotalAmount)
                .Update
                
                Call UpDatePODetails(.Fields(0).Value)
        End With
    cn.CommitTrans
    MsgBox "Record Save", vbInformation
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
                        RstInsert.Fields("Amount").Value = CDbl(ListView1.ListItems.Item(Y).SubItems(6))
                        RstInsert.Fields("From").Value = ListView1.ListItems.Item(Y).SubItems(3)
                        RstInsert.Fields("To").Value = ListView1.ListItems.Item(Y).SubItems(4)
                        RstInsert.Fields("Qty").Value = ListView1.ListItems.Item(Y).SubItems(5)
                        
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
        .txtPONumber = Empty
        
       
        .txtDate = Format(Now, "mm/dd/yyyy")
End With
End Sub


Sub LoadValues(param)
Dim RsPODetails         As New ADODB.Recordset
Dim mySQL               As String
Dim ctr                 As Integer

    With RsPO
        .MoveFirst
        .Find "[PoID]=" & param
        If .EOF Or .BOF Then
                MsgBox "Not Found!!!!", vbCritical
            Exit Sub
        End If
    End With
    

With Me
        .txtPONumber = RsPO.Fields("Po Number").Value
        .txtDate = RsPO.Fields("PO Date").Value
        .txtPayto = RsPO.Fields("Pay to").Value
        .txtAddress = RsPO.Fields("Address").Value
        .txtTotalAmount = Format(RsPO.Fields("Amount").Value, "###,##0.00")
       If RsPO.Fields("Posted").Value = "POSTED" Then
                .Combo1.ListIndex = 0
            Else
                .Combo1.ListIndex = 1
       End If
        .Tag = RsPO.Fields("PoID").Value

'//=======================================================================
'//pull out data from details and load it to list view
'//=======================================================================
mySQL = "SELECT * FROM tbl_PODetails_Domestic WHERE [POid]=" & param
Me.ListView1.ListItems.Clear
         With RsPODetails
                .Open mySQL, cn, adOpenKeyset, adLockOptimistic
               If .RecordCount > 0 Then
                .MoveFirst
                    
                    ctr = 0
                    On Error Resume Next
                    Do While Not .EOF
                    ctr = ctr + 1
                        ListView1.ListItems.Add , , .Fields("POID_Details").Value
                        ListView1.ListItems.Item(ctr).SubItems(1) = .Fields("POID").Value
                        ListView1.ListItems.Item(ctr).SubItems(2) = .Fields("Particulars").Value
                        ListView1.ListItems.Item(ctr).SubItems(3) = .Fields("from").Value
                        ListView1.ListItems.Item(ctr).SubItems(4) = .Fields("to").Value
                        ListView1.ListItems.Item(ctr).SubItems(5) = .Fields("qty").Value
                        ListView1.ListItems.Item(ctr).SubItems(6) = Format(.Fields("Amount").Value, "###,##0.00")

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


Private Sub LaVolpeButton1_Click()

End Sub

Private Sub ListView1_Click()
If Me.ListView1.ListItems.Count > 0 Then
    lngIndex = Me.ListView1.SelectedItem.Index
End If
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

