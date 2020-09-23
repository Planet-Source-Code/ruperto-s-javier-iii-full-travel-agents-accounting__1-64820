VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "Lvbuttons.ocx"
Begin VB.Form frmStatementInter 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "STATEMENT OF ACCOUNTS INTERNATIONAL"
   ClientHeight    =   10215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12120
   Icon            =   "frmStatementInter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10215
   ScaleWidth      =   12120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00400040&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   12090
      TabIndex        =   68
      Top             =   -15
      Width           =   12150
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Statement of Accounts (International Only)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   600
         TabIndex        =   70
         Top             =   120
         Width           =   9780
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Note:  Issuance of Statements for International"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   69
         Top             =   600
         Width           =   4575
      End
      Begin VB.Image Image1 
         Height          =   495
         Left            =   120
         Picture         =   "frmStatementInter.frx":08CA
         Stretch         =   -1  'True
         Top             =   120
         Width           =   495
      End
   End
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   11445
      Top             =   1710
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   35
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":170C
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":23E6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":3238
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":408A
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":4964
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":523E
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":5B18
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":64E2
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":6DBC
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":70D6
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":79B0
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":828A
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":8B64
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":8E7E
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":9758
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":A032
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":A90C
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":B1E6
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":BAC0
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":C39A
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":CC74
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":D54E
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":DE28
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":E702
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":EFDC
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":F8B6
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":10190
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":10A6A
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":11344
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":11C1E
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":124D4
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":12DAE
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":13200
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":13652
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatementInter.frx":15E04
            Key             =   "IMG35"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtAirlineRateDollar 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7635
      TabIndex        =   45
      Text            =   "0.00"
      Top             =   5370
      Width           =   2160
   End
   Begin VB.TextBox txtGrandDollar 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7590
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "0.00"
      Top             =   8895
      Width           =   2205
   End
   Begin VB.TextBox txtGrandPeso 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9870
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "0.00"
      Top             =   8895
      Width           =   2175
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1485
      Left            =   60
      TabIndex        =   24
      Top             =   2955
      Width           =   12030
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFF80&
         Caption         =   "Select Which Exchange Rate"
         Height          =   705
         Left            =   6645
         TabIndex        =   37
         Top             =   705
         Width           =   5325
         Begin VB.OptionButton OptDollar 
            BackColor       =   &H00FFFF80&
            Caption         =   "Dollar Exchange"
            Height          =   315
            Left            =   3375
            TabIndex        =   11
            Top             =   225
            Width           =   1770
         End
         Begin VB.OptionButton OptPeso 
            BackColor       =   &H00FFFF80&
            Caption         =   "Peso Exchange"
            Height          =   315
            Left            =   1275
            TabIndex        =   10
            Top             =   255
            Value           =   -1  'True
            Width           =   1770
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2715
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   3720
      End
      Begin VB.TextBox txtRoute 
         Height          =   420
         Left            =   7815
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   255
         Width           =   4110
      End
      Begin VB.TextBox txtTicketNo 
         Height          =   420
         Left            =   2715
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   735
         Width           =   3705
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "ROUTE "
         Height          =   345
         Left            =   6630
         TabIndex        =   29
         Top             =   255
         Width           =   1065
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "TICKET NOS."
         Height          =   345
         Left            =   600
         TabIndex        =   28
         Top             =   765
         Width           =   1440
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "AIRLINE NAME"
         Height          =   345
         Left            =   600
         TabIndex        =   27
         Top             =   270
         Width           =   1440
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   60
      TabIndex        =   15
      Top             =   1350
      Width           =   12030
      Begin VB.TextBox txtCVno 
         Height          =   420
         Left            =   8160
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1080
         Width           =   3705
      End
      Begin VB.TextBox txtPOno 
         Height          =   420
         Left            =   8160
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   630
         Width           =   3705
      End
      Begin VB.TextBox txtDate 
         Enabled         =   0   'False
         Height          =   420
         Left            =   8145
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   180
         Width           =   3705
      End
      Begin VB.TextBox txtTelNo 
         Height          =   420
         Left            =   2670
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1065
         Width           =   3705
      End
      Begin VB.TextBox txtAddress 
         Height          =   420
         Left            =   2670
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   615
         Width           =   3705
      End
      Begin VB.TextBox txtName 
         Height          =   420
         Left            =   2655
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   165
         Width           =   3705
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CV NO."
         Height          =   345
         Left            =   5580
         TabIndex        =   21
         Top             =   1110
         Width           =   2340
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PO NO."
         Height          =   345
         Left            =   5595
         TabIndex        =   20
         Top             =   675
         Width           =   2340
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DATE :"
         Height          =   345
         Left            =   5565
         TabIndex        =   19
         Top             =   210
         Width           =   2340
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TEL NO."
         Height          =   345
         Left            =   300
         TabIndex        =   18
         Top             =   1140
         Width           =   2040
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS :"
         Height          =   345
         Left            =   300
         TabIndex        =   17
         Top             =   675
         Width           =   2040
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NAME :"
         Height          =   345
         Left            =   270
         TabIndex        =   16
         Top             =   225
         Width           =   2040
      End
   End
   Begin VB.TextBox txtStatementNo 
      Height          =   375
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "Text7"
      Top             =   1005
      Width           =   2145
   End
   Begin LVbuttons.LaVolpeButton cmdNew 
      Height          =   480
      Left            =   225
      TabIndex        =   0
      Top             =   9600
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
      MICON           =   "frmStatementInter.frx":17086
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
      Left            =   1590
      TabIndex        =   33
      Top             =   9615
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
      MICON           =   "frmStatementInter.frx":170A2
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
      Left            =   2955
      TabIndex        =   34
      Top             =   9615
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "Post"
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
      MICON           =   "frmStatementInter.frx":170BE
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
   Begin LVbuttons.LaVolpeButton cmdRecalc 
      Height          =   480
      Left            =   4320
      TabIndex        =   35
      Top             =   9615
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "ReCalc"
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
      MICON           =   "frmStatementInter.frx":170DA
      ALIGN           =   1
      IMGLST          =   "SmallImages"
      IMGICON         =   "10"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton cmdExit 
      Height          =   480
      Left            =   7245
      TabIndex        =   36
      Top             =   9615
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
      MICON           =   "frmStatementInter.frx":170F6
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
   Begin VB.TextBox txtAirlineRatePeso 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9870
      TabIndex        =   46
      Text            =   "0.00"
      Top             =   5370
      Width           =   2160
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   60
      TabIndex        =   23
      Top             =   4815
      Width           =   12030
      Begin VB.TextBox txtDescMisc 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2985
         TabIndex        =   63
         Text            =   "0.00"
         Top             =   2955
         Width           =   4485
      End
      Begin VB.TextBox txtDescTax 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2970
         TabIndex        =   62
         Text            =   "0.00"
         Top             =   1980
         Width           =   4485
      End
      Begin VB.TextBox txtDescVusa 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2970
         TabIndex        =   61
         Text            =   "0.00"
         Top             =   1485
         Width           =   4485
      End
      Begin VB.PictureBox Picture1 
         Enabled         =   0   'False
         Height          =   3420
         Left            =   7545
         ScaleHeight     =   3360
         ScaleWidth      =   2190
         TabIndex        =   47
         Top             =   465
         Width           =   2250
         Begin VB.TextBox txtFareDollar 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   0
            TabIndex        =   53
            Text            =   "0.00"
            Top             =   495
            Width           =   2160
         End
         Begin VB.TextBox txtVUSAdollar 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   0
            TabIndex        =   48
            Text            =   "0.00"
            Top             =   975
            Width           =   2160
         End
         Begin VB.TextBox txtUSHKtaxDollar 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   0
            TabIndex        =   52
            Text            =   "0.00"
            Top             =   1470
            Width           =   2160
         End
         Begin VB.TextBox txtDomesticTicketDollar 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   0
            TabIndex        =   51
            Text            =   "0.00"
            Top             =   1950
            Width           =   2160
         End
         Begin VB.TextBox txtMiscDollar 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   0
            TabIndex        =   50
            Text            =   "0.00"
            Top             =   2445
            Width           =   2160
         End
         Begin VB.TextBox txtPhilTaxDollar 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   0
            TabIndex        =   49
            Text            =   "0.00"
            Top             =   2940
            Width           =   2160
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   3420
         Left            =   9780
         ScaleHeight     =   3360
         ScaleWidth      =   2190
         TabIndex        =   38
         Top             =   465
         Width           =   2250
         Begin VB.TextBox txtFarePeso 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   0
            TabIndex        =   44
            Text            =   "0.00"
            Top             =   495
            Width           =   2160
         End
         Begin VB.TextBox txtVUSAPeso 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   0
            TabIndex        =   39
            Text            =   "0.00"
            Top             =   975
            Width           =   2160
         End
         Begin VB.TextBox txtUSHKtaxPeso 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   0
            TabIndex        =   43
            Text            =   "0.00"
            Top             =   1455
            Width           =   2160
         End
         Begin VB.TextBox txtDomesticTicketPeso 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   0
            TabIndex        =   42
            Text            =   "0.00"
            Top             =   1935
            Width           =   2160
         End
         Begin VB.TextBox txtMiscPeso 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   0
            TabIndex        =   41
            Text            =   "0.00"
            Top             =   2430
            Width           =   2160
         End
         Begin VB.TextBox txtPhilTaxPeso 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   0
            TabIndex        =   40
            Text            =   "0.00"
            Top             =   2925
            Width           =   2160
         End
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "================================================>"
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   2985
         TabIndex        =   67
         Top             =   3555
         Width           =   4545
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "================================================>"
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   2955
         TabIndex        =   66
         Top             =   2595
         Width           =   4545
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "================================================>"
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   2985
         TabIndex        =   65
         Top             =   1065
         Width           =   4545
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "================================================>"
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   2985
         TabIndex        =   64
         Top             =   600
         Width           =   4545
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "PHIL. TRAVEL TAX"
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   180
         TabIndex        =   60
         Top             =   3450
         Width           =   2340
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "MISC."
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   180
         TabIndex        =   59
         Top             =   3045
         Width           =   2340
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "DOMESTIC TICKET"
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   180
         TabIndex        =   58
         Top             =   2550
         Width           =   2340
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "TAX AND INSURANCE"
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   180
         TabIndex        =   57
         Top             =   2070
         Width           =   2340
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "FARE"
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   180
         TabIndex        =   56
         Top             =   1050
         Width           =   2340
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "AIRLINE RATE"
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   180
         TabIndex        =   55
         Top             =   600
         Width           =   2340
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "VUSA"
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   180
         TabIndex        =   54
         Top             =   1590
         Width           =   2340
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PESO EXCHANGE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   9795
         TabIndex        =   31
         Top             =   105
         Width           =   2220
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DOLLAR EXCHANGE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   7560
         TabIndex        =   30
         Top             =   105
         Width           =   2235
      End
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL =========================>"
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   3960
      TabIndex        =   32
      Top             =   8970
      Width           =   3330
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AMOUNT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   7620
      TabIndex        =   26
      Top             =   4470
      Width           =   4470
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PARTICULARS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   60
      TabIndex        =   25
      Top             =   4470
      Width           =   7545
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "STATEMENT NO."
      Height          =   345
      Left            =   -405
      TabIndex        =   22
      Top             =   1050
      Width           =   2340
   End
End
Attribute VB_Name = "frmStatementInter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As String
Private Sub cmdExit_Click()
Unload Me
End Sub

Public Sub ClearMe()
 Dim txt As Control
 For Each txt In frmStatementInter
  If TypeOf txt Is TextBox Then txt.Text = ""
 Next
End Sub

Private Sub cmdNew_Click()
Me.txtName.SetFocus
Me.txtDate = Format(Now, "mm/dd/yyyy")
txtStatementNo.Text = AutoIncrement(GetLastNumber)
            
          
            
            txtCvno.Text = "0"
            txtPOno.Text = "0"
            
            txtAirlineRatePeso.Text = "0.00"
            txtFarePeso.Text = "0.00"
            txtUSHKtaxPeso.Text = "0.00"
            txtDomesticTicketPeso.Text = "0.00"
            txtMiscPeso.Text = "0.00"
            txtPhilTaxPeso.Text = "0.00"
            txtVUSAPeso.Text = "0.00"
            txtGrandPeso.Text = "0.00"
            
            txtAirlineRateDollar.Text = "0.00"
            txtFareDollar.Text = "0.00"
            txtUSHKtaxDollar.Text = "0.00"
            txtDomesticTicketDollar.Text = "0.00"
            txtMiscDollar.Text = "0.00"
            txtPhilTaxDollar.Text = "0.00"
            txtVUSAdollar.Text = "0.00"
            txtGrandDollar.Text = "0.00"
            
End Sub

Private Sub cmdPost_Click()
On Error GoTo ErrExit
Dim ask As Integer
Dim Rs As ADODB.Recordset
Dim RsDetail As ADODB.Recordset
Dim SQL As String

If CheckNull(txtName) Then
    With Me.txtName
    MsgBox "Passenger Name should not be blank", vbInformation, "ELS"
    .SelStart = 0
    .SelLength = Len(.Text)
    .SetFocus
    End With
    Exit Sub
End If

If CheckNull(txtPOno) Then
    With Me.txtPOno
    MsgBox "PO number should not be blank", vbInformation, "ELS"
    .SelStart = 0
    .SelLength = Len(.Text)
    .SetFocus
    End With
    Exit Sub
End If

If CheckNull(txtCvno) Then
    With Me.txtCvno
    MsgBox "CV number should not be blank", vbInformation, "ELS"
    .SelStart = 0
    .SelLength = Len(.Text)
    .SetFocus
    End With
    Exit Sub
End If

If CheckNull(Combo1) Then
    With Me.Combo1
    MsgBox "Airline should not be blank", vbInformation, "ELS"
   
    .SetFocus
    End With
    Exit Sub
End If

If CheckNull(txtRoute) Then
    With Me.txtRoute
    MsgBox "Route should not be blank", vbInformation, "ELS"
    .SelStart = 0
    .SelLength = Len(.Text)
    .SetFocus
    End With
    Exit Sub
End If

If CheckNull(txtTicketNo) Then
    With Me.txtTicketNo
    MsgBox "Ticket No. should not be blank", vbInformation, "ELS"
    .SelStart = 0
    .SelLength = Len(.Text)
    .SetFocus
    End With
    Exit Sub
End If




ask = MsgBox("Are you sure you want to save this?", vbInformation + vbYesNo)
If ask = vbYes Then
    SQL = "SELECT * FROM tbl_StaInternational"
    Set Rs = New ADODB.Recordset
    Set RsDetail = New ADODB.Recordset
    With Rs
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            cn.BeginTrans
            .AddNew
            .Fields(1).Value = txtStatementNo.Text
            .Fields(2).Value = txtName.Text
            .Fields(3).Value = txtAddress.Text
            .Fields(4).Value = txtTelNo.Text
            .Fields(5).Value = txtDate.Text
            .Fields(6).Value = txtCvno.Text
            .Fields(7).Value = txtPOno.Text
           .Fields("Branch Number").Value = WhichBranch.Fields(2).Value
           .Fields("Total Amount").Value = CDbl(txtGrandPeso)
           .Fields("Total Amount Dollar").Value = CDbl(Me.txtGrandDollar)
           .Fields("Paid").Value = False
           .Fields("Void").Value = False
           .Fields("Refund").Value = False
           .Fields("AirlineID").Value = FindAirline(Me.Combo1)
           .Fields("Balance").Value = CDbl(txtGrandPeso)
           .Fields("Credit Card Activated").Value = False
            .Update
    End With
    
    SQL = "SELECT * FROM tbl_StaInternationalDetail"
    With RsDetail
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            .AddNew
            .Fields(1).Value = Rs.Fields(0).Value
            .Fields(2).Value = FindAirline(Me.Combo1)
            .Fields(3).Value = txtTicketNo.Text
            .Fields(4).Value = txtRoute.Text
            .Fields(5).Value = txtAirlineRatePeso.Text
            .Fields(6).Value = txtFarePeso.Text
            .Fields(7).Value = txtUSHKtaxPeso.Text
            .Fields(8).Value = txtDomesticTicketPeso.Text
            .Fields(9).Value = txtMiscPeso.Text
            .Fields(10).Value = txtPhilTaxPeso.Text
            .Fields(11).Value = Me.txtVUSAPeso
            .Fields(12).Value = txtGrandPeso.Text
            
            .Fields(13).Value = txtAirlineRateDollar.Text
            .Fields(14).Value = txtFareDollar.Text
            .Fields(15).Value = txtUSHKtaxDollar.Text
            .Fields(16).Value = txtDomesticTicketDollar.Text
            .Fields(17).Value = txtMiscDollar.Text
            .Fields(18).Value = txtPhilTaxDollar.Text
            .Fields(19).Value = Me.txtVUSAdollar
            .Fields(20).Value = txtGrandDollar.Text
            .Fields(21).Value = Me.txtDescVusa
            .Fields(22).Value = Me.txtDescTax
            .Fields(23).Value = Me.txtDescMisc
            .Update
            cn.CommitTrans
    End With
    MsgBox "Record Save..."
End If
Exit Sub
ErrExit:
cn.RollbackTrans

End Sub


Function GetLastNumber() As String
Dim RsFnumber As ADODB.Recordset
Dim SQL As String

Set RsFnumber = New ADODB.Recordset
SQL = "SELECT StatementNo from tbl_StaInternational"
RsFnumber.Open SQL, cn, adOpenKeyset, adLockOptimistic

With RsFnumber
        If .RecordCount > 0 Then
                .MoveLast
               GetLastNumber = RsFnumber("StatementNo").Value
        Else
               GetLastNumber = "SI-" & WhichBranch.Fields(2).Value & "-000000000"
               GetLastNumber = "SI" & WhichBranch.Fields(2).Value & Mid(Year(Now), 3, 2) & "-" & returnMon & returnMon & "00000"
        End If
End With

End Function

Sub FillCombo()
Dim Tmp As ADODB.Recordset
Set Tmp = New ADODB.Recordset
SQL = "SELECT * FROM tbl_Airline"
Me.Combo1.Clear
With Tmp
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
    If .RecordCount > 0 Then
        Do While Not .EOF
        Me.Combo1.AddItem .Fields(1).Value
        .MoveNext
        Loop
    End If
End With
End Sub

Function FindAirline(param) As Long
Dim Tmp As ADODB.Recordset
Set Tmp = New ADODB.Recordset
SQL = "SELECT * FROM tbl_Airline WHERE [AirlineName]='" & Me.Combo1 & "'"
With Tmp
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
    If .RecordCount > 0 Then
        FindAirline = .Fields(0).Value
      Else
        FindAirline = -1
    End If
    .Close
End With
Set Tmp = Nothing
End Function

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub Form_Load()
ClearMe
Call FillCombo
End Sub

Sub ReCalc()
With Me
        
        .txtGrandDollar = 0
        .txtGrandPeso = 0
        .txtGrandDollar = CDbl(.txtFareDollar) + CDbl(.txtUSHKtaxDollar) + CDbl(txtDomesticTicketDollar) + _
        CDbl(txtMiscDollar) + CDbl(txtPhilTaxDollar) + CDbl(txtVUSAdollar)
        .txtGrandDollar = Format(.txtGrandDollar, "###,##0.00")
        .txtGrandPeso = CDbl(txtFarePeso) + CDbl(txtUSHKtaxPeso) + CDbl(txtDomesticTicketPeso) + CDbl(txtMiscPeso) + _
        CDbl(txtPhilTaxPeso) + CDbl(txtVUSAPeso) + CDbl(txtGrandPeso)
        .txtGrandPeso = Format(.txtGrandPeso, "###,##0.00")
End With
End Sub



Private Sub OptDollar_Click()
With Me
            Me.Picture2.Enabled = False
            Me.Picture1.Enabled = True
           With txtAirlineRateDollar
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
            End With

    
End With

End Sub

Private Sub OptDollar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
With Me
   
            Me.Picture1.Enabled = False
            Me.Picture2.Enabled = True
            With txtAirlineRatePeso
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
            End With
End With
End If

End Sub

Private Sub OptPeso_Click()
With Me
   
            Me.Picture1.Enabled = False
            Me.Picture2.Enabled = True
            With txtAirlineRatePeso
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
            End With
End With
End Sub

Private Sub OptPeso_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
With Me
   
            Me.Picture1.Enabled = False
            Me.Picture2.Enabled = True
            With txtAirlineRateDollar
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
            End With
End With
End If
End Sub

Private Sub txtAddress_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter (KeyCode)
End Sub

Private Sub txtAirlineRateDollar_Change()
On Error GoTo ErrExit
If Not IsNumeric(txtAirlineRateDollar) Then
    With txtAirlineRateDollar
          .Text = "0.00"
          .SelStart = 0
          .SelLength = Len(.Text)
          .SetFocus
    End With
End If
Exit Sub
ErrExit:
End Sub

Private Sub txtAirlineRateDollar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
With Me.txtAirlineRatePeso
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
End With
End If
End Sub

Private Sub txtAirlineRatePeso_Change()
On Error GoTo ErrExit
If Not IsNumeric(txtAirlineRatePeso) Then
    With txtAirlineRatePeso
          .Text = "0.00"
          .SelStart = 0
          .SelLength = Len(.Text)
          .SetFocus
    End With
End If
Exit Sub
ErrExit:

End Sub

Private Sub txtAirlineRatePeso_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
       If Me.OptPeso Then
            With Me.txtFarePeso
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
            End With
       Else
            With Me.txtFareDollar
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
            End With
       
       End If
End If
End Sub


Private Sub txtAirlineRatePeso_LostFocus()
txtAirlineRatePeso = Format(txtAirlineRatePeso, "###,##0.00")
End Sub

Private Sub txtBillingNo_Change()
On Error GoTo ErrExit
If Not IsNumeric(txtBillingNo) Then
    With txtBillingNo
          .Text = "0"
          .SelStart = 0
          .SelLength = Len(.Text)
          .SetFocus
    End With
End If
Exit Sub
ErrExit:
End Sub

Private Sub txtBillingNo_GotFocus()
With txtBillingNo
          .Text = "0"
          .SelStart = 0
          .SelLength = Len(.Text)
          .SetFocus
    End With
End Sub

Private Sub txtBillingNo_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter (KeyCode)
End Sub

Private Sub txtCVno_Change()
On Error GoTo ErrExit
If Not IsNumeric(txtCvno) Then
    With txtCvno
          .Text = "0"
          .SelStart = 0
          .SelLength = Len(.Text)
          .SetFocus
    End With
End If
Exit Sub
ErrExit:
End Sub

Private Sub txtCVno_GotFocus()
With txtCvno
          .Text = "0"
          .SelStart = 0
          .SelLength = Len(.Text)
          .SetFocus
    End With
End Sub

Private Sub txtCVno_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter (KeyCode)
End Sub

Private Sub txtDomesticTicketDollar_Change()
On Error GoTo ErrExit
If Not IsNumeric(txtDomesticTicketDollar) Then
    With txtDomesticTicketDollar
          .Text = "0.00"
          .SelStart = 0
          .SelLength = Len(.Text)
          .SetFocus
    End With
Else
Call ReCompute
End If
Exit Sub
ErrExit:
End Sub

Private Sub txtDomesticTicketDollar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
       If Me.OptDollar Then
            With Me.txtMiscDollar
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
            End With
       End If
       
       If Me.OptPeso Then
            With Me.txtMiscPeso
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
            End With
       
       End If
End If

End Sub

Private Sub txtDomesticTicketDollar_LostFocus()
txtDomesticTicketDollar = Format(txtDomesticTicketDollar, "###,##0.00")
End Sub

Private Sub txtDomesticTicketPeso_Change()
On Error GoTo ErrExit
If Not IsNumeric(txtDomesticTicketPeso) Then
    With txtDomesticTicketPeso
          .Text = "0.00"
          .SelStart = 0
          .SelLength = Len(.Text)
          .SetFocus
    End With
Else
Call ReCompute
End If
Exit Sub
ErrExit:
End Sub

Private Sub txtDomesticTicketPeso_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
       If Me.OptDollar Then
            With Me.txtMiscDollar
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
            End With
       End If
       
       If Me.OptPeso Then
            With Me.txtMiscPeso
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
            End With
       
       End If
End If
End Sub

Private Sub txtDomesticTicketPeso_LostFocus()
txtDomesticTicketPeso = Format(txtDomesticTicketPeso, "###,##0.00")
End Sub

Private Sub txtFareDollar_Change()
On Error GoTo ErrExit
If Not IsNumeric(txtFareDollar) Then
    With txtFareDollar
          .Text = "0.00"
          .SelStart = 0
          .SelLength = Len(.Text)
          .SetFocus
    End With
Else
Call ReCompute
End If
Exit Sub
ErrExit:

End Sub

Private Sub txtFareDollar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
       If Me.OptDollar Then
            With Me.txtVUSAdollar
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
            End With
       End If
       
       If Me.OptPeso Then
            With Me.txtVUSAPeso
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
            End With
       
       End If
End If
End Sub

Private Sub txtFareDollar_LostFocus()
txtFareDollar = Format(txtFareDollar, "###,##0.00")
End Sub

Private Sub txtFarePeso_Change()
On Error GoTo ErrExit
If Not IsNumeric(txtFarePeso) Then
    With txtFarePeso
          .Text = "0.00"
          .SelStart = 0
          .SelLength = Len(.Text)
          .SetFocus
    End With
Else
Call ReCompute
End If
Exit Sub
ErrExit:
End Sub

Private Sub txtFarePeso_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
       If Me.OptDollar Then
            With Me.txtVUSAdollar
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
            End With
       End If
       
       If Me.OptPeso Then
            With Me.txtVUSAPeso
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
            End With
       
       End If
End If
End Sub

Private Sub txtFarePeso_LostFocus()
txtFarePeso = Format(txtFarePeso, "###,##0.00")
End Sub


Private Sub txtGrandDollar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then

End If
End Sub

Private Sub txtInsuranceDollar_Change()
On Error GoTo ErrExit
If Not IsNumeric(txtInsuranceDollar) Then
    With txtInsuranceDollar
          .Text = "0.00"
          .SelStart = 0
          .SelLength = Len(.Text)
          .SetFocus
    End With
Else
Call ReCompute
End If
Exit Sub
ErrExit:
End Sub

Private Sub txtInsuranceDollar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call cmdPost_Click
End If
End Sub

Private Sub txtInsuranceDollar_LostFocus()
txtInsuranceDollar = Format(txtInsuranceDollar, "###,##0.00")
End Sub

Private Sub txtInsurancePeso_Change()
On Error GoTo ErrExit
If Not IsNumeric(txtInsurancePeso) Then
    With txtInsurancePeso
          .Text = "0.00"
          .SelStart = 0
          .SelLength = Len(.Text)
          .SetFocus
    End With
Else
Call ReCompute
End If
Exit Sub
ErrExit:
End Sub


Private Sub txtVUSAdollar_Change()
On Error GoTo ErrExit
If Not IsNumeric(txtVUSAdollar) Then
    With txtVUSAdollar
          .Text = "0.00"
          .SelStart = 0
          .SelLength = Len(.Text)
          .SetFocus
    End With
Else
Call ReCompute
End If
Exit Sub
ErrExit:
End Sub

Private Sub txtVUSAdollar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
       If Me.OptDollar Then
            With Me.txtUSHKtaxDollar
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
            End With
       End If
       
       If Me.OptPeso Then
            With Me.txtUSHKtaxPeso
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
            End With
       
       End If
End If

End Sub

Private Sub txtVUSAdollar_LostFocus()
txtVUSAdollar = Format(txtVUSAdollar, "###,##0.00")
End Sub

Private Sub txtMiscDollar_Change()
On Error GoTo ErrExit
If Not IsNumeric(txtMiscDollar) Then
    With txtMiscDollar
          .Text = "0.00"
          .SelStart = 0
          .SelLength = Len(.Text)
          .SetFocus
    End With
Else
Call ReCompute
End If
Exit Sub
ErrExit:
End Sub

Private Sub txtMiscDollar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
       If Me.OptDollar Then
            With Me.txtPhilTaxDollar
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
            End With
       End If
       If Me.OptPeso Then
            With Me.txtPhilTaxPeso
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
            End With
       End If
End If
End Sub



Private Sub txtMiscDollar_LostFocus()
txtMiscDollar = Format(txtMiscDollar, "###,##0.00")
End Sub

Private Sub txtMiscPeso_Change()
On Error GoTo ErrExit
If Not IsNumeric(txtMiscPeso) Then
    With txtMiscPeso
          .Text = "0.00"
          .SelStart = 0
          .SelLength = Len(.Text)
          .SetFocus
    End With
Else
Call ReCompute
End If
Exit Sub
ErrExit:
End Sub

Private Sub txtMiscPeso_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
       If Me.OptDollar Then
            With Me.txtPhilTaxDollar
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
            End With
       End If
       If Me.OptPeso Then
            With Me.txtPhilTaxPeso
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
            End With
       End If
End If
End Sub




Private Sub txtMiscPeso_LostFocus()
txtMiscPeso = Format(txtMiscPeso, "###,##0.00")
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter (KeyCode)
End Sub

Private Sub txtPhilTaxDollar_Change()
On Error GoTo ErrExit
If Not IsNumeric(txtPhilTaxDollar) Then
    With txtPhilTaxDollar
          .Text = "0.00"
          .SelStart = 0
          .SelLength = Len(.Text)
          .SetFocus
    End With
Else
Call ReCompute
End If
Exit Sub
ErrExit:
End Sub

Private Sub txtPhilTaxDollar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  Call cmdPost_Click
End If

End Sub

Private Sub txtPhilTaxDollar_LostFocus()
txtPhilTaxDollar = Format(txtPhilTaxDollar, "###,##0.00")
End Sub

Private Sub txtPhilTaxPeso_Change()
On Error GoTo ErrExit
If Not IsNumeric(txtPhilTaxPeso) Then
    With txtPhilTaxPeso
          .Text = "0.00"
          .SelStart = 0
          .SelLength = Len(.Text)
          .SetFocus
    End With
Else
Call ReCompute
End If
Exit Sub
ErrExit:
End Sub

Private Sub txtPhilTaxPeso_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  Call cmdPost_Click
End If

End Sub

Private Sub txtPhilTaxPeso_LostFocus()
txtPhilTaxPeso = Format(txtPhilTaxPeso, "###,##0.00")
End Sub

Private Sub txtPOno_Change()
On Error GoTo ErrExit
If Not IsNumeric(txtPOno) Then
    With txtPOno
          .Text = "0"
          .SelStart = 0
          .SelLength = Len(.Text)
          .SetFocus
    End With
End If
Exit Sub
ErrExit:

End Sub

Private Sub txtPOno_GotFocus()
With txtPOno
          .Text = "0"
          .SelStart = 0
          .SelLength = Len(.Text)
          .SetFocus
    End With
End Sub

Private Sub txtPOno_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter (KeyCode)
End Sub

Private Sub txtRoute_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtTelno_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter (KeyCode)
End Sub

Private Sub txtTicketNo_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtUSHKtaxDollar_Change()
On Error GoTo ErrExit
If Not IsNumeric(txtUSHKtaxDollar) Then
    With txtUSHKtaxDollar
          .Text = "0.00"
          .SelStart = 0
          .SelLength = Len(.Text)
          .SetFocus
    End With
Else
Call ReCompute
End If
Exit Sub
ErrExit:
End Sub

Private Sub txtUSHKtaxDollar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
       If Me.OptDollar Then
            With Me.txtDomesticTicketDollar
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
            End With
       End If
       
       If Me.OptPeso Then
            With Me.txtDomesticTicketPeso
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
            End With
       
       End If
End If

End Sub

Private Sub txtUSHKtaxDollar_LostFocus()
txtUSHKtaxDollar = Format(txtUSHKtaxDollar, "###,##0.00")
End Sub

Private Sub txtUSHKtaxPeso_Change()
On Error GoTo ErrExit
If Not IsNumeric(txtUSHKtaxPeso) Then
    With txtUSHKtaxPeso
          .Text = "0.00"
          .SelStart = 0
          .SelLength = Len(.Text)
          .SetFocus
    End With
Else
Call ReCompute
End If
Exit Sub
ErrExit:
End Sub

Sub ReCompute()
On Error Resume Next
With Me
    If Me.OptDollar Then
                .txtFarePeso = Format(CDbl(.txtFareDollar) * CDbl(.txtAirlineRatePeso), "###,##0.00")
                .txtUSHKtaxPeso = Format(CDbl(.txtUSHKtaxDollar) * CDbl(.txtAirlineRatePeso), "###,##0.00")
                .txtDomesticTicketPeso = Format(CDbl(.txtDomesticTicketDollar) * CDbl(.txtAirlineRatePeso), "###,##0.00")
                .txtMiscPeso = Format(CDbl(.txtMiscDollar) * CDbl(.txtAirlineRatePeso), "###,##0.00")
                .txtPhilTaxPeso = Format(CDbl(.txtPhilTaxDollar) * CDbl(.txtAirlineRatePeso), "###,##0.00")
                .txtVUSAPeso = Format(CDbl(.txtVUSAdollar) * CDbl(.txtAirlineRatePeso), "###,##0.00")
    End If
    If Me.OptPeso Then
                .txtFareDollar = Format(CDbl(.txtFarePeso) / CDbl(.txtAirlineRatePeso), "###,##0.00")
                .txtUSHKtaxDollar = Format(CDbl(.txtUSHKtaxPeso) / CDbl(.txtAirlineRatePeso), "###,##0.00")
                .txtDomesticTicketDollar = Format(CDbl(.txtDomesticTicketPeso) / CDbl(.txtAirlineRatePeso), "###,##0.00")
                .txtMiscDollar = Format(CDbl(.txtMiscPeso) / CDbl(.txtAirlineRatePeso), "###,##0.00")
                .txtPhilTaxDollar = Format(CDbl(.txtPhilTaxPeso) / CDbl(.txtAirlineRatePeso), "###,##0.00")
                .txtVUSAdollar = Format(CDbl(.txtVUSAPeso) / CDbl(.txtAirlineRatePeso), "###,##0.00")
    End If
End With
Call ReCalc
End Sub

Private Sub txtUSHKtaxPeso_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
       If Me.OptDollar Then
            With Me.txtDomesticTicketDollar
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
            End With
       End If
       
       If Me.OptPeso Then
            With Me.txtDomesticTicketPeso
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
            End With
       
       End If
End If
End Sub

Private Sub txtUSHKtaxPeso_LostFocus()
txtUSHKtaxPeso = Format(txtUSHKtaxPeso, "###,##0.00")
End Sub

Private Sub txtVUSAPeso_Change()
On Error GoTo ErrExit
If Not IsNumeric(txtVUSAPeso) Then
    With txtVUSAPeso
          .Text = "0.00"
          .SelStart = 0
          .SelLength = Len(.Text)
          .SetFocus
    End With
Else
Call ReCompute
End If
Exit Sub
ErrExit:
End Sub

Private Sub txtVUSAPeso_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
       If Me.OptDollar Then
            With Me.txtUSHKtaxDollar
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
            End With
       End If
       
       If Me.OptPeso Then
            With Me.txtUSHKtaxPeso
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
            End With
       
       End If
End If
End Sub

Private Sub txtVUSAPeso_LostFocus()
txtVUSAPeso = Format(txtVUSAPeso, "###,##0.00")
End Sub
