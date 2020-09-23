VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "Lvbuttons.ocx"
Begin VB.Form frmSetTicketPricing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Route Pricing"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14730
   Icon            =   "frmSetTicketPrice.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9210
   ScaleWidth      =   14730
   Begin VB.PictureBox Picture3 
      Height          =   495
      Left            =   45
      ScaleHeight     =   435
      ScaleWidth      =   3300
      TabIndex        =   50
      Top             =   8655
      Width           =   3360
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "RECORD :"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   60
         TabIndex        =   52
         Top             =   75
         Width           =   900
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "stattus"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   990
         TabIndex        =   51
         Top             =   75
         Width           =   900
      End
   End
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   7515
      Top             =   300
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
            Picture         =   "frmSetTicketPrice.frx":1272
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":1F4C
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":2D9E
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":3BF0
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":44CA
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":4DA4
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":567E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":6048
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":6922
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":6C3C
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":7516
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":7DF0
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":86CA
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":89E4
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":92BE
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":9B98
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":A472
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":AD4C
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":B626
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":BF00
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":C7DA
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":D0B4
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":D98E
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":E268
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":EB42
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":F41C
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":FCF6
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":105D0
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":10EAA
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":11784
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":1203A
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":12914
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":12D66
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":131B8
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetTicketPrice.frx":1596A
            Key             =   "IMG35"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Caption         =   "Current Selection"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2385
      Left            =   8250
      TabIndex        =   43
      Top             =   1110
      Width           =   6390
      Begin VB.TextBox txtShipLines 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   405
         Width           =   4275
      End
      Begin VB.TextBox txtTicketType 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   "Text1"
         Top             =   1005
         Width           =   4275
      End
      Begin VB.TextBox txtRoute 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   44
         Text            =   "Text1"
         Top             =   1590
         Width           =   4275
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Shipping / Airline"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   49
         Top             =   450
         Width           =   1890
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Ticket Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   48
         Top             =   1095
         Width           =   1890
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Route"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   47
         Top             =   1680
         Width           =   1890
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Set Ticket Type"
      Height          =   375
      Left            =   6825
      TabIndex        =   42
      Top             =   3105
      Width           =   1350
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   3570
      Width           =   3360
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00400040&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   14700
      TabIndex        =   32
      Top             =   -75
      Width           =   14760
      Begin VB.Image Image1 
         Height          =   495
         Left            =   120
         Picture         =   "frmSetTicketPrice.frx":16BEC
         Stretch         =   -1  'True
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Note:  Set ticket Routes and Price"
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
         TabIndex        =   34
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Set Ticket Route(s) and Pricing"
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
         TabIndex        =   33
         Top             =   120
         Width           =   9120
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Current Selection"
      Height          =   5610
      Left            =   3450
      TabIndex        =   15
      Top             =   3555
      Width           =   11160
      Begin VB.CommandButton Command1 
         Caption         =   "Edit Price"
         Height          =   405
         Left            =   9960
         TabIndex        =   41
         Top             =   4080
         Width           =   1140
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   405
         Left            =   8835
         TabIndex        =   40
         Top             =   4080
         Width           =   1140
      End
      Begin MSDataGridLib.DataGrid DataGrid4 
         Height          =   3420
         Left            =   6780
         TabIndex        =   39
         Top             =   630
         Width           =   4305
         _ExtentX        =   7594
         _ExtentY        =   6033
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "FareBasisID"
            Caption         =   "FareBasisID"
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
            DataField       =   "FareBasis"
            Caption         =   "FareBasis"
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
            DataField       =   "FareBasisAmount"
            Caption         =   "FareBasisAmount"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            RecordSelectors =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2745.071
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1560.189
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Insert"
         Height          =   405
         Left            =   7710
         TabIndex        =   38
         Top             =   4080
         Width           =   1140
      End
      Begin VB.CommandButton cmdNewFB 
         Caption         =   "Set Fare Basis"
         Height          =   405
         Left            =   9150
         TabIndex        =   37
         Top             =   165
         Width           =   1935
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   780
         Left            =   90
         TabIndex        =   22
         Top             =   4695
         Width           =   10995
         Begin LVbuttons.LaVolpeButton cmdAddSave 
            Height          =   480
            Left            =   105
            TabIndex        =   10
            Top             =   180
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   847
            BTYPE           =   3
            TX              =   "Add"
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
            MICON           =   "frmSetTicketPrice.frx":174B6
            ALIGN           =   1
            IMGLST          =   "SmallImages"
            IMGICON         =   "13"
            ICONAlign       =   0
            ORIENT          =   0
            STYLE           =   0
            IconSize        =   2
            SHOWF           =   -1  'True
            BSTYLE          =   0
         End
         Begin LVbuttons.LaVolpeButton cmdExit 
            Height          =   480
            Left            =   9435
            TabIndex        =   11
            Top             =   165
            Width           =   1485
            _ExtentX        =   2619
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
            MICON           =   "frmSetTicketPrice.frx":174D2
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
         Begin LVbuttons.LaVolpeButton LaVolpeButton1 
            Height          =   480
            Left            =   1605
            TabIndex        =   23
            Top             =   180
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   847
            BTYPE           =   3
            TX              =   "View All"
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
            MICON           =   "frmSetTicketPrice.frx":174EE
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
         Begin LVbuttons.LaVolpeButton cmdEvat 
            Height          =   480
            Left            =   3105
            TabIndex        =   53
            Top             =   180
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   847
            BTYPE           =   3
            TX              =   "Set EVAT"
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
            MICON           =   "frmSetTicketPrice.frx":1750A
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
         Begin LVbuttons.LaVolpeButton cmdSetComm 
            Height          =   480
            Left            =   4605
            TabIndex        =   54
            Top             =   180
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   847
            BTYPE           =   3
            TX              =   "Set COMM %"
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
            MICON           =   "frmSetTicketPrice.frx":17526
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
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   4095
         Left            =   180
         ScaleHeight     =   4035
         ScaleWidth      =   6495
         TabIndex        =   16
         Top             =   615
         Width           =   6555
         Begin VB.Frame Frame4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Important Fill this for Refund"
            ForeColor       =   &H80000008&
            Height          =   3135
            Left            =   3405
            TabIndex        =   56
            Top             =   45
            Width           =   3015
            Begin VB.TextBox txtNoShowFee 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   1380
               TabIndex        =   62
               Text            =   "0.00"
               Top             =   1455
               Width           =   1470
            End
            Begin VB.TextBox txtVoidFee 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   1380
               TabIndex        =   61
               Text            =   "0.00"
               Top             =   1110
               Width           =   1470
            End
            Begin VB.TextBox txtRefund 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   1380
               TabIndex        =   59
               Text            =   "0.00"
               Top             =   735
               Width           =   1470
            End
            Begin VB.TextBox txtServiceFee 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   1380
               TabIndex        =   58
               Text            =   "0.00"
               Top             =   360
               Width           =   1470
            End
            Begin VB.Label Label24 
               BackStyle       =   0  'Transparent
               Caption         =   "NoShow Fee :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   285
               Left            =   135
               TabIndex        =   64
               Top             =   1470
               Width           =   1260
            End
            Begin VB.Label Label23 
               BackStyle       =   0  'Transparent
               Caption         =   "Void Fee :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   285
               Left            =   135
               TabIndex        =   63
               Top             =   1095
               Width           =   1260
            End
            Begin VB.Label Label22 
               BackStyle       =   0  'Transparent
               Caption         =   "Refund Fee:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   285
               Left            =   120
               TabIndex        =   60
               Top             =   705
               Width           =   1260
            End
            Begin VB.Label Label18 
               BackStyle       =   0  'Transparent
               Caption         =   "Service Fee:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   285
               Left            =   135
               TabIndex        =   57
               Top             =   345
               Width           =   1260
            End
         End
         Begin VB.TextBox txtMisc 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   1875
            TabIndex        =   55
            Text            =   "0.00"
            Top             =   2355
            Width           =   1455
         End
         Begin VB.TextBox txtEvat 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   2445
            TabIndex        =   29
            Text            =   "0.00"
            Top             =   2850
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.TextBox txtVAT 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   1545
            TabIndex        =   9
            Text            =   "0.00"
            Top             =   2865
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.TextBox txtMeals 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   1875
            TabIndex        =   8
            Text            =   "0.00"
            Top             =   2010
            Width           =   1470
         End
         Begin VB.TextBox txtTerminalFee 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   345
            Left            =   1875
            TabIndex        =   7
            Text            =   "0.00"
            Top             =   1650
            Width           =   1470
         End
         Begin VB.TextBox txtNetFare 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   4320
            Locked          =   -1  'True
            TabIndex        =   27
            Text            =   "0.00"
            Top             =   3570
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.TextBox txtASF 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   345
            Left            =   1875
            TabIndex        =   6
            Text            =   "0.00"
            Top             =   1275
            Width           =   1470
         End
         Begin VB.TextBox txtCommision 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   345
            Left            =   2535
            TabIndex        =   4
            Text            =   "0"
            Top             =   510
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.TextBox txtInsurance 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   345
            Left            =   1875
            TabIndex        =   5
            Text            =   "0.00"
            Top             =   885
            Width           =   1470
         End
         Begin VB.TextBox txtGrossFare 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   345
            Left            =   1875
            TabIndex        =   3
            Text            =   "0.00"
            Top             =   135
            Visible         =   0   'False
            Width           =   1470
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "E-VAT %  :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   135
            TabIndex        =   31
            Top             =   2880
            Visible         =   0   'False
            Width           =   1890
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   2850
            TabIndex        =   30
            Top             =   2880
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "MISC :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   135
            TabIndex        =   28
            Top             =   2415
            Width           =   1890
         End
         Begin VB.Line Line1 
            BorderWidth     =   3
            X1              =   6345
            X2              =   1455
            Y1              =   3420
            Y2              =   3420
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Meals :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   135
            TabIndex        =   26
            Top             =   2055
            Width           =   1890
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   2190
            TabIndex        =   25
            Top             =   540
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Terminal Fee :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   135
            TabIndex        =   24
            Top             =   1695
            Width           =   1890
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Net Fare :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3405
            TabIndex        =   21
            Top             =   3645
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "ASF :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   150
            TabIndex        =   20
            Top             =   1350
            Width           =   1890
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Commision :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   150
            TabIndex        =   19
            Top             =   555
            Visible         =   0   'False
            Width           =   1890
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Insurance :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   135
            TabIndex        =   18
            Top             =   915
            Width           =   1890
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Gross Fare"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   120
            TabIndex        =   17
            Top             =   180
            Visible         =   0   'False
            Width           =   1890
         End
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   1830
      Left            =   135
      TabIndex        =   0
      Top             =   1215
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   3228
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483633
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
      Caption         =   "Select Shipping / Airline Below"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "AirlineID"
         Caption         =   "AirlineID"
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
         DataField       =   "AirlineName"
         Caption         =   "Ship / Airline Name "
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
         MarqueeStyle    =   3
         ScrollBars      =   2
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         RecordSelectors =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3254.74
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4665
      Left            =   60
      TabIndex        =   1
      Top             =   3930
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   8229
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483633
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "RouteID"
         Caption         =   "RouteID"
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
         DataField       =   "From"
         Caption         =   "From"
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
         DataField       =   "To"
         Caption         =   "To"
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
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         RecordSelectors =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column02 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   1830
      Left            =   3465
      TabIndex        =   2
      Top             =   1215
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   3228
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483633
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "TicketTypeID"
         Caption         =   "TicketTypeID"
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
         DataField       =   "Ticket Type"
         Caption         =   "Ticket Type"
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
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         RecordSelectors =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4500.284
         EndProperty
      EndProperty
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Route :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   75
      TabIndex        =   36
      Top             =   3240
      Width           =   1890
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Ticket Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   3480
      TabIndex        =   14
      Top             =   960
      Width           =   4635
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Available Routes"
      Height          =   285
      Left            =   180
      TabIndex        =   13
      Top             =   6315
      Width           =   1890
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Shipping / Airline "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   165
      TabIndex        =   12
      Top             =   960
      Width           =   3105
   End
End
Attribute VB_Name = "frmSetTicketPricing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As New ADODB.Recordset
Dim RsRoute As New ADODB.Recordset
Dim RsTicketType As New ADODB.Recordset
Dim RsSetTicketRoute As New ADODB.Recordset
Dim RstFareBasis As New ADODB.Recordset
Dim SQL As String
Dim OldBookMark
Dim Temp As String

Private Sub cmdAddSave_Click()
'On Error GoTo FailSafe_Error
Dim i As Integer

If DupPricing Then
    MsgBox "The pricing of this route is already added! Click view if you want to change pricing of this!", vbInformation
Exit Sub
End If

    Set RstFareBasis = New ADODB.Recordset
    SQL = "SELECT * FROM tbl_TmpFareBasis"
    RstFareBasis.Open SQL, cn, adOpenKeyset, adLockOptimistic
    
    If RstFareBasis.RecordCount = 0 Then
            MsgBox "There are no fare basis selected...please insert fare basis"
            RstFareBasis.Close
            Exit Sub
    End If


If Me.cmdAddSave.Caption = "Add" Then
    Me.cmdAddSave.Caption = "Save"
    
    
    Me.txtInsurance.Enabled = True
    Me.txtCommision.Enabled = True
    Me.txtASF.Enabled = True
    Me.txtNetFare.Enabled = True
    Me.txtTerminalFee.Enabled = True
   Me.txtInsurance.SetFocus
Else
    If CheckNull(Me.txtGrossFare) Then
        MsgBox "The Airline/Shipping Line Should Not be Blank"
        Exit Sub
    End If
    
           
    
    
    SQL = "SELECT * FROM tbl_RoutePricing"
    With RsSetTicketRoute
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
    
    For i = 1 To CountTempFB
       Me.txtGrossFare = RstFareBasis.Fields(2).Value
            .AddNew
                .Fields(1).Value = CDbl(Me.DataGrid1.Columns(0).Text)
                .Fields(2).Value = CDbl(Me.DataGrid2.Columns(0).Text)
                .Fields(3).Value = CDbl(Me.DataGrid3.Columns(0).Text)
                .Fields(4).Value = CDbl(Me.txtGrossFare)
                .Fields(5).Value = CDbl(Me.txtInsurance)
                .Fields(6).Value = CDbl(Me.txtCommision)
                .Fields(7).Value = CDbl(Me.txtASF)
                .Fields(8).Value = CDbl(Me.txtTerminalFee)
                .Fields(9).Value = CDbl(Me.txtNetFare)
                .Fields(10).Value = CDbl(Me.txtMeals)
                .Fields(11).Value = Format((CDbl(Me.txtVat) / 100) * CDbl(Me.txtGrossFare), "###,##0.00")
                .Fields("FareBasisID").Value = RstFareBasis.Fields(0).Value
                .Fields("FareBasis").Value = RstFareBasis.Fields(1).Value
                .Fields("Misc").Value = CDbl(Me.txtMisc)
                .Fields("Service Fee").Value = CDbl(Me.txtServiceFee)
                .Fields("Refund Fee").Value = CDbl(Me.txtRefund)
                .Fields("Void Fee").Value = CDbl(Me.txtVoidFee)
                .Fields("Noshow Fee").Value = CDbl(Me.txtNoShowFee)

            .Update
            
            If FindRouteID >= 0 Then
            .AddNew
                .Fields(1).Value = FindRouteID()
                .Fields(2).Value = CDbl(Me.DataGrid2.Columns(0).Text)
                .Fields(3).Value = CDbl(Me.DataGrid3.Columns(0).Text)
                .Fields(4).Value = CDbl(Me.txtGrossFare)
                .Fields(5).Value = CDbl(Me.txtInsurance)
                .Fields(6).Value = CDbl(Me.txtCommision)
                .Fields(7).Value = CDbl(Me.txtASF)
                .Fields(8).Value = CDbl(Me.txtTerminalFee)
                .Fields(9).Value = CDbl(Me.txtNetFare)
                .Fields(10).Value = CDbl(Me.txtMeals)
                .Fields(11).Value = Format((CDbl(Me.txtVat) / 100) * CDbl(Me.txtGrossFare), "###,##0.00")
                .Fields("FareBasisID").Value = RstFareBasis.Fields(0).Value
                .Fields("FareBasis").Value = RstFareBasis.Fields(1).Value
                .Fields("Misc").Value = CDbl(Me.txtMisc)
                .Fields("Service Fee").Value = CDbl(Me.txtServiceFee)
                .Fields("Refund Fee").Value = CDbl(Me.txtRefund)
                .Fields("Void Fee").Value = CDbl(Me.txtVoidFee)
                .Fields("Noshow Fee").Value = CDbl(Me.txtNoShowFee)
                
            .Update
            End If
                RstFareBasis.MoveNext
    Next i

            .Close
      Set RsSetTicketRoute = Nothing
      MsgBox "Record Save", vbInformation
    End With
    
    Me.cmdAddSave.Caption = "Add"
    'Me.txtGrossFare.Enabled = True
    Me.txtInsurance.Enabled = True
    Me.txtCommision.Enabled = True
    Me.txtASF.Enabled = True
    Me.txtNetFare.Enabled = True
    Me.txtTerminalFee.Enabled = True
  
End If
Exit Sub
FailSafe_Error:
MsgBox "There were errors while saving... " & Err.Description
End Sub

Function FindRouteID() As Long
Dim RsFindRt As New ADODB.Recordset
Dim SQL As String
Dim StrCriteria(1 To 2) As String

StrCriteria(1) = Me.DataGrid1.Columns(2).Text
StrCriteria(2) = Me.DataGrid1.Columns(1).Text

SQL = "SELECT * FROM tbl_Routes WHERE [FROM]= '" & _
     StrCriteria(1) & "' AND [TO]='" & StrCriteria(2) & "'"
    
With RsFindRt
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            FindRouteID = RsFindRt.Fields(0).Value
        Else
            FindRouteID = -1
        End If
End With

End Function

Private Sub cmdAddSave_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub cmdDelete_Click()
On Error Resume Next
Dim ask As Integer

ask = MsgBox("Are you sure you want to remove this Airline/Shipping Line?", vbCritical + vbYesNo)
If ask = vbYes Then
SQL = "DELETE * FROM tbl_Airline WHERE [AirlineID]=" & Me.DataGrid1.Columns(0).Text
cn.Execute SQL
OldBookMark = Rs.Bookmark
Rs.Requery
MsgBox "One Airline / Shipping Line deleted"
Rs.Bookmark = OldBookMark
End If
End Sub

Private Sub cmdEvat_Click()
    frmSetEVAT.Show 1
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNewFB_Click()
    frmFareBasis.Show 1
End Sub

Private Sub cmdRemove_Click()
On Error GoTo FailSafe
Dim ask As Integer
SQL = "DELETE * FROM tbl_TmpFareBasis WHERE [FareBasisID]=" & Me.DataGrid4.Columns(0).Text
ask = MsgBox("Sure you want to delete this?", vbCritical + vbYesNo, "ELS")
If ask = vbYes Then
    cn.Execute SQL
     Call DisplayFB
    MsgBox "One record deleted", vbInformation
   
End If
Exit Sub
FailSafe:
MsgBox "There was an error in deleting fare basis..."
End Sub

Private Sub cmdSelect_Click()
frmFareBasisSelect.Show 1
End Sub

Private Sub cmdSetComm_Click()
If MDImain.StatusBar1.Panels(3).Text <> "admin" Then
        MsgBox "You have insufficient rights to update the commission", vbCritical, "Warning!!!!"
        Exit Sub
End If

frmSetCOMM.Show 1
End Sub

Private Sub Combo1_Click()
Call FilterRoute
End Sub

Private Sub Command1_Click()
On Error GoTo FailSafe_Error
    
    frmFareBasisTmp.Tag = Me.DataGrid4.Columns(0).Text
    frmFareBasisTmp.Text1 = Me.DataGrid4.Columns(1).Text
    frmFareBasisTmp.Text2 = Me.DataGrid4.Columns(2).Text
    frmFareBasisTmp.Show 1
Exit Sub
FailSafe_Error:
End Sub

Private Sub Command2_Click()
frmAddTicketType.Show 1
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call DisplayCurrentSel
Call DisplayStat
End Sub

Private Sub DataGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call DisplayCurrentSel
Temp = Rs.Fields(2).Value
Call DisplayTicketType(Temp)
End Sub

Private Sub DataGrid3_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call DisplayCurrentSel
End Sub

Private Sub Form_Load()
On Error GoTo FailSafeErr
Set Rs = New ADODB.Recordset
SQL = "SELECT  * FROM tbl_Airline WHERE [AirlineName]<>'NONE' ORDER by [AirlineName] ASC "
Rs.Open SQL, cn, adOpenKeyset, adLockOptimistic
Set Me.DataGrid2.DataSource = Rs


Temp = Rs.Fields(2).Value
Call DisplayTicketType(Temp)

Call FillRoutes
Me.Combo1.ListIndex = 0

Call FilterRoute
Call DisplayCurrentSel
Call DisplayStat
Call KillTemp

FailSafeErr:
End Sub

Function CountTempFB() As Long
Dim RstTmpFareBasis As New ADODB.Recordset
SQL = "SELECT * FROM tbl_TmpFareBasis"
With RstTmpFareBasis
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
    If .RecordCount > 0 Then
        CountTempFB = .RecordCount
    Else
        CountTempFB = 0
    End If
    .Close
 Set RstTmpFareBasis = Nothing
End With

End Function

Sub KillTemp()
On Error GoTo FailSafeErr
Dim Rst As New ADODB.Recordset
SQL = "DELETE * FROM tbl_TmpFareBasis"
cn.BeginTrans
    cn.Execute SQL
cn.CommitTrans
FailSafeErr:
cn.RollbackTrans
End Sub
Sub FilterRoute()
If RsRoute.State = 1 Then
    RsRoute.Close
    Set RsRoute = Nothing
End If
Set RsRoute = New ADODB.Recordset
SQL = "SELECT  * FROM tbl_Routes WHERE [From]='" & Me.Combo1 & "' ORDER BY [From] ASC"
RsRoute.Open SQL, cn, adOpenKeyset, adLockOptimistic
Set Me.DataGrid1.DataSource = RsRoute
End Sub

Sub DisplayTicketType(ByVal Param As String)

If RsTicketType.State = 1 Then
    RsTicketType.Close
    Set RsTicketType = Nothing
End If
Set RsTicketType = New ADODB.Recordset
SQL = "SELECT  * FROM tbl_TicketType WHERE [AirlineShippingLine]='" & UCase(Param) & "' ORDER BY [Ticket Type] ASC"
RsTicketType.Open SQL, cn, adOpenKeyset, adLockOptimistic
Set Me.DataGrid3.DataSource = RsTicketType
End Sub

Sub DisplayCurrentSel()
On Error GoTo ErrExit
With Me
            .txtShipLines = Me.DataGrid2.Columns(1).Text
            .txtTicketType = Me.DataGrid3.Columns(1).Text
            .txtRoute = Me.DataGrid1.Columns(1) & " - " & Me.DataGrid1.Columns(2)
End With
Exit Sub
ErrExit:
End Sub

Sub DisplayStat()
Me.lblStatus = (RsRoute.AbsolutePosition) & " / " & RsRoute.RecordCount
End Sub

Function DupPricing() As Boolean
On Error GoTo FailSafe_Error
Dim Tmp As ADODB.Recordset
Set Tmp = New ADODB.Recordset
SQL = "SELECT  * FROM tbl_RoutePricing WHERE [RouteID]=" & CDbl(Me.DataGrid1.Columns(0).Text) & " AND [AirlineID]=" & CDbl(Me.DataGrid2.Columns(0).Text) & " AND [FareBasisID]=" & CDbl(Me.DataGrid4.Columns(0).Text) & ""
With Tmp
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            DupPricing = True
        Else
            DupPricing = False
        End If
End With
Exit Function
FailSafe_Error:
 DupPricing = False


End Function


Private Sub txtAirline_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.cmdAddSave.SetFocus
End If
End Sub

Private Sub LaVolpeButton1_Click()
frmDisplayRoutePricing.Tag = Me.DataGrid2.Columns(0).Text
frmDisplayRoutePricing.Show 1
End Sub

Private Sub LaVolpeButton2_Click()

End Sub

Private Sub txtASF_Change()
On Error GoTo FailSafe_Err
If IsNumeric(Me.txtASF) Then
ReCalc
End If
Exit Sub
FailSafe_Err:
End Sub

Private Sub txtASF_GotFocus()
With Me.txtASF
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
    End With
End Sub

Private Sub txtASF_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtASF_LostFocus()
Me.txtASF = Format(Me.txtASF, "###,##0.00")
End Sub

Private Sub txtCommision_Change()
If IsNumeric(Me.txtCommision) Then
ReCalc
End If
End Sub

Private Sub txtCommision_GotFocus()
With Me.txtCommision
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
    End With
End Sub

Private Sub txtCommision_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtCommision_LostFocus()
Me.txtCommision = Format(Me.txtCommision, "###,##0.00")
End Sub

Private Sub txtEvat_Change()
If IsNumeric(Me.txtEvat) Then
ReCalc
End If
End Sub

Private Sub txtGrossFare_Change()
If Not IsNumeric(Me.txtGrossFare) Then
    With Me.txtGrossFare
                .Text = "0.00"
                .SelStart = 0
                .SelLength = Len(.Text)
                .SetFocus
    End With
Else
    ReCalc
End If
End Sub

Private Sub txtGrossFare_GotFocus()
With Me.txtGrossFare
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
End With
End Sub

Private Sub txtGrossFare_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtGrossFare_LostFocus()
txtGrossFare = Format(Me.txtGrossFare, "###,##0.00")
End Sub

Private Sub txtInsurance_Change()
On Error GoTo FailSafe_Err
If IsNumeric(Me.txtInsurance) Then
    ReCalc
End If
Exit Sub
FailSafe_Err:
End Sub

Private Sub txtInsurance_GotFocus()
With Me.txtInsurance
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
    End With
End Sub

Private Sub txtInsurance_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtInsurance_LostFocus()
Me.txtInsurance = Format(Me.txtInsurance, "###,##0.00")
End Sub

Sub ReCalc()
Dim Tmp                 As Double
Dim TmpComm             As Double
Dim TmpCommWithVAT      As Double
Dim TmpVAT              As Double

With Me
        TmpComm = CDbl(Me.txtGrossFare) * (CDbl(Me.txtCommision) / 100)
        TmpCommWithVAT = TmpComm * (CDbl(Me.txtVat) / 100)
        TmpVAT = (CDbl(Me.txtGrossFare) + CDbl(txtInsurance) + CDbl(txtASF) + _
                  CDbl(txtTerminalFee)) * (CDbl(Me.txtVat) / 100)

'Tmp = CDbl(Me.txtGrossFare) - CDbl(TmpComm) + CDbl(txtInsurance) + CDbl(txtASF) + CDbl(txtTerminalFee) + CDbl(TmpCommWithVAT) + CDbl(TmpVAT)
'.txtNetFare = Format(Tmp, "###,##0.00")

        Tmp = CDbl(Me.txtGrossFare) + CDbl(txtInsurance) + CDbl(txtASF) + CDbl(txtTerminalFee) + _
              CDbl(TmpCommWithVAT) + CDbl(TmpVAT) + CDbl(Me.txtMisc)
              .txtNetFare = Format(Tmp, "###,##0.00")
            
End With

End Sub


Private Sub txtMeals_Change()
On Error GoTo FailSafe_Err
If IsNumeric(Me.txtMeals) Then
    ReCalc
End If
Exit Sub
FailSafe_Err:
End Sub

Private Sub txtMeals_GotFocus()
With Me.txtMeals
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
    End With
End Sub

Private Sub txtMeals_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtMeals_LostFocus()
Me.txtMeals = Format(Me.txtMeals, "###,##0.00")
End Sub


Private Sub txtMisc_Change()
On Error GoTo FailSafe_Err
If IsNumeric(Me.txtMisc) Then
    ReCalc
End If
Exit Sub
FailSafe_Err:
End Sub

Private Sub txtTerminalFee_Change()
On Error GoTo FailSafe_Err
If IsNumeric(Me.txtTerminalFee) Then
ReCalc
End If
Exit Sub
FailSafe_Err:

End Sub


Private Sub txtTerminalFee_GotFocus()
With Me.txtTerminalFee
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
    End With
End Sub

Private Sub txtTerminalFee_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtTerminalFee_LostFocus()
Me.txtTerminalFee = Format(Me.txtTerminalFee, "###,##0.00")
End Sub

Private Sub txtVAT_Change()
If IsNumeric(Me.txtVat) Then
ReCalc
End If
End Sub


Sub FillRoutes()
On Error GoTo FailSafeErr
Dim Tmp As ADODB.Recordset
Set Tmp = New ADODB.Recordset
SQL = "SELECT DISTINCT [From] From tbl_Routes ORDER BY [From]"
With Tmp
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
    If .RecordCount > 0 Then
        Me.Combo1.Clear
        Do While Not .EOF
            Me.Combo1.AddItem .Fields(0).Value
        .MoveNext
        Loop
    End If
End With
FailSafeErr:
End Sub


Sub DisplayFB()
Dim RstTempFB As New ADODB.Recordset
SQL = "SELECT * FROM tbl_TmpFareBasis WHERE [AirlineID]=" & CLng(Me.DataGrid2.Columns(0).Text)

RstTempFB.Open SQL, cn, adOpenKeyset, adLockOptimistic
        Set Me.DataGrid4.DataSource = RstTempFB

End Sub
