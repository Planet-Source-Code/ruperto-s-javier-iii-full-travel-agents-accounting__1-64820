VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "Lvbuttons.ocx"
Begin VB.Form frmDisplayRoutePricing 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Display Route Pricing"
   ClientHeight    =   9690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14940
   Icon            =   "frmDisplayRoutePricing.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   9690
   ScaleWidth      =   14940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmDisplayRoutePricing.frx":08CA
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":15A4
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":23F6
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":3248
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":3B22
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":43FC
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":4CD6
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":56A0
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":5F7A
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":6294
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":6B6E
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":7448
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":7D22
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":803C
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":8916
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":91F0
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":9ACA
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":A3A4
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":AC7E
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":B558
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":BE32
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":C70C
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":CFE6
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":D8C0
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":E19A
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":EA74
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":F34E
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":FC28
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":10502
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":10DDC
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":11692
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":11F6C
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":123BE
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":12810
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":14FC2
            Key             =   "IMG35"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplayRoutePricing.frx":16244
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   780
      Left            =   75
      TabIndex        =   6
      Top             =   8865
      Width           =   14790
      Begin LVbuttons.LaVolpeButton cmdAddSave 
         Height          =   480
         Left            =   105
         TabIndex        =   7
         Top             =   180
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   847
         BTYPE           =   3
         TX              =   "Edit"
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
         MICON           =   "frmDisplayRoutePricing.frx":189F6
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
         Left            =   13125
         TabIndex        =   8
         Top             =   180
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
         MICON           =   "frmDisplayRoutePricing.frx":18A12
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
      Begin LVbuttons.LaVolpeButton cmdDelete 
         Height          =   480
         Left            =   1605
         TabIndex        =   9
         Top             =   180
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   847
         BTYPE           =   3
         TX              =   "Delete"
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
         MICON           =   "frmDisplayRoutePricing.frx":18A2E
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
      Begin LVbuttons.LaVolpeButton cmdSetINS 
         Height          =   480
         Left            =   3105
         TabIndex        =   17
         Top             =   180
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   847
         BTYPE           =   3
         TX              =   "Set Insurance"
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
         MICON           =   "frmDisplayRoutePricing.frx":18A4A
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
      Begin LVbuttons.LaVolpeButton cmdEvat 
         Height          =   480
         Left            =   4605
         TabIndex        =   18
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
         MICON           =   "frmDisplayRoutePricing.frx":18A66
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
         Left            =   6105
         TabIndex        =   19
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
         MICON           =   "frmDisplayRoutePricing.frx":18A82
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
      Begin LVbuttons.LaVolpeButton cmdRefresh 
         Height          =   480
         Left            =   10050
         TabIndex        =   20
         Top             =   180
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   847
         BTYPE           =   3
         TX              =   "Refresh"
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
         MICON           =   "frmDisplayRoutePricing.frx":18A9E
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
      Begin LVbuttons.LaVolpeButton cmdPrint 
         Height          =   480
         Left            =   11550
         TabIndex        =   21
         Top             =   180
         Width           =   1485
         _ExtentX        =   2619
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
         MICON           =   "frmDisplayRoutePricing.frx":18ABA
         ALIGN           =   1
         IMGLST          =   "SmallImages"
         IMGICON         =   "36"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   3
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton cmdSetFee 
         Height          =   480
         Left            =   7605
         TabIndex        =   22
         Top             =   180
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   847
         BTYPE           =   3
         TX              =   "Set Fee(s)"
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
         MICON           =   "frmDisplayRoutePricing.frx":18AD6
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Options"
      Height          =   1515
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   14745
      Begin VB.PictureBox Picture1 
         Height          =   1185
         Left            =   5340
         ScaleHeight     =   1125
         ScaleWidth      =   9240
         TabIndex        =   3
         Top             =   195
         Width           =   9300
         Begin VB.ComboBox Combo3 
            Enabled         =   0   'False
            Height          =   315
            Left            =   6495
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   660
            Width           =   2580
         End
         Begin VB.ComboBox Combo2 
            Enabled         =   0   'False
            Height          =   315
            Left            =   3015
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   645
            Width           =   2580
         End
         Begin VB.ComboBox Combo1 
            Enabled         =   0   'False
            Height          =   315
            Left            =   3015
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   195
            Width           =   6075
         End
         Begin VB.Label Label3 
            Caption         =   "TO"
            Height          =   285
            Left            =   6045
            TabIndex        =   15
            Top             =   675
            Width           =   405
         End
         Begin VB.Label Label2 
            Caption         =   "FROM"
            Height          =   285
            Left            =   2040
            TabIndex        =   13
            Top             =   660
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Shipping / Airline :"
            Height          =   375
            Left            =   1320
            TabIndex        =   5
            Top             =   180
            Width           =   1455
         End
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Display By Airline/Shipping Line"
         Height          =   705
         Left            =   2385
         TabIndex        =   2
         Top             =   270
         Width           =   3570
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Display All"
         Height          =   705
         Left            =   480
         TabIndex        =   1
         Top             =   240
         Width           =   1320
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6705
      Left            =   120
      TabIndex        =   16
      Top             =   1740
      Width           =   14715
      _ExtentX        =   25956
      _ExtentY        =   11827
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   24
      BeginProperty Column00 
         DataField       =   "RoutePricingID"
         Caption         =   "RoutePricingID"
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
      BeginProperty Column02 
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
      BeginProperty Column03 
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
      BeginProperty Column04 
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
      BeginProperty Column05 
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
      BeginProperty Column06 
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
      BeginProperty Column07 
         DataField       =   "AirlineName"
         Caption         =   "AirlineName"
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
      BeginProperty Column08 
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
      BeginProperty Column09 
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
      BeginProperty Column10 
         DataField       =   "Gross Fare"
         Caption         =   "Gross Fare"
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
      BeginProperty Column11 
         DataField       =   "Misc"
         Caption         =   "MISC"
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
      BeginProperty Column12 
         DataField       =   "Insurance"
         Caption         =   "INS"
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
      BeginProperty Column13 
         DataField       =   "meals"
         Caption         =   "MEALS"
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
      BeginProperty Column14 
         DataField       =   "Commision"
         Caption         =   "COMM"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "ASF"
         Caption         =   "ASF"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column16 
         DataField       =   "Terminal Fee"
         Caption         =   "TF"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column17 
         DataField       =   "VAT"
         Caption         =   "VAT"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column18 
         DataField       =   "EVAT"
         Caption         =   "EVAT"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column19 
         DataField       =   "Net Fare"
         Caption         =   "Net Fare"
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
      BeginProperty Column20 
         DataField       =   "Service Fee"
         Caption         =   "Service Fee"
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
      BeginProperty Column21 
         DataField       =   "Refund Fee"
         Caption         =   "Refund Fee"
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
      BeginProperty Column22 
         DataField       =   "Void Fee"
         Caption         =   "Void Fee"
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
      BeginProperty Column23 
         DataField       =   "Noshow Fee"
         Caption         =   "Noshow Fee"
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
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1484.787
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   734.74
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   824.882
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   629.858
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   675.213
         EndProperty
         BeginProperty Column19 
            Alignment       =   1
            ColumnWidth     =   1184.882
         EndProperty
         BeginProperty Column20 
         EndProperty
         BeginProperty Column21 
         EndProperty
         BeginProperty Column22 
         EndProperty
         BeginProperty Column23 
         EndProperty
      EndProperty
   End
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
      Left            =   135
      TabIndex        =   11
      Top             =   8550
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
      Left            =   1005
      TabIndex        =   10
      Top             =   8565
      Width           =   900
   End
End
Attribute VB_Name = "frmDisplayRoutePricing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs                  As New ADODB.Recordset
Dim SQL                 As String

Private Sub cmdAddSave_Click()
frmEditRoutePricing.Tag = Me.DataGrid1.Columns(0).Text
frmEditRoutePricing.Show 1
End Sub

Private Sub cmdDelete_Click()
Dim ask As Integer
ask = MsgBox("Are you sure you want to delete this?", vbInformation + vbYesNo)
If ask = vbYes Then
SQL = "DELETE * FROM tbl_RoutePricing where [RoutePricingID]=" & Me.DataGrid1.Columns(0).Text
cn.Execute SQL
Rs.Requery
MsgBox "One route deleted..", vbInformation
End If
End Sub

Private Sub cmdEvat_Click()
frmSetEVAT.Tag = "-drp-"
frmSetEVAT.Show 1
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub


Private Sub cmdPrint_Click()
frmReportRoute.Show 1
End Sub

Private Sub cmdRefresh_Click()
Call Combo3_Click
End Sub

Private Sub cmdSetComm_Click()
If MDImain.StatusBar1.Panels(3).Text <> "admin" Then
        MsgBox "You have insufficient rights to update the commission", vbCritical, "Warning!!!!"
        Exit Sub
End If

frmSetCOMM.Show 1
End Sub

Private Sub cmdSetFee_Click()
frmSetRefundFee.Show 1
End Sub

Private Sub cmdSetINS_Click()
frmSetInsurance.Show 1
End Sub

Private Sub Combo1_Click()
Set Rs = New ADODB.Recordset
'SQL = "SELECT * from qryRoutePricing WHERE [AirlineName]='" & Me.Combo1 & "' AND [FROM]='" & Me.Combo2 & "' AND [TO]='" & Me.Combo3 & "'"
SQL = "SELECT * from qryRoutePricing WHERE [AirlineName]='" & Me.Combo1 & "' ORDER BY [FareBasis] ASC"
With Rs
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            Set Me.DataGrid1.DataSource = Rs
           
End With

End Sub

Private Sub Combo2_Click()
Set Rs = New ADODB.Recordset
SQL = "SELECT * from qryRoutePricing WHERE [AirlineName]='" & Me.Combo1 & "' AND [FROM]='" & Me.Combo2 & "' AND [TO]='" & Me.Combo3 & "' ORDER BY [FareBasis] ASC"
With Rs
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            Set Me.DataGrid1.DataSource = Rs
End With
End Sub


Private Sub Combo3_Click()
Set Rs = New ADODB.Recordset
SQL = "SELECT * from qryRoutePricing WHERE [AirlineName]='" & Me.Combo1 & "' AND [FROM]='" & Me.Combo2 & "' AND [TO]='" & Me.Combo3 & "' ORDER BY [FareBasis] ASC"
With Rs
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            Set Me.DataGrid1.DataSource = Rs
           
End With

End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
DisplayStat
End Sub

Sub DisplayStat()
Me.lblStatus = (Rs.AbsolutePosition) & " / " & Rs.RecordCount
End Sub



Private Sub Form_Load()
On Error GoTo FailSafe_Error
Call FillCombo
Call FillRoutes(1)
Call FillRoutes(2)

'Me.Combo1.ListIndex = 0

Me.Combo1 = frmSetTicketPricing.DataGrid2.Columns(1).Text
Me.Combo2.ListIndex = 0
Me.Combo3.ListIndex = 0

  
 
Set Rs = New ADODB.Recordset
SQL = "SELECT * from qryRoutePricing WHERE [AirlineName]='" & Me.Combo1 & "' ORDER by [FareBasis] ASC"
With Rs
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
Set Me.DataGrid1.DataSource = Rs
End With

If MDImain.StatusBar1.Panels(3).Text <> "admin" Then
    Me.DataGrid1.Columns(14).Width = 0
End If

Exit Sub

FailSafe_Error:
End Sub



Private Sub Option1_Click()
Me.Combo1.Enabled = False
Set Rs = New ADODB.Recordset
SQL = "SELECT * from qryRoutePricing ORDER by [FareBasis] ASC"
With Rs
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
Set Me.DataGrid1.DataSource = Rs
End With

End Sub

Private Sub Option2_Click()
Me.Combo1.Enabled = True
Me.Combo2.Enabled = True
Me.Combo3.Enabled = True
End Sub


Sub FillCombo()
Dim Tmp As New ADODB.Recordset
SQL = "SELECT * FROM tbl_Airline ORDER BY [AirlineName] ASC"
With Tmp
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
           If .RecordCount > 0 Then
                Me.Combo1.Clear
                Do While Not .EOF
                    Me.Combo1.AddItem .Fields(1).Value
                .MoveNext
                Loop
           End If
End With
End Sub


Sub FillRoutes(Param)
Dim Tmp As ADODB.Recordset
Set Tmp = New ADODB.Recordset
If Param = 1 Then
SQL = "SELECT DISTINCT [From] From qryRoutePricing ORDER BY [From]"
Else
SQL = "SELECT DISTINCT [TO] From qryRoutePricing ORDER BY [TO]"
End If
With Tmp
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
    If .RecordCount > 0 Then
   If Param = 1 Then
        Me.Combo2.Clear
        Do While Not .EOF
            Me.Combo2.AddItem .Fields(0).Value
        .MoveNext
        Loop
     Else
        Me.Combo3.Clear
        Do While Not .EOF
            Me.Combo3.AddItem .Fields(0).Value
        .MoveNext
        Loop
     
   End If
    End If
End With

End Sub

Sub Refresh_Grid()
On Error GoTo FailSafe_Error
Set Rs = New ADODB.Recordset
SQL = "SELECT * from qryRoutePricing WHERE [AirlineName]='" & Me.Combo1 & "' AND [FROM]='" & Me.Combo2 & "' AND [TO]='" & Me.Combo3 & "' ORDER by [FareBasis] ASC"
With Rs
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            Set Me.DataGrid1.DataSource = Rs
End With
DisplayStat
Exit Sub
FailSafe_Error:

End Sub


