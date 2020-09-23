VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "Lvbuttons.ocx"
Begin VB.Form frmBankAccSettings 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bank Account Settings"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15210
   Icon            =   "frmBankAccSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   15210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   13410
      Top             =   1395
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
            Picture         =   "frmBankAccSettings.frx":6852
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":752C
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":837E
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":91D0
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":9AAA
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":A384
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":AC5E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":B628
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":BF02
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":C21C
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":CAF6
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":D3D0
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":DCAA
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":DFC4
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":E89E
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":F178
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":FA52
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":1032C
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":10C06
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":114E0
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":11DBA
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":12694
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":12F6E
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":13848
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":14122
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":149FC
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":152D6
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":15BB0
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":1648A
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":16D64
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":1761A
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":17EF4
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":18346
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":18798
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":1AF4A
            Key             =   "IMG35"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":1C1CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankAccSettings.frx":1E97E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400040&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   15180
      TabIndex        =   15
      Top             =   -30
      Width           =   15240
      Begin VB.Image Image1 
         Height          =   495
         Left            =   120
         Picture         =   "frmBankAccSettings.frx":1F658
         Stretch         =   -1  'True
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Note: Settings of Bank Accounts"
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
         TabIndex        =   17
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Accounts and Settings"
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
         TabIndex        =   16
         Top             =   120
         Width           =   6525
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Select Bank and Airline"
      ForeColor       =   &H00C00000&
      Height          =   2955
      Left            =   105
      TabIndex        =   12
      Top             =   1035
      Width           =   15060
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   2700
         Left            =   9780
         TabIndex        =   13
         Top             =   165
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   4763
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Appearance      =   0
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
         Caption         =   "Airline / Shipping Line"
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
               ColumnWidth     =   4484.977
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2700
         Left            =   150
         TabIndex        =   14
         Top             =   165
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   4763
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
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
         Caption         =   "Bank Name"
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "BankID"
            Caption         =   "BankID"
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
            DataField       =   "Bank Name"
            Caption         =   "Bank Name"
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
            DataField       =   "Bank Address"
            Caption         =   "Bank Address"
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
            AllowFocus      =   0   'False
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3479.811
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   4844.977
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Account Details"
      ForeColor       =   &H80000008&
      Height          =   2670
      Left            =   9900
      TabIndex        =   4
      Top             =   4110
      Width           =   5235
      Begin VB.TextBox txtDescription 
         Height          =   405
         Left            =   135
         TabIndex        =   20
         Top             =   1200
         Width           =   4995
      End
      Begin VB.TextBox txtCurrentBalance 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   135
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   2100
         Width           =   1695
      End
      Begin VB.TextBox txtAccountName 
         Height          =   405
         Left            =   2100
         TabIndex        =   7
         Top             =   555
         Width           =   3075
      End
      Begin VB.TextBox txtAccountNumber 
         Height          =   405
         Left            =   135
         TabIndex        =   6
         Top             =   555
         Width           =   1860
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Description :"
         Height          =   390
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   1470
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Current Balance :"
         Height          =   390
         Left            =   120
         TabIndex        =   10
         Top             =   1875
         Width           =   1470
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name :"
         Height          =   390
         Left            =   2115
         TabIndex        =   8
         Top             =   315
         Width           =   1470
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Number :"
         Height          =   390
         Left            =   150
         TabIndex        =   5
         Top             =   330
         Width           =   1470
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   2625
      Left            =   105
      TabIndex        =   0
      Top             =   4065
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   4630
      _Version        =   393216
      AllowUpdate     =   -1  'True
      Appearance      =   0
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
      Caption         =   "Current Settings"
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "AccountsSettingID"
         Caption         =   "AccountsSettingID"
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
         DataField       =   "BankID"
         Caption         =   "BankID"
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
         DataField       =   "AirlineId"
         Caption         =   "AirlineId"
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
         DataField       =   "Bank Name"
         Caption         =   "Bank Name"
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
      BeginProperty Column05 
         DataField       =   "Account Number"
         Caption         =   "Account Number"
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
         DataField       =   "Account Name"
         Caption         =   "Account Name"
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
         DataField       =   "Description"
         Caption         =   "Description"
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
         DataField       =   "Current Balance"
         Caption         =   "Current Balance"
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
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1409.953
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1379.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1560.189
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1560.189
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton cmdSet 
      Height          =   480
      Left            =   270
      TabIndex        =   1
      Top             =   6945
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "Set"
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
      MICON           =   "frmBankAccSettings.frx":1FF22
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
   Begin LVbuttons.LaVolpeButton cmdDel 
      Height          =   480
      Left            =   1635
      TabIndex        =   2
      Top             =   6945
      Width           =   1350
      _ExtentX        =   2381
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
      MICON           =   "frmBankAccSettings.frx":1FF3E
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
   Begin LVbuttons.LaVolpeButton cmdExit 
      Height          =   480
      Left            =   7380
      TabIndex        =   3
      Top             =   6945
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
      MICON           =   "frmBankAccSettings.frx":1FF5A
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
   Begin LVbuttons.LaVolpeButton cmdEdit 
      Height          =   480
      Left            =   3000
      TabIndex        =   11
      Top             =   6960
      Width           =   1350
      _ExtentX        =   2381
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
      COLTYPE         =   2
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmBankAccSettings.frx":1FF76
      ALIGN           =   1
      IMGLST          =   "SmallImages"
      IMGICON         =   "16"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   480
      Left            =   4365
      TabIndex        =   18
      Top             =   6960
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "Passbook"
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
      MICON           =   "frmBankAccSettings.frx":1FF92
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
Attribute VB_Name = "frmBankAccSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsBank As ADODB.Recordset
Dim RsAirline As ADODB.Recordset
Dim Rs As ADODB.Recordset
Dim RsSet As ADODB.Recordset
Dim SQL As String

Private Sub cmdDel_Click()
On Error GoTo ErrExit
Dim ask As Integer

ask = MsgBox("Are you sure you want to delete this?", vbYesNo + vbInformation)
If ask = vbYes Then
    SQL = "DELETE * FROM tbl_AccountsSetting WHERE [AccountsSettingID]=" & Me.DataGrid3.Columns(0).Text
    cn.Execute SQL
    Rs.Requery
    MsgBox "One record deleted...", vbInformation
End If

Exit Sub
ErrExit:
End Sub

Private Sub cmdEdit_Click()
Me.cmdEdit.Enabled = False
Frame1.Enabled = True
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdNew_Click()

End Sub

Private Sub cmdSet_Click()
'On Error GoTo FailSafe_Error
Dim Criteria As String

'If checkAirline Then
'  If Me.cmdEdit.Enabled = False Then: GoTo xxx
'    MsgBox "This Account already set", vbInformation
'Else

If CheckNull(Me.txtAccountNumber) Then
    MsgBox "Account Number Should not be blank", vbCritical
    Me.txtAccountNumber.SetFocus
    Exit Sub
End If

If CheckNull(Me.txtAccountName) Then
    MsgBox "Account Name Should not be blank", vbCritical
    Me.txtAccountName.SetFocus
    Exit Sub
End If

   ' Me.txtCurrentBalance = Format(checkDupAcc, "###,##0.00")

xxx:
With RsSet
If Me.cmdEdit.Enabled = True Then
    
        .AddNew
        .Fields(1).Value = Me.DataGrid1.Columns(0).Text
        .Fields(2).Value = Me.DataGrid2.Columns(0).Text
        .Fields(3).Value = UCase(Me.txtAccountNumber)
        .Fields(4).Value = UCase(Me.txtAccountName)
        .Fields(5).Value = Format(Me.txtCurrentBalance, "###,##0.00")
        .Fields("Description").Value = UCase(Me.txtDescription)
        .Update
        
    Else
    
SQL = "UPDATE tbl_AccountsSetting SET [BankID]=" & CDbl(Me.DataGrid1.Columns(0).Text) & _
      ",[Current Balance] = " & CDbl(Me.txtCurrentBalance) & _
      ", [Account Name] = '" & UCase(Me.txtAccountName) & _
      "', [Description] = '" & UCase(Me.txtDescription) & "'" & _
      " Where ((([Account Number]) = '" & Me.txtAccountNumber & "'))"

'SQL = "UPDATE tbl_AccountsSetting SET " & _
'      "[Current Balance] = " & CDbl(Me.txtCurrentBalance) & _
'      ", [Account Name] = '" & UCase(Me.txtAccountName) & _
'      "', [Description] = '" & UCase(Me.txtDescription) & "'" & _
'      " Where ((([Account Number]) = '" & Me.txtAccountNumber & "'))"


              cn.BeginTrans
                    cn.Execute SQL
              cn.CommitTrans
              
              MsgBox "Accounts successfully updated...", vbInformation
    End If
    
End With
Me.cmdEdit.Enabled = True
Rs.Requery
'End If
Exit Sub
FailSafe_Error:
MsgBox Err.Description
cn.RollbackTrans
End Sub

Sub UpdatePassbook(Param, Criteria)
Dim Rs As New ADODB.Recordset
Dim RsCheck As New ADODB.Recordset


SQL = "SELECT * FROM tbl_BankPassbook WHERE [Account Number]='" & Param & "'"
RsCheck.Open SQL, cn, adOpenKeyset, adLockOptimistic

If RsCheck.RecordCount > 0 Then
        RsCheck.Close
      Set RsCheck = Nothing
      Exit Sub
End If
SQL = "SELECT * FROM tbl_BankPassbook"

With Rs
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If Criteria = "new" Then
            .AddNew
            .Fields(1).Value = Param
            .Fields(2).Value = Format(Now, "mm/dd/yyyy")
            .Fields(3).Value = Format(Me.txtCurrentBalance, "###,##0.00")
            .Fields(4).Value = 0
            .Fields(5).Value = Format(Me.txtCurrentBalance, "###,##0.00")
            Else
            .MoveFirst
            Do While Not .EOF
            If .Fields(1).Value = Param Then
                .Fields(5).Value = Format(Me.txtCurrentBalance, "###,##0.00")
            End If
            .MoveNext
            Loop
        End If
            .Update
        .Close
     Set Rs = Nothing
End With

End Sub

Private Sub DataGrid3_Click()
On Error GoTo ErrExit
Me.txtAccountNumber = Me.DataGrid3.Columns(5).Text
Me.txtAccountName = Me.DataGrid3.Columns(6).Text
Me.txtDescription = Me.DataGrid3.Columns(7).Text
'Me.txtCurrentBalance = Me.DataGrid3.Columns(8).Text

Exit Sub
ErrExit:
End Sub


Private Sub Form_Load()
Set RsBank = New ADODB.Recordset
Set RsAirline = New ADODB.Recordset
SQL = "SELECT * FROM tbl_Banks"
With RsBank
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
   Set Me.DataGrid1.DataSource = RsBank
End With

SQL = "SELECT * FROM tbl_Airline"
With RsAirline
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
   Set Me.DataGrid2.DataSource = RsAirline
End With


Set Rs = New ADODB.Recordset
SQL = "SELECT * FROM qryAccountsSettings ORDER by [Account Number] ASC"
With Rs
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
        Set Me.DataGrid3.DataSource = Rs
End With

Set RsSet = New ADODB.Recordset
SQL = "SELECT * FROM tbl_AccountsSetting"
With RsSet
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
End With
End Sub

Function chekDup() As Boolean
Dim RsCheck As ADODB.Recordset
Set RsCheck = New ADODB.Recordset
SQL = "SELECT * FROM tbl_AccountsSetting WHERE [BankID]=" & Me.DataGrid1.Columns(0).Text & " AND [AirlineId] = " & Me.DataGrid2.Columns(0).Text
With RsCheck
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
    If .RecordCount > 0 Then
    chekDup = True
    Else
    chekDup = False
    End If
    .Close
End With

End Function

Function checkDupAcc() As Double
Dim RsCheck As ADODB.Recordset
Set RsCheck = New ADODB.Recordset
SQL = "SELECT * FROM tbl_AccountsSetting WHERE [Account Number]='" & Me.txtAccountNumber & "'"
With RsCheck
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
    If .RecordCount > 0 Then
    checkDupAcc = .Fields(5).Value
    Else
    checkDupAcc = Format(Me.txtCurrentBalance, "###,##0.00")
    End If
    .Close
End With

End Function


Function checkAirline() As Boolean
Dim RsCheck As ADODB.Recordset
Set RsCheck = New ADODB.Recordset
SQL = "SELECT * FROM tbl_AccountsSetting WHERE [AirlineId]=" & Me.DataGrid2.Columns(0).Text
With RsCheck
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
    If .RecordCount > 0 Then
    checkAirline = True
    Else
    checkAirline = False
    End If
    .Close
End With

End Function


Private Sub LaVolpeButton1_Click()
frmPassbook.Show 1
End Sub
