VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "Lvbuttons.ocx"
Begin VB.Form frmCustomerAccounts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Accounts"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13875
   Icon            =   "frmCustomerAccounts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   13875
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReconcile 
      Caption         =   "Reconcile A/R"
      Height          =   540
      Left            =   9045
      TabIndex        =   46
      Top             =   6555
      Width           =   2760
   End
   Begin VB.CommandButton cmdSetComm 
      Caption         =   "Set Commission Percent"
      Height          =   540
      Left            =   6270
      TabIndex        =   43
      Top             =   7125
      Width           =   2760
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy Ship/Airline"
      Height          =   540
      Left            =   6270
      TabIndex        =   42
      Top             =   6555
      Width           =   2760
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   2115
      Left            =   30
      TabIndex        =   41
      Top             =   6540
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   3731
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "AccountIDDetails"
         Caption         =   "AccountIDDetails"
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
         DataField       =   "AccountID"
         Caption         =   "AccountID"
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
      BeginProperty Column03 
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
         DataField       =   "ShipAirline"
         Caption         =   "ShipAirline"
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
         DataField       =   "Commission"
         Caption         =   "Commission"
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
         BeginProperty Column00 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   3990.047
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdMovelast 
      Caption         =   ">>"
      Height          =   375
      Left            =   8760
      TabIndex        =   40
      Top             =   8760
      Width           =   1725
   End
   Begin VB.CommandButton cmdMoveNext 
      Caption         =   ">"
      Height          =   375
      Left            =   7050
      TabIndex        =   39
      Top             =   8760
      Width           =   1725
   End
   Begin VB.CommandButton cmdMovePrevious 
      Caption         =   "<"
      Height          =   375
      Left            =   5340
      TabIndex        =   38
      Top             =   8760
      Width           =   1725
   End
   Begin VB.CommandButton cmdMoveFirst 
      Caption         =   "<<"
      Height          =   375
      Left            =   3630
      TabIndex        =   37
      Top             =   8760
      Width           =   1725
   End
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   165
      Top             =   1305
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
            Picture         =   "frmCustomerAccounts.frx":0442
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":111C
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":1F6E
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":2DC0
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":369A
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":3F74
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":484E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":5218
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":5AF2
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":5E0C
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":66E6
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":6FC0
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":789A
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":7BB4
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":848E
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":8D68
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":9642
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":9F1C
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":A7F6
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":B0D0
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":B9AA
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":C284
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":CB5E
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":D438
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":DD12
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":E5EC
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":EEC6
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":F7A0
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":1007A
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":10954
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":1120A
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":11AE4
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":11F36
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":12388
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":14B3A
            Key             =   "IMG35"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerAccounts.frx":15DBC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Caption         =   "Customer Info"
      Enabled         =   0   'False
      Height          =   2520
      Left            =   45
      TabIndex        =   12
      Top             =   45
      Width           =   13770
      Begin VB.TextBox txtcustNo 
         Enabled         =   0   'False
         Height          =   300
         Left            =   11220
         TabIndex        =   32
         Top             =   480
         Width           =   2385
      End
      Begin VB.PictureBox Picture1 
         Height          =   1185
         Left            =   105
         ScaleHeight     =   1125
         ScaleWidth      =   13515
         TabIndex        =   19
         Top             =   1245
         Width           =   13575
         Begin VB.TextBox txtAccountBalanceDollar 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   10050
            TabIndex        =   44
            Text            =   "0.00"
            Top             =   570
            Width           =   3405
         End
         Begin VB.Frame Frame5 
            Caption         =   "Credit Terms"
            ForeColor       =   &H00C00000&
            Height          =   1050
            Left            =   765
            TabIndex        =   29
            Top             =   75
            Width           =   1605
            Begin VB.TextBox txtCreditTerms 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   105
               TabIndex        =   31
               Text            =   "0"
               Top             =   600
               Width           =   1350
            End
            Begin VB.Label Label4 
               Caption         =   "No. of Day(s) :"
               Height          =   255
               Left            =   120
               TabIndex        =   30
               Top             =   315
               Width           =   1095
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Credit Limit"
            ForeColor       =   &H00C00000&
            Height          =   1050
            Left            =   2430
            TabIndex        =   24
            Top             =   75
            Width           =   2670
            Begin VB.TextBox txtCreditLimitDollar 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   975
               TabIndex        =   28
               Text            =   "0.00"
               Top             =   585
               Width           =   1575
            End
            Begin VB.TextBox txtCreditLimitPeso 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   975
               TabIndex        =   27
               Text            =   "0.00"
               Top             =   225
               Width           =   1575
            End
            Begin VB.Label Label3 
               Caption         =   "Dollar :"
               Height          =   255
               Left            =   420
               TabIndex        =   26
               Top             =   705
               Width           =   600
            End
            Begin VB.Label Label2 
               Caption         =   "Peso :"
               Height          =   255
               Left            =   465
               TabIndex        =   25
               Top             =   330
               Width           =   600
            End
         End
         Begin VB.TextBox txtAccountBalancePeso 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   10050
            TabIndex        =   21
            Text            =   "0.00"
            Top             =   105
            Width           =   3405
         End
         Begin VB.Label Label5 
            Caption         =   "Current Account Balance Dollar:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   6165
            TabIndex        =   45
            Top             =   570
            Width           =   3855
         End
         Begin VB.Label Label1 
            Caption         =   "Current Account Balance Peso:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   6165
            TabIndex        =   20
            Top             =   105
            Width           =   3840
         End
      End
      Begin VB.TextBox txtAccountName 
         DataField       =   "Account Name"
         Height          =   285
         Left            =   1860
         TabIndex        =   1
         Top             =   180
         Width           =   3375
      End
      Begin VB.TextBox txtContactPerson 
         DataField       =   "Contact Person"
         Height          =   285
         Left            =   6660
         TabIndex        =   2
         Top             =   135
         Width           =   3375
      End
      Begin VB.TextBox txtMobileNo 
         DataField       =   "Mobile No"
         Height          =   285
         Left            =   1860
         TabIndex        =   3
         Top             =   540
         Width           =   3375
      End
      Begin VB.TextBox txtTelNo 
         DataField       =   "Tel No"
         Height          =   285
         Left            =   6660
         TabIndex        =   4
         Top             =   495
         Width           =   3375
      End
      Begin VB.TextBox txtEmail 
         DataField       =   "Email"
         Height          =   285
         Left            =   1860
         TabIndex        =   5
         Top             =   885
         Width           =   3375
      End
      Begin VB.TextBox txtBusinessAddress 
         DataField       =   "Business Address"
         Height          =   285
         Left            =   6675
         TabIndex        =   6
         Top             =   870
         Width           =   3375
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Customer ID :"
         Height          =   195
         Index           =   0
         Left            =   11220
         TabIndex        =   33
         Top             =   180
         Width           =   960
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Account Name:"
         Height          =   255
         Index           =   1
         Left            =   390
         TabIndex        =   18
         Top             =   195
         Width           =   1455
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Contact Person:"
         Height          =   255
         Index           =   2
         Left            =   4815
         TabIndex        =   17
         Top             =   180
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Mobile No:"
         Height          =   255
         Index           =   3
         Left            =   15
         TabIndex        =   16
         Top             =   585
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tel No:"
         Height          =   255
         Index           =   4
         Left            =   4815
         TabIndex        =   15
         Top             =   540
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Email:"
         Height          =   255
         Index           =   5
         Left            =   15
         TabIndex        =   14
         Top             =   930
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Business Address:"
         Height          =   255
         Index           =   6
         Left            =   4830
         TabIndex        =   13
         Top             =   915
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Account List"
      Height          =   3015
      Left            =   30
      TabIndex        =   10
      Top             =   3480
      Width           =   13800
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2610
         Left            =   75
         TabIndex        =   11
         Top             =   270
         Width           =   13635
         _ExtentX        =   24051
         _ExtentY        =   4604
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
         ColumnCount     =   14
         BeginProperty Column00 
            DataField       =   "AccountID"
            Caption         =   "AccountID"
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
         BeginProperty Column02 
            DataField       =   "Contact Person"
            Caption         =   "Contact Person"
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
            DataField       =   "Mobile No"
            Caption         =   "Mobile No"
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
            DataField       =   "Tel No"
            Caption         =   "Tel No"
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
            DataField       =   "Email"
            Caption         =   "Email"
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
            DataField       =   "Business Address"
            Caption         =   "Business Address"
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
         BeginProperty Column07 
            DataField       =   ""
            Caption         =   ""
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
         BeginProperty Column08 
            DataField       =   "Credit Limit Peso"
            Caption         =   "Credit Limit Peso"
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
         BeginProperty Column09 
            DataField       =   "Credit Limit Dollar"
            Caption         =   "Credit Limit Dollar"
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
         BeginProperty Column10 
            DataField       =   "Commission Percent"
            Caption         =   "Commission Percent"
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
         BeginProperty Column11 
            DataField       =   "Credit Terms"
            Caption         =   "Credit Terms"
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
         BeginProperty Column12 
            DataField       =   "Current Balance Peso"
            Caption         =   "Current Balance Peso"
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
            DataField       =   "Current Balance Dollar"
            Caption         =   "Current Balance Dollar"
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
               ColumnWidth     =   2924.788
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1844.787
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column08 
            EndProperty
            BeginProperty Column09 
            EndProperty
            BeginProperty Column10 
            EndProperty
            BeginProperty Column11 
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   2145.26
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   1980.284
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   870
      Left            =   30
      TabIndex        =   7
      Top             =   2580
      Width           =   13800
      Begin VB.TextBox txtSearch 
         DataField       =   "Account Name"
         Height          =   285
         Left            =   7170
         TabIndex        =   34
         Top             =   390
         Width           =   3375
      End
      Begin LVbuttons.LaVolpeButton cmdAddSave 
         Height          =   480
         Left            =   105
         TabIndex        =   0
         Top             =   195
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
         MICON           =   "frmCustomerAccounts.frx":1C61E
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
         Left            =   12135
         TabIndex        =   8
         Top             =   195
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
         MICON           =   "frmCustomerAccounts.frx":1C63A
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
         Top             =   195
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
         MICON           =   "frmCustomerAccounts.frx":1C656
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
         Left            =   10635
         TabIndex        =   22
         Top             =   195
         Width           =   1485
         _ExtentX        =   2619
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
         MICON           =   "frmCustomerAccounts.frx":1C672
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
      Begin LVbuttons.LaVolpeButton cmdEdit 
         Height          =   480
         Left            =   3105
         TabIndex        =   23
         Top             =   195
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
         MICON           =   "frmCustomerAccounts.frx":1C68E
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
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Type Here the Account Name to Search:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   7170
         TabIndex        =   35
         Top             =   165
         Width           =   2910
      End
   End
   Begin VB.Label lblStatus 
      Caption         =   "status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   45
      TabIndex        =   36
      Top             =   8820
      Width           =   2940
   End
End
Attribute VB_Name = "frmCustomerAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RsClone As ADODB.Recordset


Private Sub cmdAddSave_Click()
If Me.cmdAddSave.Caption = "Add" Then
    Me.cmdAddSave.Caption = "Save"
    Me.Frame3.Enabled = True
    Me.txtAccountName.SetFocus
    Call ClearTxt
Else
    If CheckNull(Me.txtAccountName) Then: MsgBox "Account Name Should not be blank!", vbInformation: Me.txtAccountName.SetFocus: Exit Sub
    If CheckNull(Me.txtContactPerson) Then: MsgBox "Contact person Should not be blank!", vbInformation: Me.txtContactPerson.SetFocus: Exit Sub
    
    If Not IsNumeric(Me.txtAccountBalancePeso) Then
        MsgBox "Account Balance Peso Should be numeric!", vbInformation
        Call kulotHL(Me.txtAccountBalancePeso)
        Exit Sub
    End If
    
    If Not IsNumeric(Me.txtAccountBalanceDollar) Then
        MsgBox "Account Balance Dollar Should be numeric!", vbInformation
        Call kulotHL(Me.txtAccountBalanceDollar)
        Exit Sub
    End If
    
    
    If Not IsNumeric(Me.txtCreditLimitPeso) Then
        MsgBox "CreditLimitPeso Should be numeric!", vbInformation
        Call kulotHL(Me.txtCreditLimitPeso)
        Exit Sub
    End If
    
    If Not IsNumeric(Me.txtCreditLimitDollar) Then
        MsgBox "CreditLimitDollar Should be numeric!", vbInformation
        Call kulotHL(Me.txtCreditLimitDollar)
        Exit Sub
    End If
    
    
    
    If Not IsNumeric(Me.txtCreditTerms) Then
        MsgBox "CreditTerms Should be numeric!", vbInformation
        Call kulotHL(Me.txtCreditTerms)
        Exit Sub
    End If
    
Dim ask As Integer
    
ask = MsgBox("Are you sure you want to save?", vbInformation + vbYesNo, "Confirm")
        If ask = vbYes Then
    
                If Me.cmdEdit.Enabled = False Then
                  Call SaveData("edit")
                  Me.cmdAddSave.Caption = "Add"
                  Me.Frame3.Enabled = False
                  Me.cmdEdit.Enabled = True
                  Exit Sub
                Else
                  GoTo xxx
                End If
                Me.cmdDelete.Enabled = True
        Else
                Exit Sub
                Me.cmdDelete.Enabled = False
        End If
    

    
ask = MsgBox("Are you sure you want to save?", vbInformation + vbYesNo, "Confirm")
        If ask = vbYes Then
xxx:
                If Not checkDupAcc Then
                  Call SaveData("new")
                Else
                  Me.cmdAddSave.Caption = "Add"
                  Me.Frame3.Enabled = False
                End If
                Me.cmdDelete.Enabled = True
        Else
                Me.cmdDelete.Enabled = False
        End If
  
End If

End Sub

Sub SaveData(ByVal userStat As String)
On Error GoTo ErrExit
cn.BeginTrans
With Rs
If userStat = "new" Then
        .AddNew
 Else
                .MoveFirst
                .Find "[AccountID]=" & CLng(Me.txtcustNo)
End If
                .Fields(1).Value = UCase(Me.txtAccountName)
                .Fields(2).Value = UCase(Me.txtContactPerson)
                .Fields(3).Value = UCase(Me.txtMobileNo)
                .Fields(4).Value = UCase(Me.txtTelNo)
                .Fields(5).Value = UCase(Me.txtEmail)
                .Fields(6).Value = UCase(Me.txtBusinessAddress)
                .Fields(7).Value = CDbl(Me.txtAccountBalancePeso)
                .Fields(8).Value = CDbl(Me.txtCreditLimitPeso)
                .Fields(9).Value = CDbl(Me.txtCreditLimitDollar)
                .Fields(11).Value = CDbl(Me.txtCreditTerms)
                .Fields(12).Value = CDbl(Me.txtAccountBalanceDollar)
        .Update
        
End With
cn.CommitTrans
MsgBox "Record Save...", vbInformation
Exit Sub
ErrExit:
cn.RollbackTrans
MsgBox "There was an error in saving the accounts! save cancelled!"
End Sub

Function checkDupAcc() As Boolean
Dim Rstmp As New ADODB.Recordset
If CheckNull(Me.txtcustNo) Then: checkDupAcc = False: Exit Function
SQL = "SELECT * FROM tbl_CustAccounts WHERE [AccountID]=" & CLng(Me.txtcustNo)
With Rstmp
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                checkDupAcc = True
            Else
                checkDupAcc = False
            End If
           .Close
        Set Rstmp = Nothing
End With
End Function

Private Sub cmdCopy_Click()
On Error GoTo FailSafe_Error
Dim ask As Integer
Dim Rstmp As New ADODB.Recordset

Dim tmpAirline As Long

ask = MsgBox("Are you sure you want to copy to this account :" & Me.txtAccountName, vbInformation + vbYesNo)
If ask = vbYes Then
    With Rstmp
            .Open "SELECT * FROM tbl_Airline", cn, adOpenKeyset, adLockOptimistic
           If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                tmpAirline = .Fields("AirlineID").Value
                               

                If Not checkDupAccDetails(tmpAirline, CLng(Me.txtcustNo)) Then
                    Call UpdateAccDetails(tmpAirline)
                End If
                .MoveNext
                Loop
                Call CustAccDetails
           End If
           
    End With
End If
Exit Sub
FailSafe_Error:
MsgBox Err.Description

End Sub


Function checkDupAccDetails(ByVal UserAirlineID As Long, ByVal UserAccountId As Long) As Boolean
Dim Rstmp As New ADODB.Recordset
SQL = "SELECT * FROM tbl_custAccDetails WHERE [AirlineID]=" & CLng(UserAirlineID) & " AND [AccountID]=" & CLng(UserAccountId)
With Rstmp
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                checkDupAccDetails = True
            Else
                checkDupAccDetails = False
            End If
           .Close
        Set Rstmp = Nothing
End With
End Function

Sub UpdateAccDetails(ByVal UserAirline As Long)
Dim rsAccTmp As New ADODB.Recordset
   With rsAccTmp
          .Open "SELECT * FROM tbl_custAccDetails", cn, adOpenKeyset, adLockOptimistic
                        cn.BeginTrans
         .AddNew
                            .Fields("AccountID").Value = CLng(Me.txtcustNo)
                            .Fields("AirlineID").Value = CLng(UserAirline)
                            .Fields("Commission").Value = 0
         .Update
                        cn.CommitTrans
   End With

End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrExit
Dim SQL As String
Dim ask As Integer

SQL = "DELETE * FROM tbl_CustAccounts WHERE [AccountID]=" & Me.DataGrid1.Columns(0).Text
ask = MsgBox("Are you sure you want to remove this account name?", vbYesNo + vbCritical)
If ask = vbYes Then
    With cn
        .BeginTrans
        .Execute SQL
        .CommitTrans
    End With
    Rs.Requery
    MsgBox "The selected Account was deleted...", vbInformation
End If
Exit Sub
ErrExit:
       cn.RollbackTrans
End Sub

Private Sub cmdEdit_Click()
Me.cmdEdit.Enabled = False
Me.Frame3.Enabled = True
Me.cmdDelete.Enabled = False
Me.cmdAddSave.Caption = "Save"
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Sub ClearTxt()
With Me
        .txtAccountName = ""
        .txtBusinessAddress = ""
        .txtContactPerson = ""
        .txtEmail = ""
        .txtMobileNo = ""
        .txtTelNo = ""
        .txtAccountBalancePeso = "0.00"
        .txtAccountBalanceDollar = "0.00"
        .txtCreditLimitPeso = "0.00"
        .txtCreditLimitDollar = "0.00"
        .txtCreditTerms = "0"
        
        .txtcustNo = ""
End With
End Sub

Sub LoadValues()
Dim myAccountName As String
Dim myContactPerson As String
Dim myMobileNo As String
Dim myTelno As String
Dim myEmail As String
Dim myBusinessAddress As String
Dim myAccountBalancePeso
Dim myAccountBalanceDollar
Dim myPesoCreditLimit
Dim myDollarCreditLimit
Dim myCommissionPercent
Dim myCreditTerms
Dim mycustNo


Dim flag As Integer

flag = IIf(IsNull(Me.DataGrid1.Columns(0).Text), 0, 1)

If flag = 0 Then
    MsgBox "Please select a customer", vbInformation
    Exit Sub
End If

mycustNo = Me.DataGrid1.Columns(0).Text

With Me.DataGrid1
        myAccountName = IIf(IsNull(.Columns(1).Text), Empty, .Columns(1).Text)
        myContactPerson = IIf(IsNull(.Columns(2).Text), Empty, .Columns(2).Text)
        myMobileNo = IIf(IsNull(.Columns(3).Text), Empty, .Columns(3).Text)
        myTelno = IIf(IsNull(.Columns(4).Text), Empty, .Columns(4).Text)
        myEmail = IIf(IsNull(.Columns(5).Text), Empty, .Columns(5).Text)
        myBusinessAddress = IIf(IsNull(.Columns(6).Text), Empty, .Columns(6).Text)
        myAccountBalancePeso = IIf(IsNull(.Columns(12).Text), "0.00", .Columns(12).Text)
        myAccountBalanceDollar = IIf(IsNull(.Columns(13).Text), "0.00", .Columns(13).Text)
        myPesoCreditLimit = IIf(IsNull(.Columns(8).Text), "0.00", .Columns(8).Text)
        myDollarCreditLimit = IIf(IsNull(.Columns(9).Text), "0.00", .Columns(9).Text)
        myCommissionPercent = IIf(IsNull(.Columns(10).Text), "0.00", .Columns(10).Text)
        myCreditTerms = IIf(IsNull(.Columns(11).Text), "0", .Columns(11).Text)
        
End With

With Me
        .txtcustNo = mycustNo
        .txtAccountName = myAccountName
        .txtBusinessAddress = myBusinessAddress
        .txtContactPerson = myContactPerson
        .txtEmail = myEmail
        .txtMobileNo = myMobileNo
        .txtTelNo = myTelno
        .txtAccountBalancePeso = IIf(CheckNull(myAccountBalancePeso), "0.00", myAccountBalancePeso)
        .txtAccountBalanceDollar = IIf(CheckNull(myAccountBalanceDollar), "0.00", myAccountBalanceDollar)
        .txtCreditLimitPeso = IIf(CheckNull(myPesoCreditLimit), "0.00", myPesoCreditLimit)
        .txtCreditLimitDollar = IIf(CheckNull(myDollarCreditLimit), "0.00", myDollarCreditLimit)
        .txtCreditTerms = IIf(CheckNull(myCreditTerms), "0.00", myCreditTerms)
        
        
End With
Call CustAccDetails
Me.lblStatus = "RECORD :" & Rs.AbsolutePosition & "/" & Rs.RecordCount
End Sub

Private Sub cmdMoveFirst_Click()
Call Navigate(Me.cmdMoveFirst.Caption)
End Sub

Private Sub cmdMovelast_Click()
Call Navigate(Me.cmdMovelast.Caption)
End Sub

Private Sub cmdMoveNext_Click()
Call Navigate(Me.cmdMoveNext.Caption)
End Sub

Private Sub cmdMovePrevious_Click()
Call Navigate(Me.cmdMovePrevious.Caption)
End Sub

Private Sub Command1_Click()

End Sub

Function Return_INTL_Due(Param, WhichBal) As Double
Dim RstDueIntl As New ADODB.Recordset
SQL = "SELECT * FROM qryCheck_SA_BALANCE_INTL WHERE [AccountID]=" & Param
With RstDueIntl
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                        If WhichBal = "DOLLAR" Then
                            Return_INTL_Due = .Fields(0).Value
                        Else
                            Return_INTL_Due = .Fields(1).Value
                        End If
            End If
            .Close
         Set RstDueIntl = Nothing
End With
End Function

Private Sub cmdReconcile_Click()
Dim Rst             As New ADODB.Recordset
Dim ask             As Integer
Dim PesoBal         As Double
Dim DollarBal       As Double

SQL = "SELECT * FROM qryCustAccountBalancePeso WHERE [AccountID]=" & CLng(Me.txtcustNo)

With Rst
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
    If .RecordCount > 0 Then
    
    PesoBal = CDbl(.Fields(1).Value) + CDbl(Return_INTL_Due(CLng(Me.txtcustNo), "PESO"))
    DollarBal = CDbl(Return_INTL_Due(CLng(Me.txtcustNo), "DOLLAR"))
    
    ask = MsgBox("Not reconciled Peso Balance :" & Me.txtAccountBalancePeso & Chr(13) & _
                 "Not reconciled Dollar Balance :" & Me.txtAccountBalanceDollar & Chr(13) & _
                 "Reconciled Peso Balance :" & RetCurrency(PesoBal) & Chr(13) & _
                 "Reconciled Dollar Balance :" & RetCurrency(DollarBal) & Chr(13) & _
                 "Are you sure you want to Reconcile?", vbInformation + vbYesNo, "Confirm")
        If ask = vbYes Then
        
        Me.txtAccountBalancePeso = RetCurrency(PesoBal)
        Me.txtAccountBalanceDollar = RetCurrency(DollarBal)
        
        Me.cmdEdit.Enabled = False
                If Me.cmdEdit.Enabled = False Then
                  Call SaveData("edit")
                  Me.cmdAddSave.Caption = "Add"
                  Me.Frame3.Enabled = False
                  Me.cmdEdit.Enabled = True
                  Exit Sub
                End If
                Me.cmdDelete.Enabled = True
        Else
                Exit Sub
                Me.cmdDelete.Enabled = False
        End If
    
    Else
    
                If CheckSA_BAL(CLng(Me.txtcustNo)) = 0 And CDbl(Me.txtAccountBalancePeso) > 0 Then
                  
                    PesoBal = CheckSA_BAL(CLng(Me.txtcustNo)) + CDbl(CheckSA_BAL_INTL(CLng(Me.txtcustNo), "PESO"))
                    DollarBal = CDbl(CheckSA_BAL_INTL(CLng(Me.txtcustNo), "DOLLAR"))
                    
                            ask = MsgBox("Not reconciled Peso Balance :" & Me.txtAccountBalancePeso & Chr(13) & _
                                         "Not reconciled Dollar Balance :" & Me.txtAccountBalanceDollar & Chr(13) & _
                                         "Reconciled Peso Balance :" & RetCurrency(PesoBal) & Chr(13) & _
                                         "Reconciled Dollar Balance :" & RetCurrency(DollarBal) & Chr(13) & _
                                         "Are you sure you want to Reconcile?", vbInformation + vbYesNo, "Confirm")
                                         
                             If ask = vbYes Then
                             
                             Me.txtAccountBalancePeso = RetCurrency(PesoBal)
                             Me.txtAccountBalanceDollar = RetCurrency(DollarBal)
                             
                             Me.cmdEdit.Enabled = False
                                     If Me.cmdEdit.Enabled = False Then
                                       Call SaveData("edit")
                                       Me.cmdAddSave.Caption = "Add"
                                       Me.Frame3.Enabled = False
                                       Me.cmdEdit.Enabled = True
                                       Exit Sub
                                     End If
                                     Me.cmdDelete.Enabled = True
                             Else
                                     Exit Sub
                                     Me.cmdDelete.Enabled = False
                             End If
                
                End If
    End If

End With
End Sub
Function CheckSA_BAL(Param) As Double
Dim RstCheck As New ADODB.Recordset
SQL = "SELECT * FROM qryCheck_SA_BALANCE WHERE [AccountID]=" & CLng(Param)
With RstCheck
        .Open SQL, cn, adOpenKeyset
        If .RecordCount > 0 Then
            CheckSA_BAL = .Fields(0).Value
          Else
            CheckSA_BAL = 0
        End If
        .Close
     Set RstCheck = Nothing
End With
End Function

Function CheckSA_BAL_INTL(Param, WhichBal) As Double
Dim RstCheck As New ADODB.Recordset
SQL = "SELECT * FROM qryCheck_SA_BALANCE_INTL WHERE [AccountID]=" & CLng(Param)
With RstCheck
        .Open SQL, cn, adOpenKeyset
        If .RecordCount > 0 Then
                        If WhichBal = "DOLLAR" Then
                            CheckSA_BAL_INTL = .Fields(0).Value
                        Else
                            CheckSA_BAL_INTL = .Fields(1).Value
                        End If
        Else
        CheckSA_BAL_INTL = 0
        End If
        .Close
     Set RstCheck = Nothing
End With
End Function


Private Sub cmdSetComm_Click()
On Error GoTo FailSafe_Error
Dim myComm
myComm = InputBox("Enter commission", "Please enter commission in numeric format")

If CheckNull(myComm) Then: Exit Sub
myComm = CDbl(myComm)

SQL = "UPDATE tbl_custAccDetails SET Commission = " & myComm & _
    " WHERE (((AccountIDDetails)=" & Me.DataGrid2.Columns(0).Text & "))"
cn.BeginTrans
    cn.Execute SQL
cn.CommitTrans
Call CustAccDetails

Exit Sub
FailSafe_Error:
MsgBox Err.Description

End Sub

Private Sub DataGrid1_Click()
LoadValues
End Sub

Private Sub Form_Load()
Set Rs = New ADODB.Recordset

SQL = "SELECT * FROM tbl_CustAccounts ORDER BY [Account Name] ASC"
Rs.Open SQL, cn, adOpenKeyset, adLockOptimistic


Set Me.DataGrid1.DataSource = Rs
Set RsClone = Rs.Clone
Me.lblStatus = "RECORD :" & Rs.AbsolutePosition & "/" & Rs.RecordCount
End Sub


Private Sub txtSearch_Change()

With RsClone
If Len(Me.txtSearch) <= 0 Then
    Call ClearTxt
    Rs.MoveFirst
    RsClone.MoveFirst
    Me.lblStatus = "RECORD :" & Rs.AbsolutePosition & "/" & Rs.RecordCount
    Exit Sub
End If
            .MoveFirst
            .Find "[Account Name] like '" & Me.txtSearch & "%'"
            If Not .EOF Then
                    Rs.MoveFirst
                    Rs.Find "[AccountID]=" & RsClone.Fields(0).Value
                    Call LoadValues
            End If
            
End With
End Sub

Sub Navigate(ByVal userAction As String)

 Select Case userAction
            Case "<<"
                Rs.MoveFirst
            Case ">>"
                Rs.MoveLast
            Case ">"
                Rs.MoveNext
                If Rs.EOF Then
                    MsgBox "Already at end of File!"
                    Rs.MoveLast
                End If
            Case "<"
                Rs.MovePrevious
                If Rs.BOF Then
                    MsgBox "Already at beginning of File!"
                    Rs.MoveFirst
                End If
        End Select
Call LoadValues
End Sub

Sub CustAccDetails()
Dim RsAccDetails As ADODB.Recordset
Set RsAccDetails = New ADODB.Recordset
SQL = "SELECT * FROM qryAccDetails WHERE [AccountID]=" & CLng(Me.txtcustNo)
RsAccDetails.Open SQL, cn, adOpenKeyset, adLockOptimistic
Set Me.DataGrid2.DataSource = RsAccDetails
End Sub
